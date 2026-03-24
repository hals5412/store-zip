"""
repack.py v2.0 - 圧縮ファイルを無圧縮ZIPに変換するツール

主な機能:
  - ZIP / RAR / 7z / tar.gz / CBZ / CBR など主要形式 → 無圧縮ZIP
  - ゴミファイル自動削除（設定ファイル + 対話的判断で管理）
  - 未分類ファイルはユーザーに確認し decisions.json に永続保存
  - ルートに単一フォルダのみの場合は剥がしてフラット化
  - 元ファイルのタイムスタンプを引き継ぐ
  - 元ファイルはゴミ箱へ
  - 文字コード問題に対応（Shift-JIS / UTF-8 / その他）
"""

import sys
import os
import io
import zipfile
import tempfile
import shutil
import subprocess
import fnmatch
import time
import traceback
import json
from pathlib import Path
from datetime import datetime

# ── Windows コンソール UTF-8 化 ──────────────────────────────────────
# PowerShell / cmd でも日本語が正しく表示されるようにする
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
        sys.stdin  = io.TextIOWrapper(sys.stdin.buffer,  encoding="utf-8", errors="replace")
    except Exception:
        pass

# ── TOML パーサー ────────────────────────────────────────────────────
try:
    import tomllib          # Python 3.11+ 標準
except ImportError:
    try:
        import tomli as tomllib   # pip install tomli  (Python 3.10 以下)
    except ImportError:
        tomllib = None


# ============================================================
# ログ出力ユーティリティ
# ============================================================
def log(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def log_error(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] ERROR: {msg}", flush=True)

def log_ok(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] OK: {msg}", flush=True)

def log_skip(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] SKIP: {msg}", flush=True)


# ============================================================
# バイト列デコード（複数エンコーディングを順に試みる）
# ============================================================
def decode_bytes(data: bytes) -> str:
    """
    7-Zip の stdout/stderr はシステムのコードページで出力される。
    日本語 Windows では cp932 が多いため UTF-8 → cp932 → latin-1 の順に試みる。
    latin-1 は必ず成功する（バイト値そのままマッピング）。
    """
    for enc in ("utf-8", "cp932", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("latin-1")


# ============================================================
# デフォルト設定
# ============================================================
DEFAULT_CONFIG: dict = {
    "junk_patterns": [
        "*.url", "*.lnk", "*.webloc",
        "Thumbs.db", "desktop.ini", "ehthumbs.db",
        ".DS_Store", ".AppleDouble", ".LSOverride", "._*",
        "folder.jpg", "folder.jpeg", "folder.png",
        "Folder.jpg", "Folder.jpeg", "Folder.png",
        "AlbumArt*.jpg", "AlbumArt*.png",
        "*.nfo", "*.sfv", "*.torrent",
        "ComicInfo.xml", "series.json",
    ],
    "junk_dirs": [
        "__MACOSX", ".git", ".svn",
    ],
    "allow_patterns": [],
    # "ask"  : 未分類ファイルをユーザーに確認（デフォルト）
    # "keep" : 未分類ファイルを自動保持
    # "junk" : 未分類ファイルを自動削除
    "unknown_file_action": "ask",
    "write_log": False,
}


# ============================================================
# 設定ファイル読み込み（config.toml + decisions.json をマージ）
# ============================================================
def load_config(exe_dir: Path) -> dict:
    config: dict = {
        "junk_patterns":       list(DEFAULT_CONFIG["junk_patterns"]),
        "junk_dirs":           list(DEFAULT_CONFIG["junk_dirs"]),
        "allow_patterns":      list(DEFAULT_CONFIG["allow_patterns"]),
        "unknown_file_action": DEFAULT_CONFIG["unknown_file_action"],
        "write_log":           DEFAULT_CONFIG["write_log"],
    }

    # ── config.toml ────────────────────────────────────────
    config_path = exe_dir / "config.toml"
    if config_path.exists() and tomllib is not None:
        try:
            with open(config_path, "rb") as f:
                cfg = tomllib.load(f)
            for k in ("junk_patterns", "junk_dirs", "allow_patterns"):
                if k in cfg:
                    config[k] = list(cfg[k])
            for k in ("unknown_file_action", "write_log"):
                if k in cfg:
                    config[k] = cfg[k]
            log(f"設定読み込み完了: {config_path}")
        except Exception as e:
            log(f"警告: config.toml の読み込みに失敗 ({e})。デフォルト設定を使用します。")
    elif config_path.exists() and tomllib is None:
        log("警告: TOML パーサーが利用できません。デフォルト設定を使用します。")
        log("  → pip install tomli  を実行するとTOMLを読み込めます。")

    # ── decisions.json（対話的判断の永続ストア）──────────
    decisions_path = exe_dir / "decisions.json"
    if decisions_path.exists():
        try:
            with open(decisions_path, "r", encoding="utf-8") as f:
                decisions = json.load(f)
            n_allow = 0
            n_junk  = 0
            for p in decisions.get("allow_patterns", []):
                if p not in config["allow_patterns"]:
                    config["allow_patterns"].append(p)
                    n_allow += 1
            for p in decisions.get("junk_patterns", []):
                if p not in config["junk_patterns"]:
                    config["junk_patterns"].append(p)
                    n_junk += 1
            if n_allow or n_junk:
                log(f"判断ファイル読み込み: allow +{n_allow}件, junk +{n_junk}件")
        except Exception as e:
            log(f"警告: decisions.json の読み込みに失敗 ({e})")

    return config


# ============================================================
# 判断の永続保存（decisions.json）
# ============================================================
def save_decision(exe_dir: Path, filename: str, decision: str) -> None:
    """
    ユーザーが対話的に決定したファイル名を decisions.json に追記する。
    decision : "allow" または "junk"
    """
    decisions_path = exe_dir / "decisions.json"
    existing: dict = {"allow_patterns": [], "junk_patterns": []}

    if decisions_path.exists():
        try:
            with open(decisions_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            pass

    key = "allow_patterns" if decision == "allow" else "junk_patterns"
    if filename not in existing.get(key, []):
        existing.setdefault(key, []).append(filename)
        try:
            with open(decisions_path, "w", encoding="utf-8") as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log(f"警告: decisions.json の保存に失敗 ({e})")


# ============================================================
# セッションキャッシュ（同一バッチ内での重複確認を防ぐ）
# キーは拡張子パターン（例: "*.txt"）。拡張子なしはファイル名そのまま。
# ============================================================
_session_cache: dict[str, str] = {}  # "*.ext" or filename → "allow" | "junk"


def _ext_pattern(fname: str) -> str:
    """
    ファイル名から拡張子パターンを生成する。
      "image.jpg"  →  "*.jpg"
      "Makefile"   →  "Makefile"  （拡張子なし → ファイル名そのまま）
    """
    suffix = Path(fname).suffix.lower()
    return f"*{suffix}" if suffix else fname


def get_file_decision(fname: str, config: dict, exe_dir: Path) -> str:
    """
    ファイル名に対する処理方針を返す。
      - "allow" : 保持する
      - "junk"  : 削除する

    判定順序:
      1. allow_patterns（junk_patterns より優先）
      2. junk_patterns
      3. 未分類 → 拡張子単位でセッションキャッシュを確認
      4.          キャッシュになければ unknown_file_action に従い自動判断 or ユーザー確認
    """
    # 1. allow_patterns（保持を強制）
    if _matches_any(fname, config["allow_patterns"]):
        return "allow"

    # 2. junk_patterns
    if _matches_any(fname, config["junk_patterns"]):
        return "junk"

    # 3. 未分類 → 拡張子パターンで判断
    pattern = _ext_pattern(fname)

    if pattern in _session_cache:
        return _session_cache[pattern]

    # 4. ユーザー確認 or 自動判断
    action = config.get("unknown_file_action", "ask")
    if action == "keep":
        decision = "allow"
        log(f"  自動保持 (unknown_file_action=keep): {pattern}")
    elif action == "junk":
        decision = "junk"
        log(f"  自動削除 (unknown_file_action=junk): {pattern}")
    else:
        decision = _ask_user_for_extension(pattern, fname)

    _session_cache[pattern] = decision
    save_decision(exe_dir, pattern, decision)
    return decision


def _ask_user_for_extension(pattern: str, example_fname: str) -> str:
    """
    未分類の拡張子についてユーザーに確認する。"allow" または "junk" を返す。
    細かいファイル名単位の設定は config.toml で行う。
    """
    print()
    print(f"  ┌─ 未分類の拡張子: {pattern}  （例: {example_fname}）")
    print( "  │  この拡張子は既存のルールに一致しませんでした。")
    print( "  │  同じ拡張子のファイルすべてに適用されます。")
    print( "  │  より細かい設定は config.toml を直接編集してください。")
    print( "  ├─ [K] 保持する（allow_patterns に追加）")
    print( "  └─ [D] 削除する（junk_patterns に追加）")

    while True:
        try:
            sys.stdout.flush()
            choice = input("  選択 [K/D]: ").strip().upper()
        except (EOFError, KeyboardInterrupt):
            log(f"  入力なし。今回は保持します: {pattern}")
            return "allow"

        if choice in ("K", "KEEP"):
            return "allow"
        if choice in ("D", "DELETE", "DEL"):
            return "junk"
        print("  K（保持）か D（削除）を入力してください。")


def _matches_any(filename: str, patterns: list) -> bool:
    return any(fnmatch.fnmatch(filename, p) for p in patterns)


# ============================================================
# 7-Zip の検出
# ============================================================
_SEVENZIP_CANDIDATES = [
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
]

def find_7zip(exe_dir: Path) -> str | None:
    # 同梱 exe と同じフォルダ
    local = exe_dir / "7z.exe"
    if local.exists():
        return str(local)
    # 標準インストール先
    for candidate in _SEVENZIP_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    # PATH
    return shutil.which("7z")


# ============================================================
# 対応拡張子
# ============================================================
SUPPORTED_EXTENSIONS: frozenset[str] = frozenset({
    ".zip", ".cbz",
    ".rar", ".cbr",
    ".7z",  ".cb7",
    ".tar", ".cbt",
    ".gz",  ".tgz",
    ".bz2", ".tbz2", ".tbz",
    ".xz",  ".txz",
    ".lzma",
    ".zst",
    ".lzh", ".lha",
    ".iso",
    ".arj",
    ".cab",
    ".rpm",
    ".deb",
    ".wim",
})


# ============================================================
# 7-Zip による展開（文字コード問題対策済み）
# ============================================================
# -mcp=932 が有効な形式（ZIP のファイル名エンコーディング指定）
_ZIP_LIKE_EXTENSIONS = frozenset({".zip", ".cbz"})


def _to_7zip_path(path: Path) -> str:
    """
    7-Zip に渡すパス文字列を返す。
    UNCパス (\\server\share\...) は \\?\UNC\server\share\... 形式に変換する。
    これにより:
      - MAX_PATH (260文字) 制限を回避
      - 日本語・特殊文字を含むUNCパスの Windows API エラーを回避
      - ファイルコピー不要で直接処理できる
    """
    s = str(path)
    if s.startswith("\\\\") and not s.startswith("\\\\?\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return s


def extract_with_7zip(sevenzip: str, archive: Path, dest_dir: Path) -> bool:
    """
    7-Zip でアーカイブを展開する。

    文字コード対策:
      -mcp=932 : ZIP のファイル名を Shift-JIS として解釈する（ZIP系のみ）。
                 RAR/7z 等には適用しない（誤動作の原因になり得るため）。
      UNCパス  : \\?\UNC\ 拡張形式に変換してパス制限・文字コード問題を回避。
    """
    cmd = [
        sevenzip, "x", _to_7zip_path(archive),
        f"-o{_to_7zip_path(dest_dir)}",
        "-y",    # すべての確認プロンプトに Yes
        "-aoa",  # 既存ファイルを上書き
    ]
    # -mcp=932 は ZIP 系のみ（EFSビットなし旧日本語ZIPのファイル名文字化け対策）
    if archive.suffix.lower() in _ZIP_LIKE_EXTENSIONS:
        cmd.append("-mcp=932")

    log(f"  展開中: {archive.name}")

    try:
        result = subprocess.run(cmd, capture_output=True)
    except FileNotFoundError:
        log_error(f"7-Zip 実行ファイルが見つかりません: {sevenzip}")
        return False
    except Exception as e:
        log_error(f"7-Zip の起動に失敗しました: {e}")
        return False

    stdout_str = decode_bytes(result.stdout) if result.stdout else ""
    stderr_str = decode_bytes(result.stderr) if result.stderr else ""

    if result.returncode != 0:
        log_error(f"7-Zip 展開失敗 (終了コード={result.returncode})")
        if stderr_str.strip():
            log_error(f"  stderr: {stderr_str.strip()[:600]}")
        if stdout_str.strip():
            # 末尾だけ表示（大量出力の中でエラー行は末尾に出ることが多い）
            log(f"  stdout (末尾300文字): {stdout_str.strip()[-300:]}")
        return False

    return True


# ============================================================
# ゴミ除去
# ============================================================
def remove_junk_from_dir(extract_dir: Path, config: dict, exe_dir: Path) -> int:
    """
    展開ディレクトリからゴミを除去する。
    未分類ファイルは get_file_decision() 経由でユーザーに確認または自動判断する。
    削除した件数を返す。
    """
    removed = 0
    junk_dirs = config["junk_dirs"]

    # ── 1. ゴミディレクトリを削除 ────────────────────────────
    for dirpath, dirnames, _ in os.walk(extract_dir, topdown=True):
        for d in list(dirnames):
            if _matches_any(d, junk_dirs):
                target = Path(dirpath) / d
                try:
                    shutil.rmtree(target, ignore_errors=True)
                    log(f"  ゴミディレクトリ削除: {d}/")
                    removed += 1
                except Exception as e:
                    log(f"  警告: ディレクトリ削除失敗 {d}: {e}")
                if d in dirnames:
                    dirnames.remove(d)

    # ── 2. ファイルを処理 ─────────────────────────────────────
    files_to_delete: list[Path] = []

    for dirpath, _, filenames in os.walk(extract_dir):
        for fname in filenames:
            fpath = Path(dirpath) / fname
            decision = get_file_decision(fname, config, exe_dir)
            if decision == "junk":
                files_to_delete.append(fpath)

    for fpath in files_to_delete:
        try:
            fpath.unlink()
            log(f"  削除: {fpath.name}")
            removed += 1
        except Exception as e:
            log(f"  警告: ファイル削除失敗 {fpath.name}: {e}")

    # ── 3. 空ディレクトリを掃除 ──────────────────────────────
    for dirpath, _, _ in os.walk(extract_dir, topdown=False):
        if dirpath == str(extract_dir):
            continue
        p = Path(dirpath)
        try:
            if not any(p.iterdir()):
                p.rmdir()
        except Exception:
            pass

    return removed


# ============================================================
# ルートフォルダ剥がし
# ============================================================
def strip_root_folder(extract_dir: Path) -> Path:
    """
    extract_dir 直下にフォルダが1つだけあり、ファイルが0件の場合、
    そのフォルダを新しいルートとする（再帰的に適用）。
    """
    while True:
        entries = list(extract_dir.iterdir())
        dirs  = [e for e in entries if e.is_dir()]
        files = [e for e in entries if e.is_file()]
        if len(dirs) == 1 and len(files) == 0:
            log(f"  ルートフォルダ除去: {dirs[0].name}/")
            extract_dir = dirs[0]
        else:
            break
    return extract_dir


# ============================================================
# 無圧縮 ZIP 作成
# ============================================================
def make_store_zip(source_dir: Path, output_zip: Path) -> None:
    """source_dir 以下のファイルを STORE（無圧縮）ZIP に格納する。"""
    with zipfile.ZipFile(
        output_zip, "w",
        compression=zipfile.ZIP_STORED,
        allowZip64=True,
    ) as zf:
        for fpath in sorted(source_dir.rglob("*")):
            if fpath.is_file():
                arcname = fpath.relative_to(source_dir)
                zf.write(fpath, arcname)

    size = output_zip.stat().st_size
    log(f"  無圧縮ZIP生成完了: {output_zip.name}  ({size:,} bytes)")


# ============================================================
# タイムスタンプ復元
# ============================================================
def apply_timestamp(target: Path, mtime: float) -> None:
    os.utime(target, (mtime, mtime))
    dt_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
    log(f"  タイムスタンプ復元: {dt_str}")


# ============================================================
# ゴミ箱送り（Windows）
# ============================================================
def send_to_recycle_bin(path: Path) -> bool:
    # send2trash ライブラリ（推奨）
    try:
        import send2trash
        send2trash.send2trash(str(path))
        return True
    except ImportError:
        pass

    # PowerShell フォールバック
    # パス内のシングルクォートをエスケープ
    safe_path = str(path).replace("'", "''")
    ps_script = (
        "Add-Type -AssemblyName Microsoft.VisualBasic; "
        "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile("
        f"'{safe_path}', 'OnlyErrorDialogs', 'SendToRecycleBin')"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True, timeout=30,
        )
        if result.returncode == 0:
            return True
    except Exception as e:
        log(f"  警告: PowerShell ゴミ箱送り失敗 ({e})。直接削除します。")

    # 最終手段: 直接削除
    try:
        path.unlink()
        return True
    except Exception as e:
        log_error(f"元ファイルの削除に失敗しました: {e}")
        return False


# ============================================================
# 1ファイルの処理
# ============================================================
def process_file(
    archive_path: Path,
    sevenzip: str,
    config: dict,
    exe_dir: Path,
) -> bool:
    print("-" * 60)
    log(f"処理開始: {archive_path.name}")

    ext = archive_path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        log_skip(f"非対応形式: {ext}")
        return True

    if not archive_path.exists():
        log_error(f"ファイルが見つかりません: {archive_path}")
        return False

    # 元ファイルのタイムスタンプを保存
    original_mtime = archive_path.stat().st_mtime

    with tempfile.TemporaryDirectory(prefix="repack_") as tmp_root:
        tmp_path = Path(tmp_root)
        extract_dir = tmp_path / "extracted"
        extract_dir.mkdir()

        # ── 展開 ──────────────────────────────────────────────
        if not extract_with_7zip(sevenzip, archive_path, extract_dir):
            return False

        # ── ゴミ除去 ──────────────────────────────────────────
        removed = remove_junk_from_dir(extract_dir, config, exe_dir)
        if removed:
            log(f"  合計 {removed} 件を削除しました")

        # ── 空チェック ────────────────────────────────────────
        remaining_files = [f for f in extract_dir.rglob("*") if f.is_file()]
        if not remaining_files:
            log("  ゴミ除去後にファイルが残りませんでした。元ファイルは変更しません。")
            return True

        # ── ルートフォルダ剥がし ──────────────────────────────
        actual_root = strip_root_folder(extract_dir)

        # ── 無圧縮ZIP生成（一時ファイル）─────────────────────
        tmp_zip = tmp_path / (archive_path.stem + ".zip")
        make_store_zip(actual_root, tmp_zip)

        # ── 元ファイルをゴミ箱へ ─────────────────────────────
        log(f"  ゴミ箱へ送信: {archive_path.name}")
        send_to_recycle_bin(archive_path)

        # ── 出力先パスを決定 ─────────────────────────────────
        output_path = archive_path.with_suffix(".zip")

        if output_path.exists():
            backup = output_path.with_suffix(".zip.bak")
            output_path.rename(backup)
            log(f"  既存ファイルを退避: {backup.name}")

        # ── 一時ZIP を最終パスへ移動 ─────────────────────────
        shutil.move(str(tmp_zip), str(output_path))

        # ── タイムスタンプ復元 ───────────────────────────────
        apply_timestamp(output_path, original_mtime)

    log_ok(f"完了: {output_path.name}")
    return True


# ============================================================
# エントリポイント
# ============================================================
def main() -> None:
    # exe のディレクトリを取得（PyInstaller でビルドした場合は sys.executable の親）
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent

    print("=" * 60)
    print("  repack v2.0 - 無圧縮ZIP変換ツール")
    print("=" * 60)
    print()

    args = sys.argv[1:]
    if not args:
        print("使い方:")
        print("  repack.exe <圧縮ファイル> [<圧縮ファイル2> ...]")
        print()
        print("  install_sendto.bat を実行すると SendTo に登録できます。")
        print("  エクスプローラーで圧縮ファイルを右クリック → 送る → repack")
        print()
        input("Enter キーで終了...")
        sys.exit(0)

    # 設定読み込み
    config = load_config(exe_dir)

    # 7-Zip 検出
    sevenzip = find_7zip(exe_dir)
    if not sevenzip:
        log_error("7-Zip (7z.exe) が見つかりません。")
        log_error("  https://www.7-zip.org/ からインストールするか、")
        log_error("  7z.exe を repack.exe と同じフォルダに置いてください。")
        input("Enter キーで終了...")
        sys.exit(1)
    log(f"7-Zip: {sevenzip}")
    print()

    # ファイルを順次処理
    errors: list[str] = []
    for arg in args:
        p = Path(arg)
        try:
            ok = process_file(p, sevenzip, config, exe_dir)
            if not ok:
                errors.append(p.name)
        except Exception:
            log_error(f"{p.name}: 予期しないエラー\n{traceback.format_exc()}")
            errors.append(p.name)

    # 結果表示
    print()
    print("=" * 60)
    if errors:
        print(f"  [!] エラーが発生したファイル: {len(errors)} 件")
        for name in errors:
            print(f"      - {name}")
        print()
        input("Enter キーで終了...")
        sys.exit(1)
    else:
        processed = sum(
            1 for a in args
            if Path(a).suffix.lower() in SUPPORTED_EXTENSIONS
        )
        print(f"  すべての処理が完了しました。（{processed} ファイル）")
        print("  3秒後に自動で閉じます...")
        time.sleep(3)
        sys.exit(0)


if __name__ == "__main__":
    main()
