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
import threading
import ctypes
from concurrent.futures import ThreadPoolExecutor, as_completed
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


# ── Tee ライター（コンソールとログファイルに同時出力）────────────────
class _TeeWriter:
    """stdout を置き換えてコンソールとログファイルに同時に書き込む。"""
    def __init__(self, console, logfile):
        self._console = console
        self._logfile = logfile

    def write(self, data: str) -> int:
        self._console.write(data)
        self._logfile.write(data)
        return len(data)

    def flush(self) -> None:
        self._console.flush()
        self._logfile.flush()

    # TextIOWrapper が要求する属性を委譲
    @property
    def encoding(self):  return self._console.encoding
    @property
    def errors(self):    return self._console.errors
    def fileno(self):    return self._console.fileno()
    def isatty(self):    return False


def setup_log_file(exe_dir: Path) -> None:
    """write_log=true のとき呼び出す。repack.log への Tee を設定する。"""
    log_path = exe_dir / "repack.log"
    try:
        lf = open(log_path, "a", encoding="utf-8")
        lf.write(f"\n{'='*60}\n")
        lf.write(f"  session: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        lf.write(f"{'='*60}\n")
        sys.stdout = _TeeWriter(sys.stdout, lf)
    except Exception as e:
        print(f"警告: ログファイルを開けません ({e})", flush=True)

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
# スレッドローカルバッファが設定されているときはバッファに蓄積し、
# ファイル処理完了後に _print_lock を取得して一括出力する。
# これにより並列処理時もファイル単位でログがまとまって表示される。
# ============================================================
_thread_local = threading.local()


def _emit(line: str) -> None:
    """1行出力。バッファ中ならバッファへ、そうでなければ直接 stdout へ。"""
    buf: list | None = getattr(_thread_local, "buffer", None)
    if buf is not None:
        buf.append(line)
    else:
        sys.stdout.write(line)
        sys.stdout.flush()


def log(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")

def log_error(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {msg}\n")

def log_ok(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] OK: {msg}\n")

def log_skip(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] SKIP: {msg}\n")


def _flush_buffer() -> None:
    """スレッドローカルバッファの内容を _output_lock を取得して一括出力する。"""
    buf: list | None = getattr(_thread_local, "buffer", None)
    if not buf:
        return
    with _output_lock:
        for line in buf:
            sys.stdout.write(line)
        sys.stdout.flush()
    buf.clear()


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
    # "counter" : archive (1).zip, archive (2).zip ...
    # "date"    : archive_20241225.zip, archive_20241225 (1).zip ...
    "duplicate_name_style": "counter",
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
        "unknown_file_action":  DEFAULT_CONFIG["unknown_file_action"],
        "duplicate_name_style": DEFAULT_CONFIG["duplicate_name_style"],
        "write_log":            DEFAULT_CONFIG["write_log"],
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
            for k in ("unknown_file_action", "duplicate_name_style", "write_log"):
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
    並列処理時の競合を防ぐため _decisions_lock で保護する。
    """
    with _decisions_lock:
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
# 並列処理時のスレッド安全のためロックで保護する。
# ============================================================
_session_cache  = {}             # "*.ext" or filename → "allow" | "junk"
_cache_lock     = threading.Lock()  # _session_cache の読み書き保護
_decisions_lock = threading.Lock()  # decisions.json の読み書き保護
# stdout への書き込みをすべて排他制御する1本のロック。
# バッファフラッシュとプロンプトで共用することで、
# プロンプト表示中に他スレッドの出力が割り込まなくなる。
_output_lock    = threading.Lock()


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

    with _cache_lock:
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
        # _ask_user_for_extension 内でキャッシュに入っていた場合は None が返る
        if decision is None:
            with _cache_lock:
                return _session_cache[pattern]

    with _cache_lock:
        _session_cache[pattern] = decision
    save_decision(exe_dir, pattern, decision)
    return decision


def _ask_user_for_extension(pattern: str, example_fname: str) -> str | None:
    """
    未分類の拡張子についてユーザーに確認する。
    _output_lock をプロンプト全体（input() 中も含む）で保持するため、
    ユーザーが回答するまで他スレッドの出力はブロックされる。

    戻り値: "allow" / "junk"、またはロック待機中に別スレッドが回答済みなら None
    """
    with _output_lock:
        # _output_lock 待機中に別スレッドが同じ拡張子を回答した可能性を確認
        with _cache_lock:
            if pattern in _session_cache:
                return None  # 呼び出し元でキャッシュ値を使う

        # スレッドローカルバッファをここでフラッシュ（ロック内なので安全）
        buf: list | None = getattr(_thread_local, "buffer", None)
        if buf:
            for line in buf:
                sys.stdout.write(line)
            sys.stdout.flush()
            buf.clear()

        sys.stdout.write("\n")
        sys.stdout.write(f"  ┌─ 未分類の拡張子: {pattern}  （例: {example_fname}）\n")
        sys.stdout.write( "  │  この拡張子は既存のルールに一致しませんでした。\n")
        sys.stdout.write( "  │  同じ拡張子のファイルすべてに適用されます。\n")
        sys.stdout.write( "  │  より細かい設定は config.toml を直接編集してください。\n")
        sys.stdout.write( "  ├─ [K] 保持する（allow_patterns に追加）\n")
        sys.stdout.write( "  └─ [D] 削除する（junk_patterns に追加）\n")
        sys.stdout.flush()

        while True:
            try:
                choice = input("  選択 [K/D]: ").strip().upper()
            except (EOFError, KeyboardInterrupt):
                sys.stdout.write(f"  入力なし。今回は保持します: {pattern}\n")
                sys.stdout.flush()
                return "allow"

            if choice in ("K", "KEEP"):
                return "allow"
            if choice in ("D", "DELETE", "DEL"):
                return "junk"
            sys.stdout.write("  K（保持）か D（削除）を入力してください。\n")
            sys.stdout.flush()


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

# マジックバイト → 実際のフォーマット拡張子
_FORMAT_SIGNATURES: list[tuple[bytes, str]] = [
    (b"\x52\x61\x72\x21\x1a\x07", ".rar"),  # RAR4 / RAR5
    (b"\x37\x7a\xbc\xaf\x27\x1c", ".7z"),   # 7-Zip
    (b"\x1f\x8b",                  ".gz"),   # gzip / tar.gz
    (b"\x42\x5a\x68",             ".bz2"),  # bzip2
    (b"\xfd\x37\x7a\x58\x5a\x00", ".xz"),   # xz
    (b"\x50\x4b",                  ".zip"),  # ZIP (PK signature)
]


def _detect_real_ext(archive: Path) -> str:
    """マジックバイトからファイルの実際の拡張子を返す。読み取り失敗時は元の拡張子。"""
    try:
        with open(str(archive), "rb") as f:
            header = f.read(8)
        for sig, ext in _FORMAT_SIGNATURES:
            if header.startswith(sig):
                return ext
    except Exception:
        pass
    return archive.suffix.lower()


def _to_7zip_path(path: Path) -> str:
    """
    7-Zip に渡すパス文字列を返す。
    NAS 等の UNC パスは拡張 UNC 形式に変換し、
    MAX_PATH 制限や日本語パスの Windows API エラーを回避する。
    """
    s = str(path)
    if s.startswith("\\\\") and not s.startswith("\\\\?\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return s


def extract_with_7zip(sevenzip: str, archive: Path, dest_dir: Path) -> bool:
    """
    7-Zip でアーカイブを展開する。

    この関数を呼ぶ前にパスの [ ] チェックを行い、必要なら
    _extract_via_temp_copy 経由でブラケットのないローカルパスに変換しておくこと。
    -mcp=932 は ZIP 系のみ適用（RAR/7z 等での誤動作防止）。
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
            log(f"  stdout (末尾300文字): {stdout_str.strip()[-300:]}")
        return False

    return True


# ============================================================
# ZIP/CBZ を Python の zipfile で直接展開
# ============================================================
def extract_zip_python(archive: Path, dest_dir: Path) -> bool:
    """
    Python の zipfile モジュールで ZIP/CBZ を展開する。

    7-Zip を使わないため、UNC パスや角括弧を含むパスでも正常に動作する。
    EFS フラグなしの旧日本語 ZIP は CP932 でファイル名を再デコードする。
    """
    log(f"  展開中: {archive.name}")
    try:
        with zipfile.ZipFile(str(archive), "r") as zf:
            for info in zf.infolist():
                fname = info.filename

                # EFS フラグ (bit 11) がない場合、ファイル名は CP437 バイト列として
                # 格納されている。日本語 ZIP は CP932 が多いので再デコードを試みる。
                if not (info.flag_bits & 0x800):
                    try:
                        fname = fname.encode("cp437").decode("cp932")
                    except (UnicodeEncodeError, UnicodeDecodeError):
                        pass  # デコード失敗時はそのまま使用

                # パストラバーサル対策
                target = (dest_dir / fname).resolve()
                try:
                    target.relative_to(dest_dir.resolve())
                except ValueError:
                    log(f"  スキップ（不正パス）: {fname}")
                    continue

                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                else:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with zf.open(info) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
        return True
    except zipfile.BadZipFile:
        raise  # 呼び出し元が 7-Zip へのフォールバックを処理する
    except Exception as e:
        log_error(f"ZIP 展開失敗: {e}")
        return False


def _has_wildcard_chars(path: Path) -> bool:
    """7-Zip がワイルドカードとして解釈する文字 [ ] がパスに含まれるか確認する。"""
    s = str(path)
    return "[" in s or "]" in s


def _extract_via_temp_copy(sevenzip: str, archive: Path, dest_dir: Path) -> bool:
    """
    アーカイブをローカル一時フォルダにコピーしてから 7-Zip で展開する。
    UNC パスや角括弧を含むパスで 7-Zip が直接アクセスできない場合のフォールバック。
    """
    # 拡張子が実態と異なる場合（例: RAR を .cbz にリネームしたファイル）に備え、
    # マジックバイトで実際のフォーマットを判定して正しい拡張子のファイル名にする。
    # これにより extract_with_7zip が誤って -mcp=932 を付けることを防ぐ。
    real_ext = _detect_real_ext(archive)
    if real_ext != archive.suffix.lower():
        log(f"  実際のフォーマット: {archive.suffix} → {real_ext}")
    log(f"  ローカル一時コピーを経由して展開します...")
    with tempfile.TemporaryDirectory(prefix="repack_tmp_") as tmp_dir:
        safe_name = "archive" + real_ext
        tmp_archive = Path(tmp_dir) / safe_name
        try:
            shutil.copy2(str(archive), str(tmp_archive))
        except Exception as e:
            log_error(f"  一時コピー作成失敗: {e}")
            return False
        return extract_with_7zip(sevenzip, tmp_archive, dest_dir)


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
    if config.get("remove_empty_dirs", True):
        for dirpath, _, _ in os.walk(extract_dir, topdown=False):
            if dirpath == str(extract_dir):
                continue
            p = Path(dirpath)
            try:
                if not any(p.iterdir()):
                    p.rmdir()
                    log(f"  空フォルダ削除: {p.relative_to(extract_dir)}/")
            except Exception:
                pass

    return removed


# ============================================================
# 無圧縮ZIP判定
# ============================================================
def is_already_store_zip(archive_path: Path) -> bool:
    """
    ZIPファイルの全エントリが STORE（無圧縮）なら True を返す。
    ZIP/CBZ 以外は常に False（処理対象）。
    """
    if archive_path.suffix.lower() not in {".zip", ".cbz"}:
        return False
    try:
        with zipfile.ZipFile(archive_path, "r") as zf:
            infos = zf.infolist()
            if not infos:
                return False
            return all(info.compress_type == zipfile.ZIP_STORED for info in infos)
    except Exception:
        return False


# ============================================================
# 重複しない出力パスの生成
# ============================================================
def _unique_path(path: Path, style: str = "counter") -> Path:
    """
    path が存在しなければそのまま返す。存在する場合は style に従い回避する。
    既存ファイルは一切変更しない。

    style="counter" : archive (1).zip, archive (2).zip ...
    style="date"    : archive_20241225.zip,
                      archive_20241225 (1).zip, archive_20241225 (2).zip ...
    """
    if not path.exists():
        return path
    parent = path.parent
    stem   = path.stem
    suffix = path.suffix

    if style == "date":
        date_str = datetime.now().strftime("%Y%m%d")
        dated = parent / f"{stem}_{date_str}{suffix}"
        if not dated.exists():
            return dated
        # 日付付きでも衝突するなら日付ベースに連番を付ける
        stem = f"{stem}_{date_str}"

    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


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
) -> str:
    """
    戻り値:
      "converted" : 変換完了
      "skipped"   : スキップ（無圧縮ZIP済み・非対応形式・空アーカイブ）
      "error"     : 失敗
    """
    _emit("-" * 60 + "\n")
    log(f"処理開始: {archive_path.name}")

    ext = archive_path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        log_skip(f"非対応形式: {ext}")
        return "skipped"

    if not archive_path.exists():
        log_error(f"ファイルが見つかりません: {archive_path}")
        return "error"

    if is_already_store_zip(archive_path):
        log_skip(f"既に無圧縮ZIPです。スキップします。")
        return "skipped"

    # 元ファイルのタイムスタンプを保存
    original_mtime = archive_path.stat().st_mtime

    with tempfile.TemporaryDirectory(prefix="repack_") as tmp_root:
        tmp_path = Path(tmp_root)
        extract_dir = tmp_path / "extracted"
        extract_dir.mkdir()

        # ── 展開 ──────────────────────────────────────────────
        # ZIP/CBZ → Python zipfile で直接展開（ワイルドカード問題なし）。
        #           BadZipFile（実体が RAR 等）の場合は 7-Zip にフォールバック。
        # RAR/7z 等 → パスに [ ] があればローカル一時コピー経由、なければ直接。
        ok = False
        if archive_path.suffix.lower() in _ZIP_LIKE_EXTENSIONS:
            try:
                ok = extract_zip_python(archive_path, extract_dir)
            except zipfile.BadZipFile:
                log(f"  ZIP 形式ではありません。7-Zip で再試行します...")
                if _has_wildcard_chars(archive_path):
                    ok = _extract_via_temp_copy(sevenzip, archive_path, extract_dir)
                else:
                    ok = extract_with_7zip(sevenzip, archive_path, extract_dir)
        else:
            if _has_wildcard_chars(archive_path):
                ok = _extract_via_temp_copy(sevenzip, archive_path, extract_dir)
            else:
                ok = extract_with_7zip(sevenzip, archive_path, extract_dir)
        if not ok:
            return "error"

        # ── ゴミ除去 ──────────────────────────────────────────
        removed = remove_junk_from_dir(extract_dir, config, exe_dir)
        if removed:
            log(f"  合計 {removed} 件を削除しました")

        # ── 空チェック ────────────────────────────────────────
        remaining_files = [f for f in extract_dir.rglob("*") if f.is_file()]
        if not remaining_files:
            log("  ゴミ除去後にファイルが残りませんでした。元ファイルは変更しません。")
            return "skipped"

        # ── ルートフォルダ剥がし ──────────────────────────────
        actual_root = strip_root_folder(extract_dir)

        # ── 無圧縮ZIP生成（一時ファイル）─────────────────────
        tmp_zip = tmp_path / (archive_path.stem + ".zip")
        make_store_zip(actual_root, tmp_zip)

        # ── 元ファイルをゴミ箱へ ─────────────────────────────
        log(f"  ゴミ箱へ送信: {archive_path.name}")
        send_to_recycle_bin(archive_path)

        # ── 出力先パスを決定（同名ファイルがあれば設定に従い回避）──
        style = config.get("duplicate_name_style", "counter")
        output_path = _unique_path(archive_path.with_suffix(".zip"), style)
        if output_path != archive_path.with_suffix(".zip"):
            log(f"  同名ファイルが存在するため変更: {output_path.name}")

        # ── 一時ZIP を最終パスへ移動 ─────────────────────────
        shutil.move(str(tmp_zip), str(output_path))

        # ── タイムスタンプ ───────────────────────────────────
        if config.get("preserve_timestamp", True):
            apply_timestamp(output_path, original_mtime)
        else:
            log("  タイムスタンプ: 処理日時のまま")

    log_ok(f"完了: {output_path.name}")
    return "converted"


# ============================================================
# エントリポイント
# ============================================================
# ── スリープ抑止（Windows のみ）────────────────────────────────────────
_ES_CONTINUOUS     = 0x80000000
_ES_SYSTEM_REQUIRED = 0x00000001

def _prevent_sleep() -> None:
    """処理中にシステムスリープが発生しないよう Windows に通知する。"""
    if sys.platform == "win32":
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(
                _ES_CONTINUOUS | _ES_SYSTEM_REQUIRED
            )
        except Exception:
            pass

def _allow_sleep() -> None:
    """スリープ抑止を解除する。"""
    if sys.platform == "win32":
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(_ES_CONTINUOUS)
        except Exception:
            pass


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

    if config.get("write_log"):
        setup_log_file(exe_dir)

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

    # ファイルを並列処理
    # I/Oバウンドな処理なのでスレッドで十分。
    # ワーカー数はファイル数と上限(4)の小さい方。
    errors: list[str] = []
    n_workers = min(4, len(args))

    def _run(arg: str) -> tuple[str, str]:
        """1ファイルをバッファ付きで処理し (name, status) を返す。"""
        p = Path(arg)
        _thread_local.buffer = []
        try:
            status = process_file(p, sevenzip, config, exe_dir)
        except Exception:
            log_error(f"予期しないエラー\n{traceback.format_exc()}")
            status = "error"
        finally:
            _flush_buffer()
            _thread_local.buffer = None
        return p.name, status

    results: list[tuple[str, str]] = []  # (name, status)

    _prevent_sleep()
    try:
        if n_workers == 1:
            results.append(_run(args[0]))
        else:
            log(f"並列処理開始: {len(args)} ファイル / {n_workers} ワーカー")
            print()
            with ThreadPoolExecutor(max_workers=n_workers) as executor:
                futures = {executor.submit(_run, a): a for a in args}
                for future in as_completed(futures):
                    try:
                        results.append(future.result())
                    except Exception:
                        log_error(f"予期しないエラー\n{traceback.format_exc()}")
                        results.append((Path(futures[future]).name, "error"))
    finally:
        _allow_sleep()

    # ── サマリー表示 ────────────────────────────────────────
    converted = [(n, s) for n, s in results if s == "converted"]
    skipped   = [(n, s) for n, s in results if s == "skipped"]
    errors    = [(n, s) for n, s in results if s == "error"]

    print()
    print("=" * 60)
    print(f"  [結果] 変換: {len(converted)}件  スキップ: {len(skipped)}件  エラー: {len(errors)}件")
    print()

    if converted:
        print(f"  変換完了 ({len(converted)}件):")
        for name, _ in converted:
            print(f"    [OK]   {name}")
    if skipped:
        print(f"  スキップ ({len(skipped)}件):")
        for name, _ in skipped:
            print(f"    [--]   {name}")
    if errors:
        print(f"  エラー ({len(errors)}件):")
        for name, _ in errors:
            print(f"    [ERR]  {name}")

    print()
    print("=" * 60)
    input("Enter キーで終了...")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
