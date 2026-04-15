"""
store-zip.py v2.0 - 圧縮ファイルを無圧縮ZIPに変換するツール

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
import re
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
import hashlib
import unicodedata
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
    """write_log=true のとき呼び出す。store-zip.log への Tee を設定する。"""
    log_path = exe_dir / "store-zip.log"
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
_LOG_STATUS_WIDTH = 5
_LOG_LABEL_WIDTH = 22
_LOG_LABEL_MAX_WIDTH = 30


def _emit(line: str) -> None:
    """1行出力。バッファ中ならバッファへ、そうでなければ直接 stdout へ。"""
    buf: list | None = getattr(_thread_local, "buffer", None)
    if buf is not None:
        buf.append(line)
    else:
        sys.stdout.write(line)
        sys.stdout.flush()


def _display_width(text: str) -> int:
    width = 0
    for ch in text:
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def _pad_display(text: str, width: int) -> str:
    return text + " " * max(0, width - _display_width(text))


def _format_structured_message(msg: str, label_width: int = _LOG_LABEL_WIDTH) -> str:
    """`ラベル: 詳細` 形式のログを見やすく桁揃えする。"""
    if "\n" in msg:
        return msg

    indent_len = len(msg) - len(msg.lstrip(" "))
    indent = msg[:indent_len]
    body = msg[indent_len:]
    if ":" not in body:
        return msg

    label, detail = body.split(":", 1)
    label = label.rstrip()
    detail = detail.lstrip()
    if not label or _display_width(label) > _LOG_LABEL_MAX_WIDTH:
        return msg

    return f"{indent}{_pad_display(label, label_width)} : {detail}"


def _format_status_message(status: str, msg: str) -> str:
    status_prefix = f"{status:<{_LOG_STATUS_WIDTH}} "
    label_width = max(1, _LOG_LABEL_WIDTH - _display_width(status_prefix))
    return f"{status_prefix}{_format_structured_message(msg, label_width)}"


def log(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] {_format_structured_message(msg)}\n")

def log_error(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] {_format_status_message('ERROR', msg)}\n")

def log_ok(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] {_format_status_message('OK', msg)}\n")

def log_skip(msg: str) -> None:
    _emit(f"[{datetime.now().strftime('%H:%M:%S')}] {_format_status_message('SKIP', msg)}\n")


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
    "preserve_timestamp": True,
    "remove_empty_dirs": True,
    "write_log": False,
    # 0 = 制限なし
    "file_list_limit": 0,
    # "zip" : 無圧縮 ZIP（デフォルト）
    # "rar" : 無圧縮 RAR + リカバリーレコード（WinRAR の rar.exe が必要）
    "output_format": "zip",
    # リカバリーレコードの割合（%）。output_format = "rar" のときのみ有効。
    "rar_recovery_record": 5,
    # rar.exe のパスを明示指定する場合のみ設定（空文字 = 自動検索）
    "rar_exe_path": "",
    # 同名・同サイズの出力ファイルが既に存在する場合にスキップするか
    # true  : 既存ファイルと同サイズなら重複とみなし新ファイルを破棄してスキップ
    # false : 通常の重複回避処理（duplicate_name_style に従いリネーム）
    "skip_if_same_size": True,
}

_VALID_UNKNOWN_FILE_ACTIONS = {"ask", "keep", "junk"}
_VALID_DUPLICATE_NAME_STYLES = {"counter", "date"}
_VALID_OUTPUT_FORMATS = {"zip", "rar"}


# ============================================================
# 設定ファイル読み込み（config.toml + decisions.json をマージ）
# ============================================================
def load_config(exe_dir: Path) -> dict:
    config: dict = {
        "junk_patterns":       list(DEFAULT_CONFIG["junk_patterns"]),
        "junk_dirs":           list(DEFAULT_CONFIG["junk_dirs"]),
        "allow_patterns":      list(DEFAULT_CONFIG["allow_patterns"]),
        "unknown_file_action":  DEFAULT_CONFIG["unknown_file_action"],
        "duplicate_name_style":  DEFAULT_CONFIG["duplicate_name_style"],
        "preserve_timestamp":    DEFAULT_CONFIG["preserve_timestamp"],
        "remove_empty_dirs":     DEFAULT_CONFIG["remove_empty_dirs"],
        "write_log":             DEFAULT_CONFIG["write_log"],
        "file_list_limit":       DEFAULT_CONFIG["file_list_limit"],
        "output_format":         DEFAULT_CONFIG["output_format"],
        "rar_recovery_record":   DEFAULT_CONFIG["rar_recovery_record"],
        "rar_exe_path":          DEFAULT_CONFIG["rar_exe_path"],
        "skip_if_same_size":     DEFAULT_CONFIG["skip_if_same_size"],
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
            for k in ("unknown_file_action", "duplicate_name_style", "write_log",
                      "preserve_timestamp", "remove_empty_dirs", "file_list_limit",
                      "output_format", "rar_recovery_record", "rar_exe_path",
                      "skip_if_same_size"):
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

    return _normalize_config(config)


def _normalize_config(config: dict) -> dict:
    """設定値を検証し、不正値は安全なデフォルトへ戻す。"""
    normalized = dict(config)

    if normalized.get("unknown_file_action") not in _VALID_UNKNOWN_FILE_ACTIONS:
        log("警告: unknown_file_action が不正です。'ask' に戻します。")
        normalized["unknown_file_action"] = DEFAULT_CONFIG["unknown_file_action"]

    if normalized.get("duplicate_name_style") not in _VALID_DUPLICATE_NAME_STYLES:
        log("警告: duplicate_name_style が不正です。'counter' に戻します。")
        normalized["duplicate_name_style"] = DEFAULT_CONFIG["duplicate_name_style"]

    if normalized.get("output_format") not in _VALID_OUTPUT_FORMATS:
        log("警告: output_format が不正です。'zip' に戻します。")
        normalized["output_format"] = DEFAULT_CONFIG["output_format"]

    for key in ("preserve_timestamp", "remove_empty_dirs", "write_log", "skip_if_same_size"):
        normalized[key] = bool(normalized.get(key, DEFAULT_CONFIG[key]))

    try:
        normalized["file_list_limit"] = max(0, int(normalized.get("file_list_limit", 0)))
    except (TypeError, ValueError):
        log("警告: file_list_limit が不正です。0 に戻します。")
        normalized["file_list_limit"] = DEFAULT_CONFIG["file_list_limit"]

    try:
        rr = int(normalized.get("rar_recovery_record", DEFAULT_CONFIG["rar_recovery_record"]))
    except (TypeError, ValueError):
        log("警告: rar_recovery_record が不正です。5 に戻します。")
        rr = DEFAULT_CONFIG["rar_recovery_record"]
    normalized["rar_recovery_record"] = min(100, max(1, rr))

    for key in ("junk_patterns", "junk_dirs", "allow_patterns"):
        value = normalized.get(key, DEFAULT_CONFIG[key])
        if isinstance(value, list):
            normalized[key] = [str(item) for item in value]
        else:
            log(f"警告: {key} は配列である必要があります。デフォルトを使用します。")
            normalized[key] = list(DEFAULT_CONFIG[key])

    normalized["rar_exe_path"] = str(normalized.get("rar_exe_path", ""))
    return normalized


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
# WinRAR (rar.exe) の検出
# ============================================================
_WINRAR_CANDIDATES = [
    r"C:\Program Files\WinRAR\rar.exe",
    r"C:\Program Files (x86)\WinRAR\rar.exe",
]

def find_rar(exe_dir: Path, configured_path: str = "") -> str | None:
    # config.toml で明示指定されている場合はそちらを優先
    if configured_path:
        p = Path(configured_path)
        if p.exists():
            return str(p)
        return None  # 指定されているのに見つからない場合は None を返してエラーにする
    # 自動検索
    local = exe_dir / "rar.exe"
    if local.exists():
        return str(local)
    for candidate in _WINRAR_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return shutil.which("rar")


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
        dest_root = dest_dir.resolve()
        # ファイルハンドルを外側の with で管理する。
        # zipfile.ZipFile(filename) は __init__ で BadZipFile を投げると
        # __exit__ が呼ばれずハンドルが残るため、open() を外側に置いて確実に閉じる。
        with open(str(archive), "rb") as fp:
            with zipfile.ZipFile(fp) as zf:
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
                        target.relative_to(dest_root)
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
        raise  # 呼び出し元が 7-Zip へのフォールバックを処理する（fp は既に閉じ済み）
    except RuntimeError as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            log_error("パスワード付きアーカイブは処理できません。")
        else:
            log_error(f"ZIP 展開失敗: {e}")
        return False
    except Exception as e:
        log_error(f"ZIP 展開失敗: {e}")
        return False



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
    with tempfile.TemporaryDirectory(prefix="store_zip_tmp_") as tmp_dir:
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

    # 拡張子単位でグループ化して削除・ログ出力
    # { ext: (example_name, count) }
    ext_groups: dict[str, tuple[str, int]] = {}
    for fpath in files_to_delete:
        try:
            # 読み取り専用ファイルは削除前に書き込み権限を付与する
            if not os.access(fpath, os.W_OK):
                os.chmod(fpath, fpath.stat().st_mode | 0o200)
            fpath.unlink()
            removed += 1
            ext = fpath.suffix.lower() or fpath.name  # 拡張子なしはファイル名そのもの
            if ext not in ext_groups:
                ext_groups[ext] = (fpath.name, 1)
            else:
                ext_groups[ext] = (ext_groups[ext][0], ext_groups[ext][1] + 1)
        except Exception as e:
            log(f"  警告: ファイル削除失敗 {fpath.name}: {e}")

    for ext, (example, count) in ext_groups.items():
        if count == 1:
            log(f"  削除: {example}")
        else:
            log(f"  削除: {example} など {count} 件 (*{ext})")

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
        # open() を外側に置いてファイルハンドルを確実に閉じる。
        # ZipFile(filename) は __init__ で例外が起きると __exit__ が呼ばれないため。
        with open(str(archive_path), "rb") as fp:
            with zipfile.ZipFile(fp) as zf:
                infos = zf.infolist()
                if not infos:
                    return False
                return all(info.compress_type == zipfile.ZIP_STORED for info in infos)
    except Exception:
        return False


def _store_zip_needs_root_strip(archive_path: Path) -> bool:
    """
    ZIP のルート直下に単一フォルダのみ存在し、ファイルが一切ない場合に True を返す。
    この場合のみルートフォルダ除去のため再パックが必要。
    ルートに複数フォルダ、またはファイルが存在する場合は False。
    """
    try:
        with open(str(archive_path), "rb") as fp:
            with zipfile.ZipFile(fp) as zf:
                root_dirs = set()
                for info in zf.infolist():
                    name = info.filename.replace("\\", "/")
                    parts = name.split("/")
                    if not info.is_dir() and len(parts) == 1:
                        return False  # ルート直下にファイルがある
                    if len(parts) >= 2:
                        root_dirs.add(parts[0])
                        if len(root_dirs) > 1:
                            return False  # ルートフォルダが複数ある
                return len(root_dirs) == 1
    except Exception:
        pass
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
def make_store_zip(source_dir: Path, output_zip: Path, original_size: int = 0) -> None:
    """source_dir 以下のファイルを STORE（無圧縮）ZIP に格納する。"""
    # strict_timestamps=False: 1980年以前のタイムスタンプを自動的に 1980-01-01 に切り上げる。
    # ファイル自体は変更しない（read-only ファイルも安全に処理できる）。
    with zipfile.ZipFile(
        output_zip, "w",
        compression=zipfile.ZIP_STORED,
        allowZip64=True,
        strict_timestamps=False,
    ) as zf:
        for dirpath, dirnames, filenames in os.walk(source_dir):
            dirnames.sort()
            filenames.sort()
            base_dir = Path(dirpath)
            for fname in filenames:
                fpath = base_dir / fname
                arcname = fpath.relative_to(source_dir)
                zf.write(fpath, arcname)

    size = output_zip.stat().st_size
    if original_size > 0:
        ratio = size / original_size * 100
        log(f"  無圧縮ZIP生成完了: {output_zip.name}  ({size:,} bytes / 元 {original_size:,} bytes / {ratio:.2f}%)")
    else:
        log(f"  無圧縮ZIP生成完了: {output_zip.name}  ({size:,} bytes)")


# ============================================================
# 無圧縮 RAR 作成（WinRAR の rar.exe を使用）
# ============================================================
def make_store_rar(
    rar_exe: str,
    source_dir: Path,
    output_rar: Path,
    recovery_record: int,
    original_size: int = 0,
) -> bool:
    """source_dir 以下のファイルを無圧縮 RAR に格納する。"""
    cmd = [
        rar_exe, "a",
        "-m0",                       # 無圧縮（Store）
        f"-rr{recovery_record}",     # リカバリーレコード
        "-ep1",                      # ベースディレクトリをパスから除外
        "-r",                        # サブディレクトリを再帰処理
        "-o+",                       # 既存ファイルを上書き
        str(output_rar),
        str(source_dir) + "\\",
    ]
    log(f"  無圧縮RAR生成中 (リカバリーレコード {recovery_record}%)...")
    try:
        result = subprocess.run(cmd, capture_output=True)
    except FileNotFoundError:
        log_error(f"WinRAR (rar.exe) が見つかりません: {rar_exe}")
        return False
    except Exception as e:
        log_error(f"RAR 作成失敗: {e}")
        return False

    if result.returncode not in (0, 1):  # WinRAR は警告時に 1 を返す場合がある
        stderr_str = decode_bytes(result.stderr) if result.stderr else ""
        log_error(f"RAR 作成失敗 (終了コード={result.returncode})")
        if stderr_str.strip():
            log_error(f"  stderr: {stderr_str.strip()[:300]}")
        return False

    size = output_rar.stat().st_size
    if original_size > 0:
        ratio = size / original_size * 100
        log(f"  無圧縮RAR生成完了: {output_rar.name}  ({size:,} bytes / 元 {original_size:,} bytes / {ratio:.2f}%)")
    else:
        log(f"  無圧縮RAR生成完了: {output_rar.name}  ({size:,} bytes)")
    return True


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
    # NAS/SMB パスでは SMB オポチュニスティックロックや Synology インデクサー等が
    # ファイルを掴んでいて WinError 32 になることがある。
    # send2trash → PowerShell の順に試み、各ステップで最大 2 回リトライする。
    # 失敗時に直接削除へフォールバックすると「ゴミ箱へ送る」という期待を破るため行わない。

    def _try_send2trash() -> bool:
        try:
            import send2trash
            send2trash.send2trash(str(path))
            return True
        except Exception:
            return False

    def _try_powershell() -> bool:
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
            return result.returncode == 0
        except Exception:
            return False

    for attempt in range(3):
        if attempt > 0:
            time.sleep(2)  # SMB ロック解放を待つ
        if _try_send2trash():
            return True
        if _try_powershell():
            return True

    log(f"  警告: 元ファイルを削除できませんでした（NAS の SMB ロックが継続中の可能性）。"
        f"\n         手動で削除してください: {path.name}")
    return False


def _hash_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """ファイル全体の SHA-256 を返す。"""
    digest = hashlib.sha256()
    with open(path, "rb") as fp:
        while True:
            chunk = fp.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def files_are_identical(path1: Path, path2: Path) -> bool:
    """サイズ一致に加え、SHA-256 でも同一性を確認する。"""
    try:
        if path1.stat().st_size != path2.stat().st_size:
            return False
        return _hash_file(path1) == _hash_file(path2)
    except Exception as e:
        log(f"  警告: 重複判定に失敗したため別ファイルとして扱います: {e}")
        return False


# ============================================================
# 1ファイルの処理
# ============================================================
def process_file(
    archive_path: Path,
    sevenzip: str,
    config: dict,
    exe_dir: Path,
    rar_exe: str | None = None,
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

    if config.get("output_format", "zip") == "zip" and is_already_store_zip(archive_path):
        if archive_path.suffix.lower() == ".zip":
            if not _store_zip_needs_root_strip(archive_path):
                log_skip(f"既に無圧縮ZIPです。スキップします。")
                return "skipped"
            # ルートに単一フォルダのみ → ルートフォルダ除去のため処理続行
            log(f"  既に無圧縮ZIPですが、ルートフォルダを除去して再パックします。")
        else:
            # 無圧縮ZIPだが拡張子が .zip でない（.cbz 等）→ .zip にコピーしてゴミ箱へ
            log(f"  既に無圧縮ZIPです。拡張子を .zip に修正します。")
            original_mtime = archive_path.stat().st_mtime
            style = config.get("duplicate_name_style", "counter")
            output_path = _unique_path(archive_path.with_suffix(".zip"), style)
            if output_path != archive_path.with_suffix(".zip"):
                log(f"  同名ファイルが存在するため変更: {output_path.name}")
            shutil.copy2(str(archive_path), str(output_path))
            if config.get("preserve_timestamp", True):
                apply_timestamp(output_path, original_mtime)
            log(f"  ゴミ箱へ送信: {archive_path.name}")
            if not send_to_recycle_bin(archive_path):
                log(f"  ※ 元ファイルは手動で削除してください。")
            log_ok(f"完了: {output_path.name}")
            return "converted"

    # 元ファイルのタイムスタンプ・サイズを保存
    stat = archive_path.stat()
    original_mtime = stat.st_mtime
    original_size  = stat.st_size

    with tempfile.TemporaryDirectory(prefix="store_zip_") as tmp_root:
        tmp_path = Path(tmp_root)
        extract_dir = tmp_path / "extracted"
        extract_dir.mkdir()

        # ── 展開 ──────────────────────────────────────────────
        # ZIP/CBZ → Python zipfile で直接展開（ワイルドカード問題なし）。
        #           BadZipFile（実体が RAR 等）の場合は一時コピー経由で再試行。
        #           一時コピー時にマジックバイトで実フォーマットを検出し
        #           正しい拡張子でコピーすることで -mcp=932 の誤適用を防ぐ。
        # RAR/7z 等 → 直接 7-Zip で展開（UNC/角括弧パスも問題なし）。
        ok = False
        if archive_path.suffix.lower() in _ZIP_LIKE_EXTENSIONS:
            try:
                ok = extract_zip_python(archive_path, extract_dir)
            except zipfile.BadZipFile:
                log(f"  ZIP 形式ではありません。7-Zip で再試行します...")
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
        if not any(f.is_file() for f in extract_dir.rglob("*")):
            log("  ゴミ除去後にファイルが残りませんでした。元ファイルは変更しません。")
            return "skipped"

        # ── ルートフォルダ剥がし ──────────────────────────────
        actual_root = strip_root_folder(extract_dir)

        # ── 出力ファイル生成（一時ファイル）──────────────────
        output_format = config.get("output_format", "zip")
        if output_format == "rar":
            tmp_out = tmp_path / (archive_path.stem + ".rar")
            if not make_store_rar(
                rar_exe, actual_root, tmp_out,
                config.get("rar_recovery_record", 5),
                original_size,
            ):
                return "error"
            out_suffix = ".rar"
        else:
            tmp_out = tmp_path / (archive_path.stem + ".zip")
            make_store_zip(actual_root, tmp_out, original_size)
            out_suffix = ".zip"

        intended_path = archive_path.with_suffix(out_suffix)

        if intended_path == archive_path:
            # 入力と出力が同一パス（例: a.zip → a.zip）の場合
            # 一時ファイルをアーカイブと同じフォルダの中間パスに移動してから
            # 元ファイルを削除し、中間ファイルを最終名にリネームする。
            # こうしないとゴミ箱送り中に _unique_path が a(1).zip を返してしまう。
            intermediate = archive_path.parent / (archive_path.stem + "._tmp" + out_suffix)
            shutil.move(str(tmp_out), str(intermediate))
            log(f"  ゴミ箱へ送信: {archive_path.name}")
            if not send_to_recycle_bin(archive_path):
                style = config.get("duplicate_name_style", "counter")
                output_path = _unique_path(intended_path, style)
                log(f"  ※ 元ファイルを保持したまま、変換結果を別名保存します: {output_path.name}")
                shutil.move(str(intermediate), str(output_path))
            else:
                shutil.move(str(intermediate), str(intended_path))
                output_path = intended_path
        else:
            # 入力と出力が異なるパスの場合（例: a.rar → a.zip）
            # 元ファイルを先に削除してから出力先を確定する。
            log(f"  ゴミ箱へ送信: {archive_path.name}")
            if not send_to_recycle_bin(archive_path):
                log(f"  ※ 変換は完了しています。元ファイルは手動で削除してください。")

            if (config.get("skip_if_same_size", True)
                    and intended_path.exists()
                    and files_are_identical(intended_path, tmp_out)):
                tmp_out.unlink()
                log(f"  同名・同内容のファイルが既に存在します: {intended_path.name}")
                log(f"  ※ 前回変換済みの可能性があります。元ファイルの削除に失敗していないか確認してください。")
                log_ok(f"スキップ（重複）: {intended_path.name}")
                return "skipped"

            style = config.get("duplicate_name_style", "counter")
            output_path = _unique_path(intended_path, style)
            if output_path != intended_path:
                log(f"  同名ファイルが存在するため変更: {output_path.name}")

            shutil.move(str(tmp_out), str(output_path))

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


# ============================================================
# 設定メニュー（引数なし起動時）
# ============================================================

_SETTINGS_DEFS = [
    # (表示名,                             key,                    type,   choices)
    ("出力フォーマット",                   "output_format",        "enum", ["zip", "rar"]),
    ("タイムスタンプ保持",                 "preserve_timestamp",   "bool", None),
    ("空フォルダを削除",                   "remove_empty_dirs",    "bool", None),
    ("ログファイル出力",                   "write_log",            "bool", None),
    ("重複サイズスキップ",                 "skip_if_same_size",    "bool", None),
    ("未分類ファイルの処理",               "unknown_file_action",  "enum", ["ask", "keep", "junk"]),
    ("重複ファイルの命名規則",             "duplicate_name_style", "enum", ["counter", "date"]),
    ("ファイル一覧の表示上限 (0=無制限)",  "file_list_limit",      "int",  None),
    ("RARリカバリーレコード (%)",          "rar_recovery_record",  "int",  None),
    ("rar.exe のパス (空=自動検索)",       "rar_exe_path",         "str",  None),
]


def _fmt_val(key: str, value) -> str:
    """メニュー表示用の値文字列を返す。"""
    if key == "rar_exe_path":
        return value if value else "(自動検索)"
    if isinstance(value, bool):
        return "有効" if value else "無効"
    return str(value)


def _update_toml_value(content: str, key: str, value) -> str:
    """TOML テキスト内の `key = ...` を書き換える。キーがなければ末尾に追加。"""
    if isinstance(value, bool):
        new_val = "true" if value else "false"
    elif isinstance(value, int):
        new_val = str(value)
    elif isinstance(value, str):
        new_val = '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'
    else:
        return content
    new_content, n = re.subn(
        rf'^({re.escape(key)}\s*=\s*).*$',
        rf'\g<1>{new_val}',
        content,
        flags=re.MULTILINE,
    )
    if n == 0:
        new_content = content.rstrip("\n") + f"\n{key} = {new_val}\n"
    return new_content


def _save_config_scalars(config_path: Path, config: dict) -> None:
    """config.toml のスカラー設定のみ更新して保存する。配列・コメントは保持。"""
    content = config_path.read_text(encoding="utf-8") if config_path.exists() else ""
    for _, key, _, _ in _SETTINGS_DEFS:
        content = _update_toml_value(content, key, config.get(key, DEFAULT_CONFIG[key]))
    config_path.write_text(content, encoding="utf-8")


def _settings_menu(exe_dir: Path) -> None:
    """引数なし起動時の設定変更メニュー。"""
    config = load_config(exe_dir)
    config_path = exe_dir / "config.toml"
    changed = False

    while True:
        print()
        print("-" * 60)
        print("  設定メニュー  (config.toml)")
        print("-" * 60)
        print()
        for i, (label, key, _, _) in enumerate(_SETTINGS_DEFS, 1):
            val = _fmt_val(key, config.get(key, DEFAULT_CONFIG[key]))
            print(f"  [{i:>2}] {label:<34} {val}")
        print()
        print("  [s] 保存して終了  [q] 保存せず終了")
        print()
        choice = input("選択 > ").strip().lower()

        if choice == "q":
            if changed:
                print("変更は保存されませんでした。")
            break

        if choice == "s":
            _save_config_scalars(config_path, config)
            print(f"保存しました: {config_path}")
            break

        try:
            idx = int(choice) - 1
            if not (0 <= idx < len(_SETTINGS_DEFS)):
                raise ValueError
        except ValueError:
            print("  無効な入力です。番号か s / q を入力してください。")
            continue

        label, key, typ, choices = _SETTINGS_DEFS[idx]
        current = config.get(key, DEFAULT_CONFIG[key])

        if typ == "bool":
            config[key] = not current
            changed = True
            print(f"  {label}: {_fmt_val(key, current)} → {_fmt_val(key, config[key])}")

        elif typ == "enum":
            opts = " / ".join(f"[{c}]" if c == current else c for c in choices)
            raw = input(f"  {label} ({opts}): ").strip()
            if raw in choices:
                config[key] = raw
                changed = True
            elif raw == "":
                pass  # 変更なし
            else:
                print(f"  無効な値です。{choices} のいずれかを入力してください。")

        elif typ == "int":
            raw = input(f"  {label} (現在: {current}、空欄でキャンセル): ").strip()
            if raw == "":
                pass
            else:
                try:
                    config[key] = int(raw)
                    changed = True
                except ValueError:
                    print("  整数を入力してください。")

        elif typ == "str":
            raw = input(f"  {label} (現在: {current!r}、空欄=自動検索): ").strip()
            config[key] = raw
            changed = True


def main() -> None:
    # exe のディレクトリを取得（PyInstaller でビルドした場合は sys.executable の親）
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent

    print("=" * 60)
    print("  store-zip v2.0 - 無圧縮ZIP変換ツール")
    print("=" * 60)
    print()

    args = sorted(sys.argv[1:], key=lambda p: Path(p).name.lower())
    if not args:
        print("引数なしで起動しました。")
        print("SendTo から使う場合: エクスプローラーで圧縮ファイルを右クリック → 送る → store-zip")
        _settings_menu(exe_dir)
        input("\nEnter キーで終了...")
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
        log_error("  7z.exe を store-zip.exe と同じフォルダに置いてください。")
        input("Enter キーで終了...")
        sys.exit(1)
    log(f"7-Zip: {sevenzip}")

    # WinRAR 検出（output_format = "rar" のときのみ必須）
    rar_exe = None
    if config.get("output_format", "zip") == "rar":
        rar_exe = find_rar(exe_dir, config.get("rar_exe_path", ""))
        if not rar_exe:
            log_error("WinRAR (rar.exe) が見つかりません。")
            log_error("  https://www.rarlab.com/ からインストールするか、")
            log_error("  config.toml の rar_exe_path にフルパスを指定してください。")
            input("Enter キーで終了...")
            sys.exit(1)
        log(f"WinRAR: {rar_exe}")
    print()

    # 処理対象ファイル一覧を表示
    limit = config.get("file_list_limit", 0)
    print(f"処理対象: {len(args)} ファイル")
    show_args = args if limit <= 0 else args[:limit]
    for i, arg in enumerate(show_args, 1):
        print(f"  [{i:>3}] {Path(arg).name}")
    if limit > 0 and len(args) > limit:
        print(f"         ... 他 {len(args) - limit} ファイル")
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
            status = process_file(p, sevenzip, config, exe_dir, rar_exe)
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
