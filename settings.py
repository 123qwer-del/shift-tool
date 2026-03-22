"""
settings.py  v1.0
==================
設定の読み書きを一元管理するモジュール。
設定は input.xlsx の「設定」シートに保存される。

シート構成（各テーブルを縦に並べる）:
  ブロック1: 従業員名簿
  ブロック2: 固定ワーカー設定
  ブロック3: シフト種類・勤務時間
  ブロック4: 制約パラメータ
"""

from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===========================================================================
# デフォルト設定値（初回起動時・設定シート不在時のフォールバック）
# ===========================================================================

DEFAULT_ROSTER: List[str] = [
    "末吉 弘一",
    "伊藤 晶俊",
    "吉村 智",
    "南 英俊",
    "杉田 孝行",
    "山田 誠",
    "大西 信一",
    "村主 博",
    "河内 拳",
]

DEFAULT_FIXED_WORKER: str = "末吉 弘一"

DEFAULT_SHIFT_HOURS: Dict[str, int] = {
    "日勤":  8,
    "夜勤A": 11,
    "夜勤B": 10,
    "夜勤C": 11,
    "休日":  0,
}

DEFAULT_CONSTRAINTS: Dict[str, int] = {
    "月間上限時間":         178,
    "日勤必要人数":           2,
    "夜勤必要人数":           3,
    "最大連続勤務日数":       4,
    "週休判定ウィンドウ幅":   7,
}

CONSTRAINT_DESCRIPTIONS: Dict[str, str] = {
    "月間上限時間":         "1人あたりの月間最大労働時間 (h)",
    "日勤必要人数":         "1日に必要な日勤担当者数",
    "夜勤必要人数":         "1日に必要な夜勤担当者数 (A+B+C の合計)",
    "最大連続勤務日数":     "連続して勤務できる最大日数",
    "週休判定ウィンドウ幅": "週1休を判定するスライディングウィンドウの幅 (日)",
}

SHEET_NAME = "設定"


# ===========================================================================
# 設定クラス
# ===========================================================================

class Settings:
    """
    アプリ全体の設定を保持するデータクラス。
    load() / save() で input.xlsx の「設定」シートと同期する。
    """

    def __init__(self):
        self.roster: List[str]          = list(DEFAULT_ROSTER)
        self.fixed_worker: str          = DEFAULT_FIXED_WORKER
        self.shift_hours: Dict[str, int] = dict(DEFAULT_SHIFT_HOURS)
        self.constraints: Dict[str, int] = dict(DEFAULT_CONSTRAINTS)

    # -----------------------------------------------------------------------
    # 読み込み
    # -----------------------------------------------------------------------
    @classmethod
    def load(cls, filepath: Path) -> "Settings":
        """
        input.xlsx の「設定」シートから設定を読み込む。
        シートが存在しない場合はデフォルト値で初期化した Settings を返す。
        """
        s = cls()
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
        except FileNotFoundError:
            return s  # ファイル不在 → デフォルト返却

        if SHEET_NAME not in wb.sheetnames:
            return s  # 設定シート未作成 → デフォルト返却

        ws = wb[SHEET_NAME]
        rows = [[cell.value for cell in row] for row in ws.iter_rows()]

        # ブロックごとにパース
        s._parse_roster(rows)
        s._parse_fixed_worker(rows)
        s._parse_shift_hours(rows)
        s._parse_constraints(rows)
        return s

    def _find_block(self, rows: list, header: str) -> Optional[int]:
        """指定ヘッダー行のインデックスを返す。見つからなければ None。"""
        for i, row in enumerate(rows):
            if row and str(row[0]).strip() == header:
                return i
        return None

    def _parse_roster(self, rows: list):
        idx = self._find_block(rows, "■ 従業員名簿")
        if idx is None:
            return
        roster = []
        for row in rows[idx + 2:]:  # ヘッダー行をスキップ
            if not row or row[0] is None or str(row[0]).strip() == "":
                break
            name = str(row[0]).strip()
            if name:
                roster.append(name)
        if roster:
            self.roster = roster

    def _parse_fixed_worker(self, rows: list):
        idx = self._find_block(rows, "■ 固定ワーカー設定")
        if idx is None:
            return
        # idx+1 行目: ※注釈行（スキップ）
        # idx+2 行目: 列ヘッダー「設定項目 / 値」（スキップ）
        # idx+3 行目: データ行（「固定ワーカー名」| 実際の名前）
        for row in rows[idx + 3:]:
            if not row or row[0] is None or str(row[0]).strip() == "":
                break
            # 「設定項目」列が「固定ワーカー名」で、値列に名前が入っている
            val = row[1] if len(row) > 1 else None
            if val is not None and str(val).strip():
                self.fixed_worker = str(val).strip()
            break  # データ行は1行のみ

    def _parse_shift_hours(self, rows: list):
        idx = self._find_block(rows, "■ シフト種類・勤務時間")
        if idx is None:
            return
        shift_hours = {}
        for row in rows[idx + 2:]:
            if not row or row[0] is None or str(row[0]).strip() == "":
                break
            shift = str(row[0]).strip()
            try:
                hours = int(row[1]) if len(row) > 1 and row[1] is not None else 0
                shift_hours[shift] = hours
            except (ValueError, TypeError):
                pass
        if shift_hours:
            self.shift_hours = shift_hours

    def _parse_constraints(self, rows: list):
        idx = self._find_block(rows, "■ 制約パラメータ")
        if idx is None:
            return
        for row in rows[idx + 2:]:
            if not row or row[0] is None or str(row[0]).strip() == "":
                break
            key = str(row[0]).strip()
            try:
                val = int(row[1]) if len(row) > 1 and row[1] is not None else None
                if key in self.constraints and val is not None:
                    self.constraints[key] = val
            except (ValueError, TypeError):
                pass

    # -----------------------------------------------------------------------
    # 保存
    # -----------------------------------------------------------------------
    def save(self, filepath: Path):
        """
        設定を input.xlsx の「設定」シートに書き込む。
        ファイルが存在しない場合は新規作成する。
        既存シートがあれば上書き（内容をクリアして再描画）。
        """
        try:
            wb = openpyxl.load_workbook(filepath)
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]

        # 既存の設定シートを削除して再作成
        if SHEET_NAME in wb.sheetnames:
            del wb[SHEET_NAME]
        ws = wb.create_sheet(SHEET_NAME, 0)  # 先頭に挿入

        writer = _SheetWriter(ws)
        writer.write_roster(self.roster)
        writer.write_fixed_worker(self.fixed_worker, self.roster)
        writer.write_shift_hours(self.shift_hours)
        writer.write_constraints(self.constraints)
        writer.adjust_columns()

        wb.save(filepath)

    # -----------------------------------------------------------------------
    # バリデーション
    # -----------------------------------------------------------------------
    def validate(self) -> List[str]:
        """
        設定値の整合性チェック。
        問題があればエラーメッセージのリストを返す（空なら正常）。
        """
        errors = []

        if not self.roster:
            errors.append("従業員名簿が空です。")

        if self.fixed_worker and self.fixed_worker not in self.roster:
            errors.append(
                f"固定ワーカー「{self.fixed_worker}」が従業員名簿に存在しません。"
            )

        for shift in ["日勤", "休日"]:
            if shift not in self.shift_hours:
                errors.append(f"シフト種類「{shift}」が定義されていません。")

        req_day   = self.constraints.get("日勤必要人数", 2)
        req_night = self.constraints.get("夜勤必要人数", 3)
        n_shift_workers = len([w for w in self.roster if w != self.fixed_worker])
        if n_shift_workers < req_day + req_night:
            errors.append(
                f"シフト対象者({n_shift_workers}名) < "
                f"必要人数(日勤{req_day} + 夜勤{req_night} = {req_day+req_night}名)"
            )

        return errors

    # -----------------------------------------------------------------------
    # optimizer.py 向けの変換ヘルパー
    # -----------------------------------------------------------------------
    @property
    def night_shifts(self) -> List[str]:
        """日勤・休日以外のシフトをすべて夜勤とみなして返す。"""
        return [s for s in self.shift_hours if s not in ("日勤", "休日")]

    @property
    def shift_types(self) -> List[str]:
        """シフト種類のリスト（休日を末尾に固定）。"""
        non_rest = [s for s in self.shift_hours if s != "休日"]
        return non_rest + ["休日"]


# ===========================================================================
# Excelシート書き込みヘルパー
# ===========================================================================

class _SheetWriter:
    """設定シートへの装飾付き書き込みを担うプライベートクラス。"""

    # スタイル定義
    COLOR_HEADER_BG = "1A1A2E"   # タイトル行背景（濃紺）
    COLOR_HEADER_FG = "FFFFFF"   # タイトル行文字（白）
    COLOR_COL_BG    = "3A3A5C"   # 列ヘッダー背景（中紺）
    COLOR_COL_FG    = "FFFFFF"   # 列ヘッダー文字（白）
    COLOR_DATA_ALT  = "F0F4FF"   # 偶数行の縞模様（薄青）
    COLOR_NOTE      = "888888"   # 説明文字（灰色）

    def __init__(self, ws):
        self.ws  = ws
        self.row = 1  # 現在の書き込み行（1始まり）

    # ---- ブロック書き込み ----

    def write_roster(self, roster: List[str]):
        self._section_header("■ 従業員名簿")
        self._col_headers(["名前"])
        for i, name in enumerate(roster):
            self._data_row([name], i)
        self.row += 1  # ブロック間の空行

    def write_fixed_worker(self, fixed_worker: str, roster: List[str]):
        self._section_header("■ 固定ワーカー設定")
        self._note_row("平日は日勤固定・土日は休日固定となる従業員を指定します（1名のみ・空欄で無効）")
        self._col_headers(["設定項目", "値"])
        self._data_row(["固定ワーカー名", fixed_worker], 0)
        self.row += 1

    def write_shift_hours(self, shift_hours: Dict[str, int]):
        self._section_header("■ シフト種類・勤務時間")
        self._note_row("シフト名を変更する場合は optimizer.py 内の参照箇所も合わせて確認してください")
        self._col_headers(["シフト名", "勤務時間 (h)"])
        for i, (shift, hours) in enumerate(shift_hours.items()):
            self._data_row([shift, hours], i)
        self.row += 1

    def write_constraints(self, constraints: Dict[str, int]):
        self._section_header("■ 制約パラメータ")
        self._col_headers(["パラメータ名", "値", "説明"])
        for i, (key, val) in enumerate(constraints.items()):
            desc = CONSTRAINT_DESCRIPTIONS.get(key, "")
            self._data_row([key, val, desc], i)
        self.row += 1

    # ---- ロープライベートユーティリティ ----

    def _section_header(self, title: str):
        cell = self.ws.cell(self.row, 1, title)
        cell.font      = Font(bold=True, color=self.COLOR_HEADER_FG, size=11)
        cell.fill      = PatternFill("solid", fgColor=self.COLOR_HEADER_BG)
        cell.alignment = Alignment(vertical="center", indent=1)
        # 3列分をスタイル塗りつぶし（結合はしない）
        for c in range(2, 4):
            self.ws.cell(self.row, c).fill = PatternFill("solid", fgColor=self.COLOR_HEADER_BG)
        self.ws.row_dimensions[self.row].height = 22
        self.row += 1

    def _note_row(self, text: str):
        cell = self.ws.cell(self.row, 1, f"  ※ {text}")
        cell.font      = Font(color=self.COLOR_NOTE, italic=True, size=9)
        cell.alignment = Alignment(vertical="center")
        self.row += 1

    def _col_headers(self, headers: List[str]):
        for c, h in enumerate(headers, start=1):
            cell = self.ws.cell(self.row, c, h)
            cell.font      = Font(bold=True, color=self.COLOR_COL_FG)
            cell.fill      = PatternFill("solid", fgColor=self.COLOR_COL_BG)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        self.ws.row_dimensions[self.row].height = 18
        self.row += 1

    def _data_row(self, values: list, index: int):
        bg = self.COLOR_DATA_ALT if index % 2 == 1 else None
        thin = Side(style="thin", color="CCCCCC")
        border = Border(bottom=Side(style="hair", color="DDDDDD"))
        for c, val in enumerate(values, start=1):
            cell = self.ws.cell(self.row, c, val)
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)
            cell.border    = border
            cell.alignment = Alignment(vertical="center", indent=1)
        self.row += 1

    def adjust_columns(self):
        """列幅を内容に合わせて自動調整する。"""
        col_widths = {}
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                col = cell.column
                length = len(str(cell.value)) * 2  # 日本語考慮で×2
                col_widths[col] = max(col_widths.get(col, 10), min(length, 50))
        for col, width in col_widths.items():
            self.ws.column_dimensions[get_column_letter(col)].width = width + 2
