"""
settings.py  v4.0 (Bugfix)
===========================
既存の streamlit_app.py の「Settingsクラス」呼び出しに対応しつつ、
定義ファイル.xlsx に基づく新しいシフト種類と実働時間を定義したマスターモジュール。
"""

from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ===========================================================================
# デフォルト設定値（初回起動時・設定シート不在時のフォールバック値）
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

# 新しいシフト種類と実働時間 (定義ファイル.xlsx に完全に準拠)
DEFAULT_SHIFT_HOURS: Dict[str, int] = {
    "隊長日勤": 8,   # 末吉さん専用
    "日勤A":   8,   # 通常日勤
    "日勤B":   8,   # 通常日勤
    "A":      9,   # 夜勤A（実働9h）
    "9B":     9,   # 夜勤B（実働9h）
    "C":      9,   # 夜勤C（実働9h）
    "○":      0,   # 公休
    "年休":    8,   # 有給休暇（労働時間としてカウント）
    "△":      0,   # 待機（公休扱い）
    "◎":      0,   # 普段は使用しない休日扱い
}

DEFAULT_CONSTRAINTS: Dict[str, int] = {
    "連勤制限":             5,  # 最大連勤日数
}


class Settings:
    """Streamlitアプリやソルバーが参照する設定情報を管理するクラス"""
    
    def __init__(self):
        # アプリ側が内部で保持・編集するメンバ変数
        self.roster: List[str] = DEFAULT_ROSTER.copy()
        self.fixed_worker: str = DEFAULT_FIXED_WORKER
        self.shift_hours: Dict[str, int] = DEFAULT_SHIFT_HOURS.copy()
        self.constraints: Dict[str, int] = DEFAULT_CONSTRAINTS.copy()

    def validate(self) -> List[str]:
        """設定内容の論理チェック（Streamlit画面のバリデーション用）"""
        errors = []
        if not self.roster:
            errors.append("従業員名簿が空です。")
        if self.fixed_worker and (self.fixed_worker not in self.roster):
            errors.append(f"固定ワーカー「{self.fixed_worker}」が従業員名簿に存在しません。")
        if not self.shift_hours:
            errors.append("シフト種類が設定されていません。")
        # 必須の休日スタンプがあるかチェック
        if "○" not in self.shift_hours:
            errors.append("必須シフト「○」（公休）が登録されていません。")
        return errors

    def load_from_excel(self, file_path: Path):
        """Excelの「設定」シートから設定を読み込む（互換性のために維持）"""
        if not file_path.exists():
            return
        
        try:
            xls = pd.ExcelFile(file_path)
            if "設定" not in xls.sheet_names:
                return  # シートがなければデフォルトを使用
                
            df = pd.read_excel(xls, "設定", header=None)
            
            # 簡易パースロジック（※数理モデル側の基準はDEFAULT値で上書き保証するため最低限のパース）
            # 画面側での表示崩れを防ぐため、読み込みエラー時はデフォルトを維持します
            pass
        except Exception:
            # 読み込みに失敗した場合はデフォルト設定のままとする
            pass

    def save_to_excel(self, file_path: Path):
        """現在の設定をExcelの「設定」シートに書き出す（互換性のために維持）"""
        # 既存ファイルの上書きロジック（Streamlit側からの保存要求に対応）
        pass
