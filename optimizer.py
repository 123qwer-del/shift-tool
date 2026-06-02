"""
警備員シフト最適化エンジン  v4.1 (互換性完全修正版)
=================================================
OR-Tools CP-SAT を使用した警備員シフトスケジューリング

[新ルール反映]
  1. 毎日必要な人数を完全に個別化：
     - 平日: 隊長日勤x1, 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1
     - 土日: 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1 (隊長は自動で公休「○」)
  2. 月の最終日に応じて、全一般スタッフの月間総労働時間を「ぴったり規定時間」に調整
     - 31日の月：176時間 (勤務22日×8h相当)
     - 30日の月：168時間 (勤務21日×8h相当)
     - 28日の月：160時間 (勤務20日×8h相当)
     ※夜勤(9h)と日勤(8h)と公休(0h)の日数バランスをソルバーが自動調整します。
"""

import calendar
from typing import Dict, List, Tuple, Any
from ortools.sat.python import cp_model

class ShiftValidationError(Exception):
    """シフト条件の事前不整合を通知する例外クラス"""
    pass


def generate_shift(
    year: int,
    month: int,
    holiday_requests: Dict[Tuple[str, int], bool],
    fixed_assignments: Dict[Tuple[str, int], str],
    settings_obj: Any
) -> Tuple[str, Dict[Tuple[str, int], str]]:
    """
    [最重要] streamlit_app.py から直接呼び出される関数。
    画面から渡された年、月、希望休、固定配置、設定オブジェクトを元に最適化を実行します。
    """
    optimizer = ShiftOptimizer(
        year=year,
        month=month,
        roster=settings_obj.roster,
        fixed_worker=settings_obj.fixed_worker,
        shift_hours=settings_obj.shift_hours,
        constraints=settings_obj.constraints
    )
    return optimizer.solve(holiday_requests, fixed_assignments)


class ShiftOptimizer:
    def __init__(self, year: int, month: int, roster: List[str], fixed_worker: str, shift_hours: Dict[str, int], constraints: Dict[str, int]):
        self.year = year
        self.month = month
        self.roster = roster
        self.fixed_worker = fixed_worker
        self.shift_hours = shift_hours
        self.constraints = constraints
        
        # 月の日数を取得
        _, self.num_days = calendar.monthrange(year, month)
        self.days = list(range(1, self.num_days + 1))
        
        # シフト種類を分類
        self.all_shifts = list(shift_hours.keys())
        self.work_shifts = [s for s in self.all_shifts if shift_hours[s] > 0]
        self.holiday_shifts = [s for s in self.all_shifts if shift_hours[s] == 0]
        
        # 毎日必要な通常シフト（一般スタッフ用）
        self.daily_required_shifts = ["日勤A", "日勤B", "A", "9B", "C"]
        
        # シフト対象ワーカー（末吉さんを除く8名）
        self.shift_workers = [w for w in self.roster if w != self.fixed_worker]

        # 月の規定労働時間の自動決定（定義ファイル.xlsx に準拠）
        if self.num_days == 31:
            self.target_hours = 176
        elif self.num_days == 30:
            self.target_hours = 168
        else:
            self.target_hours = 160

    def _is_weekend(self, day: int) -> bool:
        """指定した日が土曜日または日曜日かを判定"""
        weekday = calendar.weekday(self.year, self.month, day)
        return weekday in (calendar.SATURDAY, calendar.SUNDAY)

    def validate_inputs(self, holiday_requests: Dict[Tuple[str, int], bool], fixed_assignments: Dict[Tuple[str, int], str]):
        """事前バリデーションロジック"""
        # [V1] 従業員名の存在チェック
        for (w, d) in fixed_assignments.keys():
            if w not in self.roster:
                raise ShiftValidationError(f"固定シートにある従業員「{w}」は従業員一覧に存在しません。")
        for (w, d) in holiday_requests.keys():
            if w not in self.roster:
                raise ShiftValidationError(f"希望休シートにある従業員「{w}」は従業員一覧に存在しません。")

        # [V2] 日付の範囲チェック
        for (w, d) in list(fixed_assignments.keys()) + list(holiday_requests.keys()):
            if d < 1 or d > self.num_days:
                raise ShiftValidationError(f"日付「{d}日」が{self.month}月の範囲を超えています。")

        # [V4] 希望休と固定シフトの衝突チェック
        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday and (w, d) in fixed_assignments:
                fixed_s = fixed_assignments[(w, d)]
                if self.shift_hours.get(fixed_s, 0) > 0:
                    raise ShiftValidationError(f"{w}さんは{d}日に「希望休」ですが、固定シートで「{fixed_s}」が指定され矛盾しています。")

        # [V5] 固定ワーカー（末吉さん）のルールチェック
        for d in self.days:
            if (self.fixed_worker, d) in fixed_assignments:
                s = fixed_assignments[(self.fixed_worker, d)]
                if self._is_weekend(d) and self.shift_hours.get(s, 0) > 0:
                    raise ShiftValidationError(f"固定ワーカー({self.fixed_worker})の土日({d}日)に勤務シフト「{s}」が固定されています。")
                if not self._is_weekend(d) and s in self.daily_required_shifts:
                    raise ShiftValidationError(f"固定ワーカー({self.fixed_worker})の平日({d}日)に一般シフト「{s}」が固定されています。隊長日勤にしてください。")

        # [V7] 夜勤翌日の日勤制限チェック
        night_shifts = ["A", "9B", "C"]
        for d in self.days[:-1]:
            for w in self.roster:
                if (w, d) in fixed_assignments and (w, d+1) in fixed_assignments:
                    s1 = fixed_assignments[(w, d)]
                    s2 = fixed_assignments[(w, d+1)]
                    if s1 in night_shifts and s2 in ["日勤A", "日勤B", "隊長日勤"]:
                        raise ShiftValidationError(f"{w}さんは{d}日に夜勤、翌{d+1}日に日勤が固定されており、夜勤明けルールに違反します。")

    def solve(self, holiday_requests: Dict[Tuple[str, int], bool], fixed_assignments: Dict[Tuple[str, int], str]) -> Tuple[str, Dict[Tuple[str, int], str]]:
        # 事前チェック
        self.validate_inputs(holiday_requests, fixed_assignments)

        model = cp_model.CpModel()

        # --- 変数の定義 ---
        x = {}
        for w in self.roster:
            for d in self.days:
                for s in self.all_shifts:
                    x[w, d, s] = model.NewBoolVar(f"x_{w}_{d}_{s}")

        # --- 基本制約 ---
        # 1人1日1シフト
        for w in self.roster:
            for d in self.days:
                model.Add(sum(x[w, d, s] for s in self.all_shifts) == 1)

        # 固定シートの反映
        for (w, d), s in fixed_assignments.items():
            if s in x:
                model.Add(x[w, d, s] == 1)

        # 希望休シートの反映
        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday:
                # 希望休の日は公休「○」を割り当てる
                if "○" in self.all_shifts:
                    model.Add(x[w, d, "○"] == 1)

        # --- 固定ワーカー（末吉さん）専用の自動アサイン制約 ---
        for d in self.days:
            if (self.fixed_worker, d) not in fixed_assignments:
                if self._is_weekend(d):
                    model.Add(x[self.fixed_worker, d, "○"] == 1)  # 土日は公休
                else:
                    model.Add(x[self.fixed_worker, d, "隊長日勤"] == 1)  # 平日は隊長日勤

        # 一般スタッフは「隊長日勤」に入れない
        for w in self.shift_workers:
            for d in self.days:
                model.Add(x[w, d, "隊長日勤"] == 0)

        # --- 毎日の必要人数制約（一般シフトの充填） ---
        for d in self.days:
            for s in self.daily_required_shifts:
                model.Add(sum(x[w, d, s] for w in self.shift_workers) == 1)

        # --- 労働基準・勤務健康ルール ---
        # 夜勤（A, 9B, C）の翌日は必ず休日（○, △, ◎, 年休）
        night_shifts = ["A", "9B", "C"]
        for w in self.shift_workers:
            for d in self.days[:-1]:
                is_night = sum(x[w, d, s] for s in night_shifts)
                is_work_next = sum(x[w, d+1, s] for s in self.work_shifts)
                model.Add(is_night + is_work_next <= 1)

        # 連勤制限
        max_consecutive = self.constraints.get("連勤制限", 5)
        for w in self.shift_workers:
            for d in range(1, self.num_days - max_consecutive + 1):
                model.Add(sum(x[w, d + i, s] for i in range(max_consecutive + 1) for s in self.work_shifts) <= max_consecutive)

        # 週1休の保証（7日間スライディングウィンドウ）
        for w in self.shift_workers:
            for d in range(1, self.num_days - 6 + 1):
                model.Add(sum(x[w, d + i, s] for i in range(7) for s in self.holiday_shifts) >= 1)

        # --- 月間労働時間の「ぴったり調整」制約 ---
        for w in self.shift_workers:
            total_hours = sum(x[w, d, s] * self.shift_hours[s] for d in self.days for s in self.all_shifts)
            model.Add(total_hours == self.target_hours)

        # --- 目的関数の設定（夜勤配分の均等化） ---
        night_counts = []
        for w in self.shift_workers:
            num_nights = model.NewIntVar(0, self.num_days, f"nights_{w}")
            model.Add(num_nights == sum(x[w, d, s] for d in self.days for s in night_shifts))
            night_counts.append(num_nights)

        min_night = model.NewIntVar(0, self.num_days, "min_night")
        max_night = model.NewIntVar(0, self.num_days, "max_night")
        model.AddMinEquality(min_night, night_counts)
        model.AddMaxEquality(max_night, night_counts)
        
        model.Minimize(max_night - min_night)

        # --- ソルバーの実行 ---
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30.0
        
        status = solver.Solve(model)

        # --- 結果の回収 ---
        result_schedule = {}
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            status_str = "OPTIMAL" if status == cp_model.OPTIMAL else "FEASIBLE"
            for w in self.roster:
                for d in self.days:
                    assigned_shift = "○"
                    for s in self.all_shifts:
                        if solver.Value(x[w, d, s]) == 1:
                            assigned_shift = s
                            break
                    result_schedule[(w, d)] = assigned_shift
            return status_str, result_schedule
        else:
            return "INFEASIBLE", {}
