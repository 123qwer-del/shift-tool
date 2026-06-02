"""
警備員シフト最適化エンジン  v4.4 (DataFrame出力対応版)
=================================================
OR-Tools CP-SAT を使用した警備員シフトスケジューリング

[新ルール反映]
  1. 毎日必要な人数を完全に個別化：
     - 平日: 隊長日勤x1, 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1
     - 土日: 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1 (隊長は自動で公休「○」)
  2. 月の最終日に応じて、全一般スタッフの月間総労働時間を「ぴったり規定時間」に調整
"""

import calendar
import pandas as pd
from typing import Dict, List, Tuple, Any, Optional
from ortools.sat.python import cp_model

class ShiftValidationError(Exception):
    """シフト条件の事前不整合を通知する例外クラス"""
    pass


def generate_shift(
    year: int,
    month: int,
    holiday_requests: Dict[Tuple[str, int], bool],
    fixed_assignments: Dict[Tuple[str, int], str],
    settings: Any,
    prev_month_tail: Optional[Dict[str, List[str]]] = None
) -> pd.DataFrame:
    """
    [最重要] streamlit_app.py から呼び出される関数。
    ソルバーの計算結果をPandas DataFrame（表データ）に変換して返します。
    """
    optimizer = ShiftOptimizer(
        year=year,
        month=month,
        roster=settings.roster,
        fixed_worker=settings.fixed_worker,
        shift_hours=settings.shift_hours,
        constraints=settings.constraints,
        prev_month_tail=prev_month_tail
    )
    status, schedule = optimizer.solve(holiday_requests, fixed_assignments)

    # 物理的にシフトが組めない条件だった場合のエラーハンドリング
    if status == "INFEASIBLE" or not schedule:
        raise Exception("条件が厳しすぎるため、シフトを作成できませんでした（解なし）。固定シフトの矛盾や、休み希望が多すぎないか確認してください。")

    # 成功した場合、結果を表（DataFrame）に変換
    df = pd.DataFrame(index=settings.roster, columns=list(range(1, optimizer.num_days + 1)))
    for w in settings.roster:
        for d in range(1, optimizer.num_days + 1):
            df.loc[w, d] = schedule.get((w, d), "○")
            
    return df


class ShiftOptimizer:
    def __init__(
        self, 
        year: int, 
        month: int, 
        roster: List[str], 
        fixed_worker: str, 
        shift_hours: Dict[str, int], 
        constraints: Dict[str, int],
        prev_month_tail: Optional[Dict[str, List[str]]] = None
    ):
        self.year = year
        self.month = month
        self.roster = roster
        self.fixed_worker = fixed_worker
        self.shift_hours = shift_hours
        self.constraints = constraints
        self.prev_month_tail = prev_month_tail if prev_month_tail is not None else {}
        
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

        # 月の規定労働時間の自動決定
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
        for (w, d) in fixed_assignments.keys():
            if w not in self.roster:
                raise ShiftValidationError(f"固定シートにある従業員「{w}」は従業員一覧に存在しません。")
        for (w, d) in holiday_requests.keys():
            if w not in self.roster:
                raise ShiftValidationError(f"希望休シートにある従業員「{w}」は従業員一覧に存在しません。")

        for (w, d) in list(fixed_assignments.keys()) + list(holiday_requests.keys()):
            if d < 1 or d > self.num_days:
                raise ShiftValidationError(f"日付「{d}日」が{self.month}月の範囲を超えています。")

        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday and (w, d) in fixed_assignments:
                fixed_s = fixed_assignments[(w, d)]
                if self.shift_hours.get(fixed_s, 0) > 0:
                    raise ShiftValidationError(f"{w}さんは{d}日に「希望休」ですが、固定シートで「{fixed_s}」が指定され矛盾しています。")

        for d in self.days:
            if (self.fixed_worker, d) in fixed_assignments:
                s = fixed_assignments[(self.fixed_worker, d)]
                if self._is_weekend(d) and self.shift_hours.get(s, 0) > 0:
                    raise ShiftValidationError(f"固定ワーカー({self.fixed_worker})の土日({d}日)に勤務シフト「{s}」が固定されています。")
                if not self._is_weekend(d) and s in self.daily_required_shifts:
                    raise ShiftValidationError(f"固定ワーカー({self.fixed_worker})の平日({d}日)に一般シフト「{s}」が固定されています。隊長日勤にしてください。")

        night_shifts = ["A", "9B", "C"]
        for d in self.days[:-1]:
            for w in self.roster:
                if (w, d) in fixed_assignments and (w, d+1) in fixed_assignments:
                    s1 = fixed_assignments[(w, d)]
                    s2 = fixed_assignments[(w, d+1)]
                    if s1 in night_shifts and s2 in ["日勤A", "日勤B", "隊長日勤"]:
                        raise ShiftValidationError(f"{w}さんは{d}日に夜勤、翌{d+1}日に日勤が固定されており、夜勤明けルールに違反します。")

    def solve(self, holiday_requests: Dict[Tuple[str, int], bool], fixed_assignments: Dict[Tuple[str, int], str]) -> Tuple[str, Dict[Tuple[str, int], str]]:
        self.validate_inputs(holiday_requests, fixed_assignments)

        model = cp_model.CpModel()

        # --- 変数の定義 ---
        x = {}
        for w in self.roster:
            for d in self.days:
                for s in self.all_shifts:
                    x[w, d, s] = model.NewBoolVar(f"x_{w}_{d}_{s}")

        # --- 基本制約 ---
        for w in self.roster:
            for d in self.days:
                model.Add(sum(x[w, d, s] for s in self.all_shifts) == 1)

        for (w, d), s in fixed_assignments.items():
            if s in x:
                model.Add(x[w, d, s] == 1)

        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday and "○" in self.all_shifts:
                model.Add(x[w, d, "○"] == 1)

        # --- 固定ワーカー制約 ---
        for d in self.days:
            if (self.fixed_worker, d) not in fixed_assignments:
                if self._is_weekend(d):
                    model.Add(x[self.fixed_worker, d, "○"] == 1)
                else:
                    model.Add(x[self.fixed_worker, d, "隊長日勤"] == 1)

        for w in self.shift_workers:
            for d in self.days:
                model.Add(x[w, d, "隊長日勤"] == 0)

        # --- 毎日の必要人数制約 ---
        for d in self.days:
            for s in self.daily_required_shifts:
                model.Add(sum(x[w, d, s] for w in self.shift_workers) == 1)

        # --- 労働基準・勤務健康ルール ---
        night_shifts = ["A", "9B", "C"]
        for w in self.shift_workers:
            for d in self.days[:-1]:
                is_night = sum(x[w, d, s] for s in night_shifts)
                is_work_next = sum(x[w, d+1, s] for s in self.work_shifts)
                model.Add(is_night + is_work_next <= 1)

        for w in self.shift_workers:
            if w in self.prev_month_tail and len(self.prev_month_tail[w]) > 0:
                last_shift_prev = self.prev_month_tail[w][-1]
                if last_shift_prev in ["A", "9B", "C", "夜勤A", "夜勤B", "夜勤C"]:
                    model.Add(sum(x[w, 1, s] for s in self.work_shifts) == 0)

        max_consecutive = self.constraints.get("連勤制限", 5)
        for w in self.shift_workers:
            for d in range(1, self.num_days - max_consecutive + 1):
                model.Add(sum(x[w, d + i, s] for i in range(max_consecutive + 1) for s in self.work_shifts) <= max_consecutive)

        for w in self.shift_workers:
            for d in range(1, self.num_days - 6 + 1):
                model.Add(sum(x[w, d + i, s] for i in range(7) for s in self.holiday_shifts) >= 1)

        # --- 月間労働時間の調整制約 ---
        for w in self.shift_workers:
            total_hours = sum(x[w, d, s] * self.shift_hours[s] for d in self.days for s in self.all_shifts)
            model.Add(total_hours == self.target_hours)

        # --- 目的関数の設定 ---
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
