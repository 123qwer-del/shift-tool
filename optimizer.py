"""
警備員シフト最適化エンジン  v4.5
=================================================
OR-Tools CP-SAT を使用した警備員シフトスケジューリング

[シフト体系]
  平日: 隊長日勤x1, 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1
  土日: 日勤Ax1, 日勤Bx1, Ax1, 9Bx1, Cx1  (隊長は自動で○)
  10B : 毎日必須ではなく「時間調整用オプション」。
        月間労働時間を上限に収めるため、ソルバーが必要に応じて割り当てる。

[v4.5 変更点]
  - 10B シフト追加（実働10h、夜勤扱い・夜勤明け翌日禁止・夜勤均等化対象）
  - 月間労働時間制約を「ぴったり規定時間」→「上限以内（<=）」に緩和
  - 目的関数: 夜勤回数の最大-最小差最小化（10B含む）
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
    prev_month_tail: Optional[Dict[str, List[str]]] = None,
) -> pd.DataFrame:
    """streamlit_app.py から呼び出されるメイン関数。"""
    optimizer = ShiftOptimizer(
        year=year,
        month=month,
        roster=settings.roster,
        fixed_worker=settings.fixed_worker,
        shift_hours=settings.shift_hours,
        constraints=settings.constraints,
        prev_month_tail=prev_month_tail,
    )
    status, schedule = optimizer.solve(holiday_requests, fixed_assignments)

    if status == "INFEASIBLE" or not schedule:
        raise Exception(
            "条件が厳しすぎるため、シフトを作成できませんでした（解なし）。"
            "固定シフトの矛盾や、休み希望が多すぎないか確認してください。"
        )

    df = pd.DataFrame(
        index=settings.roster,
        columns=list(range(1, optimizer.num_days + 1)),
    )
    for w in settings.roster:
        for d in range(1, optimizer.num_days + 1):
            df.loc[w, d] = schedule.get((w, d), "○")
    return df


class ShiftOptimizer:
    # 毎日1名ずつ必ず割り当てる一般スタッフ向けシフト（10B は含めない）
    DAILY_REQUIRED_SHIFTS = ["日勤A", "日勤B", "A", "9B", "C"]

    # 夜勤扱い（翌日勤務禁止・均等化対象）。10B も夜勤扱い。
    NIGHT_SHIFTS = ["A", "9B", "10B", "C"]

    def __init__(
        self,
        year: int,
        month: int,
        roster: List[str],
        fixed_worker: str,
        shift_hours: Dict[str, int],
        constraints: Dict[str, int],
        prev_month_tail: Optional[Dict[str, List[str]]] = None,
    ):
        self.year         = year
        self.month        = month
        self.roster       = roster
        self.fixed_worker = fixed_worker
        self.shift_hours  = shift_hours
        self.constraints  = constraints
        self.prev_month_tail = prev_month_tail or {}

        _, self.num_days = calendar.monthrange(year, month)
        self.days = list(range(1, self.num_days + 1))

        self.all_shifts     = list(shift_hours.keys())
        self.work_shifts    = [s for s in self.all_shifts if shift_hours[s] > 0]
        self.holiday_shifts = [s for s in self.all_shifts if shift_hours[s] == 0]

        self.shift_workers = [w for w in self.roster if w != self.fixed_worker]

        # 月間上限時間（設定値を優先、なければ月日数から自動決定）
        default_max = {31: 176, 30: 168}.get(self.num_days, 160)
        self.max_hours = constraints.get("月間上限時間", default_max)

    def _is_weekend(self, day: int) -> bool:
        return calendar.weekday(self.year, self.month, day) in (
            calendar.SATURDAY, calendar.SUNDAY
        )

    def validate_inputs(
        self,
        holiday_requests: Dict[Tuple[str, int], bool],
        fixed_assignments: Dict[Tuple[str, int], str],
    ):
        for (w, d) in fixed_assignments:
            if w not in self.roster:
                raise ShiftValidationError(
                    f"固定シートにある従業員「{w}」は従業員一覧に存在しません。"
                )
        for (w, d) in holiday_requests:
            if w not in self.roster:
                raise ShiftValidationError(
                    f"希望休シートにある従業員「{w}」は従業員一覧に存在しません。"
                )
        for (w, d) in list(fixed_assignments) + list(holiday_requests):
            if d < 1 or d > self.num_days:
                raise ShiftValidationError(
                    f"日付「{d}日」が{self.month}月の範囲を超えています。"
                )
        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday and (w, d) in fixed_assignments:
                fs = fixed_assignments[(w, d)]
                if self.shift_hours.get(fs, 0) > 0:
                    raise ShiftValidationError(
                        f"{w}さんは{d}日に「希望休」ですが、"
                        f"固定シートで「{fs}」が指定され矛盾しています。"
                    )
        for d in self.days:
            if (self.fixed_worker, d) in fixed_assignments:
                s = fixed_assignments[(self.fixed_worker, d)]
                if self._is_weekend(d) and self.shift_hours.get(s, 0) > 0:
                    raise ShiftValidationError(
                        f"固定ワーカー({self.fixed_worker})の土日({d}日)に"
                        f"勤務シフト「{s}」が固定されています。"
                    )
                if not self._is_weekend(d) and s in self.DAILY_REQUIRED_SHIFTS:
                    raise ShiftValidationError(
                        f"固定ワーカー({self.fixed_worker})の平日({d}日)に"
                        f"一般シフト「{s}」が固定されています。隊長日勤にしてください。"
                    )
        for d in self.days[:-1]:
            for w in self.roster:
                if (w, d) in fixed_assignments and (w, d + 1) in fixed_assignments:
                    s1 = fixed_assignments[(w, d)]
                    s2 = fixed_assignments[(w, d + 1)]
                    if s1 in self.NIGHT_SHIFTS and s2 in ["日勤A", "日勤B", "隊長日勤"]:
                        raise ShiftValidationError(
                            f"{w}さんは{d}日に夜勤、翌{d+1}日に日勤が固定されており、"
                            f"夜勤明けルールに違反します。"
                        )

    def solve(
        self,
        holiday_requests: Dict[Tuple[str, int], bool],
        fixed_assignments: Dict[Tuple[str, int], str],
    ) -> Tuple[str, Dict[Tuple[str, int], str]]:

        self.validate_inputs(holiday_requests, fixed_assignments)

        model = cp_model.CpModel()

        # ── 変数定義 ──────────────────────────────────────────────────────────
        x = {
            (w, d, s): model.NewBoolVar(f"x_{w}_{d}_{s}")
            for w in self.roster
            for d in self.days
            for s in self.all_shifts
        }

        # ── 1日1シフト ────────────────────────────────────────────────────────
        for w in self.roster:
            for d in self.days:
                model.Add(sum(x[w, d, s] for s in self.all_shifts) == 1)

        # ── 固定シフト・希望休 ────────────────────────────────────────────────
        for (w, d), s in fixed_assignments.items():
            if (w, d, s) in x:
                model.Add(x[w, d, s] == 1)

        for (w, d), is_holiday in holiday_requests.items():
            if is_holiday and "○" in self.all_shifts:
                model.Add(x[w, d, "○"] == 1)

        # ── 固定ワーカー制約 ───────────────────────────────────────────────────
        for d in self.days:
            if (self.fixed_worker, d) not in fixed_assignments:
                if self._is_weekend(d):
                    model.Add(x[self.fixed_worker, d, "○"] == 1)
                else:
                    model.Add(x[self.fixed_worker, d, "隊長日勤"] == 1)

        for w in self.shift_workers:
            for d in self.days:
                model.Add(x[w, d, "隊長日勤"] == 0)

        # ── 毎日の必要人数（10B は含めない） ────────────────────────────────
        for d in self.days:
            for s in self.DAILY_REQUIRED_SHIFTS:
                model.Add(sum(x[w, d, s] for w in self.shift_workers) == 1)

        # ── 夜勤明け翌日勤務禁止 ──────────────────────────────────────────────
        for w in self.shift_workers:
            for d in self.days[:-1]:
                model.Add(
                    sum(x[w, d, s] for s in self.NIGHT_SHIFTS)
                    + sum(x[w, d + 1, s] for s in self.work_shifts)
                    <= 1
                )

        # 前月末が夜勤なら1日目は休み
        for w in self.shift_workers:
            tail = self.prev_month_tail.get(w, [])
            if tail and tail[-1] in self.NIGHT_SHIFTS:
                model.Add(sum(x[w, 1, s] for s in self.work_shifts) == 0)

        # ── 連続勤務制限 ─────────────────────────────────────────────────────
        max_consec = self.constraints.get("連勤制限", 5)
        for w in self.shift_workers:
            for d in range(1, self.num_days - max_consec + 1):
                model.Add(
                    sum(
                        x[w, d + i, s]
                        for i in range(max_consec + 1)
                        for s in self.work_shifts
                    ) <= max_consec
                )

        # ── 週1休保証 ────────────────────────────────────────────────────────
        for w in self.shift_workers:
            for d in range(1, self.num_days - 6 + 1):
                model.Add(
                    sum(
                        x[w, d + i, s]
                        for i in range(7)
                        for s in self.holiday_shifts
                    ) >= 1
                )

        # ── 月間労働時間：上限以内（<= max_hours） ───────────────────────────
        for w in self.shift_workers:
            total_hours = sum(
                x[w, d, s] * self.shift_hours[s]
                for d in self.days
                for s in self.all_shifts
            )
            model.Add(total_hours <= self.max_hours)

        # ── 目的関数：夜勤回数の最大-最小差を最小化 ─────────────────────────
        night_counts = []
        for w in self.shift_workers:
            cnt = model.NewIntVar(0, self.num_days, f"nights_{w}")
            model.Add(
                cnt == sum(x[w, d, s] for d in self.days for s in self.NIGHT_SHIFTS)
            )
            night_counts.append(cnt)

        min_night = model.NewIntVar(0, self.num_days, "min_night")
        max_night = model.NewIntVar(0, self.num_days, "max_night")
        model.AddMinEquality(min_night, night_counts)
        model.AddMaxEquality(max_night, night_counts)
        model.Minimize(max_night - min_night)

        # ── ソルバー実行 ─────────────────────────────────────────────────────
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 60.0
        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            status_str = "OPTIMAL" if status == cp_model.OPTIMAL else "FEASIBLE"
            result = {}
            for w in self.roster:
                for d in self.days:
                    assigned = "○"
                    for s in self.all_shifts:
                        if solver.Value(x[w, d, s]) == 1:
                            assigned = s
                            break
                    result[(w, d)] = assigned
            return status_str, result
        else:
            return "INFEASIBLE", {}
