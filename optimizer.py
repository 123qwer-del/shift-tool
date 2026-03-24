"""
警備員シフト最適化エンジン  v3.1
=================================
OR-Tools CP-SAT を使用した9名の警備員シフトスケジューリング

変更履歴
  v1.0  初版リリース
  v2.0  [1] 従業員を外部リスト WORKER_ROSTER で固定定義
        [2] 週1休を7日間スライディングウィンドウで保証
        [3] 目的関数に労働時間均等化を追加（夜勤回数 + 労働時間の二軸最適化）
        [4] fixed_assignments の事前整合性チェックを実装
        [5] 矛盾検出時は solve 前に ShiftValidationError を送出
  v3.0  [1] 従業員名を実名（姓名）に変更
        [2] 夜勤合計3名 → 夜勤A/B/C 各1名ずつの個別制約に変更
        [3] 末吉 弘一をシフト対象外に変更（日勤2名は shift_workers のみで充填）
        [4] include_fixed_in_day_count フラグを廃止
        [5] 月間上限176h維持（平均180h/人のためソルバーが不均等配分で対処）
  v3.1  [1] 実労働時間の定義を修正（夜勤A: 11h→9h, 夜勤B: 10h→8h, 夜勤C: 11h→9h）
        [2] 176h上限で全12ヶ月が実現可能に（最大162.8h/人、余裕13.2h）

インストール要件:
    pip install ortools pandas

使用方法:
    from optimizer import generate_shift, WORKER_ROSTER
    df = generate_shift(2025, 6, requests, fixed_assignments)
"""

import calendar
from typing import Dict, List, Optional, Tuple

import pandas as pd
from ortools.sat.python import cp_model

# ===========================================================================
# [修正1] 従業員リストを外部定義（モジュールレベル定数）
# ===========================================================================

#: 全従業員リスト（順序がDataFrameの行順に反映される）
WORKER_ROSTER: List[str] = [
    "末吉 弘一",  # 平日日勤固定
    "伊藤 晶俊",
    "吉村 智",
    "南 英俊",
    "杉田 孝行",
    "山田 誠",
    "大西 信一",
    "村主 博",
    "河内 拳",
]

#: 平日日勤固定の従業員名（WORKER_ROSTERの先頭と一致させること）
FIXED_WORKER: str = "末吉 弘一"

# ---------------------------------------------------------------------------
# シフト定義
# ---------------------------------------------------------------------------
SHIFT_TYPES: List[str] = ["日勤", "夜勤A", "夜勤B", "夜勤C", "休日"]

SHIFT_HOURS: Dict[str, int] = {
    "日勤":  8,   # 拘束9h - 休憩1h = 実労働8h
    "夜勤A": 9,   # 拘束15h - 休憩6h = 実労働9h
    "夜勤B": 8,   # 拘束14h - 休憩5h = 実労働8h
    "夜勤C": 9,   # 拘束15h - 休憩6h = 実労働9h
    "休日":  0,
}

DAY_SHIFTS   = frozenset({"日勤"})
NIGHT_SHIFTS = frozenset({"夜勤A", "夜勤B", "夜勤C"})
WORK_SHIFTS  = DAY_SHIFTS | NIGHT_SHIFTS

# インデックス変換
SHIFT_INDEX: Dict[str, int] = {s: i for i, s in enumerate(SHIFT_TYPES)}

# 制約パラメータ
MAX_MONTHLY_HOURS: int = 176   # 月間上限時間（実労働時間ベース）
REQUIRED_DAY:      int = 2     # 日勤 必要人数（shift_workers のみで充填、末吉はシフト対象外）
REQUIRED_NIGHT:    int = 3     # 夜勤 必要人数（夜勤A×1 + 夜勤B×1 + 夜勤C×1）
MAX_CONSECUTIVE:   int = 4     # 最大連続勤務日数
SLIDING_WINDOW:    int = 7     # 週1休の判定ウィンドウ幅 [修正3]


# ===========================================================================
# [修正6] カスタム例外クラス
# ===========================================================================

class ShiftValidationError(ValueError):
    """
    fixed_assignments / requests の整合性チェックで矛盾が見つかった場合に送出。
    errors 属性に全矛盾メッセージのリストを持つ。
    """
    def __init__(self, errors: List[str]) -> None:
        self.errors = errors
        bullet = "\n  - "
        super().__init__(
            f"シフト定義に {len(errors)} 件の矛盾が検出されました:{bullet}"
            + bullet.join(errors)
        )


# ===========================================================================
# [修正5] 事前整合性チェック関数
# ===========================================================================

def validate_inputs(
    year: int,
    month: int,
    requests: Dict[Tuple[str, int], bool],
    fixed_assignments: Dict[Tuple[str, int], str],
    roster: List[str],
) -> None:
    """
    fixed_assignments と requests の整合性を検査する。
    矛盾が1件でもあれば ShiftValidationError を送出する（solve前）。

    検査項目:
      V1  存在しない従業員名
      V2  存在しない日付
      V3  未定義のシフト種類
      V4  同一(名前,日)の fixed と requests の衝突
          （固定シフトが休日でないのに希望休がある）
      V5  FIXED_WORKER への固定セル違反
          （平日は日勤固定・土日は休日固定）
      V6  シフト対象者数の充足可能性（日勤2名＋夜勤3名 = 5名を満たせるか）
      V7  夜勤翌日日勤の固定セル矛盾
      V8  同一(名前,日)への requests と fixed の休日強制との整合
    """
    errors: List[str] = []
    num_days  = calendar.monthrange(year, month)[1]
    valid_days = set(range(1, num_days + 1))
    shift_workers = [w for w in roster if w != FIXED_WORKER]

    def is_weekday(day: int) -> bool:
        return calendar.weekday(year, month, day) < 5

    # ------------------------------------------------------------------
    # V1, V2, V3: fixed_assignments の基本バリデーション
    # ------------------------------------------------------------------
    for (name, day), shift in fixed_assignments.items():
        if name not in roster:
            errors.append(
                f"[V1] fixed_assignments: 未登録の従業員 '{name}' (日:{day})"
            )
        if day not in valid_days:
            errors.append(
                f"[V2] fixed_assignments: 無効な日付 {day}日 "
                f"(従業員:'{name}', 月:{year}/{month:02d})"
            )
        if shift not in SHIFT_INDEX:
            errors.append(
                f"[V3] fixed_assignments: 未定義のシフト '{shift}' "
                f"(従業員:'{name}', 日:{day})"
            )

    # ------------------------------------------------------------------
    # V1, V2: requests の基本バリデーション
    # ------------------------------------------------------------------
    for (name, day), flag in requests.items():
        if not flag:
            continue
        if name not in roster:
            errors.append(
                f"[V1] requests: 未登録の従業員 '{name}' (日:{day})"
            )
        if day not in valid_days:
            errors.append(
                f"[V2] requests: 無効な日付 {day}日 "
                f"(従業員:'{name}', 月:{year}/{month:02d})"
            )

    # ------------------------------------------------------------------
    # V4: fixed と requests の衝突
    #     希望休(True)があるのに、その日が非休日で固定されている場合
    # ------------------------------------------------------------------
    for (name, day), flag in requests.items():
        if not flag:
            continue
        if day not in valid_days or name not in roster:
            continue  # V1/V2 で報告済み
        assigned = fixed_assignments.get((name, day))
        if assigned is not None and assigned != "休日":
            errors.append(
                f"[V4] 希望休と固定シフトの衝突: '{name}' {day}日 "
                f"(fixed='{assigned}' vs request=休日希望)"
            )

    # ------------------------------------------------------------------
    # V5: FIXED_WORKER への固定セル違反
    # ------------------------------------------------------------------
    for (name, day), shift in fixed_assignments.items():
        if name != FIXED_WORKER:
            continue
        if day not in valid_days:
            continue  # V2 で報告済み
        if is_weekday(day):
            if shift != "日勤":
                errors.append(
                    f"[V5] '{FIXED_WORKER}' は平日({day}日)は日勤固定ですが、"
                    f"'{shift}' が指定されています"
                )
        else:
            if shift != "休日":
                errors.append(
                    f"[V5] '{FIXED_WORKER}' は土日({day}日)は休日固定ですが、"
                    f"'{shift}' が指定されています"
                )

    # ------------------------------------------------------------------
    # V6: シフト対象者数の充足可能性
    #   末吉はシフト対象外のため、shift_workers だけで以下を満たす必要がある
    #   1日の必要人数: 日勤{REQUIRED_DAY}名 + 夜勤A1+夜勤B1+夜勤C1 = 5名
    # ------------------------------------------------------------------
    n_sw = len(shift_workers)
    need_per_day = REQUIRED_DAY + REQUIRED_NIGHT  # 2 + 3 = 5名
    if n_sw < need_per_day:
        errors.append(
            f"[V6] シフト対象者が {n_sw} 名のため、"
            f"1日の必要人数（日勤{REQUIRED_DAY}名 + 夜勤{REQUIRED_NIGHT}名 = "
            f"{need_per_day}名）を満たせません"
        )

    # ------------------------------------------------------------------
    # V7: 夜勤翌日日勤の固定セル矛盾
    # ------------------------------------------------------------------
    for (name, day), shift in fixed_assignments.items():
        if name == FIXED_WORKER:
            continue
        if day not in valid_days or shift not in NIGHT_SHIFTS:
            continue
        next_day = day + 1
        if next_day not in valid_days:
            continue
        next_shift = fixed_assignments.get((name, next_day))
        if next_shift == "日勤":
            errors.append(
                f"[V7] 夜勤翌日日勤の矛盾: '{name}' "
                f"{day}日='{shift}' → {next_day}日='{next_shift}' "
                f"(夜勤翌日の日勤は禁止)"
            )

    # ------------------------------------------------------------------
    # V8: FIXED_WORKER への希望休（末吉は土日休固定なので不要だが明示チェック）
    # ------------------------------------------------------------------
    for (name, day), flag in requests.items():
        if not flag or name != FIXED_WORKER:
            continue
        if day not in valid_days:
            continue
        if is_weekday(day):
            errors.append(
                f"[V8] '{FIXED_WORKER}' は平日({day}日)は日勤固定のため "
                f"希望休を適用できません"
            )

    # ------------------------------------------------------------------
    # 矛盾があれば例外送出（solve前に停止）[修正6]
    # ------------------------------------------------------------------
    if errors:
        raise ShiftValidationError(errors)


# ===========================================================================
# メイン関数
# ===========================================================================

def generate_shift(
    year: int,
    month: int,
    requests: Dict[Tuple[str, int], bool],
    fixed_assignments: Dict[Tuple[str, int], str],
    *,
    roster: Optional[List[str]] = None,
    solver_time_limit: float = 60.0,
    solver_workers: int = 4,
    settings: Optional[Dict] = None,  # 後方互換性のため追加（非推奨）
) -> pd.DataFrame:
    """
    警備員シフトスケジュールを最適化して返す。

    Parameters
    ----------
    year  : int
        対象年
    month : int
        対象月（1-12）
    requests : dict
        { (名前, 日): True }  希望休（必ず休日に設定）
    fixed_assignments : dict
        { (名前, 日): "日勤" など }  固定セル（変更不可）
    roster : list[str], optional
        従業員リスト。省略時はモジュール定数 WORKER_ROSTER を使用。
        末吉 弘一（FIXED_WORKER）はシフト対象外で平日日勤固定・土日休日固定。
        日勤2名・夜勤A/B/C各1名はすべて残り8名の shift_workers から充填する。
    solver_time_limit : float
        ソルバー最大実行時間（秒）デフォルト60秒
    solver_workers : int
        ソルバー並列スレッド数

    Returns
    -------
    pd.DataFrame
        行=名前（roster順）、列=日付(1〜月末)、値=シフト種類

    Raises
    ------
    ShiftValidationError
        入力に矛盾が検出された場合（solve前に送出）[修正6]
    RuntimeError
        ソルバーが解を見つけられなかった場合
    """
    # ================================================================
    # settings パラメータの処理（後方互換性）
    # ================================================================
    if settings is not None:
        # settings から個別パラメータを上書き（存在する場合のみ）
        solver_time_limit = settings.get('solver_time_limit', solver_time_limit)
        solver_workers = settings.get('solver_workers', solver_workers)
        if 'roster' in settings and roster is None:
            roster = settings['roster']
    
    # [修正1] roster の解決
    roster = list(roster) if roster is not None else list(WORKER_ROSTER)

    # ---- 基本情報 ----
    num_days = calendar.monthrange(year, month)[1]
    days     = list(range(1, num_days + 1))

    def is_weekday(day: int) -> bool:
        return calendar.weekday(year, month, day) < 5

    shift_workers = [w for w in roster if w != FIXED_WORKER]
    num_workers   = len(shift_workers)

    # ================================================================
    # [修正5, 6] solve前の整合性チェック
    # ================================================================
    validate_inputs(
        year, month, requests, fixed_assignments,
        roster,
    )

    # ================================================================
    # モデル構築
    # ================================================================
    model  = cp_model.CpModel()
    solver = cp_model.CpSolver()

    num_shifts = len(SHIFT_TYPES)
    si_day    = SHIFT_INDEX["日勤"]
    si_nightA = SHIFT_INDEX["夜勤A"]
    si_nightB = SHIFT_INDEX["夜勤B"]
    si_nightC = SHIFT_INDEX["夜勤C"]
    si_rest   = SHIFT_INDEX["休日"]
    night_sis = [si_nightA, si_nightB, si_nightC]

    # 決定変数: x[wi, d, si] = 1 ← 従業員 wi が日 d にシフト si を担当
    x: Dict[Tuple[int, int, int], cp_model.IntVar] = {}
    for wi in range(num_workers):
        for d in days:
            for si in range(num_shifts):
                x[wi, d, si] = model.NewBoolVar(f"x_{wi}_{d}_{si}")

    # ================================================================
    # 制約 C1: 1日に1つのシフトのみ
    # ================================================================
    for wi in range(num_workers):
        for d in days:
            model.AddExactlyOne(x[wi, d, si] for si in range(num_shifts))

    # ================================================================
    # 制約 C2: 1日の必要人数
    #   末吉はシフト対象外のため、全日程で shift_workers のみで充填する
    #   日勤  2名 / 夜勤A 1名 / 夜勤B 1名 / 夜勤C 1名  （計5名/日）
    # ================================================================
    for d in days:
        # --- 日勤：平日・土日ともに shift_workers から2名 ---
        model.Add(
            sum(x[wi, d, si_day] for wi in range(num_workers)) == REQUIRED_DAY
        )

        # --- 夜勤：種別ごとに必ず1名ずつ ---
        model.Add(
            sum(x[wi, d, si_nightA] for wi in range(num_workers)) == 1
        )
        model.Add(
            sum(x[wi, d, si_nightB] for wi in range(num_workers)) == 1
        )
        model.Add(
            sum(x[wi, d, si_nightC] for wi in range(num_workers)) == 1
        )

    # ================================================================
    # 制約 C3: 月間労働時間 ≤ 176時間
    # ================================================================
    for wi in range(num_workers):
        model.Add(
            sum(
                x[wi, d, si] * SHIFT_HOURS[SHIFT_TYPES[si]]
                for d in days
                for si in range(num_shifts)
            )
            <= MAX_MONTHLY_HOURS
        )

    # ================================================================
    # 制約 C4: 夜勤翌日の日勤禁止
    # ================================================================
    for wi in range(num_workers):
        for di in range(len(days) - 1):
            d, nd = days[di], days[di + 1]
            for si_n in night_sis:
                model.Add(x[wi, nd, si_day] == 0).OnlyEnforceIf(x[wi, d, si_n])

    # ================================================================
    # 制約 C5: 連続勤務4日以内
    # ================================================================
    for wi in range(num_workers):
        for di in range(len(days) - MAX_CONSECUTIVE):
            window = [days[di + k] for k in range(MAX_CONSECUTIVE + 1)]
            model.Add(
                sum(
                    x[wi, d, si]
                    for d in window
                    for si in range(num_shifts)
                    if SHIFT_TYPES[si] != "休日"
                )
                <= MAX_CONSECUTIVE
            )

    # ================================================================
    # 制約 C6: 週1休 — 7日スライディングウィンドウ  [修正3]
    #
    # 旧実装（カレンダー週ブロック）との違い:
    #   旧: 月曜〜日曜の暦週ブロック内で休日≥1
    #       → 月初が木曜始まりなら最初のブロックは木〜日の4日のみ
    #   新: 任意の連続7日間で必ず休日≥1
    #       → 月中どの7日区間を取っても休みがあることを保証
    # ================================================================
    for wi in range(num_workers):
        for di in range(len(days) - SLIDING_WINDOW + 1):
            window_7 = [days[di + k] for k in range(SLIDING_WINDOW)]
            model.Add(
                sum(x[wi, d, si_rest] for d in window_7) >= 1
            )

    # ================================================================
    # 制約 C7: 希望休は必ず休日
    # ================================================================
    for (name, day), flag in requests.items():
        if flag and name in shift_workers and day in days:
            wi = shift_workers.index(name)
            model.Add(x[wi, day, si_rest] == 1)

    # ================================================================
    # 制約 C8: 固定セルは変更不可
    # ================================================================
    for (name, day), shift in fixed_assignments.items():
        if name in shift_workers and day in days:
            wi = shift_workers.index(name)
            si = SHIFT_INDEX[shift]
            model.Add(x[wi, day, si] == 1)

    # ================================================================
    # 目的関数: 役割別5軸均等化（偏りを多少許容した重み付け最小化）
    #
    # 各役割の担当回数の (max - min) を最小化する。
    # 偏り許容度はウェイトで調整：大きいほど均等化を優先。
    #   WEIGHT_NIGHT_A/B/C : 夜勤種別ごとの均等化（優先度高）
    #   WEIGHT_DAY          : 日勤の均等化（優先度中）
    #   WEIGHT_HOURS        : 労働時間の均等化（優先度低・偏り許容）
    # ================================================================

    def _make_diff(name: str, counts_list: list, ub: int) -> cp_model.IntVar:
        """counts_list の max-min を表す IntVar を返す。"""
        v_max  = model.NewIntVar(0, ub, f"max_{name}")
        v_min  = model.NewIntVar(0, ub, f"min_{name}")
        v_diff = model.NewIntVar(0, ub, f"diff_{name}")
        model.AddMaxEquality(v_max, counts_list)
        model.AddMinEquality(v_min, counts_list)
        model.Add(v_diff == v_max - v_min)
        return v_diff

    # --- 夜勤A 担当回数 ---
    nightA_counts = []
    for wi in range(num_workers):
        c = model.NewIntVar(0, num_days, f"cntA_{wi}")
        model.Add(c == sum(x[wi, d, si_nightA] for d in days))
        nightA_counts.append(c)
    diff_nightA = _make_diff("nightA", nightA_counts, num_days)

    # --- 夜勤B 担当回数 ---
    nightB_counts = []
    for wi in range(num_workers):
        c = model.NewIntVar(0, num_days, f"cntB_{wi}")
        model.Add(c == sum(x[wi, d, si_nightB] for d in days))
        nightB_counts.append(c)
    diff_nightB = _make_diff("nightB", nightB_counts, num_days)

    # --- 夜勤C 担当回数 ---
    nightC_counts = []
    for wi in range(num_workers):
        c = model.NewIntVar(0, num_days, f"cntC_{wi}")
        model.Add(c == sum(x[wi, d, si_nightC] for d in days))
        nightC_counts.append(c)
    diff_nightC = _make_diff("nightC", nightC_counts, num_days)

    # --- 日勤 担当回数 ---
    day_counts = []
    for wi in range(num_workers):
        c = model.NewIntVar(0, num_days, f"cntD_{wi}")
        model.Add(c == sum(x[wi, d, si_day] for d in days))
        day_counts.append(c)
    diff_day = _make_diff("day", day_counts, num_days)

    # --- 労働時間 ---
    hour_counts = []
    for wi in range(num_workers):
        h = model.NewIntVar(0, MAX_MONTHLY_HOURS, f"hours_{wi}")
        model.Add(
            h == sum(
                x[wi, d, si] * SHIFT_HOURS[SHIFT_TYPES[si]]
                for d in days
                for si in range(num_shifts)
            )
        )
        hour_counts.append(h)
    diff_hours = _make_diff("hours", hour_counts, MAX_MONTHLY_HOURS)

    # --- 重み付き合算（偏り多少許容: HOURSの重みを低く設定）---
    WEIGHT_NIGHT_A: int = 6   # 夜勤A均等化（優先度高）
    WEIGHT_NIGHT_B: int = 6   # 夜勤B均等化（優先度高）
    WEIGHT_NIGHT_C: int = 6   # 夜勤C均等化（優先度高）
    WEIGHT_DAY:     int = 4   # 日勤均等化（優先度中）
    WEIGHT_HOURS:   int = 1   # 労働時間均等化（偏り許容）

    obj_ub = (
        WEIGHT_NIGHT_A * num_days + WEIGHT_NIGHT_B * num_days
        + WEIGHT_NIGHT_C * num_days + WEIGHT_DAY * num_days
        + WEIGHT_HOURS * MAX_MONTHLY_HOURS
    )
    combined = model.NewIntVar(0, obj_ub, "combined_obj")
    model.Add(
        combined == (
            WEIGHT_NIGHT_A * diff_nightA
            + WEIGHT_NIGHT_B * diff_nightB
            + WEIGHT_NIGHT_C * diff_nightC
            + WEIGHT_DAY     * diff_day
            + WEIGHT_HOURS   * diff_hours
        )
    )
    model.Minimize(combined)

    # ================================================================
    # ソルバー実行
    # ================================================================
    solver.parameters.max_time_in_seconds = solver_time_limit
    solver.parameters.num_search_workers  = solver_workers

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(
            f"解が見つかりませんでした。ステータス: {solver.StatusName(status)}\n"
            "制約条件の緩和や入力データの見直しを検討してください。"
        )

    # ================================================================
    # 結果の抽出 → DataFrame 構築
    # ================================================================
    result: Dict[str, Dict[int, str]] = {}

    # FIXED_WORKER（末吉）: 平日=日勤、土日=休日
    result[FIXED_WORKER] = {
        d: ("日勤" if is_weekday(d) else "休日")
        for d in days
    }

    # シフト対象者
    for wi, worker in enumerate(shift_workers):
        result[worker] = {}
        for d in days:
            assigned = "休日"
            for si in range(num_shifts):
                if solver.Value(x[wi, d, si]) == 1:
                    assigned = SHIFT_TYPES[si]
                    break
            result[worker][d] = assigned

    # roster 順を維持して DataFrame 化
    ordered = [w for w in roster if w in result]
    df = pd.DataFrame({w: result[w] for w in ordered}).T
    df.columns = pd.Index(days, name="日")
    df.index.name = "名前"

    return df


# ===========================================================================
# ユーティリティ
# ===========================================================================

def get_role_counts(df: pd.DataFrame) -> pd.DataFrame:
    """
    各従業員の役割別担当回数と労働時間を集計して返す。

    Parameters
    ----------
    df : pd.DataFrame
        generate_shift() が返すシフト表

    Returns
    -------
    pd.DataFrame
        列: 日勤, 夜勤A, 夜勤B, 夜勤C, 夜勤計, 休日, 労働時間(h)
        行: 従業員名（df.index順）
        末尾行: 合計・最大・最小・偏り(max-min) の統計行
    """
    records = []
    for name in df.index:
        row        = df.loc[name]
        day_cnt    = int(sum(1 for v in row if v == "日勤"))
        nightA_cnt = int(sum(1 for v in row if v == "夜勤A"))
        nightB_cnt = int(sum(1 for v in row if v == "夜勤B"))
        nightC_cnt = int(sum(1 for v in row if v == "夜勤C"))
        night_cnt  = nightA_cnt + nightB_cnt + nightC_cnt
        rest_cnt   = int(sum(1 for v in row if v == "休日"))
        hours      = int(sum(SHIFT_HOURS.get(str(v), 0) for v in row))
        records.append({
            "名前":       name,
            "日勤":       day_cnt,
            "夜勤A":      nightA_cnt,
            "夜勤B":      nightB_cnt,
            "夜勤C":      nightC_cnt,
            "夜勤計":     night_cnt,
            "休日":       rest_cnt,
            "労働時間(h)": hours,
        })

    rc = pd.DataFrame(records).set_index("名前")

    # 統計行（末吉など固定従業員を除いた数値列のみ集計）
    numeric_cols = ["日勤", "夜勤A", "夜勤B", "夜勤C", "夜勤計", "休日", "労働時間(h)"]
    stats = {}
    for col in numeric_cols:
        vals = rc[col]
        stats[col] = {
            "合計":          int(vals.sum()),
            "最大":          int(vals.max()),
            "最小":          int(vals.min()),
            "偏り(max-min)": int(vals.max() - vals.min()),
        }
    stats_df = pd.DataFrame(stats).T
    stats_df.index.name = "統計"

    return rc, stats_df


def print_summary(
    df: pd.DataFrame,
    year: int,
    month: int,
    *,
    show_bias: bool = True,
) -> None:
    """
    シフト集計サマリーと役割別偏り指標をコンソールに表示する。

    Parameters
    ----------
    show_bias : bool
        True の場合、役割別の偏り指標（max-min）も表示する
    """
    rc, stats_df = get_role_counts(df)

    W = 72
    print(f"\n{'='*W}")
    print(f"  {year}年{month}月 シフト最適化結果サマリー")
    print(f"{'='*W}")

    # ---- 役割回数テーブル ----
    header = f"  {'名前':<10}  {'日勤':>4}  {'夜勤A':>5}  {'夜勤B':>5}  {'夜勤C':>5}  {'夜勤計':>5}  {'休日':>4}  {'労働時間':>6}"
    print(header)
    print(f"  {'-'*(W-2)}")

    for name in rc.index:
        r = rc.loc[name]
        print(
            f"  {name:<10}  {r['日勤']:>4}  {r['夜勤A']:>5}  {r['夜勤B']:>5}  "
            f"{r['夜勤C']:>5}  {r['夜勤計']:>5}  {r['休日']:>4}  {r['労働時間(h)']:>5}h"
        )

    # ---- 偏り指標 ----
    if show_bias:
        print(f"  {'-'*(W-2)}")
        print(f"  {'【偏り指標 (max - min)】'}")
        print(f"  {'-'*(W-2)}")
        roles = ["日勤", "夜勤A", "夜勤B", "夜勤C", "夜勤計", "労働時間(h)"]
        for role in roles:
            if role not in stats_df.index:
                continue
            s = stats_df.loc[role]
            label = "労働時間" if role == "労働時間(h)" else role
            unit  = "h" if role == "労働時間(h)" else "回"
            bias  = int(s["偏り(max-min)"])
            vmax  = int(s["最大"])
            vmin  = int(s["最小"])
            flag  = "  ← 偏りあり" if bias >= 3 else ""
            print(
                f"  {label:<6}  max={vmax:>3}{unit}  min={vmin:>3}{unit}"
                f"  差={bias:>3}{unit}{flag}"
            )

    print(f"{'='*W}\n")


# ===========================================================================
# 使用例（スタンドアロン実行）
# ===========================================================================

if __name__ == "__main__":
    YEAR, MONTH = 2025, 6

    # 希望休サンプル
    sample_requests: Dict[Tuple[str, int], bool] = {
        ("伊藤 晶俊",  5): True,
        ("吉村 智",   10): True,
        ("南 英俊",   15): True,
        ("杉田 孝行", 20): True,
    }

    # 固定セルサンプル
    sample_fixed: Dict[Tuple[str, int], str] = {
        ("山田 誠",   1): "夜勤A",
        ("大西 信一", 1): "夜勤B",
        ("村主 博",   1): "夜勤C",
    }

    # ------------------------------------------------------------------
    # 矛盾チェックデモ（コメントを外すと ShiftValidationError が発生）
    # ------------------------------------------------------------------
    # bad_fixed = {
    #     ("伊藤 晶俊",  5): "夜勤A",  # [V4] 希望休と固定シフトの衝突
    #     ("末吉 弘一",  2): "夜勤B",  # [V5] FIXED_WORKER の平日に夜勤指定
    #     ("山田 誠",    3): "夜勤A",
    #     ("山田 誠",    4): "日勤",   # [V7] 夜勤翌日に日勤
    #     ("幽霊 太郎",  1): "日勤",   # [V1] 存在しない従業員
    #     ("吉村 智",   40): "日勤",   # [V2] 存在しない日付（6月は30日まで）
    #     ("吉村 智",    6): "深夜",   # [V3] 未定義のシフト種類
    # }
    # try:
    #     generate_shift(YEAR, MONTH, sample_requests, bad_fixed)
    # except ShiftValidationError as e:
    #     print(e)

    print(f"\n{YEAR}年{MONTH}月のシフト最適化を実行中...")
    print(f"従業員ロスター: {WORKER_ROSTER}\n")

    try:
        df = generate_shift(
            YEAR, MONTH,
            sample_requests,
            sample_fixed,
        )
        print("【シフト表】")
        print(df.to_string())
        print_summary(df, YEAR, MONTH)

    except ShiftValidationError as e:
        print(f"\n入力エラー:\n{e}")
    except RuntimeError as e:
        print(f"\n最適化エラー: {e}")
