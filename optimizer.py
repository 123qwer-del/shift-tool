"""
警備員シフト最適化エンジン  v4.0
=================================
OR-Tools CP-SAT を使用した警備員シフトスケジューリング

変更履歴
  v1.0  初版リリース
  v2.0  従業員外部リスト・週1休スライディングウィンドウ・目的関数二軸化
  v3.0  従業員を実名化・夜勤A/B/C個別制約・FIXED_WORKERシフト対象外化
  v4.0  [設定外部化] ハードコード値をすべて Settings オブジェクト経由に変更
        WORKER_ROSTER / FIXED_WORKER / SHIFT_HOURS / 制約パラメータが
        settings.py の Settings クラスから動的に供給される。
        後方互換: settings 省略時はデフォルト値を使用。

インストール要件:
    pip install ortools pandas openpyxl

使用方法:
    from settings import Settings
    from optimizer import generate_shift
    s = Settings.load(Path("input.xlsx"))
    df = generate_shift(2025, 6, requests, fixed_assignments, settings=s)
"""

import calendar
from typing import Dict, List, Optional, Tuple

import pandas as pd
from ortools.sat.python import cp_model

import unicodedata
import difflib

from settings import Settings


# ===========================================================================
# 後方互換用: デフォルト設定インスタンス
# ===========================================================================

_default_settings = Settings()

# 後方互換のため WORKER_ROSTER / FIXED_WORKER をモジュールレベルで参照可能に残す
WORKER_ROSTER: List[str] = _default_settings.roster
FIXED_WORKER:  str       = _default_settings.fixed_worker


# ===========================================================================
# 名前の正規化・自動補正
# ===========================================================================

def normalize_name(name: str) -> str:
    """全角→半角、空白除去、小文字化。"""
    name = unicodedata.normalize("NFKC", str(name))
    name = name.replace(" ", "").replace("\u3000", "")
    return name.strip()


def auto_correct_name(input_name: str, roster: Optional[List[str]] = None) -> str:
    """
    roster 内の名前に最も近いものを返す。
    完全一致→正規化一致→あいまい一致（cutoff=0.6）の順で探す。
    """
    roster = roster or WORKER_ROSTER
    normalized = {normalize_name(n): n for n in roster}
    key = normalize_name(input_name)

    if key in normalized:
        return normalized[key]

    matches = difflib.get_close_matches(key, normalized.keys(), n=1, cutoff=0.6)
    if matches:
        return normalized[matches[0]]

    raise ValueError(f"登録されていない名前です: {input_name}")


# ===========================================================================
# カスタム例外
# ===========================================================================

class ShiftValidationError(ValueError):
    """
    fixed_assignments / requests の整合性チェックで矛盾が検出された場合に送出。
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
# 事前整合性チェック
# ===========================================================================

def validate_inputs(
    year: int,
    month: int,
    requests: Dict[Tuple[str, int], bool],
    fixed_assignments: Dict[Tuple[str, int], str],
    settings: Settings,
) -> None:
    """
    fixed_assignments と requests の整合性を検査する。
    矛盾が1件でもあれば ShiftValidationError を送出する（solve前）。

    検査項目:
      V1  存在しない従業員名
      V2  存在しない日付
      V3  未定義のシフト種類
      V4  同一(名前,日)の fixed と requests の衝突
      V5  FIXED_WORKER への固定セル違反（平日=日勤・土日=休日）
      V6  シフト対象者数の充足可能性
      V7  夜勤翌日日勤の固定セル矛盾
      V8  FIXED_WORKER への希望休（平日）
    """
    errors: List[str] = []
    roster       = settings.roster
    fixed_worker = settings.fixed_worker
    shift_index  = {s: i for i, s in enumerate(settings.shift_types)}
    night_shifts = set(settings.night_shifts)
    req_day      = settings.constraints["日勤必要人数"]
    req_night    = settings.constraints["夜勤必要人数"]

    num_days   = calendar.monthrange(year, month)[1]
    valid_days = set(range(1, num_days + 1))
    shift_workers = [w for w in roster if w != fixed_worker]

    def is_weekday(day: int) -> bool:
        return calendar.weekday(year, month, day) < 5

    # V1, V2, V3: fixed_assignments の基本バリデーション
    for (name, day), shift in fixed_assignments.items():
        if name not in roster:
            errors.append(f"[V1] fixed_assignments: 未登録の従業員 '{name}' (日:{day})")
        if day not in valid_days:
            errors.append(
                f"[V2] fixed_assignments: 無効な日付 {day}日 "
                f"(従業員:'{name}', 月:{year}/{month:02d})"
            )
        if shift not in shift_index:
            errors.append(
                f"[V3] fixed_assignments: 未定義のシフト '{shift}' "
                f"(従業員:'{name}', 日:{day})"
            )

    # V1, V2: requests の基本バリデーション
    for (name, day), flag in requests.items():
        if not flag:
            continue
        if name not in roster:
            errors.append(f"[V1] requests: 未登録の従業員 '{name}' (日:{day})")
        if day not in valid_days:
            errors.append(
                f"[V2] requests: 無効な日付 {day}日 "
                f"(従業員:'{name}', 月:{year}/{month:02d})"
            )

    # V4: fixed と requests の衝突
    for (name, day), flag in requests.items():
        if not flag or day not in valid_days or name not in roster:
            continue
        assigned = fixed_assignments.get((name, day))
        if assigned is not None and assigned != "休日":
            errors.append(
                f"[V4] 希望休と固定シフトの衝突: '{name}' {day}日 "
                f"(fixed='{assigned}' vs request=休日希望)"
            )

    # V5: FIXED_WORKER への固定セル違反
    if fixed_worker:
        for (name, day), shift in fixed_assignments.items():
            if name != fixed_worker or day not in valid_days:
                continue
            if is_weekday(day):
                if shift != "日勤":
                    errors.append(
                        f"[V5] '{fixed_worker}' は平日({day}日)は日勤固定ですが、"
                        f"'{shift}' が指定されています"
                    )
            else:
                if shift != "休日":
                    errors.append(
                        f"[V5] '{fixed_worker}' は土日({day}日)は休日固定ですが、"
                        f"'{shift}' が指定されています"
                    )

    # V6: シフト対象者数の充足可能性
    n_sw = len(shift_workers)
    need_per_day = req_day + req_night
    if n_sw < need_per_day:
        errors.append(
            f"[V6] シフト対象者が {n_sw} 名のため、"
            f"1日の必要人数（日勤{req_day}名 + 夜勤{req_night}名 = "
            f"{need_per_day}名）を満たせません"
        )

    # V7: 夜勤翌日日勤の固定セル矛盾
    for (name, day), shift in fixed_assignments.items():
        if name == fixed_worker or day not in valid_days:
            continue
        if shift not in night_shifts:
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

    # V8: FIXED_WORKER への希望休（平日）
    if fixed_worker:
        for (name, day), flag in requests.items():
            if not flag or name != fixed_worker or day not in valid_days:
                continue
            if is_weekday(day):
                errors.append(
                    f"[V8] '{fixed_worker}' は平日({day}日)は日勤固定のため "
                    f"希望休を適用できません"
                )

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
    settings: Optional[Settings] = None,
    roster: Optional[List[str]] = None,   # 後方互換（非推奨）
    solver_time_limit: float = 60.0,
    solver_workers: int = 4,
) -> pd.DataFrame:
    """
    警備員シフトスケジュールを最適化して返す。

    Parameters
    ----------
    year, month        : 対象年月
    requests           : { (名前, 日): True }  希望休
    fixed_assignments  : { (名前, 日): "日勤" など }  固定セル
    settings           : Settings オブジェクト（省略時はデフォルト値）
    roster             : 後方互換引数（非推奨・settings.roster が優先）
    solver_time_limit  : ソルバー最大実行時間（秒）
    solver_workers     : ソルバー並列スレッド数

    Returns
    -------
    pd.DataFrame  行=名前（roster順）、列=日付(1〜月末)、値=シフト種類

    Raises
    ------
    ShiftValidationError  入力に矛盾が検出された場合（solve前）
    RuntimeError          ソルバーが解を見つけられなかった場合
    """
    if settings is None:
        settings = _default_settings

    # 後方互換: roster 引数が渡された場合は settings を上書き
    if roster is not None:
        settings = Settings()
        settings.roster = list(roster)

    roster_list  = settings.roster
    fixed_worker = settings.fixed_worker
    shift_types  = settings.shift_types
    shift_hours  = settings.shift_hours
    night_shifts = settings.night_shifts
    constraints  = settings.constraints

    max_monthly_hours = constraints["月間上限時間"]
    required_day      = constraints["日勤必要人数"]
    required_night    = constraints["夜勤必要人数"]
    max_consecutive   = constraints["最大連続勤務日数"]
    sliding_window    = constraints["週休判定ウィンドウ幅"]

    num_days = calendar.monthrange(year, month)[1]
    days     = list(range(1, num_days + 1))

    def is_weekday(day: int) -> bool:
        return calendar.weekday(year, month, day) < 5

    shift_workers = [w for w in roster_list if w != fixed_worker]
    num_workers   = len(shift_workers)
    num_shifts    = len(shift_types)

    shift_index: Dict[str, int] = {s: i for i, s in enumerate(shift_types)}
    si_day    = shift_index["日勤"]
    si_rest   = shift_index["休日"]
    night_sis = [shift_index[s] for s in night_shifts]

    # ================================================================
    # solve前の整合性チェック
    # ================================================================
    validate_inputs(year, month, requests, fixed_assignments, settings)

    # ================================================================
    # モデル構築
    # ================================================================
    model  = cp_model.CpModel()
    solver = cp_model.CpSolver()

    x: Dict[Tuple[int, int, int], cp_model.IntVar] = {}
    for wi in range(num_workers):
        for d in days:
            for si in range(num_shifts):
                x[wi, d, si] = model.NewBoolVar(f"x_{wi}_{d}_{si}")

    # C1: 1日に1シフトのみ
    for wi in range(num_workers):
        for d in days:
            model.AddExactlyOne(x[wi, d, si] for si in range(num_shifts))

    # C2: 1日の必要人数
    for d in days:
        model.Add(sum(x[wi, d, si_day] for wi in range(num_workers)) == required_day)
        for si_n in night_sis:
            model.Add(sum(x[wi, d, si_n] for wi in range(num_workers)) == 1)

    # C3: 月間労働時間上限
    for wi in range(num_workers):
        model.Add(
            sum(
                x[wi, d, si] * shift_hours.get(shift_types[si], 0)
                for d in days for si in range(num_shifts)
            ) <= max_monthly_hours
        )

    # C4: 夜勤翌日の日勤禁止
    for wi in range(num_workers):
        for di in range(len(days) - 1):
            d, nd = days[di], days[di + 1]
            for si_n in night_sis:
                model.Add(x[wi, nd, si_day] == 0).OnlyEnforceIf(x[wi, d, si_n])

    # C5: 連続勤務上限
    for wi in range(num_workers):
        for di in range(len(days) - max_consecutive):
            window = [days[di + k] for k in range(max_consecutive + 1)]
            model.Add(
                sum(
                    x[wi, d, si]
                    for d in window for si in range(num_shifts)
                    if shift_types[si] != "休日"
                ) <= max_consecutive
            )

    # C6: スライディングウィンドウ週1休
    for wi in range(num_workers):
        for di in range(len(days) - sliding_window + 1):
            w7 = [days[di + k] for k in range(sliding_window)]
            model.Add(sum(x[wi, d, si_rest] for d in w7) >= 1)

    # C7: 希望休
    for (name, day), flag in requests.items():
        if flag and name in shift_workers and day in days:
            wi = shift_workers.index(name)
            model.Add(x[wi, day, si_rest] == 1)

    # C8: 固定セル
    for (name, day), shift in fixed_assignments.items():
        if name in shift_workers and day in days and shift in shift_index:
            wi = shift_workers.index(name)
            model.Add(x[wi, day, shift_index[shift]] == 1)

    # ================================================================
    # 目的関数: 役割別均等化（重み付き最小化）
    # ================================================================

    def _make_diff(label: str, counts_list: list, ub: int) -> cp_model.IntVar:
        v_max  = model.NewIntVar(0, ub, f"max_{label}")
        v_min  = model.NewIntVar(0, ub, f"min_{label}")
        v_diff = model.NewIntVar(0, ub, f"diff_{label}")
        model.AddMaxEquality(v_max, counts_list)
        model.AddMinEquality(v_min, counts_list)
        model.Add(v_diff == v_max - v_min)
        return v_diff

    WEIGHT_NIGHT = 6
    WEIGHT_DAY   = 4
    WEIGHT_HOURS = 1

    obj_terms = []

    for si_n in night_sis:
        counts = []
        for wi in range(num_workers):
            c = model.NewIntVar(0, num_days, f"cnt_n{si_n}_{wi}")
            model.Add(c == sum(x[wi, d, si_n] for d in days))
            counts.append(c)
        obj_terms.append((WEIGHT_NIGHT, _make_diff(f"night{si_n}", counts, num_days)))

    day_counts = []
    for wi in range(num_workers):
        c = model.NewIntVar(0, num_days, f"cnt_day_{wi}")
        model.Add(c == sum(x[wi, d, si_day] for d in days))
        day_counts.append(c)
    obj_terms.append((WEIGHT_DAY, _make_diff("day", day_counts, num_days)))

    hour_counts = []
    for wi in range(num_workers):
        h = model.NewIntVar(0, max_monthly_hours, f"hours_{wi}")
        model.Add(h == sum(
            x[wi, d, si] * shift_hours.get(shift_types[si], 0)
            for d in days for si in range(num_shifts)
        ))
        hour_counts.append(h)
    obj_terms.append((WEIGHT_HOURS, _make_diff("hours", hour_counts, max_monthly_hours)))

    obj_ub = (
        WEIGHT_NIGHT * len(night_sis) * num_days
        + WEIGHT_DAY * num_days
        + WEIGHT_HOURS * max_monthly_hours
    )
    combined = model.NewIntVar(0, obj_ub + 1, "combined_obj")
    model.Add(combined == sum(w * d for w, d in obj_terms))
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

    if fixed_worker:
        result[fixed_worker] = {
            d: ("日勤" if is_weekday(d) else "休日") for d in days
        }

    for wi, worker in enumerate(shift_workers):
        result[worker] = {}
        for d in days:
            assigned = "休日"
            for si in range(num_shifts):
                if solver.Value(x[wi, d, si]) == 1:
                    assigned = shift_types[si]
                    break
            result[worker][d] = assigned

    ordered = [w for w in roster_list if w in result]
    df = pd.DataFrame({w: result[w] for w in ordered}).T
    df.columns = pd.Index(days, name="日")
    df.index.name = "名前"

    return df


# ===========================================================================
# ユーティリティ
# ===========================================================================

def get_role_counts(
    df: pd.DataFrame,
    settings: Optional[Settings] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """各従業員の役割別担当回数と労働時間を集計して返す。"""
    if settings is None:
        settings = _default_settings

    shift_hours  = settings.shift_hours
    night_shifts = settings.night_shifts

    records = []
    for name in df.index:
        row = df.loc[name]
        row_dict: Dict = {"名前": name, "日勤": 0, "夜勤計": 0, "休日": 0, "労働時間(h)": 0}
        for s in night_shifts:
            row_dict[s] = 0

        for v in row:
            v = str(v)
            if v == "日勤":
                row_dict["日勤"] += 1
            elif v in night_shifts:
                row_dict[v] = row_dict.get(v, 0) + 1
                row_dict["夜勤計"] += 1
            elif v == "休日":
                row_dict["休日"] += 1
            row_dict["労働時間(h)"] += shift_hours.get(v, 0)

        records.append(row_dict)

    rc = pd.DataFrame(records).set_index("名前")

    numeric_cols = ["日勤"] + night_shifts + ["夜勤計", "休日", "労働時間(h)"]
    numeric_cols = [c for c in numeric_cols if c in rc.columns]
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
    settings: Optional[Settings] = None,
    show_bias: bool = True,
) -> None:
    """シフト集計サマリーをコンソールに表示する。"""
    if settings is None:
        settings = _default_settings

    rc, stats_df = get_role_counts(df, settings)
    night_shifts = settings.night_shifts

    W = 72
    print(f"\n{'='*W}")
    print(f"  {year}年{month}月 シフト最適化結果サマリー")
    print(f"{'='*W}")

    for name in rc.index:
        r = rc.loc[name]
        night_vals = "  ".join(f"{r.get(s, 0):>5}" for s in night_shifts)
        print(f"  {name:<10}  {r['日勤']:>4}  {night_vals}  {r['夜勤計']:>5}  {r['労働時間(h)']:>5}h")

    if show_bias:
        print(f"  {'-'*(W-2)}")
        for role in ["日勤"] + night_shifts + ["夜勤計", "労働時間(h)"]:
            if role not in stats_df.index:
                continue
            s    = stats_df.loc[role]
            label = "労働時間" if role == "労働時間(h)" else role
            unit  = "h" if role == "労働時間(h)" else "回"
            bias  = int(s["偏り(max-min)"])
            flag  = "  ← 偏りあり" if bias >= 3 else ""
            print(f"  {label:<6}  max={int(s['最大']):>3}{unit}  min={int(s['最小']):>3}{unit}  差={bias:>3}{unit}{flag}")

    print(f"{'='*W}\n")


# ===========================================================================
# スタンドアロン実行サンプル
# ===========================================================================

if __name__ == "__main__":
    from pathlib import Path

    YEAR, MONTH = 2025, 6
    s = Settings.load(Path("input.xlsx"))

    sample_requests = {("伊藤 晶俊", 5): True, ("吉村 智", 10): True}
    sample_fixed    = {("山田 誠", 1): "夜勤A", ("大西 信一", 1): "夜勤B", ("村主 博", 1): "夜勤C"}

    try:
        df = generate_shift(YEAR, MONTH, sample_requests, sample_fixed, settings=s)
        print(df.to_string())
        print_summary(df, YEAR, MONTH, settings=s)
    except ShiftValidationError as e:
        print(f"\n入力エラー:\n{e}")
    except RuntimeError as e:
        print(f"\n最適化エラー: {e}")
