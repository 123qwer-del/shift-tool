"""
警備員シフト自動生成 Web版 (Streamlit)
========================================
app.py (tkinter GUI) を Streamlit に移植したバージョン。
"""

import calendar
import io
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ─── ページ設定 ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="警備員シフト自動生成",
    page_icon="🛡️",
    layout="wide",
)

# ─── CSSカスタマイズ ─────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0f0f1a; }
  [data-testid="stSidebar"] { background: #1a1a2e; }
  h1, h2, h3 { color: #e0e0f0; }
  .stDataFrame { border-radius: 8px; }
  .shift-day  { background:#1e40af; color:#fff; padding:2px 6px; border-radius:4px; font-size:12px; }
  .shift-nightA { background:#6d28d9; color:#fff; padding:2px 6px; border-radius:4px; font-size:12px; }
  .shift-nightB { background:#7c3aed; color:#fff; padding:2px 6px; border-radius:4px; font-size:12px; }
  .shift-nightC { background:#8b5cf6; color:#fff; padding:2px 6px; border-radius:4px; font-size:12px; }
  .shift-rest { background:#374151; color:#9ca3af; padding:2px 6px; border-radius:4px; font-size:12px; }
</style>
""", unsafe_allow_html=True)


# ─── Settings / optimizer を動的インポート ───────────────────────────────────
# ユーザーがアップロードしたファイルをメモリ内で扱う

def _load_module_from_text(name: str, source: str):
    """文字列ソースからモジュールを動的ロードする。"""
    import types
    mod = types.ModuleType(name)
    exec(compile(source, f"<{name}>", "exec"), mod.__dict__)
    sys.modules[name] = mod
    return mod


def _load_settings_from_bytes(s_mod, data: bytes):
    """xlsxバイト列からSettingsを読み込む。"""
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(data)
        tmp_path = f.name
    try:
        return s_mod.Settings.load(Path(tmp_path))
    finally:
        os.unlink(tmp_path)


def _load_requests_fixed_from_bytes(data: bytes):
    """xlsxバイト列から希望休・固定シートを読み込む。"""
    def _read_sheet(xls, sheet_name):
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        header_row = 0
        for idx, row in raw.iterrows():
            vals = [str(v).strip() for v in row if pd.notna(v)]
            if "名前" in vals:
                header_row = idx
                break
        return pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

    xls = io.BytesIO(data)

    try:
        df_req = _read_sheet(xls, "希望休")
        requests = {}
        for _, row in df_req.iterrows():
            if pd.isna(row.get("名前")) or pd.isna(row.get("日")):
                continue
            name = str(row["名前"]).strip()
            day  = int(row["日"])
            if name and day:
                requests[(name, day)] = True
    except Exception:
        requests = {}

    try:
        df_fix = _read_sheet(xls, "固定")
        fixed = {}
        for _, row in df_fix.iterrows():
            if pd.isna(row.get("名前")) or pd.isna(row.get("日")) or pd.isna(row.get("シフト")):
                continue
            name  = str(row["名前"]).strip()
            day   = int(row["日"])
            shift = str(row["シフト"]).strip()
            if name and day and shift:
                fixed[(name, day)] = shift
    except Exception:
        fixed = {}

    return requests, fixed


# ─── セッションステート初期化 ────────────────────────────────────────────────
for key, default in {
    "settings_mod":    None,
    "optimizer_mod":   None,
    "settings_obj":    None,
    "input_bytes":     None,
    "result_df":       None,
    "result_df_orig":  None,
    "last_year":       None,
    "last_month":      None,
    "last_requests":   None,
    "last_file_fixed": None,
    "log_lines":       [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_lines.append(f"[{ts}] {msg}")


# ═══════════════════════════════════════════════════════════════════════════════
# 起動時自動読み込み: settings.py / optimizer.py はリポジトリから直接インポート
# ═══════════════════════════════════════════════════════════════════════════════

def _patch_settings_src(src: str) -> str:
    """get_base_dir() をWeb環境用にパッチする。"""
    return src.replace(
        "return Path(sys.executable).parent",
        "return Path('.')"
    ).replace(
        "return Path(__file__).parent",
        "return Path('.')"
    )


def _auto_load_modules():
    """
    同じリポジトリにある settings.py / optimizer.py を起動時に自動読み込みする。
    すでに読み込み済みならスキップ。
    """
    if st.session_state.settings_mod and st.session_state.optimizer_mod:
        return  # 読み込み済み

    base = Path(__file__).parent
    s_path = base / "settings.py"
    o_path = base / "optimizer.py"

    missing = []
    if not s_path.exists():
        missing.append("settings.py")
    if not o_path.exists():
        missing.append("optimizer.py")
    if missing:
        st.error(f"リポジトリに {', '.join(missing)} が見つかりません。GitHubに追加してください。")
        st.stop()

    try:
        s_src = _patch_settings_src(s_path.read_text(encoding="utf-8"))
        o_src = o_path.read_text(encoding="utf-8")

        s_mod = _load_module_from_text("settings",  s_src)
        o_mod = _load_module_from_text("optimizer", o_src)

        st.session_state.settings_mod  = s_mod
        st.session_state.optimizer_mod = o_mod

        # input.xlsx がリポジトリにあればデフォルトとして読み込む
        i_path = base / "input.xlsx"
        if i_path.exists() and st.session_state.settings_obj is None:
            data = i_path.read_bytes()
            st.session_state.input_bytes  = data
            st.session_state.settings_obj = _load_settings_from_bytes(s_mod, data)
        elif st.session_state.settings_obj is None:
            st.session_state.settings_obj = s_mod.Settings()

        log("モジュール自動読み込み完了")
    except Exception as e:
        st.error(f"モジュール読み込みエラー: {e}")
        log(f"モジュール読み込みエラー: {e}")
        st.stop()


_auto_load_modules()


# ═══════════════════════════════════════════════════════════════════════════════
# サイドバー: input.xlsx のアップロードのみ
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🛡️ シフト自動生成")
    st.caption("警備員スケジューリング v3.2 Web版")
    st.divider()

    st.subheader("📂 入力ファイル")
    st.caption("settings.py / optimizer.py はリポジトリから自動読み込みされます")

    input_xlsx = st.file_uploader("input.xlsx（希望休・固定シート）", type="xlsx", key="up_input")

    if input_xlsx:
        if st.button("📥 input.xlsx を読み込む", use_container_width=True):
            try:
                data = input_xlsx.read()
                st.session_state.input_bytes  = data
                st.session_state.settings_obj = _load_settings_from_bytes(
                    st.session_state.settings_mod, data
                )
                log(f"input.xlsx を読み込みました")
                st.success("✅ 読み込み完了")
            except Exception as e:
                st.error(f"読み込みエラー: {e}")
                log(f"input.xlsx 読み込みエラー: {e}")

    st.divider()
    st.success("✅ settings.py 読み込み済み")
    st.success("✅ optimizer.py 読み込み済み")
    if st.session_state.input_bytes:
        st.success("✅ input.xlsx 読み込み済み")
    else:
        st.info("ℹ️ input.xlsx 未読み込み（デフォルト設定で動作）")


# ═══════════════════════════════════════════════════════════════════════════════
# メインエリア: タブ構成
# ═══════════════════════════════════════════════════════════════════════════════
tab_shift, tab_settings, tab_log = st.tabs(["🗓️ シフト生成", "⚙️ 設定確認", "📋 ログ"])


# ─── タブ1: シフト生成 ────────────────────────────────────────────────────────
with tab_shift:
    st.header("シフト自動生成")

    col1, col2, col3 = st.columns([1, 1, 2])
    now = datetime.now()
    with col1:
        year  = st.number_input("対象年", min_value=2020, max_value=2035, value=now.year)
    with col2:
        month = st.number_input("対象月", min_value=1, max_value=12, value=now.month)

    st.divider()

    # 希望休入力
    st.subheader("📋 希望休の入力")
    settings_obj = st.session_state.settings_obj
    roster = settings_obj.roster if settings_obj else []

    num_days_preview = calendar.monthrange(int(year), int(month))[1]

    if roster:
        req_data = {}
        cols = st.columns(min(len(roster), 3))
        for i, worker in enumerate(roster):
            with cols[i % 3]:
                days_off = st.multiselect(
                    f"🧑 {worker}",
                    options=list(range(1, num_days_preview + 1)),
                    key=f"req_{worker}",
                )
                for d in days_off:
                    req_data[(worker, d)] = True
    else:
        st.info("先にサイドバーからファイルを読み込んでください")
        req_data = {}

    st.divider()

    run_disabled = not (st.session_state.optimizer_mod and st.session_state.settings_mod)
    if st.button("🚀 シフトを自動生成", disabled=run_disabled, type="primary", use_container_width=True):
        o_mod = st.session_state.optimizer_mod
        s_obj = st.session_state.settings_obj or st.session_state.settings_mod.Settings()

        # input.xlsx から希望休・固定を追加読み込み
        file_requests, file_fixed = {}, {}
        if st.session_state.input_bytes:
            try:
                file_requests, file_fixed = _load_requests_fixed_from_bytes(
                    st.session_state.input_bytes
                )
            except Exception as e:
                st.warning(f"input.xlsx の読み込みをスキップ: {e}")

        # GUIからの希望休をマージ（上書き優先）
        merged_requests = {**file_requests, **req_data}

        with st.spinner("CP-SATソルバーで最適化中..."):
            try:
                log(f"最適化開始: {year}年{month}月")
                log(f"希望休: {len(merged_requests)}件  固定: {len(file_fixed)}件")

                df = o_mod.generate_shift(
                    int(year), int(month),
                    merged_requests, file_fixed,
                    settings=s_obj,
                )
                st.session_state.result_df       = df
                st.session_state.result_df_orig  = df.copy()
                st.session_state.last_year       = int(year)
                st.session_state.last_month      = int(month)
                st.session_state.last_requests   = merged_requests
                st.session_state.last_file_fixed = file_fixed
                log("最適化完了")
                st.success("✅ シフト生成完了！")
            except Exception as e:
                err_type = type(e).__name__
                msg = str(e)
                st.error(f"**{err_type}**\n\n{msg}")
                log(f"エラー: {err_type}: {msg}")

    # ─── 結果表示・手動修整 ───────────────────────────────────────────────────
    if st.session_state.result_df is not None:
        df_orig = st.session_state.result_df_orig
        df_cur  = st.session_state.result_df

        SHIFT_TYPES_OPTIONS = ["日勤", "夜勤A", "夜勤B", "夜勤C", "休日"]

        # 手動修整の差分を検出
        def _get_manual_edits(orig, cur):
            fixed_worker = getattr(
                st.session_state.optimizer_mod, "FIXED_WORKER", "末吉 弘一"
            )
            edits = {}
            for name in orig.index:
                if name == fixed_worker:
                    continue
                for col in orig.columns:
                    if str(orig.loc[name, col]) != str(cur.loc[name, col]):
                        edits[(name, int(col))] = str(cur.loc[name, col])
            return edits

        manual_edits = _get_manual_edits(df_orig, df_cur)
        n_edits = len(manual_edits)

        # ヘッダー + 修整カウントバッジ
        hcol1, hcol2 = st.columns([3, 1])
        with hcol1:
            st.subheader("📊 生成結果（セルを直接クリックして編集できます）")
        with hcol2:
            if n_edits > 0:
                st.markdown(
                    f"<div style='background:#b45309;color:white;border-radius:6px;"
                    f"padding:8px 12px;margin-top:8px;font-weight:bold;text-align:center;'>"
                    f"✏️ {n_edits} セル手動修整済み</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    "<div style='background:#1f4e3d;color:#6ee7b7;border-radius:6px;"
                    "padding:8px 12px;margin-top:8px;text-align:center;'>"
                    "自動生成値（未修整）</div>",
                    unsafe_allow_html=True,
                )

        # 編集可能テーブル（日付列をドロップダウンに）
        day_col_cfg = {
            str(c): st.column_config.SelectboxColumn(
                label=str(c),
                options=SHIFT_TYPES_OPTIONS,
                required=True,
                width="small",
            )
            for c in df_cur.columns
        }
        col_cfg = {
            "名前": st.column_config.TextColumn("名前", disabled=True),
            **day_col_cfg,
        }

        edited_df = st.data_editor(
            df_cur,
            column_config=col_cfg,
            use_container_width=True,
            key="shift_editor",
            hide_index=False,
        )

        # 変更があればセッションに反映してすぐ再描画
        if not edited_df.equals(df_cur):
            st.session_state.result_df = edited_df
            st.rerun()

        # 修整セル一覧（折りたたみ）
        if n_edits > 0:
            with st.expander(f"✏️ 手動修整セル一覧（{n_edits}件）"):
                edit_records = [
                    {
                        "名前": name,
                        "日": day,
                        "変更前": df_orig.loc[name, day],
                        "変更後": shift,
                    }
                    for (name, day), shift in sorted(
                        manual_edits.items(), key=lambda x: (x[0][0], x[0][1])
                    )
                ]
                st.dataframe(
                    pd.DataFrame(edit_records),
                    use_container_width=True,
                    hide_index=True,
                )

        # アクションボタン行
        btn1, btn2, btn3 = st.columns([2, 1, 1])

        with btn1:
            recalc_disabled = (
                n_edits == 0
                or st.session_state.last_year is None
                or not st.session_state.optimizer_mod
            )
            if st.button(
                "🔄 手動修整を固定して再計算",
                disabled=recalc_disabled,
                type="primary",
                use_container_width=True,
            ):
                o_mod  = st.session_state.optimizer_mod
                s_obj  = st.session_state.settings_obj or st.session_state.settings_mod.Settings()
                y      = st.session_state.last_year
                m      = st.session_state.last_month
                reqs   = st.session_state.last_requests or {}
                merged_fixed = {
                    **(st.session_state.last_file_fixed or {}),
                    **manual_edits,
                }
                with st.spinner(f"手動修整を固定して再最適化中（{n_edits}セル固定）..."):
                    try:
                        log(f"再計算: {y}年{m}月  固定セル={len(merged_fixed)}件")
                        new_df = o_mod.generate_shift(
                            y, m, reqs, merged_fixed, settings=s_obj,
                        )
                        st.session_state.result_df       = new_df
                        st.session_state.result_df_orig  = new_df.copy()
                        st.session_state.last_file_fixed = merged_fixed
                        log("再計算完了")
                        st.success("✅ 再計算完了！")
                        st.rerun()
                    except Exception as e:
                        err_type = type(e).__name__
                        st.error(f"**{err_type}**\n\n{e}")
                        log(f"再計算エラー: {err_type}: {e}")

        with btn2:
            if st.button(
                "↩️ 修整をリセット",
                disabled=(n_edits == 0),
                use_container_width=True,
            ):
                st.session_state.result_df = st.session_state.result_df_orig.copy()
                log("手動修整をリセット")
                st.rerun()

        with btn3:
            buf = io.BytesIO()
            df_cur.to_excel(buf, index=True)
            lm = st.session_state.last_month or 1
            ly = st.session_state.last_year or 2025
            st.download_button(
                label="⬇️ Excel出力",
                data=buf.getvalue(),
                file_name=f"shift_{ly}{lm:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()

        # サマリー
        if hasattr(st.session_state.optimizer_mod, "get_role_counts"):
            try:
                rc, stats_df = st.session_state.optimizer_mod.get_role_counts(df_cur)
                st.subheader("📈 役割回数サマリー")
                st.dataframe(rc, use_container_width=True)
                st.subheader("📉 偏り指標（固定ワーカー除外）")
                st.dataframe(stats_df, use_container_width=True)
            except Exception:
                pass


# ─── タブ2: 設定編集 ──────────────────────────────────────────────────────────
with tab_settings:
    st.header("設定編集")
    s = st.session_state.settings_obj
    if s is None:
        st.info("サイドバーからファイルを読み込むと設定が表示されます")
    else:
        col_a, col_b = st.columns(2)

        # ── 左列: 従業員名簿 & 固定ワーカー & シフト種類 ──────────────────────
        with col_a:
            st.subheader("👥 従業員名簿")
            st.caption("1行1名。順番がシフト表の行順になります。")
            roster_text = st.text_area(
                label="従業員名簿（1行1名）",
                value="\n".join(s.roster),
                height=220,
                key="edit_roster",
                label_visibility="collapsed",
            )

            st.subheader("⭐ 固定ワーカー")
            st.caption("平日=日勤固定・土日=休日固定にする従業員名（空欄で無効）")
            fixed_worker_input = st.text_input(
                label="固定ワーカー名",
                value=s.fixed_worker or "",
                key="edit_fixed_worker",
                label_visibility="collapsed",
            )

            st.subheader("🕐 シフト種類・勤務時間")
            st.caption("シフト名と時間(h)を編集できます。「休日」は必須です。")
            sh_df_edit = pd.DataFrame(
                [(k, v) for k, v in s.shift_hours.items()],
                columns=["シフト名", "勤務時間(h)"]
            )
            edited_shifts = st.data_editor(
                sh_df_edit,
                num_rows="dynamic",
                use_container_width=True,
                key="edit_shift_hours",
                column_config={
                    "シフト名":    st.column_config.TextColumn("シフト名", required=True),
                    "勤務時間(h)": st.column_config.NumberColumn("勤務時間(h)", min_value=0, max_value=24, step=1),
                },
            )

        # ── 右列: 制約パラメータ ───────────────────────────────────────────────
        with col_b:
            st.subheader("📐 制約パラメータ")
            st.caption("数値を直接クリックして編集できます。")

            CONSTRAINT_DESCRIPTIONS = {
                "月間上限時間":         "1人あたりの月間最大労働時間 (h)",
                "日勤必要人数":         "1日に必要な日勤担当者数",
                "夜勤必要人数":         "1日に必要な夜勤担当者数 (A+B+C の合計)",
                "最大連続勤務日数":     "連続して勤務できる最大日数",
                "週休判定ウィンドウ幅": "週1休を判定するスライディングウィンドウの幅 (日)",
            }
            c_df_edit = pd.DataFrame([
                {
                    "パラメータ名": k,
                    "値": v,
                    "説明": CONSTRAINT_DESCRIPTIONS.get(k, ""),
                }
                for k, v in s.constraints.items()
            ])
            edited_constraints = st.data_editor(
                c_df_edit,
                use_container_width=True,
                key="edit_constraints",
                disabled=["パラメータ名", "説明"],
                column_config={
                    "パラメータ名": st.column_config.TextColumn("パラメータ名"),
                    "値":           st.column_config.NumberColumn("値", min_value=0, step=1),
                    "説明":         st.column_config.TextColumn("説明"),
                },
            )

        st.divider()

        # ── 適用ボタン ─────────────────────────────────────────────────────────
        if st.button("✅ 設定を適用する", type="primary", use_container_width=True):
            try:
                # 名簿パース
                new_roster = [
                    name.strip()
                    for name in roster_text.splitlines()
                    if name.strip()
                ]
                if not new_roster:
                    st.error("従業員名簿が空です。")
                    st.stop()

                # シフト時間パース
                new_shift_hours = {}
                for _, row in edited_shifts.iterrows():
                    name_val  = str(row["シフト名"]).strip()
                    hours_val = int(row["勤務時間(h)"] or 0)
                    if name_val:
                        new_shift_hours[name_val] = hours_val
                if not new_shift_hours:
                    st.error("シフト種類が空です。")
                    st.stop()
                if "休日" not in new_shift_hours:
                    st.error("「休日」シフトは必須です。")
                    st.stop()

                # 制約パース
                new_constraints = {}
                for _, row in edited_constraints.iterrows():
                    key = str(row["パラメータ名"]).strip()
                    val = int(row["値"] or 0)
                    new_constraints[key] = val

                # Settingsオブジェクトを更新
                s.roster       = new_roster
                s.fixed_worker = fixed_worker_input.strip()
                s.shift_hours  = new_shift_hours
                s.constraints  = new_constraints
                st.session_state.settings_obj = s

                # バリデーション
                errors = s.validate()
                if errors:
                    for e in errors:
                        st.error(e)
                else:
                    log("設定を更新しました")
                    st.success("✅ 設定を適用しました。次回のシフト生成から反映されます。")

            except Exception as e:
                st.error(f"設定の適用に失敗しました: {e}")


# ─── タブ3: ログ ──────────────────────────────────────────────────────────────
with tab_log:
    st.header("実行ログ")
    if st.button("🗑️ ログをクリア"):
        st.session_state.log_lines = []
    log_text = "\n".join(st.session_state.log_lines) or "（ログなし）"
    st.code(log_text, language=None)
