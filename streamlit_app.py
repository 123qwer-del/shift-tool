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
    "settings_mod": None,
    "optimizer_mod": None,
    "settings_obj": None,
    "input_bytes": None,
    "result_df": None,
    "log_lines": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_lines.append(f"[{ts}] {msg}")


# ═══════════════════════════════════════════════════════════════════════════════
# サイドバー: ファイルアップロード & モジュール読み込み
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🛡️ シフト自動生成")
    st.caption("警備員スケジューリング v3.1 Web版")
    st.divider()

    st.subheader("📂 ファイルアップロード")

    settings_file = st.file_uploader("settings.py", type="py", key="up_settings")
    optimizer_file = st.file_uploader("optimizer.py", type="py", key="up_optimizer")
    input_xlsx     = st.file_uploader("input.xlsx",   type="xlsx", key="up_input")

    if st.button("⚙️ モジュールを読み込む", use_container_width=True):
        errors = []
        if not settings_file:
            errors.append("settings.py をアップロードしてください")
        if not optimizer_file:
            errors.append("optimizer.py をアップロードしてください")
        if errors:
            for e in errors:
                st.error(e)
        else:
            try:
                settings_src  = settings_file.read().decode("utf-8")
                optimizer_src = optimizer_file.read().decode("utf-8")

                # get_base_dir() をWeb環境用にパッチ
                settings_src = settings_src.replace(
                    "return Path(sys.executable).parent",
                    "return Path('.')"
                ).replace(
                    "return Path(__file__).parent",
                    "return Path('.')"
                )

                s_mod = _load_module_from_text("settings",  settings_src)
                o_mod = _load_module_from_text("optimizer", optimizer_src)

                st.session_state.settings_mod  = s_mod
                st.session_state.optimizer_mod = o_mod

                if input_xlsx:
                    data = input_xlsx.read()
                    st.session_state.input_bytes = data
                    st.session_state.settings_obj = s_mod.Settings.load_from_bytes(data) \
                        if hasattr(s_mod.Settings, "load_from_bytes") \
                        else _load_settings_from_bytes(s_mod, data)
                else:
                    st.session_state.settings_obj = s_mod.Settings()

                log("モジュール読み込み完了")
                st.success("✅ 読み込み完了")
            except Exception as e:
                st.error(f"読み込みエラー: {e}")
                log(f"読み込みエラー: {e}")

    st.divider()
    if st.session_state.settings_mod:
        st.success("✅ settings.py 読み込み済み")
    else:
        st.warning("⚠️ settings.py 未読み込み")
    if st.session_state.optimizer_mod:
        st.success("✅ optimizer.py 読み込み済み")
    else:
        st.warning("⚠️ optimizer.py 未読み込み")


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
                st.session_state.result_df = df
                log("最適化完了")
                st.success("✅ シフト生成完了！")
            except Exception as e:
                err_type = type(e).__name__
                msg = str(e)
                st.error(f"**{err_type}**\n\n{msg}")
                log(f"エラー: {err_type}: {msg}")

    # 結果表示
    if st.session_state.result_df is not None:
        df = st.session_state.result_df
        st.subheader("📊 生成結果")

        # カラーマップ
        SHIFT_COLORS = {
            "日勤":  "#1e40af",
            "夜勤A": "#6d28d9",
            "夜勤B": "#7c3aed",
            "夜勤C": "#8b5cf6",
            "休日":  "#374151",
        }

        def color_shift(val):
            color = SHIFT_COLORS.get(str(val), "#374151")
            return f"background-color: {color}; color: white; text-align: center; font-size: 11px;"

        styled = df.style.applymap(color_shift)
        st.dataframe(styled, use_container_width=True)

        # サマリー
        if hasattr(st.session_state.optimizer_mod, "get_role_counts"):
            try:
                rc, stats_df = st.session_state.optimizer_mod.get_role_counts(
                    df, st.session_state.settings_obj
                )
                st.subheader("📈 役割回数サマリー")
                st.dataframe(rc, use_container_width=True)
                st.subheader("📉 偏り指標")
                st.dataframe(stats_df, use_container_width=True)
            except Exception:
                pass

        # ダウンロード
        buf = io.BytesIO()
        df.to_excel(buf, index=True)
        st.download_button(
            label="⬇️ Excelでダウンロード",
            data=buf.getvalue(),
            file_name=f"output_shift_{year}{month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


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
