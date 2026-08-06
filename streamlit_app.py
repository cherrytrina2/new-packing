from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pandas as pd
import streamlit as st

from pack_planner import (
    DEFAULT_SIZE_SAFETY_MARGIN_MM,
    box_specs_to_rows,
    generate_outputs,
    rows_to_box_specs,
)
from packing_parser import SUPPLIER_OPTIONS, fill_template_from_parsed, get_fixed_template_path, parse_packing_list

SCENARIO_OPTIONS = {
    "方案1（HQ+FR）": "scenario1",
    "方案2（仅FR）": "scenario2",
    "方案3（全部箱型：GP+HQ+FR）": "scenario3",
    "自动推荐": "auto",
    "一次生成全部方案": "all",
    "自定义箱型": "custom",
}
BOX_ORDER = ["20GP", "40GP", "40HQ", "20FR", "40FR", "710板", "880板"]

BOX_LIMIT_NOTE_ROWS = [
    {"箱型": "20GP", "备注": "门高预留可配置；韩国叉尺约2m，仅可接长<4m货物"},
    {"箱型": "40GP", "备注": "自动模式默认禁用"},
    {"箱型": "40HQ", "备注": "场景允许时可用；并列/堆叠优先在此箱型"},
    {"箱型": "20FR", "备注": "并列拼箱后宽度受并列宽限约束"},
    {"箱型": "40FR", "备注": "并列拼箱后宽度受并列宽限约束；果园港不接受VGM>40T/W>4500"},
    {"箱型": "710板", "备注": "重件且标准FR/GP无法承载时启用"},
    {"箱型": "880板", "备注": "重件且标准FR/GP无法承载时启用"},
]


def build_zip_bytes(paths: list[Path]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            zf.write(p, arcname=p.name)
    return buf.getvalue()


def _init_box_config_state():
    if "box_spec_rows" not in st.session_state:
        st.session_state.box_spec_rows = pd.DataFrame(box_specs_to_rows())
    if "size_safety_margin_mm" not in st.session_state:
        st.session_state.size_safety_margin_mm = float(DEFAULT_SIZE_SAFETY_MARGIN_MM)


def main():
    st.set_page_config(page_title="智能配箱网页端", page_icon="📦", layout="wide")
    if "show_box_limit_table" not in st.session_state:
        st.session_state.show_box_limit_table = True
    _init_box_config_state()

    st.markdown(
        """
        <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden; height: 0;}
        footer {visibility: hidden; height: 0;}
        [data-testid="stToolbar"] {display: none !important;}
        .block-container {padding-top: 0.9rem; padding-bottom: 1.1rem; max-width: 1080px;}
        .stApp {background: linear-gradient(180deg, #f3f8ff 0%, #eef4ff 100%);}
        .hero {
            background: linear-gradient(120deg, #0b3f86, #1b64c6);
            border-radius: 12px;
            padding: 16px 18px;
            color: #ffffff;
            margin-bottom: 12px;
            box-shadow: 0 8px 24px rgba(22, 70, 145, 0.18);
        }
        .hero-sub {opacity: 0.92; font-size: 13px; margin-top: 4px;}
        div[data-testid="stSegmentedControl"] [role="radiogroup"] {
            background: #e9eef7; border-radius: 10px; padding: 4px;
        }
        div[data-testid="stSegmentedControl"] label {
            border-radius: 8px !important;
            color: #42526a !important;
            font-weight: 600 !important;
        }
        div[data-testid="stSegmentedControl"] label[data-checked="true"] {
            background: linear-gradient(90deg,#0b3f86,#1b64c6) !important;
            color: #ffffff !important;
        }
        div[data-testid="stFileUploaderDropzone"] {
            border: 1px dashed #6fa0e8 !important;
            background: #f5f9ff !important;
            border-radius: 10px !important;
        }
        div[data-testid="stSelectbox"] > div, div[data-testid="stMultiSelect"] > div {
            border-radius: 10px !important;
        }
        div.stButton > button {
            border-radius: 10px;
            font-weight: 700;
            height: 2.85rem;
            border: 0;
            color: #ffffff;
            background: linear-gradient(90deg,#0b3f86,#1b64c6);
            box-shadow: 0 8px 18px rgba(24, 92, 184, 0.25);
            transition: transform .16s ease, box-shadow .16s ease, filter .16s ease;
        }
        div.stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 12px 22px rgba(24, 92, 184, 0.30);
            filter: brightness(0.97);
        }
        div.stButton > button:active {transform: translateY(0);}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="hero"><div style="font-size:22px;font-weight:800;">智能配箱网页端</div><div class="hero-sub">上传 Excel，选择模块后一键生成结果</div></div>',
        unsafe_allow_html=True,
    )

    module_options = ["配箱", "解析 Packing List", "箱型限制表"]
    if hasattr(st, "segmented_control"):
        mode = st.segmented_control("功能模块", module_options, default="配箱")
    else:
        mode = st.radio("功能模块", module_options, horizontal=True)

    if mode == "箱型限制表":
        st.subheader("箱型限制表（可配置）")
        st.caption("货物大多不可横放；请尽量减少横放。所有货物包装后总高不大于4500mm。")
        st.caption("下表数值为**名义尺寸**；实际配箱硬约束 = 名义尺寸 − 安全余量。修改后点「应用配置」生效，配箱时会自动使用。")
        margin = st.number_input(
            "尺寸安全余量 (mm)",
            min_value=0.0,
            max_value=500.0,
            value=float(st.session_state.size_safety_margin_mm),
            step=5.0,
            help="从每个箱型的限长/限宽/限高中统一扣除的余量，默认 75mm。",
        )
        edited = st.data_editor(
            st.session_state.box_spec_rows,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="box_spec_editor_main",
        )
        st.dataframe(pd.DataFrame(BOX_LIMIT_NOTE_ROWS), use_container_width=True, hide_index=True)
        cols = st.columns(2)
        with cols[0]:
            if st.button("应用配置", type="primary", use_container_width=True):
                st.session_state.box_spec_rows = pd.DataFrame(edited)
                st.session_state.size_safety_margin_mm = float(margin)
                st.success("已应用箱型尺寸与安全余量，后续配箱将使用此配置。")
        with cols[1]:
            if st.button("恢复默认", use_container_width=True):
                st.session_state.box_spec_rows = pd.DataFrame(box_specs_to_rows())
                st.session_state.size_safety_margin_mm = float(DEFAULT_SIZE_SAFETY_MARGIN_MM)
                st.success("已恢复默认箱型尺寸与 75mm 余量。")
                st.rerun()
        return

    uploaded = st.file_uploader("输入文件（.xlsx）", type=["xlsx"])
    scenario = "scenario1"
    auto_use_hq = False
    enable_stack = False
    custom_boxes: list[str] = []
    if mode == "配箱":
        scenario_label = st.selectbox("配箱场景", list(SCENARIO_OPTIONS.keys()), index=0)
        scenario = SCENARIO_OPTIONS[scenario_label]
        auto_use_hq = st.checkbox(
            "自动模式启用 HQ",
            value=False,
            disabled=scenario not in ("auto", "all"),
            help="仅在“自动模式/全部场景”时生效。",
        )
        enable_stack = st.checkbox(
            "启用堆叠",
            value=False,
            help="与并列同级：在配箱前按尺寸条件将两件货垂直合成一个单元，再分配箱型。"
            "仅封闭箱型（20GP/40GP/40HQ）支持；FR 与定制板不支持堆叠。",
        )
        margin_now = st.number_input(
            "尺寸安全余量 (mm)",
            min_value=0.0,
            max_value=500.0,
            value=float(st.session_state.size_safety_margin_mm),
            step=5.0,
            key="pack_margin_input",
            help="从箱型限长/限宽/限高统一扣除。余量=75 时 40HQ 有效高=2425，"
            "1250+1250 无法堆叠；余量=0 时有效高=2500，可以两层堆。",
        )
        st.session_state.size_safety_margin_mm = float(margin_now)
        hq_eff_h = 2500.0 - float(margin_now)
        st.caption(
            f"当前生效：余量 **{margin_now:g} mm** → 40HQ 有效限高 **{hq_eff_h:g} mm**"
            f"（1250+1250 堆叠：{'可以' if 2500 <= hq_eff_h else '不可以'}）"
        )
        if scenario == "custom":
            custom_boxes = st.multiselect(
                "自定义箱型",
                BOX_ORDER,
                default=["40FR", "20FR"],
                help="仅在“自定义箱型组合”场景生效。",
            )
            priority_default = ",".join(custom_boxes) if custom_boxes else "40FR,20FR"
            priority_text = st.text_input(
                "箱型优先级顺序（逗号分隔）",
                value=priority_default,
                help="示例：40HQ,40FR,20FR,20GP。仅保留已勾选箱型，按你填写顺序配箱。",
            )
            typed = [x.strip() for x in priority_text.replace("，", ",").split(",") if x.strip()]
            if typed:
                ordered = [x for x in typed if x in custom_boxes]
                for c in custom_boxes:
                    if c not in ordered:
                        ordered.append(c)
                custom_boxes = ordered
        run = st.button("生成配箱结果", type="primary", use_container_width=True)
    else:
        supplier_label = st.selectbox("List 供应商模板", list(SUPPLIER_OPTIONS.keys()), index=0)
        run = st.button("解析并标准化输出", type="primary", use_container_width=True)

    if run:
        if uploaded is None:
            st.error("请先上传 Excel 文件。")
            return
        if mode == "配箱" and scenario == "custom" and not custom_boxes:
            st.error("自定义模式下，请至少选择一个箱型。")
            return

        out_root = Path.cwd() / "out_web_streamlit"
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:6]
        job_dir = out_root / stamp
        job_dir.mkdir(parents=True, exist_ok=True)

        input_path = job_dir / uploaded.name
        input_path.write_bytes(uploaded.getbuffer())

        if mode == "配箱":
            with st.spinner("正在生成，请稍候..."):
                try:
                    box_specs = rows_to_box_specs(st.session_state.box_spec_rows)
                    outputs = generate_outputs(
                        input_path=input_path,
                        outdir=job_dir,
                        scenario=scenario,
                        custom_boxes=custom_boxes,
                        auto_use_hq=auto_use_hq,
                        enable_stack=enable_stack,
                        box_specs=box_specs,
                        size_safety_margin_mm=float(st.session_state.size_safety_margin_mm),
                    )
                except Exception as exc:
                    st.error(f"生成失败：{exc}")
                    return

            st.success(
                f"生成完成（安全余量 {st.session_state.size_safety_margin_mm:g} mm，"
                f"40HQ 有效限高 {2500 - float(st.session_state.size_safety_margin_mm):g} mm）"
            )
            output_paths = [p for _, p in outputs]
            for name, p in outputs:
                data = p.read_bytes()
                st.download_button(
                    label=f"下载：{name}",
                    data=data,
                    file_name=p.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            zip_name = f"packing-results-{stamp}.zip"
            zip_data = build_zip_bytes(output_paths)
            st.download_button(
                label="下载全部结果（ZIP）",
                data=zip_data,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
            )
        else:
            with st.spinner("正在解析，请稍候..."):
                try:
                    supplier = SUPPLIER_OPTIONS[supplier_label]
                    out = parse_packing_list(input_path, job_dir / f"{input_path.stem}-parsed.xlsx", supplier=supplier)
                except Exception as exc:
                    st.error(f"解析失败：{exc}")
                    return
            st.success("解析完成")
            st.download_button(
                label="下载解析结果",
                data=out.read_bytes(),
                file_name=out.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            template_path = get_fixed_template_path()
            if template_path.exists():
                try:
                    filled = fill_template_from_parsed(out, template_path, job_dir / f"{input_path.stem}-template-filled.xlsx")
                    st.download_button(
                        label="下载模板填充结果",
                        data=filled.read_bytes(),
                        file_name=filled.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.warning(f"解析已完成，但模板填充失败：{exc}")
            else:
                st.warning(f"固定模板不存在：{template_path}。当前仅输出解析结果。")
        st.caption(f"结果目录：{job_dir}")

    with st.sidebar:
        if st.button("箱型限制表格", use_container_width=True):
            st.session_state.show_box_limit_table = not st.session_state.show_box_limit_table
        if st.session_state.show_box_limit_table:
            st.markdown("### 箱型限制表（当前配置）")
            st.caption(f"安全余量：{st.session_state.size_safety_margin_mm:g} mm（硬约束 = 名义 − 余量）")
            st.caption("货物大多不可横放；总高不大于4500mm。完整编辑请切到「箱型限制表」模块。")
            show_cols = ["箱型", "限长L(mm)", "限宽W(mm)", "限高H(mm)", "载重(kg)"]
            df_show = st.session_state.box_spec_rows
            st.dataframe(df_show[[c for c in show_cols if c in df_show.columns]], use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
