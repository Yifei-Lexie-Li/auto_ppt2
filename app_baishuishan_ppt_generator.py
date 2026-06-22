from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

from app_bus_ppt_generator import (
    build_ppt_from_template_pages,
    extract_archive,
    find_filename_groups,
)


def crop_plate_preview(image_path: Path) -> bytes:
    with Image.open(image_path) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        width, height = image.size
        left = int(width * 0.34)
        top = int(height * 0.55)
        right = int(width * 0.66)
        bottom = int(height * 0.78)
        crop = image.crop((left, top, right, bottom))
        crop.thumbnail((360, 220), Image.Resampling.LANCZOS)
        output = io.BytesIO()
        crop.save(output, format="JPEG", quality=88)
        return output.getvalue()


def normalize_tail(value: str) -> str:
    clean = value.strip().upper().replace("辽B", "").replace(" ", "")
    return clean


def create_excel_template() -> bytes:
    output = io.BytesIO()
    frame = pd.DataFrame(
        [
            {"自编号": "6656", "车牌号": "辽B09379D"},
            {"自编号": "3917", "车牌号": "辽B06850D"},
            {"自编号": "6604", "车牌号": "辽B01158D"},
        ]
    )
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        frame.to_excel(writer, sheet_name="车牌对照表", index=False)
    output.seek(0)
    return output.getvalue()


def read_plate_excel(uploaded_file) -> dict[str, str]:
    if uploaded_file is None:
        return {}

    try:
        frame = pd.read_excel(uploaded_file, dtype=str)
    except ImportError as exc:
        raise RuntimeError("缺少 openpyxl 依赖，无法读取 Excel。") from exc

    normalized_columns = {str(column).strip(): column for column in frame.columns}
    bus_column = normalized_columns.get("自编号") or normalized_columns.get("编号")
    plate_column = normalized_columns.get("车牌号") or normalized_columns.get("车牌") or normalized_columns.get("后五位")
    if bus_column is None or plate_column is None:
        raise RuntimeError("Excel 需要包含 `自编号` 和 `车牌号` 两列。")

    mapping: dict[str, str] = {}
    for _, row in frame.iterrows():
        bus_no = str(row.get(bus_column, "")).strip()
        plate = str(row.get(plate_column, "")).strip()
        if bus_no and plate and bus_no.lower() != "nan" and plate.lower() != "nan":
            clean_bus_no = bus_no.replace(" ", "")
            mapping[clean_bus_no] = normalize_tail(plate)
            if "-" in clean_bus_no:
                mapping[clean_bus_no.rsplit("-", 1)[1]] = normalize_tail(plate)
    return mapping


def render_excel_review(groups: list[tuple[str, str, list[Path]]], excel_plate_map: dict[str, str], plate_prefix: str) -> dict[str, str]:
    st.subheader("Excel 对照表核对")
    st.caption("Excel 会自动填入车牌号；如果发现不对，可以在这里临时改正，生成 PPT 会使用你最终确认的值。")

    reviewed_plate_map: dict[str, str] = {}
    for route, bus_no, images in groups:
        row_key = f"{route}-{bus_no}"
        plate_tail = excel_plate_map.get(row_key) or excel_plate_map.get(bus_no, "")
        with st.container(border=True):
            cols = st.columns([1.0, 1.3, 1.0, 1.2])
            cols[0].metric("线路", f"{route}路")
            cols[1].metric("自编号", bus_no)
            cols[2].metric("照片数", len(images))
            with cols[3]:
                corrected_plate = st.text_input(
                    "车牌号",
                    value=f"{plate_prefix}{plate_tail}" if plate_tail else "",
                    key=f"review_plate_{row_key}",
                    placeholder=f"例如 {plate_prefix}01158D",
                )
                reviewed_tail = normalize_tail(corrected_plate)
                if reviewed_tail:
                    reviewed_plate_map[row_key] = reviewed_tail
                    reviewed_plate_map[bus_no] = reviewed_tail

            try:
                st.image(crop_plate_preview(images[0]), caption=f"{row_key} 首图车牌区域预览")
            except Exception:
                st.info("这辆车的预览图生成失败。")
    return reviewed_plate_map


def render_app() -> None:
    st.set_page_config(page_title="百岁山车体 PPT 生成", layout="wide")
    st.title("百岁山车体监测 PPT 生成")

    st.markdown(
        """
        适合照片名为 `线路-自编号-照片序号.JPG` 的资料，例如 `15-6604-1.JPG` 到 `15-6604-6.JPG`。
        每辆车按 6 张照片分成 3 页，每页 2 张，并复制模版 PPT 的背景、字体颜色和图片位置。
        """
    )

    photo_archive = st.file_uploader("上传照片压缩包", type=["zip", "rar"])
    template_file = st.file_uploader("上传 PPT 模版", type=["pptx"])
    plate_excel = st.file_uploader("上传 Excel 车牌对照表", type=["xlsx", "xls"])
    plate_prefix = st.text_input("车牌前缀", value="辽B")

    st.download_button(
        "下载 Excel 对照表模版",
        data=create_excel_template(),
        file_name="百岁山车牌对照表模版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("图片压缩设置"):
        max_image_side = st.slider("图片最大边长", min_value=1000, max_value=2600, value=1800, step=100)
        image_quality = st.slider("图片质量", min_value=70, max_value=95, value=85, step=1)

    if photo_archive is None:
        st.info("先上传照片压缩包和 Excel 对照表，网站会生成核对预览。")
        return

    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        archive_path = temp_dir / photo_archive.name
        archive_path.write_bytes(photo_archive.getbuffer())
        extract_dir = temp_dir / "photos"
        extract_dir.mkdir()

        try:
            extract_archive(archive_path, extract_dir)
            groups = find_filename_groups(extract_dir)
        except Exception as exc:
            st.error(f"照片包读取失败：{exc}")
            return

        if not groups:
            st.error("没有找到类似 `15-6604-1.JPG` 这种文件名格式的图片。")
            return

        if plate_excel is None:
            st.warning("请上传 Excel 车牌对照表后再核对和生成 PPT。")
            return

        try:
            excel_plate_map = read_plate_excel(plate_excel)
        except Exception as exc:
            st.error(str(exc))
            return

        if not excel_plate_map:
            st.error("Excel 里没有读取到有效车牌。请确认包含 `自编号` 和 `车牌号` 两列，并且车牌号列已填写。")
            return

        st.success(f"读取到 {len(groups)} 辆车，共 {sum(len(images) for _, _, images in groups)} 张照片。")
        st.info(f"已从 Excel 读取 {len(excel_plate_map)} 条车牌对照。")

        plate_map = render_excel_review(groups, excel_plate_map, plate_prefix)

        missing = [
            f"{route}-{bus_no}"
            for route, bus_no, _images in groups
            if f"{route}-{bus_no}" not in plate_map and bus_no not in plate_map
        ]
        if missing:
            st.warning(f"还有 {len(missing)} 辆车没有车牌：{', '.join(missing[:8])}")

        output_name = st.text_input("导出文件名", value="百岁山车体监测报告.pptx")
        if not output_name.lower().endswith(".pptx"):
            output_name = f"{output_name}.pptx"

        can_generate = template_file is not None and not missing
        if st.button("确认无误，生成 PPT", type="primary", disabled=not can_generate):
            template_path = temp_dir / template_file.name
            template_path.write_bytes(template_file.getbuffer())
            output_path = temp_dir / output_name

            try:
                page_count = build_ppt_from_template_pages(
                    template_path=template_path,
                    groups=groups,
                    plate_prefix=plate_prefix,
                    plate_map=plate_map,
                    output_path=output_path,
                    max_image_side=max_image_side,
                    image_quality=image_quality,
                )
            except Exception as exc:
                st.error(f"PPT 生成失败：{exc}")
                return

            st.success(f"已生成 {page_count} 页。")
            st.download_button(
                "下载生成的 PPT",
                data=output_path.read_bytes(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )


if __name__ == "__main__":
    render_app()
