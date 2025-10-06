import io
import json
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Excel Merge Tool", layout="wide")

st.title("🔗 Excel Merge & Enrichment Tool")
st.write(
    """
Tải lên 2 file Excel có cùng header:

- File 1: Nhiều record hơn nhưng thiếu dữ liệu.
- File 2: Ít record hơn nhưng dữ liệu đầy đủ hơn.

Chọn các cột khóa để đối chiếu. Tool sẽ:
1. Ghép (merge) theo các cột khóa.
2. Với các ô bị thiếu trong File 1 sẽ được bổ sung bằng dữ liệu từ File 2 nếu có.
3. Ghi đè các ô rỗng hoặc giá trị null trong File 1 bằng giá trị tương ứng ở File 2.
4. Thêm các record còn thiếu (có trong File 2 nhưng không có trong File 1).

Sau khi xử lý, có thể tải về file kết quả.
"""
)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))


with st.sidebar:
    st.header("📤 Upload Files")
    f1 = st.file_uploader("File 1 (thiếu dữ liệu)", type=["xls", "xlsx"], key="file1")
    f2 = st.file_uploader("File 2 (đầy đủ dữ liệu)", type=["xls", "xlsx"], key="file2")

    merge_mode = st.selectbox(
        "Chế độ bổ sung",
        [
            "Chỉ điền vào ô trống ở File 1",
            "Ghi đè nếu khác (File 2 ưu tiên)",
        ],
    )

    add_missing_rows = st.checkbox("Thêm record mới từ File 2 vào File 1", value=True)

    st.markdown("---")
    st.caption("© 2025 Merge Tool")

if f1 and f2:
    try:
        df1 = load_excel(f1.getvalue())
        df2 = load_excel(f2.getvalue())
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        st.stop()

    if df1.empty or df2.empty:
        st.warning("Một trong hai file không có dữ liệu.")
        st.stop()

    st.subheader("👁️ Xem trước dữ liệu")
    with st.expander("File 1 - Thiếu dữ liệu", expanded=False):
        st.dataframe(df1.head(200))
    with st.expander("File 2 - Đầy đủ dữ liệu", expanded=False):
        st.dataframe(df2.head(200))

    common_columns = list(set(df1.columns).intersection(df2.columns))
    if not common_columns:
        st.error("Hai file không có cột chung nào để xử lý.")
        st.stop()

    st.subheader("🔑 Chọn cột khóa")
    key_cols = st.multiselect(
        "Chọn các cột dùng làm khóa (nên chọn đủ để xác định duy nhất 1 record)",
        options=common_columns,
    )

    if key_cols:
        if not all(col in df1.columns for col in key_cols) or not all(
            col in df2.columns for col in key_cols
        ):
            st.error("Cột khóa không tồn tại trong cả hai file.")
            st.stop()

        # Chuẩn hóa khóa (strip + lower cho text) để tăng khả năng khớp
        def normalize_key_cols(df: pd.DataFrame, cols):
            norm_df = df.copy()
            for c in cols:
                if pd.api.types.is_object_dtype(norm_df[c]):
                    norm_df[c] = norm_df[c].astype(str).str.strip().str.lower()
            return norm_df

        norm_df1 = normalize_key_cols(df1, key_cols)
        norm_df2 = normalize_key_cols(df2, key_cols)
        key_name = "__merge_key__"
        norm_df1[key_name] = norm_df1[key_cols].astype(str).agg("|".join, axis=1)
        norm_df2[key_name] = norm_df2[key_cols].astype(str).agg("|".join, axis=1)

        # Map từ key chuẩn hóa -> index gốc để cập nhật df1 sau khi merge
        key_to_index_df1 = norm_df1[key_name].to_dict()

        st.info(f"Số bản ghi File 1: {len(df1)} | File 2: {len(df2)}")

        # ================== MAPPING CỘT THỦ CÔNG ==================
        st.subheader("🧬 Mapping cột cập nhật / ghi đè")
        st.caption(
            "Bạn có thể chọn cột ở File 1 (đích) và xác định cột tương ứng ở File 2 (nguồn) để bổ sung hoặc ghi đè. Ví dụ: 'họ tên' (File 2) ➝ 'Họ và tên' (File 1). Nếu không bật, tool chỉ xử lý các cột trùng tên (loại trừ cột khóa). Có thể lưu / tải lại cấu hình mapping dưới dạng JSON."
        )
        manual_mapping_enabled = st.checkbox(
            "Bật tùy chỉnh mapping cột khác tên (File 2 ➝ File 1)", value=False
        )
        # Upload file mapping JSON (áp dụng trước khi render controls)
        mapping_upload = st.file_uploader(
            "Tải file mapping (.json) nếu có", type=["json"], key="mapping_upload"
        )

        # Các cột ứng viên (loại bỏ cột khóa)
        common_non_key_cols = [c for c in common_columns if c not in key_cols]
        dest_candidates = [c for c in df1.columns if c not in key_cols]
        source_candidates = [c for c in df2.columns if c not in key_cols]

        mapping: dict[str, str] = {}
        loaded_mapping = None
        if mapping_upload is not None:
            try:
                raw_txt = mapping_upload.getvalue().decode("utf-8")
                loaded_json = json.loads(raw_txt)
                if isinstance(loaded_json, list):
                    # Expect list of {"dest":..., "src":...}
                    loaded_mapping = {
                        item["dest"]: item["src"]
                        for item in loaded_json
                        if isinstance(item, dict) and "dest" in item and "src" in item
                    }
                elif isinstance(loaded_json, dict):
                    loaded_mapping = {str(k): str(v) for k, v in loaded_json.items()}
                else:
                    raise ValueError(
                        "Định dạng JSON không hợp lệ (chỉ hỗ trợ object hoặc list)."
                    )
                # Lọc hợp lệ
                loaded_mapping = {
                    d: s
                    for d, s in loaded_mapping.items()
                    if d in dest_candidates and s in source_candidates
                }
                if not loaded_mapping:
                    st.warning(
                        "File mapping JSON không có cặp hợp lệ với dữ liệu hiện tại."
                    )
                else:
                    # Set session state để auto chọn
                    st.session_state["dest_selected_cols"] = list(loaded_mapping.keys())
                    for d, s in loaded_mapping.items():
                        st.session_state[f"map_src_for_{d}"] = s
                    st.success(f"Đã nạp mapping JSON: {len(loaded_mapping)} cặp.")
            except Exception as e:
                st.error(f"Lỗi đọc mapping JSON: {e}")

        if manual_mapping_enabled:
            # Nếu đã upload mapping thì dùng nó làm default; nếu không thì dùng mặc định cột trùng tên có trong cả hai
            default_dest = (
                list(loaded_mapping.keys())
                if loaded_mapping
                else [c for c in dest_candidates if c in source_candidates]
            )
            dest_selected = st.multiselect(
                "Chọn các cột đích ở File 1 cần cập nhật (cột nhận dữ liệu)",
                options=dest_candidates,
                default=default_dest,
                key="dest_selected_cols",
            )
            if dest_selected:
                st.markdown("**Chọn cột nguồn tương ứng từ File 2 (ghi đè ➝ đích):**")
                for dest in dest_selected:
                    # Nếu có mapping load sẵn => ưu tiên
                    default_src = (
                        (loaded_mapping.get(dest) if loaded_mapping else None)
                        or (dest if dest in source_candidates else None)
                        or (source_candidates[0] if source_candidates else None)
                    )
                    if not source_candidates:
                        st.error("File 2 không có cột nào để mapping.")
                        break
                    src = st.selectbox(
                        f"Nguồn cho '{dest}'",
                        options=source_candidates,
                        index=(
                            source_candidates.index(default_src)
                            if default_src in source_candidates
                            else 0
                        ),
                        key=f"map_src_for_{dest}",
                    )
                    mapping[dest] = src
                if mapping:
                    mapping_preview = pd.DataFrame(
                        [
                            {"Cột File 1 (đích)": d, "Cột File 2 (nguồn)": s}
                            for d, s in mapping.items()
                        ]
                    )
                    st.dataframe(mapping_preview, use_container_width=True)
                    # Download mapping JSON
                    mapping_json_str = json.dumps(mapping, ensure_ascii=False, indent=2)
                    st.download_button(
                        "💾 Tải mapping JSON",
                        data=mapping_json_str.encode("utf-8"),
                        file_name="mapping_config.json",
                        mime="application/json",
                    )
            else:
                st.info("Chưa chọn cột nào để mapping.")
        else:
            mapping = {c: c for c in common_non_key_cols}
            # Cho phép tải mapping mặc định
            with st.expander("Xem / tải mapping mặc định (trùng tên)", expanded=False):
                if mapping:
                    mapping_preview = pd.DataFrame(
                        [
                            {"Cột File 1 (đích)": d, "Cột File 2 (nguồn)": s}
                            for d, s in mapping.items()
                        ]
                    )
                    st.dataframe(mapping_preview, use_container_width=True)
                    mapping_json_str = json.dumps(mapping, ensure_ascii=False, indent=2)
                    st.download_button(
                        "💾 Tải mapping mặc định JSON",
                        data=mapping_json_str.encode("utf-8"),
                        file_name="mapping_default.json",
                        mime="application/json",
                    )

        if st.button("🚀 Thực hiện merge", type="primary"):
            with st.spinner("Đang xử lý..."):
                df1_merged = df1.copy()

                # Thêm cột key vào bản gốc File 2 để tra cứu trực tiếp không bị lệch index
                df2_with_key = df2.copy()
                df2_with_key[key_name] = norm_df2[key_name]
                df2_indexed = df2_with_key.set_index(key_name)

                norm_df1_keys = norm_df1[key_name].tolist()
                existing_keys_set = set(norm_df1_keys)

                # Cảnh báo khóa trùng
                dup1 = norm_df1[norm_df1.duplicated(key_name, keep=False)][
                    key_name
                ].unique()
                dup2 = norm_df2[norm_df2.duplicated(key_name, keep=False)][
                    key_name
                ].unique()
                if len(dup1) > 0:
                    st.warning(
                        f"File 1 có {len(dup1)} khóa trùng (sẽ cập nhật tuần tự, bản ghi xuất hiện sau có thể ghi đè kết quả trước)."
                    )
                if len(dup2) > 0:
                    st.warning(
                        f"File 2 có {len(dup2)} khóa trùng (giữ bản ghi cuối cùng cho mỗi khóa)."
                    )
                    df2_indexed = df2_indexed[
                        ~df2_indexed.index.duplicated(keep="last")
                    ]

                updated_count = 0
                filled_cells = 0
                overwritten_cells = 0
                added_rows = 0
                if manual_mapping_enabled and not mapping:
                    st.error("Không có mapping hợp lệ để thực hiện cập nhật.")
                    st.stop()

                mapping_pairs = list(mapping.items())  # (dest_col, src_col)

                for row_idx, key in enumerate(norm_df1_keys):
                    if key in df2_indexed.index:
                        row2 = df2_indexed.loc[key]
                        any_updated = False
                        for dest_col, src_col in mapping_pairs:
                            if (
                                dest_col not in df1_merged.columns
                                or src_col not in row2.index
                            ):
                                continue
                            val1 = df1_merged.at[row_idx, dest_col]
                            val2 = row2[src_col]
                            if pd.isna(val2) or (
                                isinstance(val2, str) and val2.strip() == ""
                            ):
                                continue
                            if merge_mode == "Chỉ điền vào ô trống ở File 1":
                                if pd.isna(val1) or val1 == "":
                                    df1_merged.at[row_idx, dest_col] = val2
                                    filled_cells += 1
                                    any_updated = True
                            else:  # Ghi đè nếu khác
                                if pd.isna(val1) or val1 == "":
                                    df1_merged.at[row_idx, dest_col] = val2
                                    filled_cells += 1
                                    any_updated = True
                                elif val1 != val2:
                                    df1_merged.at[row_idx, dest_col] = val2
                                    overwritten_cells += 1
                                    any_updated = True
                        if any_updated:
                            updated_count += 1

                if add_missing_rows:
                    missing_keys = [
                        k for k in df2_indexed.index if k not in existing_keys_set
                    ]
                    if missing_keys:
                        rows_to_append = df2_indexed.loc[missing_keys].copy()
                        # Bỏ key kỹ thuật trước khi append
                        if key_name in rows_to_append.columns:
                            rows_to_append = rows_to_append.drop(columns=[key_name])
                        # Nếu mapping thủ công: đảm bảo các cột đích lấy dữ liệu từ nguồn tương ứng
                        if mapping:
                            for dest_col, src_col in mapping.items():
                                if (
                                    dest_col in df1_merged.columns
                                    and src_col in rows_to_append.columns
                                ):
                                    rows_to_append[dest_col] = rows_to_append[src_col]
                        df1_merged = pd.concat(
                            [df1_merged, rows_to_append], ignore_index=True
                        )
                        added_rows = len(rows_to_append)

                # Xuất file kết quả
                out_buffer = io.BytesIO()
                output_filename = (
                    f"merged_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
                with pd.ExcelWriter(out_buffer, engine="openpyxl") as writer:
                    df1_merged.to_excel(writer, index=False, sheet_name="Merged")

                st.success("Hoàn thành!")
                st.write(
                    f"Cập nhật từ File 2: {updated_count} record | Điền ô trống: {filled_cells} | Ghi đè: {overwritten_cells} | Thêm mới: {added_rows}"
                )
                st.download_button(
                    "⬇️ Tải file kết quả",
                    data=out_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        st.info("Hãy chọn ít nhất 1 cột khóa để tiếp tục.")
else:
    st.warning("Hãy upload cả 2 file Excel để bắt đầu.")

st.markdown("---")
st.caption("Developed with ❤️ using Streamlit & pandas")
