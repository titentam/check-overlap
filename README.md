# Excel Merge & Enrichment Tool

Ứng dụng Streamlit giúp hợp nhất 2 file Excel cùng cấu trúc header theo các cột khóa để bổ sung dữ liệu còn thiếu.

## Tính năng

- Upload 2 file Excel
- Chọn cột khóa (1 hoặc nhiều cột)
- Hai chế độ merge:
  1. Chỉ điền vào ô trống ở File 1
  2. Ghi đè nếu khác (File 2 ưu tiên)
- Thêm record mới có trong File 2 nhưng không có trong File 1
- Thống kê số ô được điền, số ô bị ghi đè, số bản ghi thêm mới
- Tải về file Excel kết quả

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy ứng dụng

```bash
streamlit run app.py
```

## Logic xử lý (tóm tắt)

1. Chuẩn hóa giá trị khóa (strip + lower cho cột dạng text)
2. Tạo key ghép bằng ký tự `|`
3. Với mỗi record trùng key giữa File 1 và File 2:
   - Nếu chế độ chỉ điền ô trống: chỉ cập nhật ô rỗng/null
   - Nếu chế độ ghi đè: cập nhật cả ô trống và ô khác giá trị
4. Thêm các record còn thiếu nếu bật tùy chọn
5. Xuất ra Excel

## Ghi chú

- Chỉ xử lý trên các cột chung giữa 2 file.
- Không thay đổi giá trị cột khóa.
- Có thể mở rộng để cấu hình thêm rule tùy biến.

---

Developed with ❤️ using Streamlit & pandas
