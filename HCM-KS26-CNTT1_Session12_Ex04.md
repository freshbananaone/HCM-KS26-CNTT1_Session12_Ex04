# Bài 4: Chuẩn hóa quy trình quản lý bệnh án điện tử

Clinic-Cloud quản lý tập trung 1.000 hồ sơ, hỗ trợ khám bệnh và kiểm soát quyền truy cập thông tin bệnh nhân.

## Phần 1. Phân tích: Chức năng và phi chức năng

**Yêu cầu chức năng** mô tả hệ thống làm gì: tạo hồ sơ, ghi nhận lần khám, lưu chẩn đoán và đơn thuốc. **Yêu cầu phi chức năng** quy định điều kiện vận hành: bảo mật, toàn vẹn và khả dụng của dữ liệu.

Ví dụ, “bác sĩ xem lịch sử khám” là chức năng; “chỉ xem hồ sơ được phân công” là ràng buộc bảo mật của chức năng đó.

## Phần 2. Triển khai: Hệ thống Clinic-Cloud tích hợp

### 1. SRS rút gọn cho module Bệnh án điện tử

**Phạm vi:** quản lý bệnh nhân, lần khám và báo cáo tổng hợp. Người dùng gồm lễ tân, bác sĩ, quản trị hệ thống và người phụ trách báo cáo.

| Mã | Yêu cầu chức năng |
|---|---|
| FR-01 | Tạo và cập nhật hồ sơ; mỗi bệnh nhân có một mã duy nhất, không để trống hoặc tái sử dụng cho người khác. |
| FR-02 | Lưu từng lần khám gồm ngày khám, bác sĩ, triệu chứng, chẩn đoán và điều trị; liên kết về đúng bệnh nhân. |
| FR-03 | Tìm hồ sơ theo mã hoặc thông tin định danh; xem lịch sử khám và lịch sử chỉnh sửa. |
| FR-04 | Nhập Excel sau khi kiểm tra dữ liệu; báo rõ dòng lỗi. Xuất dữ liệu theo quyền và tổng hợp số ca bệnh theo kỳ. |

Các ngưỡng sau là **yêu cầu thiết kế đề xuất**, dùng làm tiêu chí nghiệm thu:

| Mã | Yêu cầu phi chức năng về bảo mật và vận hành |
|---|---|
| NFR-01 | Mỗi nhân viên dùng tài khoản riêng và xác thực đa yếu tố; phiên tự hết hạn sau 15 phút không hoạt động. |
| NFR-02 | Phân quyền tại máy chủ: lễ tân quản lý thông tin hành chính; bác sĩ truy cập hồ sơ được phân công; người làm báo cáo chỉ xem số liệu tổng hợp. Quản trị kỹ thuật không mặc nhiên có quyền đọc bệnh án. |
| NFR-03 | Mã hóa đường truyền bằng TLS 1.2 trở lên; mã hóa dữ liệu lưu trữ và bản sao lưu bằng AES-256; quản lý khóa tách khỏi dữ liệu. |
| NFR-04 | Ghi nhật ký xem, sửa, xuất và truy cập bị từ chối: ai, hồ sơ nào, lúc nào, thao tác và kết quả. Chống sửa/xóa log trái phép; không ghi mật khẩu hoặc toàn văn bệnh án. |
| NFR-05 | Đính chính hồ sơ đã chốt phải lưu phiên bản, người sửa và lý do. Thời hạn lưu và quyền hủy theo chính sách được phê duyệt, phù hợp quy định áp dụng. |
| NFR-06 | Sao lưu hằng ngày; kiểm tra phục hồi mỗi quý. Mục tiêu mất dữ liệu tối đa 24 giờ và khôi phục dịch vụ trong 4 giờ. |

Cơ chế phân quyền và nhật ký tham khảo hướng dẫn [OWASP về kiểm soát truy cập](https://cheatsheetseries.owasp.org/cheatsheets/Authorization_Cheat_Sheet.html) và [ghi log an toàn](https://cheatsheetseries.owasp.org/cheatsheets/Logging_Cheat_Sheet.html).

Luồng tích hợp: **Tiếp nhận → khám và cập nhật → lưu tập trung → báo cáo**. Clinic-Cloud lưu bản chính; Excel phục vụ nhập/xuất và rà soát theo quyền. File xuất cần giới hạn truy cập và mã hóa; Protect Sheet chỉ hỗ trợ chống sửa nhầm.

### 2. Excel: quản lý 1.000 hồ sơ theo Database Mindset

Sheet `BenhNhan` có tám cột A:H: **MaBN, HoTen, NgaySinh, GioiTinh, DienThoai, DiaChi, TrangThaiThongTin, NgayTao**. Mỗi dòng là một bệnh nhân; mã và điện thoại dùng kiểu Text, ngày dùng kiểu Date. Chỉ có một hàng tiêu đề, không gộp ô, không chèn dòng tổng hoặc hàng/cột trống để phân nhóm.

Tách sheet `LuotKham`: mỗi lần khám có `MaLuotKham` riêng, liên kết bằng `MaBN`. Không tạo các cột “Lần khám 1”, “Lần khám 2”… trong bảng bệnh nhân.

**Xử lý bẫy ô trống:** ô thiếu thông tin riêng lẻ không làm AutoFilter ngừng hoạt động. Vấn đề thường nằm ở việc Excel tự nhận vùng dữ liệu bị giới hạn bởi hàng/cột trống hoàn toàn.

1. Chọn chính xác `A1:H1001` bằng Name Box, bao gồm tiêu đề và đủ 1.000 hồ sơ.
2. Dùng **Insert → Table** hoặc **Ctrl + T** trên Windows, chọn **My table has headers**, đặt tên `tblBenhNhan`. Bộ lọc sẽ áp dụng trên toàn bảng. [Hướng dẫn Microsoft](https://support.microsoft.com/en-us/excel/get-started/filter-data-in-a-range-or-table-in-excel?referral=LP).
3. Giữ ô trống nếu chưa biết giá trị, đánh dấu “Cần bổ sung” tại `TrangThaiThongTin`. Không điền 0 hoặc ngày giả để lấp ô. `MaBN` bắt buộc đầy đủ và không trùng.
4. Lọc **(Blanks)** để rà soát thông tin thiếu. Kiểm tra Table mở rộng khi thêm hồ sơ; không xóa bệnh nhân chỉ vì thiếu thông tin tùy chọn.

### 3. Donut Chart: ít chi tiết, rõ thông điệp

Số liệu minh họa trong kỳ: **100 ca** gồm cúm 50, sốt xuất huyết 25, tay chân miệng 15 và bệnh khác 10. Mỗi ca thuộc một nhóm theo chẩn đoán chính; ca chưa xác định được báo riêng. Biểu đồ phản ánh cơ cấu ca ghi nhận tại phòng khám.

Tiêu đề: **“Cúm chiếm 50% số ca bệnh truyền nhiễm đã phân loại”**. Dùng Donut Chart bốn phần, lỗ giữa 65%, ghi “100 ca” ở tâm. Tô cúm màu xanh, phần còn lại dùng sắc xám; gắn tên nhóm và tỷ lệ trực tiếp. Bỏ 3D, viền và chú giải dư thừa; chỉ dùng dữ liệu tổng hợp.

## Phần 3. Kiểm thử

### 1. Sửa đoạn văn bị gán nhầm Heading Styles

Giả sử đoạn mô tả FR-02 bị gán **Heading 2**, khiến nó xuất hiện nhầm trong mục lục:

1. Đặt con trỏ trong đoạn bị nhầm; nếu nhiều đoạn, chọn các đoạn cần sửa.
2. Nhấn **Ctrl + Shift + N** trong Word trên Windows để áp dụng **Normal**, đưa đoạn về văn bản thường mà vẫn giữ nội dung. Trên Mac dùng **Command + Shift + N**. [Phím tắt Word — Microsoft](https://support.microsoft.com/en-us/accessibility/word/keyboard-shortcuts-in-word).
3. Kiểm tra Styles là Normal, đoạn biến mất khỏi Navigation Pane; cập nhật toàn bộ mục lục. Các tiêu đề thật vẫn giữ Heading 1/2. Nếu còn định dạng chữ thủ công không mong muốn, dùng **Ctrl + Space** trên Windows cho phần đó.

### 2. Các tình huống nghiệm thu bổ sung

| Tình huống | Kết quả mong đợi |
|---|---|
| Có nhiều ô thiếu thông tin, hồ sơ cuối nằm ở dòng 1001 | Lọc được hồ sơ cuối; bỏ lọc vẫn có đủ 1.000 mã bệnh nhân. |
| Dữ liệu thử có 200 bệnh nhân chưa có điện thoại | Lọc (Blanks) tại cột điện thoại trả đúng 200 hồ sơ. |
| Lễ tân yêu cầu đọc nội dung bệnh án, kể cả truy cập trực tiếp qua API | Máy chủ từ chối và ghi nhật ký. |
| Phục hồi dữ liệu từ bản sao lưu | Đáp ứng giới hạn mất dữ liệu 24 giờ và thời gian phục hồi 4 giờ. |
| Kiểm tra nguồn Donut Chart | Tổng nhóm bằng 100 ca và 100%; không đếm trùng ca. |

Các tình huống trên là kế hoạch kiểm thử; kết quả thực tế cần ghi nhận khi triển khai Clinic-Cloud.
