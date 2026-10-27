# Học liệu dự án Smart-Home

## 1. Bảng Excel xấu và đúng

Field là cột thể hiện một thuộc tính, ví dụ Ma_BDS, Loai, Gia_ban_VND. Record là một dòng thể hiện một bất động sản. Khóa Ma_BDS duy nhất giúp phân biệt các hồ sơ.

| Tiêu chí | Bảng xấu | Bảng đúng |
|---|---|---|
| Tiêu đề | Nhiều tầng, gộp ô | Một hàng, mỗi cột một tên riêng |
| Dòng dữ liệu | Trộn tiêu đề phụ, tổng cộng, dòng trống | Mỗi dòng một record, vùng liên tục |
| Ô dữ liệu | Một ô chứa “80m², 2 tỷ, căn hộ” | Diện tích, giá, loại tách thành field |
| Giá và tỷ lệ | Nhập “2 tỷ”, “giảm 5 phần trăm” dạng chữ | 2000000000 và 5% dạng số, dùng Number Format |
| Mã tài sản | Để trống hoặc trùng | Có mã ổn định, duy nhất |
| Loại tài sản | “Căn hộ”, “can ho”, “CH” tùy người nhập | Danh mục CH, NP, DAT bằng Data Validation |
| Công thức | Mỗi hàng nhập tay, vùng tra trôi | Công thức nhất quán, khóa vùng tra bằng $ |
| Lọc và sắp xếp | Chỉ sắp xếp cột giá, lệch địa chỉ | Dùng Excel Table để thao tác cả record |

Quy tắc vàng: không Merge Cell trong bảng dữ liệu; không chèn dòng hoặc cột trống chia cắt bảng; mỗi cột một kiểu dữ liệu và mỗi hàng một đối tượng. Ô trống cho thông tin không áp dụng vẫn được phép. Bọc toàn vùng bằng Ctrl+T để phạm vi lọc rõ ràng, kể cả khi có ô trống. Không nhập số 0 thay cho dữ liệu chưa biết.

Ví dụ xấu: một hàng ghi “Căn hộ”, hàng tiếp theo ghi “SH001 / 70m² / 2 tỷ”, rồi bỏ trống mã và loại ở các dòng sau vì “giống bên trên”. Ví dụ đúng: SH001 nằm ở Ma_BDS, CH ở Loai, 70 ở Dien_tich_m2 và 2000000000 ở Gia_ban_VND trên cùng một record.

## 2. Cách sử dụng Word

Tệp SRS dùng Title và Heading 1, 2, 3 thật. Heading 1 là chương, Heading 2 là nhóm yêu cầu, Heading 3 là nội dung chi tiết. View > Navigation Pane để điều hướng. Nếu cần mục lục tự động, đặt con trỏ đầu tài liệu, References > Table of Contents > Automatic Table. Sau khi sửa chọn Update entire table.

SRS bao gồm phạm vi, vai trò, yêu cầu chức năng, quy tắc đấu giá, dữ liệu, phi chức năng và tiêu chí nghiệm thu. Chỉ tiêu trong tài liệu là yêu cầu thiết kế, chưa phải kết quả kiểm thử phần mềm đã chạy.

## 3. Cấu trúc Excel

- DanhSach: 20 bất động sản giả định, 10 bán và 10 thuê. Cột A là khóa tài sản. Chữ xanh là nhập liệu, chữ đen là công thức.
- HoaHong: bảng A1:B7, trong đó A2:B7 là 6 dòng tra cứu. Mỗi khóa xuất hiện đúng một lần.
- KiemThu: so sánh công thức vùng tương đối với vùng tuyệt đối khi kéo xuống. IFNA trả thông báo khi không có mã.

Tất cả địa chỉ, giá, chiết khấu và tỷ lệ hoa hồng chỉ phục vụ thực hành. Các cột giá bán hoặc thuê không áp dụng được để trống. Đơn vị tiền là VND, diện tích là m², thời hạn thuê tính bằng tháng.

| Cột | Ý nghĩa |
|---|---|
| C, D, E | Loại tài sản, giao dịch và khóa tra ghép |
| G, H, I | Giá bán, thuê/tháng và chiết khấu nhập liệu |
| J, K | Giá bán và thuê/tháng sau chiết khấu |
| L, M | Số tháng hợp đồng và tổng tiền thuê |
| N, O, P | Tỷ lệ môi giới, cơ sở tính và tiền hoa hồng |

## 4. Công thức và lý do sử dụng

Các công thức dưới đây nhập ở dòng 2 rồi kéo tới dòng 21. Nếu Excel của bạn dùng dấu chấm phẩy, thay dấu phẩy ngăn cách đối số bằng dấu chấm phẩy.

### 4.1 TRIM và ghép khóa

```excel
=TRIM(C2)&"_"&TRIM(D2)
```

TRIM bỏ khoảng trắng thừa ở đầu và cuối, đồng thời rút gọn chuỗi khoảng trắng thông thường ở giữa. Dấu & nối chuỗi. CH và BAN tạo khóa CH_BAN. Ghép cả loại và giao dịch để không dùng nhầm hoa hồng thuê cho giao dịch bán. TRIM không đảm bảo loại mọi ký tự khoảng trắng đặc biệt từ web, cần kiểm tra nguồn nếu vẫn không khớp.

### 4.2 IF và ROUND tính giá sau chiết khấu

```excel
=IF(D2="BAN",ROUND(G2*(1-I2),0),"")
=IF(D2="THUE",ROUND(H2*(1-I2),0),"")
```

IF chỉ tính cột áp dụng cho loại giao dịch, nếu không thì hiển thị trống. 1 − tỷ lệ chiết khấu là phần giá còn phải trả. ROUND(...,0) làm tròn đến đồng. Ví dụ giá 2 tỷ, giảm 5% thì còn 1,9 tỷ. Các giá này là giá đề xuất trong danh mục; không tự giảm giá trúng phiên đấu giá.

### 4.3 Tổng tiền thuê

```excel
=IF(D2="THUE",K2*L2,"")
```

Lấy giá thuê sau giảm mỗi tháng nhân số tháng. Ví dụ 10 triệu/tháng giảm 5%, thuê 12 tháng: 9,5 triệu × 12 = 114 triệu. Chưa bao gồm thuế, tiền cọc, phí dịch vụ và điều chỉnh giá giữa kỳ.

### 4.4 VLOOKUP tra tỷ lệ hoa hồng

```excel
=IFNA(VLOOKUP(E2,'HoaHong'!$A$2:$B$7,2,FALSE),"Thieu ma")
```

VLOOKUP tìm E2 ở cột đầu tiên của A2:B7 và trả giá trị ở cột thứ 2. FALSE bắt buộc khớp chính xác, phù hợp tra mã. IFNA chỉ xử lý trường hợp không tìm thấy, không giấu mọi lỗi bằng 0. Bảng khóa phải duy nhất vì VLOOKUP chỉ trả kết quả khớp đầu tiên. Cột tỷ lệ của bảng tra phải có số hợp lệ.

### 4.5 Cơ sở và tiền hoa hồng

```excel
=IF(D2="BAN",J2,M2)
=IF(ISNUMBER(N2),ROUND(O2*N2,0),"Kiem tra ma")
```

O là cơ sở tính: giá bán sau giảm hoặc tổng thuê cả hợp đồng. ISNUMBER ngăn thiếu tỷ lệ trở thành số hoa hồng tưởng như hợp lệ. SH001 có giá 2 tỷ, giảm 5%, tỷ lệ CH_BAN là 2% nên hoa hồng bằng 1,9 tỷ × 2% = 38 triệu. Với thuê, quy ước bài tập tính tỷ lệ trên cả hợp đồng; không phải mức phí môi giới ngoài thị trường.

## 5. Kiểm thử lỗi trôi vùng và dấu $

Công thức thiếu khóa:

```excel
=VLOOKUP(E2,'HoaHong'!A2:B7,2,FALSE)
```

Khi kéo xuống dòng 3, công thức thành:

```excel
=VLOOKUP(E3,'HoaHong'!A3:B8,2,FALSE)
```

Mã CH_BAN nằm ở A2 đã rơi ra ngoài vùng tra. Kết quả có thể là #N/A hoặc sai khi bảng có dữ liệu khác. Sửa thành:

```excel
=VLOOKUP(E2,'HoaHong'!$A$2:$B$7,2,FALSE)
```

Kéo xuống thì E2 đổi thành E3 vì cần đọc record kế tiếp, còn $A$2:$B$7 giữ nguyên vì luôn dùng chung bảng hoa hồng.

| Tham chiếu | Khóa gì | Khi kéo xuống 1 hàng | Khi kéo phải 1 cột |
|---|---|---|---|
| A2 | Không khóa | A3 | B2 |
| $A$2 | Cột A và hàng 2 | $A$2 | $A$2 |
| $A2 | Chỉ cột A | $A3 | $A2 |
| A$2 | Chỉ hàng 2 | A$2 | B$2 |

Chọn phần tham chiếu trong thanh công thức rồi nhấn F4, một số laptop cần Fn+F4. Dấu $ giữ vị trí khi sao chép, không tự mở rộng vùng khi thêm khóa mới. Nếu thêm dòng hoa hồng thứ 7, cần mở rộng $A$2:$B$7 thành vùng mới hoặc chuyển sang tham chiếu Table phù hợp.

### Bảng kiểm thử thao tác

| Ca | Thao tác | Mong đợi |
|---|---|---|
| 1 | Xem KiemThu!C2 và C3 | C2 = 2%, C3 = Khong tim thay vì mất A2 |
| 2 | Xem KiemThu!D2:D7 | Tất cả 2% vì khóa vùng |
| 3 | Xem dòng 8 của KiemThu | Mã không có trả Khong tim thay, không phải 0% |
| 4 | Đổi DanhSach!I2 thành 10% | J2 = 1,8 tỷ; P2 = 36 triệu |
| 5 | Đổi DanhSach!C2 thành XX | N2 = Thieu ma; P2 = Kiem tra ma |
| 6 | Khôi phục C2 = CH và I2 = 5% | P2 trở về 38 triệu |
| 7 | Sắp xếp cả Table theo giá | Mã, địa chỉ và công thức vẫn cùng record |
| 8 | Tạo mã Ma_BDS trùng | Conditional Formatting làm nổi bật mã trùng |

Không sửa công thức để ép ra kết quả đúng. Kiểm tra mã đầu vào, vùng tra, kiểu dữ liệu và tham chiếu trước. Các ca kiểm thử phần mềm đấu giá nằm ở chương 6 SRS, chỉ là ca nghiệm thu dự kiến.

## 6. Pitch Deck và Rule 10/20/30

Deck có đúng 10 slide kể cả trang mở đầu, tổng thời lượng dự kiến 20 phút và chữ hiển thị nhỏ nhất 30 pt. Mỗi slide có một ý chính, chữ hỗ trợ ngắn và khoảng trống rộng. Diễn giải đặt trong Speaker Notes để tránh dồn lên slide. Không có bằng chứng doanh thu hay người dùng thật; mục tiêu, ngân sách và mô hình phí đều ghi là giả định.

| Slide | Thông điệp | Phút |
|---|---|---:|
| 1 | Smart-Home đề xuất nền tảng đấu giá trực tuyến | 1 |
| 2 | Cần kiểm chứng nhu cầu về hồ sơ và mức giá đáng tin | 2 |
| 3 | Pilot bắt đầu ở một nhóm tài sản tại TP.HCM | 2 |
| 4 | Hồ sơ tập trung và lịch sử trả giá là giá trị sản phẩm | 2 |
| 5 | Máy chủ quyết định giá hợp lệ | 2 |
| 6 | Ví dụ doanh thu phí nền tảng 5 triệu/giao dịch | 2 |
| 7 | Mục tiêu pilot có 10 chủ tài sản | 2 |
| 8 | Sáu tháng đi từ MVP tới pilot | 2 |
| 9 | Chỉ mở rộng sau khi chứng minh khả năng vận hành | 2 |
| 10 | Đề xuất ngân sách giả định 600 triệu | 3 |
| Tổng | | 20 |

Phí nền tảng giả định: 2 tỷ × 0,25% = 5 triệu đồng/giao dịch. Nếu đạt 3 giao dịch thì doanh thu minh họa là 15 triệu, chưa phải lợi nhuận và chưa bù ngân sách 600 triệu. Ngân sách 360 + 150 + 90 = 600 triệu. Không dùng tỷ lệ hoa hồng môi giới Excel để gọi thành doanh thu nền tảng.

## 7. Bài tập thực hành

1. Thêm một hồ sơ CH, giao dịch THUE, giá 8 triệu/tháng, giảm 5%, thuê 12 tháng. Tổng thuê đúng là 91,2 triệu và hoa hồng theo bảng mẫu là 4,56 triệu.
2. Chứng minh đổi A2:B7 thành $A$2:$B$7 giúp kéo xuống 20 dòng an toàn. Giải thích vì sao không khóa E2.
3. Tạo mục 2.3.1 trong Word bằng Heading 3 rồi kiểm tra Navigation Pane.
4. Đổi chiết khấu SH001 về 0%. Giá sau giảm phải bằng 2 tỷ và hoa hồng 40 triệu.
5. Trình bày slide 6 trong 2 phút, phân biệt doanh thu, hoa hồng môi giới và lợi nhuận.

## 8. Tự kiểm trước khi nộp

- Word có Heading 1, 2, 3, đủ nội dung SRS và điều hướng được.
- Excel không Merge Cell, dữ liệu dạng Table, một hàng một record và một cột một field.
- Tỷ lệ là số %, giá là số VND, không ghép số với chữ trong ô dữ liệu.
- VLOOKUP dùng FALSE và khóa bảng tra bằng $; thử đầu, giữa và cuối vùng.
- PowerPoint đủ 10 slide, một slide một thông điệp, có khoảng trắng và chữ từ 30 pt.
- Các dữ liệu mẫu và mục tiêu kinh doanh được xác định là giả định.

Nguồn nội dung: đề bài Smart-Home do người học cung cấp. Hình minh họa liên kết trong đề không truy cập được khi soạn; học liệu dựa trên yêu cầu văn bản. Bộ này là học liệu và đặc tả đề xuất, không chứa phần mềm nền tảng đã triển khai.
