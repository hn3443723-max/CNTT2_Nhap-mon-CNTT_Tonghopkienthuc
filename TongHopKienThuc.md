**Họ tên:Nguyễn Đức Hiếu**

**Lớp:CNTT2**

# PHẦN TỔNG HỢP KIẾN THỨC QUAN TRỌNG 
## PHẦN 1: CÁC LỆNH TERMINAL
### Cấu trúc một câu lệnh 
**Command Options/Flags Argument**
- Command(Lệnh chính):là một động từ
- Options/Flags(Tùy chọn):là một tính từ
- Argument(Đầu vào):là một danh từ

### Các lệnh cơ bản và nâng cao 
| Công dụng | Windows PowerShell | MacOS/Linux |
| :--- | :--- | :--- |
| Xem vị trí hiện tại | pwd | pwd |
| Liệt kê nội dung | ls | ls |
| Liệt kê file và thư mục | Get-ChildItem -Recurse -Force | ls -aR |
| Chuyển thư mục | cd folder | cd folder |
| Lùi về thư mục cha | cd .. | cd .. |
| Xoá hết các lệnh | cls | cls |
| Hiển thị nội dung tệp | Get-Content | Cat |
| Hiển thị số dòng đầu tiên | Get-Content file.log -Head | head -n |
| Hiển thị số dòng cuối cùng |  Get-Content file.log -Tail | tail -n | 
| Tạo thư mục | mkdir | mkdir |
| Tạo tệp | ni | touch | 
| Sao chép tệp | cp | cp |
| Di chuyển | mv | mv |
| Xóa tệp | rm | rm |
| In nội dung ra màn hình | Write-Output | Echo |
| Lọc dòng chứa từ khoá | Select-String | Grep |
| Đếm số dòng | Measure-Object | wc -l |
| Ghi đè | > | > |
| Ghi nối tiếp | >> | >> |
| Chạy script đặt tại thư mục hiện tại | .\file.ps1 | bash file.sh |
| Tìm tệp | Get-ChildItem -Recurse | file .-name |
| Đọc file lưu lượng lớn | Get-Content file.log | less file.log |
## CÁC ĐỊNH DẠNG TRONG MARKDOWN 
| Công dụng | Định dạng Markdown |
| :--- | :--- |
| Tiêu đề | Dùng ký tự # từ H1 đến H6 kèm 1 khoảng trắng |
| In đậm | Dùng kí tự "**" |
| In nghiêng | Dùng kí tự "*" |
| Danh sách có thứ tự | 1.;2. |
| Danh sách không có thứ tự | - |
| Cú pháp siêu liên kết | [Chữ hiển thị](Đường dẫn URL) |
| Chèn hình ảnh | ![Mô tả ảnh](đường link ảnh) |
| Trích dẫn | > |
| Ngắt trang | --- |
| Code block | ```bash |
| Bọc tên biến, tên hàm, câu lệnh ngắn | `inline code` |
| Căn giữa | :---: |
| Căn trái | :--- |
| Căn phải | ---: | 
| In ký tự đặc biệt ra màn hình | \* |
| Task list | - [ ] NỘI DUNG |

### PHẦN 3: Tin học văn phòng 

**Word**

#### 1.1. Heading Styles  

- Cách 1: Dùng Phím Tắt Nhanh (Khuyên dùng): 

Ctrl + Alt + 1: Gán nhãn tiêu đề lớn cấp 1 (Heading 1 / Level 1) 

Ctrl + Alt + 2: Gán nhãn tiểu mục con cấp 2 (Heading 2 / Level 2) 

Ctrl + Shift + N: Hủy nhãn, đưa về chữ thường (Normal) 

- Cách 2: Chọn Trên Thanh Công Cụ (Menu): 

Tab Home → Nhấp vào ô Heading 1 hoặc Heading 2 trên thanh Styles Gallery. 

Hoặc Tab References → Nhấp Add Text → Chọn Level 1, Level 2. 

#### 1.2. Mục lục tự động 

- Chèn Mục Lục Tự Động (Auto TOC - Table of Contents) 

Đưa con trỏ chuột lên trang đầu tiên của tài liệu (nơi muốn đặt mục lục), vào tab References → Nhấp chọn Table of Contents → Chọn mẫu mục lục tự động (Automatic Table 1 hoặc 2). Bộ máy Word sẽ tự động quét toàn bộ tiêu đề và sinh ra bảng mục lục thụt dòng hoàn chỉnh kèm số trang chính xác. 

**Excel**

#### 2.1. AutoFilter 

- Bật Bộ Lọc AutoFilter Cho Toàn Bảng: 

Nhấp chuột vào bất kỳ ô nào bên trong bảng dữ liệu (ví dụ ô A1), sau đó nhấn phím tắt Ctrl + Shift + L (trên macOS là Cmd + Shift + F) hoặc vào tab Data [1] → Nhấp nút Filter [2]. 

 
#### 2.2. Cú pháp và cơ chế vận hành hàm VLOOKUP 

- Nhập Công Thức Tra Cứu VLOOKUP: =VLOOKUP(Lookup_value, table_array, Col_index_num, Range_lookup)

Bước 1: Lookup_value 

Bước 2: Table_array 

Bước 3: Col_index_num  

Bước 4: Range_lookup  

**PowerPoint**

- Quy tắc thiết kế bài trình chiếu: 

Tối Đa 10 Slide: Não người không tiếp nhận quá 10 ý niệm mới trong 1 buổi. Hãy cô đọng nội dung. 

Thời Lượng 20 Phút: Ngưỡng tập trung cao nhất. 10 slide trong 20 phút = mỗi slide có đúng 2 phút để nói 

Cỡ Chữ Tối Thiểu 30pt: Nếu phải giảm font xuống 14pt để nhét hết chữ, đó là do bạn đang nhét quá nhiều chữ. 

- Biểu đồ dữ liệu: 

Thay vì dán nguyên bảng Excel 50 ô số lên slide, hãy vào tab Insert [1] → Nhấp chọn Chart [2] → Chọn biểu đồ cột để mắt người xem nhận biết ngay hạng mục dẫn đầu chỉ trong 0.3 giây. 

 
