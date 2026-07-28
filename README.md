
# YVF Adoption Dashboard – CS HAD

Dashboard Streamlit gồm 5 trang:

1. Overview  
2. Customer Adoption  
3. Booking Details  
4. User Issues  
5. Enhancement Requests  

Dashboard không sử dụng logo công ty để tránh nhầm với dashboard YVF chính thức.

## 1. Cấu trúc file

```text
YVF_Adoption_Dashboard_GitHub/
├── app.py
├── requirements.txt
├── README.md
└── YVF_Adoption_Dashboard_CS_HAD.xlsx
```

## 2. Chạy trên máy tính

Mở Terminal tại thư mục project và chạy:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sau đó trình duyệt sẽ tự mở dashboard.

## 3. Upload lên GitHub

1. Tạo một Repository mới trên GitHub.
2. Chọn **Add file → Upload files**.
3. Upload toàn bộ 4 file trong thư mục project.
4. Chọn **Commit changes**.

## 4. Deploy trên Streamlit Community Cloud

1. Truy cập Streamlit Community Cloud.
2. Chọn **Create app**.
3. Chọn GitHub Repository vừa tạo.
4. Main file path nhập:

```text
app.py
```

5. Chọn **Deploy**.

## 5. File Excel cần có các sheet

- `Data_Booking`
- `Data_SI`
- `Data_CustomerActive`
- `Data_Issue`
- `Customer_Feedback`

Dashboard sẽ tự đọc file `YVF_Adoption_Dashboard_CS_HAD.xlsx` đặt cùng thư mục với `app.py`.

Người dùng cũng có thể upload một file Excel mới trực tiếp từ thanh bên trái của dashboard.
