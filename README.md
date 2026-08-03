# Quản lý Kho NVL & Thành phẩm

App quản lý xuất nhập tồn kho — Danh mục gốc → Nhập hàng → Xuất hàng → Báo cáo → Tồn kho.
React + Vite + Tailwind, dữ liệu lưu trên Supabase (Postgres).

## 1. Tạo dự án Supabase

1. Vào [supabase.com](https://supabase.com) → **New project** (miễn phí)
2. Vào **SQL Editor** → **New query** → dán toàn bộ nội dung file `supabase/schema.sql` → **Run**
   → Tạo đủ các bảng, các hàm bảo mật đăng nhập (băm mật khẩu bcrypt), và 2 tài khoản mẫu:

   | Tên đăng nhập | Mật khẩu | Vai trò |
   |---|---|---|
   | `nvkho1` | `123456` | Nhân viên kho |
   | `quanly1` | `123456` | Quản lý |

   **Đổi mật khẩu 2 tài khoản này ngay** sau khi triển khai thật (đăng nhập → bấm icon khoá ở góc trên để tự đổi, hoặc dùng tài khoản Quản lý vào mục "Tài khoản" để đặt lại).

3. Lấy thông tin kết nối: **Project Settings → API** (hoặc **Data API**), copy **Project URL** và **anon/publishable key**.

## 2. Cấu hình biến môi trường & Deploy qua Vercel

1. Đẩy toàn bộ code này lên 1 repo GitHub mới
2. Vào [vercel.com](https://vercel.com) → **Add New Project** → chọn repo vừa tạo
3. Ở phần **Environment Variables**, thêm:
   - `VITE_SUPABASE_URL` = Project URL
   - `VITE_SUPABASE_ANON_KEY` = anon/publishable key
4. Bấm **Deploy**

## 3. Cách dùng — theo đúng 6 tác vụ đã thống nhất

1. **Danh mục**: thêm Nhà cung cấp / NVL & Thành phẩm / Mã doanh thu / Mã xuất. Mã tự động gợi ý nhảy tiếp theo mã lớn nhất hiện có (bạn có thể sửa lại nếu muốn).
2. **Nhập hàng**: chọn NCC + Sản phẩm + Số lượng + Đơn giá → các trường còn lại (ngày, mã phiếu, tên NCC/SP, đơn vị, thành tiền, tình trạng thanh toán) tự động điền.
3. **Báo cáo nhập**: lọc theo ngày / NCC / nhóm sản phẩm / tình trạng thanh toán.
4. **Xuất hàng**: 1 phiếu xuất dùng chung Đơn số + Mã doanh thu + Mã xuất, bên trong thêm nhiều dòng NVL (đơn giá tự tính theo bình quân gia quyền) và/hoặc TP (tự nhập đơn giá bán).
5. **Báo cáo xuất**: theo Mã doanh thu và theo Mã xuất — tự tính giá vốn (từ dòng NVL), doanh thu (từ dòng TP), tỉ trọng doanh thu, lợi nhuận, tỉ suất lợi nhuận.
6. **Tồn kho**: báo cáo Tồn đầu — Nhập — Xuất — Tồn cuối theo từng sản phẩm trong khoảng ngày bất kỳ. Quản lý có thể chốt/điều chỉnh mốc tồn đầu khi kiểm kho thực tế.

## 4. Cách tính giá NVL khi xuất kho (bình quân gia quyền)

```
Giá bình quân = (Giá trị tồn đầu kỳ + Giá trị nhập từ mốc tồn đầu tới nay)
              ÷ (Số lượng tồn đầu kỳ + Số lượng nhập từ mốc tồn đầu tới nay)
```

Đây là **bình quân gia quyền theo kỳ** (tính lại dựa trên mốc tồn đầu gần nhất bạn chốt + toàn bộ nhập vào sau đó) — cách này đơn giản, minh bạch, và là cách phổ biến nhất trong thực tế kinh doanh nhỏ. Nếu bạn cần cách tính khác (bình quân di động theo từng giao dịch...), báo lại để điều chỉnh.

## 5. Lưu ý về bảo mật

- Mật khẩu được **băm bcrypt** ngay từ đầu, không lưu dạng đọc được
- Tài khoản mới tạo hoặc bị reset đều **bắt buộc đổi mật khẩu trong 24h**, quá hạn tự khoá
- Row Level Security đang ở chế độ mở (`allow all`) cho các bảng nghiệp vụ — phù hợp công cụ nội bộ ít người dùng, nhưng **không phù hợp nếu công khai link ra ngoài Internet**. Nếu cần khoá chặt hơn theo từng vai trò, báo lại để nâng cấp.

## 6. Các thư mục chính

```
├── supabase/schema.sql   → chạy 1 lần trong Supabase SQL Editor
├── src/
│   ├── App.jsx            → toàn bộ giao diện & logic nghiệp vụ
│   ├── supabaseClient.js  → kết nối Supabase
│   ├── main.jsx
│   └── index.css
├── package.json
└── vite.config.js
```
