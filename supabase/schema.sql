-- =============================================================================
-- QUẢN LÝ KHO NVL & THÀNH PHẨM — Schema Supabase
-- Chạy toàn bộ file này trong Supabase Dashboard → SQL Editor → New query → Run.
-- =============================================================================

create extension if not exists pgcrypto;

-- -----------------------------------------------------------------------------
-- TÀI KHOẢN ĐĂNG NHẬP — mật khẩu băm bcrypt ngay từ đầu (không lưu dạng thường)
-- 2 vai trò: nhan_vien_kho (nhập/xuất hàng ngày), quan_ly (báo cáo + sửa danh mục)
-- -----------------------------------------------------------------------------
create table employees (
  id uuid primary key default gen_random_uuid(),
  username text unique not null,
  password text not null,
  name text not null,
  role text not null check (role in ('nhan_vien_kho','quan_ly')),
  must_change_password boolean not null default false,
  password_change_deadline timestamptz,
  created_at timestamptz default now()
);

-- -----------------------------------------------------------------------------
-- DANH MỤC GỐC — có thể thêm mới bất cứ lúc nào
-- -----------------------------------------------------------------------------

-- Nhà cung cấp. payment_type quyết định "Tình trạng thanh toán" tự động khi
-- nhập hàng: tien_mat (thanh toán ngay), cong_no (ghi nợ, NCC thật), noi_bo
-- (điều chuyển nội bộ, không phát sinh tiền).
create table suppliers (
  id uuid primary key default gen_random_uuid(),
  code text unique not null,
  name text not null,
  payment_type text not null default 'cong_no' check (payment_type in ('tien_mat','cong_no','noi_bo')),
  created_at timestamptz default now()
);

-- Mã doanh thu (điểm tiêu thụ: phòng VIP, hội trường...) — dùng để gắn vào
-- phiếu xuất, phục vụ báo cáo doanh thu/lợi nhuận theo điểm.
create table revenue_codes (
  id uuid primary key default gen_random_uuid(),
  code text unique not null,
  name text not null,
  created_at timestamptz default now()
);

-- Mã xuất (loại khách / lý do xuất: khách quen, xuất chế biến, không doanh thu...)
create table export_codes (
  id uuid primary key default gen_random_uuid(),
  code text unique not null,
  name text not null,
  created_at timestamptz default now()
);

-- Danh mục sản phẩm: dùng chung cho cả Nguyên vật liệu (NL) và Thành phẩm (TP),
-- phân biệt bằng cột classification. Mã SP (code) gợi ý tự tăng theo từng
-- classification riêng (NL và TP có 2 dải mã độc lập).
create table products (
  id uuid primary key default gen_random_uuid(),
  code text unique not null,
  name text not null,
  unit text not null,
  group_code text,
  group_name text,
  classification text not null default 'NL' check (classification in ('NL','TP')),
  created_at timestamptz default now()
);

-- -----------------------------------------------------------------------------
-- TỒN ĐẦU KỲ — mốc kiểm/chốt kho do Quản lý nhập, dùng làm gốc để tính tồn
-- kho hiện tại (Tồn đầu + Nhập - Xuất = Tồn cuối) và tính giá bình quân gia
-- quyền. Mỗi lần chốt kho mới sẽ ghi thêm 1 dòng mới (giữ lịch sử).
-- -----------------------------------------------------------------------------
create table stock_opening (
  id uuid primary key default gen_random_uuid(),
  product_id uuid not null references products(id),
  as_of_date date not null,
  quantity numeric not null default 0,
  unit_price numeric not null default 0,
  note text,
  created_by uuid references employees(id),
  created_at timestamptz default now()
);

-- -----------------------------------------------------------------------------
-- PHIẾU NHẬP HÀNG — mỗi dòng = 1 lần nhập 1 sản phẩm từ 1 nhà cung cấp.
-- Nhân viên chỉ nhập: order_number, supplier_id, product_id, quantity,
-- unit_price. Các trường còn lại (receipt_code, total_amount, payment_type,
-- import_date) do ứng dụng tự tính/tự sinh.
-- -----------------------------------------------------------------------------
create table import_records (
  id uuid primary key default gen_random_uuid(),
  order_number text,
  receipt_code text not null,
  supplier_id uuid not null references suppliers(id),
  product_id uuid not null references products(id),
  quantity numeric not null,
  unit_price numeric not null,
  total_amount numeric not null,
  payment_type text not null,
  import_date date not null default current_date,
  created_by uuid references employees(id),
  created_at timestamptz default now()
);

-- -----------------------------------------------------------------------------
-- PHIẾU XUẤT HÀNG — mỗi dòng = 1 NVL tiêu hao HOẶC 1 TP bán ra, các dòng
-- cùng 1 lần xuất dùng chung order_number/receipt_code/revenue_code/export_code.
-- Dòng NVL: nhân viên chỉ nhập quantity, unit_price tự tính theo bình quân
-- gia quyền tại thời điểm xuất. Dòng TP: nhân viên nhập cả quantity lẫn
-- unit_price (giá bán).
-- -----------------------------------------------------------------------------
create table export_records (
  id uuid primary key default gen_random_uuid(),
  order_number text,
  receipt_code text not null,
  revenue_code_id uuid references revenue_codes(id),
  export_code_id uuid references export_codes(id),
  product_id uuid not null references products(id),
  line_type text not null check (line_type in ('NVL','TP')),
  quantity numeric not null,
  unit_price numeric not null,
  total_amount numeric not null,
  export_date date not null default current_date,
  created_by uuid references employees(id),
  created_at timestamptz default now()
);

-- -----------------------------------------------------------------------------
-- Row Level Security — mở quyền đọc/ghi qua anon key cho công cụ nội bộ,
-- giống các app trước. Mật khẩu tài khoản được bảo vệ riêng ở phần dưới.
-- -----------------------------------------------------------------------------
alter table employees enable row level security;
alter table suppliers enable row level security;
alter table revenue_codes enable row level security;
alter table export_codes enable row level security;
alter table products enable row level security;
alter table stock_opening enable row level security;
alter table import_records enable row level security;
alter table export_records enable row level security;

create policy "allow all - suppliers" on suppliers for all using (true) with check (true);
create policy "allow all - revenue_codes" on revenue_codes for all using (true) with check (true);
create policy "allow all - export_codes" on export_codes for all using (true) with check (true);
create policy "allow all - products" on products for all using (true) with check (true);
create policy "allow all - stock_opening" on stock_opening for all using (true) with check (true);
create policy "allow all - import_records" on import_records for all using (true) with check (true);
create policy "allow all - export_records" on export_records for all using (true) with check (true);

-- Bảng employees: chỉ cho đọc các cột không nhạy cảm qua API thông thường —
-- mật khẩu KHÔNG bao giờ lộ ra ngoài. KHÔNG cấp quyền insert/update trực tiếp
-- qua API — mọi thay đổi (tạo tài khoản, đổi/đặt lại mật khẩu) đều phải đi qua
-- các hàm RPC bên dưới (đã tự băm mật khẩu, có xác thực Quản lý).
create policy "allow select - employees" on employees for select using (true);
revoke select (password) on employees from anon, authenticated;

-- -----------------------------------------------------------------------------
-- XÁC THỰC ĐĂNG NHẬP — mật khẩu so sánh ngay trong database bằng bcrypt,
-- không bao giờ gửi (kể cả bản băm) ra ngoài cho ứng dụng.
-- -----------------------------------------------------------------------------
create or replace function verify_employee_login(p_username text, p_password text)
returns table (
  id uuid, username text, name text, role text,
  must_change_password boolean, password_change_deadline timestamptz, locked boolean
)
language sql
security definer
set search_path = public, extensions
as $$
  select
    e.id, e.username, e.name, e.role,
    e.must_change_password, e.password_change_deadline,
    (e.must_change_password and e.password_change_deadline is not null and now() > e.password_change_deadline) as locked
  from employees e
  where e.username = p_username
    and e.password = crypt(p_password, e.password)
  limit 1;
$$;
grant execute on function verify_employee_login(text, text) to anon, authenticated;

-- Bắt đầu đếm ngược 24h khi tài khoản bị bắt buộc đổi mật khẩu — chỉ set 1 lần.
create or replace function start_password_deadline(p_employee_id uuid)
returns timestamptz
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  v_deadline timestamptz;
begin
  update employees
  set password_change_deadline = coalesce(password_change_deadline, now() + interval '24 hours')
  where id = p_employee_id
  returning password_change_deadline into v_deadline;
  return v_deadline;
end;
$$;
grant execute on function start_password_deadline(uuid) to anon, authenticated;

-- Tự đổi mật khẩu (yêu cầu đúng mật khẩu hiện tại).
create or replace function change_own_password(p_employee_id uuid, p_old_password text, p_new_password text)
returns boolean
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  v_ok boolean;
begin
  select (password = crypt(p_old_password, password)) into v_ok
  from employees where id = p_employee_id;

  if not coalesce(v_ok, false) then
    return false;
  end if;

  update employees
  set password = crypt(p_new_password, gen_salt('bf')),
      must_change_password = false,
      password_change_deadline = null
  where id = p_employee_id;

  return true;
end;
$$;
grant execute on function change_own_password(uuid, text, text) to anon, authenticated;

-- Quản lý đặt lại mật khẩu cho người khác — yêu cầu xác nhận đúng mật khẩu
-- CỦA CHÍNH quản lý đang thao tác (chống giả mạo qua API).
create or replace function admin_reset_password(
  p_admin_id uuid, p_admin_password text,
  p_target_employee_id uuid, p_new_password text
)
returns boolean
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  v_admin_ok boolean;
begin
  select (role = 'quan_ly' and password = crypt(p_admin_password, password))
  into v_admin_ok
  from employees where id = p_admin_id;

  if not coalesce(v_admin_ok, false) then
    return false;
  end if;

  update employees
  set password = crypt(p_new_password, gen_salt('bf')),
      must_change_password = true,
      password_change_deadline = null
  where id = p_target_employee_id;

  return true;
end;
$$;
grant execute on function admin_reset_password(uuid, text, uuid, text) to anon, authenticated;

-- Đặt mật khẩu thủ công qua SQL Editor (dùng khi cần reset khẩn, không qua app).
create or replace function set_employee_password(p_employee_id uuid, p_new_password text)
returns void
language sql
security definer
set search_path = public, extensions
as $$
  update employees set password = crypt(p_new_password, gen_salt('bf')) where id = p_employee_id;
$$;
-- Không cấp quyền cho anon/authenticated — chỉ dùng được khi bạn tự chạy qua SQL Editor.

-- Quản lý tạo tài khoản mới — yêu cầu xác nhận đúng mật khẩu CỦA CHÍNH quản lý
-- đang thao tác (chống giả mạo qua API), mật khẩu ban đầu được băm ngay khi tạo.
create or replace function admin_create_employee(
  p_admin_id uuid, p_admin_password text,
  p_username text, p_name text, p_role text, p_initial_password text
)
returns uuid
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  v_admin_ok boolean;
  v_new_id uuid;
begin
  select (role = 'quan_ly' and password = crypt(p_admin_password, password))
  into v_admin_ok
  from employees where id = p_admin_id;

  if not coalesce(v_admin_ok, false) then
    raise exception 'ADMIN_PASSWORD_INCORRECT';
  end if;

  insert into employees (username, password, name, role, must_change_password)
  values (p_username, crypt(p_initial_password, gen_salt('bf')), p_name, p_role, true)
  returning id into v_new_id;

  return v_new_id;
end;
$$;
grant execute on function admin_create_employee(uuid, text, text, text, text, text) to anon, authenticated;

-- -----------------------------------------------------------------------------
-- 2 tài khoản mẫu để đăng nhập thử — ĐỔI MẬT KHẨU NGAY sau khi triển khai thật.
-- -----------------------------------------------------------------------------
insert into employees (username, password, name, role) values
  ('nvkho1', crypt('123456', gen_salt('bf')), 'Nhân viên kho 1', 'nhan_vien_kho'),
  ('quanly1', crypt('123456', gen_salt('bf')), 'Quản lý kho', 'quan_ly');
