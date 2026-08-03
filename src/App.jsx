import React, { useState, useMemo, useCallback, useEffect } from "react";
import { supabase } from "./supabaseClient";
import { generateBusinessRegistrationDoc, generateHouseholdBusinessDoc, soThanhChuTien } from "./docGen";
import NGANH_NGHE_CAP4 from "./nganhNgheData";
import {
  Users, ClipboardList, BarChart3, Bell, LogOut, CheckCircle2, XCircle,
  UserPlus, TrendingUp, Wallet, Building2, AlertTriangle, Clock, Camera,
  ChevronRight, Inbox, ClipboardCheck, FileText, DollarSign, HeartHandshake,
  Plus, Search, ArrowLeft, ShieldCheck, Receipt, Download, Loader2, Lock, User,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  PieChart, Pie, Cell, AreaChart, Area, Legend,
} from "recharts";

/* ============================================================================
   QUẢN LÝ ĐƠN HÀNG NMT — bản triển khai thật (React + Supabase).
   Dữ liệu (khách hàng, đơn hàng, chi phí, tài khoản) lưu trong Supabase Postgres.
   Ảnh CCCD và file PDF giấy phép lưu trong Supabase Storage.
   ========================================================================= */

// ---------------------------------------------------------------------------
// HẰNG SỐ NGHIỆP VỤ
// ---------------------------------------------------------------------------
const PROCEDURE_TYPES = [
  { key: "mo_hkd", label: "Mở HKD" },
  { key: "mo_cty", label: "Mở Công ty" },
  { key: "cham_dut_mst", label: "Chấm dứt MST KD" },
];
const procedureLabel = (k) => PROCEDURE_TYPES.find((p) => p.key === k)?.label || k;

const ORDER_STATUS = {
  cho_xu_ly: { label: "Chờ xử lý", color: "bg-slate-100 text-slate-600 border-slate-200" },
  da_tiep_nhan: { label: "Đã tiếp nhận", color: "bg-sky-50 text-sky-700 border-sky-200" },
  trinh_lanh_dao: { label: "Trình lãnh đạo", color: "bg-amber-50 text-amber-700 border-amber-200" },
  duoc_chap_thuan: { label: "Được chấp thuận", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
  gui_giay_phep: { label: "Đã gửi giấy phép", color: "bg-teal-50 text-teal-700 border-teal-200" },
  hoan_thanh: { label: "Đã hoàn thành", color: "bg-indigo-50 text-indigo-700 border-indigo-200" },
  da_thanh_toan: { label: "Đã thanh toán", color: "bg-emerald-100 text-emerald-800 border-emerald-300" },
  chua_duoc_chap_thuan: { label: "Chưa được chấp thuận", color: "bg-rose-50 text-rose-700 border-rose-200" },
};
const STATUS_FLOW = ["cho_xu_ly", "da_tiep_nhan", "trinh_lanh_dao", "duoc_chap_thuan", "gui_giay_phep"];
const RESOLVED_STATUSES = ["duoc_chap_thuan", "gui_giay_phep", "hoan_thanh", "da_thanh_toan", "chua_duoc_chap_thuan"];

const CARE_STEPS = {
  mo_hkd: ["Mở tài khoản HKD", "Gửi sổ cho khách", "Hướng dẫn khách kê khai thuế", "Tư vấn kế toán thuế"],
  mo_cty: ["Mở tài khoản HKD", "Mua chữ ký số", "Kê khai thuế ban đầu", "Tư vấn kế toán thuế"],
  cham_dut_mst: ["Hoàn thành chấm dứt đăng ký kinh doanh"],
};

const WARDS = ["Phường Thành Sen", "Phường Bắc Hà", "Phường Thạch Hà"];
const VNID_BUCKET = "vnid-photos";
const LICENSE_BUCKET = "license-pdfs";

// ---------------------------------------------------------------------------
// HÀM TIỆN ÍCH
// ---------------------------------------------------------------------------
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("vi-VN") + " đ";
}
function shortMoney(n) {
  const v = Number(n) || 0;
  if (v >= 1_000_000_000) return (v / 1_000_000_000).toFixed(1).replace(/\.0$/, "") + " tỷ";
  if (v >= 1_000_000) return (v / 1_000_000).toFixed(1).replace(/\.0$/, "") + " tr";
  if (v >= 1_000) return (v / 1_000).toFixed(0) + " k";
  return String(v);
}
function fmtDate(d) {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("vi-VN");
}
function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
function isWeekend(d) {
  const day = d.getDay();
  return day === 0 || day === 6;
}
function addBusinessDays(date, n) {
  let d = startOfDay(date);
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    if (!isWeekend(d)) added++;
  }
  return d;
}
function businessDaysSince(date) {
  let d = startOfDay(date);
  const now = startOfDay(new Date());
  let count = 0;
  while (d < now) {
    d.setDate(d.getDate() + 1);
    if (d <= now && !isWeekend(d)) count++;
  }
  return count;
}
function monthKey(d) {
  const dt = new Date(d);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`;
}
const CHART_COLORS = ["#4338ca", "#d97706", "#0891b2", "#e11d48", "#65a30d", "#9333ea"];
const CHART_TOOLTIP_STYLE = {
  contentStyle: { borderRadius: 12, border: "1px solid #e2e8f0", boxShadow: "0 4px 16px rgba(15,23,42,.08)", fontSize: 12.5 },
  labelStyle: { color: "#334155", fontWeight: 600, marginBottom: 4 },
};

// ---------------------------------------------------------------------------
// MAPPER: chuyển đổi giữa cột snake_case của Supabase và object camelCase dùng
// trong giao diện
// ---------------------------------------------------------------------------
function customerFromRow(r) {
  return {
    id: r.id, name: r.name, cccd: r.cccd, phone: r.phone, address: r.address, ward: r.ward,
    industries: Array.isArray(r.industries) ? r.industries : [], referrer: r.referrer, vnidPhoto: r.vnid_photo_url,
    employeeId: r.employee_id, createdAt: r.created_at,
  };
}
function orderFromRow(r) {
  return {
    id: r.id, orderCode: r.order_code, customerId: r.customer_id, employeeId: r.employee_id,
    procedureType: r.procedure_type, status: r.status, createdAt: r.created_at, receivedAt: r.received_at,
    leaderAt: r.leader_at, approvedAt: r.approved_at, rejectedAt: r.rejected_at,
    overdueReason: r.overdue_reason || "", licensePdfData: r.license_pdf_url, licensePdfName: r.license_pdf_name,
    licenseSentAt: r.license_sent_at, confirmedAt: r.confirmed_at, revenue: r.revenue, cost: r.cost,
    laborFee: r.labor_fee, completedAt: r.completed_at, careSteps: r.care_steps || {},
    companyName: r.company_name, capital: r.capital, ownerDob: r.owner_dob, ownerGender: r.owner_gender,
    ownerEmail: r.owner_email, ownerProvince: r.owner_province || "Hà Tĩnh",
    hqAddress: r.hq_address, hqWard: r.hq_ward, hqProvince: r.hq_province || "Hà Tĩnh",
    industries: Array.isArray(r.industries) ? r.industries : [],
  };
}
function expenseFromRow(r) {
  return { id: r.id, date: r.date, description: r.description, amount: Number(r.amount), createdAt: r.created_at };
}

// Tải file lên Supabase Storage, trả về { url, name }
async function uploadFile(bucket, file) {
  const ext = file.name.includes(".") ? file.name.split(".").pop() : "dat";
  const path = `${crypto.randomUUID()}.${ext}`;
  const { error } = await supabase.storage.from(bucket).upload(path, file);
  if (error) throw error;
  const { data } = supabase.storage.from(bucket).getPublicUrl(path);
  return { url: data.publicUrl, name: file.name };
}

// ---------------------------------------------------------------------------
// UI DÙNG CHUNG
// ---------------------------------------------------------------------------
function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-shadow duration-200 ${className}`}>{children}</div>;
}
function SectionTitle({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex items-start gap-3 mb-4">
      <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-700 to-indigo-900 text-white flex items-center justify-center shrink-0 shadow-sm shadow-indigo-900/20">
        <Icon size={19} />
      </div>
      <div>
        <h2 className="text-lg font-semibold text-slate-800 tracking-tight">{title}</h2>
        {subtitle && <p className="text-sm text-slate-500">{subtitle}</p>}
      </div>
    </div>
  );
}
function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-slate-400">
      <div className="w-14 h-14 rounded-2xl bg-slate-50 flex items-center justify-center mb-3">
        <Icon size={26} className="opacity-50" />
      </div>
      <p className="text-sm">{text}</p>
    </div>
  );
}
function MetricCard({ label, value, icon: Icon, accent = "indigo" }) {
  const a = {
    indigo: { text: "text-indigo-700", bg: "bg-indigo-50", bar: "bg-indigo-600" },
    amber: { text: "text-amber-700", bg: "bg-amber-50", bar: "bg-amber-600" },
    emerald: { text: "text-emerald-700", bg: "bg-emerald-50", bar: "bg-emerald-600" },
    rose: { text: "text-rose-700", bg: "bg-rose-50", bar: "bg-rose-600" },
  }[accent];
  return (
    <div className="relative bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4 flex items-center gap-3 overflow-hidden">
      <span className={`absolute left-0 top-0 bottom-0 w-1 ${a.bar}`} />
      <div className={`w-11 h-11 rounded-xl flex items-center justify-center shrink-0 ${a.bg} ${a.text}`}>
        <Icon size={19} />
      </div>
      <div className="min-w-0">
        <p className="text-xs text-slate-500 truncate">{label}</p>
        <p className="text-lg font-semibold text-slate-800 truncate tracking-tight">{value}</p>
      </div>
    </div>
  );
}
function Badge({ children, className = "" }) {
  return <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium border ${className}`}>{children}</span>;
}
function StatusBadge({ status }) {
  const s = ORDER_STATUS[status] || ORDER_STATUS.cho_xu_ly;
  return <Badge className={s.color}>{s.label}</Badge>;
}
function PrimaryButton({ children, className = "", ...props }) {
  return (
    <button {...props} className={`inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-gradient-to-b from-indigo-700 to-indigo-800 text-white text-sm font-medium shadow-sm hover:from-indigo-800 hover:to-indigo-900 active:scale-[0.98] transition disabled:opacity-40 disabled:cursor-not-allowed ${className}`}>
      {children}
    </button>
  );
}
function GhostButton({ children, className = "", ...props }) {
  return (
    <button {...props} className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-700 text-sm font-medium hover:bg-slate-50 active:scale-[0.98] transition disabled:opacity-40 ${className}`}>
      {children}
    </button>
  );
}
function TextField({ label, hint, ...props }) {
  return (
    <label className="block">
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <input {...props} className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-indigo-500/40 focus:border-indigo-600" />
      {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}
function SelectField({ label, children, ...props }) {
  return (
    <label className="block">
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <select {...props} className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white transition focus:outline-none focus:ring-2 focus:ring-indigo-500/40 focus:border-indigo-600">
        {children}
      </select>
    </label>
  );
}

// Bỏ dấu tiếng Việt để tìm kiếm không phân biệt có dấu/không dấu
function stripDiacritics(str) {
  return (str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d").replace(/Đ/g, "D")
    .toLowerCase();
}

// Ô "Ngành nghề kinh doanh" có gợi ý tự động: gõ mã ngành cấp 4 (VD 0111) hoặc
// gõ tên/từ khoá ngành (có dấu hoặc không dấu đều được) sẽ hiện danh sách gợi ý.
function IndustryField({ label, value, onChange, className = "" }) {
  const [query, setQuery] = useState(value || "");
  const [open, setOpen] = useState(false);
  const [highlight, setHighlight] = useState(0);

  useEffect(() => setQuery(value || ""), [value]);

  const suggestions = useMemo(() => {
    const q = query.trim();
    if (!q) return [];
    const qNorm = stripDiacritics(q);
    const byCode = q.replace(/\D/g, "").length >= 2 ? NGANH_NGHE_CAP4.filter((n) => n.code.startsWith(q.replace(/\D/g, ""))) : [];
    const byName = NGANH_NGHE_CAP4.filter((n) => stripDiacritics(n.name).includes(qNorm));
    const merged = [...byCode, ...byName.filter((n) => !byCode.includes(n))];
    return merged.slice(0, 8);
  }, [query]);

  const pick = (item) => {
    const text = `${item.code} - ${item.name}`;
    setQuery(text);
    onChange(text);
    setOpen(false);
  };

  const handleInput = (e) => {
    const v = e.target.value;
    setQuery(v);
    onChange(v);
    setOpen(true);
    setHighlight(0);
  };

  const handleKeyDown = (e) => {
    if (!open || suggestions.length === 0) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setHighlight((h) => Math.min(h + 1, suggestions.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setHighlight((h) => Math.max(h - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); pick(suggestions[highlight]); }
    else if (e.key === "Escape") { setOpen(false); }
  };

  return (
    <label className={`block relative ${className}`}>
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <input
        value={query}
        onChange={handleInput}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        onKeyDown={handleKeyDown}
        placeholder="Gõ mã ngành cấp 4 (VD: 0111) hoặc tên ngành..."
        className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-indigo-500/40 focus:border-indigo-600"
      />
      {open && suggestions.length > 0 && (
        <div className="absolute z-20 left-0 right-0 mt-1 bg-white rounded-xl border border-slate-200 shadow-lg max-h-64 overflow-y-auto">
          {suggestions.map((item, i) => (
            <button
              type="button"
              key={item.code}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => pick(item)}
              className={`w-full text-left px-3 py-2 text-sm flex items-start gap-2 ${i === highlight ? "bg-indigo-50" : "hover:bg-slate-50"}`}
            >
              <span className="text-indigo-600 font-medium shrink-0">{item.code}</span>
              <span className="text-slate-700">{item.name}</span>
            </button>
          ))}
        </div>
      )}
    </label>
  );
}

// Danh sách ngành, nghề kinh doanh — cho phép thêm nhiều ngành, mỗi ngành có ô
// "Chi tiết ngành nghề" riêng, và đánh dấu đúng 1 ngành là "Ngành chính" (radio
// chọn 1 trong nhiều), khớp đúng cấu trúc bảng trong hồ sơ đăng ký thật.
function IndustryListEditor({ industries, onChange }) {
  const [picker, setPicker] = useState("");

  const addIndustry = (value) => {
    const m = (value || "").trim().match(/^(\d{3,5})\s*-\s*(.+)$/);
    if (!m) { setPicker(value); return; }
    const code = m[1];
    const name = m[2].trim();
    if (industries.some((it) => it.code === code)) { setPicker(""); return; }
    const next = [...industries, { code, name, detail: "", isPrimary: industries.length === 0 }];
    onChange(next);
    setPicker("");
  };

  const removeIndustry = (idx) => {
    const removed = industries[idx];
    let next = industries.filter((_, i) => i !== idx);
    if (removed?.isPrimary && next.length > 0 && !next.some((it) => it.isPrimary)) {
      next = next.map((it, i) => (i === 0 ? { ...it, isPrimary: true } : it));
    }
    onChange(next);
  };

  const setPrimary = (idx) => {
    onChange(industries.map((it, i) => ({ ...it, isPrimary: i === idx })));
  };

  const setDetail = (idx, detail) => {
    onChange(industries.map((it, i) => (i === idx ? { ...it, detail } : it)));
  };

  return (
    <div className="border border-slate-200 rounded-xl overflow-hidden">
      <div className="px-4 py-3 bg-slate-50 border-b border-slate-200">
        <p className="text-sm font-semibold text-slate-800">Danh sách ngành nghề kinh doanh</p>
        <p className="text-xs text-slate-400">{industries.length} ngành nghề đã thêm</p>
      </div>

      {industries.length > 0 && (
        <div className="hidden sm:grid grid-cols-[80px_1fr_140px_70px] gap-3 px-4 py-2 text-xs font-medium text-slate-400 border-b border-slate-100">
          <span>MÃ NGÀNH</span><span>TÊN NGÀNH, NGHỀ</span><span>NGÀNH CHÍNH</span><span></span>
        </div>
      )}

      <div className="divide-y divide-slate-100">
        {industries.map((it, idx) => (
          <div key={idx} className="px-4 py-3">
            <div className="grid sm:grid-cols-[80px_1fr_140px_70px] gap-3 items-start">
              <span className="font-semibold text-teal-700 text-sm">{it.code || "—"}</span>
              <span className="text-sm text-slate-700">{it.name}</span>
              <label className="flex items-center gap-1.5 text-xs text-slate-600 cursor-pointer">
                <input type="radio" checked={it.isPrimary} onChange={() => setPrimary(idx)} className="accent-teal-700" />
                {it.isPrimary ? "Ngành chính" : "Chọn"}
              </label>
              <button type="button" onClick={() => removeIndustry(idx)} className="text-xs text-rose-600 border border-rose-200 rounded-lg px-2 py-1 hover:bg-rose-50 justify-self-start sm:justify-self-auto">Xóa</button>
            </div>
            <div className="mt-2">
              <span className="block text-xs font-medium text-slate-600 mb-1">Chi tiết ngành nghề</span>
              <textarea
                value={it.detail}
                onChange={(e) => setDetail(idx, e.target.value)}
                placeholder={`Nhập nội dung hoạt động cụ thể của ngành ${it.code || ""}...`}
                rows={2}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-600"
              />
            </div>
          </div>
        ))}
      </div>

      <div className="p-4 bg-slate-50/60">
        <IndustryField label="Thêm ngành nghề (mã cấp 4)" value={picker} onChange={addIndustry} />
      </div>
    </div>
  );
}

function Toast({ message, onClose }) {
  React.useEffect(() => {
    const t = setTimeout(onClose, 3000);
    return () => clearTimeout(t);
  }, [onClose]);
  return (
    <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-sm px-4 py-2.5 rounded-xl shadow-lg shadow-slate-900/30 flex items-center gap-2">
      <CheckCircle2 size={15} className="text-emerald-400 shrink-0" /> {message}
    </div>
  );
}
function FullScreenLoader() {
  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50">
      <Loader2 size={28} className="text-indigo-600 animate-spin" />
    </div>
  );
}

// ---------------------------------------------------------------------------
// LOGIC DẪN XUẤT: hạn xử lý, cảnh báo quá hạn, công nợ
// ---------------------------------------------------------------------------
const LATE_PENALTY_AMOUNT = 100_000;

function useDerivedOrders(orders) {
  return useMemo(() => {
    return orders.map((o) => {
      let approvalDeadline = null, approvalOverdue = false, approvalDaysElapsed = 0;
      if (o.receivedAt) {
        approvalDeadline = addBusinessDays(o.receivedAt, 3);
        approvalDaysElapsed = businessDaysSince(o.receivedAt);
        approvalOverdue = !RESOLVED_STATUSES.includes(o.status) && approvalDaysElapsed >= 3;
      }
      let paymentDeadline = null, paymentOverdue = false;
      if (o.approvedAt) {
        paymentDeadline = addBusinessDays(o.approvedAt, 2);
        paymentOverdue = o.status !== "da_thanh_toan" && businessDaysSince(o.approvedAt) >= 2;
      }
      const isCompleted = o.status === "da_thanh_toan";
      const isLatePenalty = approvalOverdue && !o.overdueReason;
      const latePenaltyAmount = isLatePenalty ? LATE_PENALTY_AMOUNT : 0;
      return { ...o, approvalDeadline, approvalOverdue, paymentDeadline, paymentOverdue, isCompleted, isLatePenalty, latePenaltyAmount };
    });
  }, [orders]);
}

function useNotifications(derivedOrders, customers) {
  return useMemo(() => {
    const list = [];
    derivedOrders.forEach((o) => {
      const cust = customers.find((c) => c.id === o.customerId);
      if (o.status === "chua_duoc_chap_thuan") {
        list.push({
          id: `${o.id}_rejected`, orderId: o.id, type: "chua_chap_thuan",
          message: `Đơn ${o.orderCode} (${cust?.name || "?"}) đã bị đánh dấu CHƯA ĐƯỢC CHẤP THUẬN. Lý do: "${o.overdueReason || "(chưa điền)"}".`,
          severity: "danger", createdAt: o.rejectedAt || o.approvalDeadline,
        });
      } else if (o.approvalOverdue) {
        list.push({
          id: `${o.id}_approval`, orderId: o.id, type: "qua_han_chap_thuan",
          message: `Đơn ${o.orderCode} (${cust?.name || "?"}) đã quá 3 ngày làm việc kể từ khi tiếp nhận mà chưa được chấp thuận.` +
            (o.overdueReason ? ` Lý do nhân viên báo cáo: "${o.overdueReason}".` : ` Nhân viên chưa điền lý do — đã tự động ghi nhận chậm tiến độ và trừ ${fmtMoney(LATE_PENALTY_AMOUNT)} tiền công.`),
          severity: o.overdueReason ? "warning" : "danger", createdAt: o.approvalDeadline,
        });
      }
      if (o.paymentOverdue) {
        list.push({
          id: `${o.id}_debt`, orderId: o.id, type: "cong_no",
          message: `Đơn ${o.orderCode} (${cust?.name || "?"}) đã quá 2 ngày làm việc kể từ khi được chấp thuận mà chưa hoàn tất thanh toán — cần thu hồi công nợ.`,
          severity: "warning", createdAt: o.paymentDeadline,
        });
      }
    });
    return list.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  }, [derivedOrders, customers]);
}

function useEmployeeNotifications(myDerivedOrders) {
  return useMemo(() => {
    const list = [];
    myDerivedOrders.forEach((o) => {
      if (o.status === "hoan_thanh") {
        list.push({
          id: `${o.id}_confirmed`, orderId: o.id, type: "xac_nhan",
          message: `Đơn ${o.orderCode} đã được quản lý xác nhận hoàn thành. Đang chờ cập nhật doanh thu / chi phí / tiền công.`,
          severity: "info", createdAt: o.confirmedAt || o.createdAt,
        });
      }
      if (o.status === "da_thanh_toan") {
        list.push({
          id: `${o.id}_paid`, orderId: o.id, type: "da_thanh_toan",
          message: `Đơn ${o.orderCode} đã hoàn tất và được ghi nhận thanh toán. Tiền công của bạn: ${fmtMoney(o.laborFee)}.`,
          severity: "success", createdAt: o.completedAt || o.createdAt,
        });
      }
    });
    return list.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  }, [myDerivedOrders]);
}

// ---------------------------------------------------------------------------
// MÀN HÌNH ĐĂNG NHẬP — tài khoản thật (bảng employees trong Supabase)
// ---------------------------------------------------------------------------
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password) return;
    setError(""); setLoading(true);
    const { data, error: qErr } = await supabase
      .from("employees")
      .select("id,name,username,role")
      .eq("username", username.trim())
      .eq("password", password)
      .maybeSingle();
    setLoading(false);
    if (qErr) { setError("Không thể kết nối máy chủ. Vui lòng thử lại."); return; }
    if (!data) { setError("Sai tên đăng nhập hoặc mật khẩu."); return; }
    onLogin(data);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-b from-indigo-50 via-slate-50 to-slate-50 px-4 py-10">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-indigo-700 to-indigo-900 mx-auto mb-3 flex items-center justify-center shadow-lg shadow-indigo-900/25">
            <FileText size={26} className="text-white" />
          </div>
          <h1 className="text-2xl font-semibold text-slate-800 tracking-tight">Quản lý đơn hàng NMT</h1>
          <p className="text-sm text-slate-500 mt-1">Đăng nhập bằng tài khoản nhân viên</p>
        </div>
        <form onSubmit={handleSubmit} className="bg-white rounded-2xl border border-slate-200/80 shadow-xl shadow-slate-900/5 p-5 space-y-4">
          <label className="block">
            <span className="block text-xs font-medium text-slate-600 mb-1">Tên đăng nhập</span>
            <div className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-300 focus-within:ring-2 focus-within:ring-indigo-500/40 focus-within:border-indigo-600">
              <User size={15} className="text-slate-400 shrink-0" />
              <input value={username} onChange={(e) => setUsername(e.target.value)} className="flex-1 text-sm outline-none" placeholder="vd: nv1" autoFocus />
            </div>
          </label>
          <label className="block">
            <span className="block text-xs font-medium text-slate-600 mb-1">Mật khẩu</span>
            <div className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-300 focus-within:ring-2 focus-within:ring-indigo-500/40 focus-within:border-indigo-600">
              <Lock size={15} className="text-slate-400 shrink-0" />
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} className="flex-1 text-sm outline-none" placeholder="••••••" />
            </div>
          </label>
          {error && <p className="text-xs text-rose-600 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
          <PrimaryButton type="submit" disabled={loading} className="w-full justify-center">
            {loading ? <Loader2 size={15} className="animate-spin" /> : <ChevronRight size={15} />} Đăng nhập
          </PrimaryButton>
        </form>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// NHÂN VIÊN — PHÂN HỆ 1: THÔNG TIN KHÁCH HÀNG
// ---------------------------------------------------------------------------
function CustomerFormCard({ onSubmit }) {
  const [form, setForm] = useState({ name: "", cccd: "", phone: "", address: "", ward: WARDS[0], industries: [], referrer: "" });
  const [photoFile, setPhotoFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [saving, setSaving] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const handlePhoto = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPhotoFile(file);
    const reader = new FileReader();
    reader.onload = () => setPreview(reader.result);
    reader.readAsDataURL(file);
  };

  const submit = async () => {
    if (!form.phone || !form.address || saving) return;
    setSaving(true);
    await onSubmit({ ...form, photoFile });
    setSaving(false);
    setForm({ name: "", cccd: "", phone: "", address: "", ward: WARDS[0], industries: [], referrer: "" });
    setPhotoFile(null);
    setPreview(null);
  };

  return (
    <Card className="p-4 sm:p-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Thêm khách hàng mới</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <TextField label="Tên khách hàng" value={form.name} onChange={set("name")} placeholder="Không bắt buộc, giúp dễ nhận diện" />
        <TextField label="Số CCCD" value={form.cccd} onChange={set("cccd")} placeholder="12 số trên căn cước công dân" />
        <TextField label="Số điện thoại *" value={form.phone} onChange={set("phone")} placeholder="09xxxxxxxx" />
        <TextField label="Địa chỉ *" value={form.address} onChange={set("address")} placeholder="Số nhà, đường..." />
        <TextField label="Phường" value={form.ward} onChange={set("ward")} placeholder="Phường..." />
        <div className="sm:col-span-2">
          <IndustryListEditor industries={form.industries} onChange={(list) => setForm((f) => ({ ...f, industries: list }))} />
        </div>
        <TextField label="Người giới thiệu" value={form.referrer} onChange={set("referrer")} placeholder="Tên người giới thiệu khách hàng này (nếu có)" className="sm:col-span-2" />
      </div>
      <div className="mb-4">
        <span className="block text-xs font-medium text-slate-600 mb-1">Ảnh chụp VNID (căn cước công dân)</span>
        <label className="flex items-center gap-3 p-3 rounded-xl border border-dashed border-slate-300 hover:border-indigo-400 hover:bg-indigo-50/40 transition cursor-pointer">
          <span className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center shrink-0 overflow-hidden">
            {preview ? <img src={preview} alt="VNID" className="w-full h-full object-cover" /> : <Camera size={17} className="text-slate-400" />}
          </span>
          <span className="text-xs text-slate-500">{preview ? "Đã chọn ảnh — bấm để đổi ảnh khác" : "Bấm để chụp / tải ảnh VNID lên"}</span>
          <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handlePhoto} />
        </label>
      </div>
      <PrimaryButton onClick={submit} disabled={saving}>
        {saving ? <Loader2 size={15} className="animate-spin" /> : <UserPlus size={15} />} Lưu khách hàng
      </PrimaryButton>
    </Card>
  );
}

function CustomerModule({ currentUser, customers, onAddCustomer }) {
  const mine = customers.filter((c) => c.employeeId === currentUser.id);
  const [q, setQ] = useState("");
  const filtered = mine.filter((c) => !q || c.phone.includes(q) || (c.cccd || "").includes(q) || (c.name || "").toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="space-y-5">
      <SectionTitle icon={Users} title="Thông tin khách hàng" subtitle="Khách hàng bạn phụ trách" />
      <CustomerFormCard onSubmit={onAddCustomer} />
      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex items-center gap-2">
          <Search size={15} className="text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm theo tên, số điện thoại hoặc CCCD..." className="flex-1 text-sm outline-none" />
        </div>
        {filtered.length === 0 ? (
          <EmptyState icon={Users} text="Chưa có khách hàng nào." />
        ) : (
          <div className="divide-y divide-slate-100">
            {filtered.map((c) => (
              <div key={c.id} className="p-4 flex items-center gap-3">
                <span className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center shrink-0 overflow-hidden">
                  {c.vnidPhoto ? <img src={c.vnidPhoto} alt="" className="w-full h-full object-cover" /> : <Users size={16} className="text-slate-400" />}
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium text-slate-800 truncate">{c.name || "(Chưa đặt tên)"} · {c.phone}</p>
                  <p className="text-xs text-slate-400 truncate">CCCD: {c.cccd || "—"} · {c.address}, {c.ward} · {c.industry || "—"}</p>
                  {c.referrer && <p className="text-xs text-indigo-500 truncate">Người giới thiệu: {c.referrer}</p>}
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// NHÂN VIÊN — PHÂN HỆ 2: TẠO ĐƠN HÀNG
// ---------------------------------------------------------------------------
function OrderCreateModule({ currentUser, customers, orders, onCreateOrder }) {
  const mineCustomers = customers.filter((c) => c.employeeId === currentUser.id);
  const [customerId, setCustomerId] = useState(mineCustomers[0]?.id || "");
  const [procedureType, setProcedureType] = useState(PROCEDURE_TYPES[0].key);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!customerId && mineCustomers[0]) setCustomerId(mineCustomers[0].id);
  }, [mineCustomers, customerId]);

  const nextCode = () => `NMT-${String(orders.length + 1).padStart(4, "0")}`;

  const submit = async () => {
    if (!customerId || saving) return;
    setSaving(true);
    await onCreateOrder(customerId, procedureType, nextCode());
    setSaving(false);
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={ClipboardList} title="Tạo đơn hàng" subtitle="Mã đơn hàng được tự động sinh" />
      <Card className="p-4 sm:p-5 space-y-3">
        {mineCustomers.length === 0 ? (
          <EmptyState icon={Users} text="Bạn cần thêm khách hàng trước khi tạo đơn hàng." />
        ) : (
          <>
            <SelectField label="Khách hàng" value={customerId} onChange={(e) => setCustomerId(e.target.value)}>
              {mineCustomers.map((c) => <option key={c.id} value={c.id}>{c.name || c.phone} — {c.phone}</option>)}
            </SelectField>
            <SelectField label="Loại hình thủ tục" value={procedureType} onChange={(e) => setProcedureType(e.target.value)}>
              {PROCEDURE_TYPES.map((p) => <option key={p.key} value={p.key}>{p.label}</option>)}
            </SelectField>
            <div className="text-xs text-slate-500">Mã đơn hàng dự kiến: <span className="font-semibold text-slate-700">{nextCode()}</span></div>
            <PrimaryButton onClick={submit} disabled={saving}>
              {saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Tạo đơn hàng
            </PrimaryButton>
          </>
        )}
      </Card>
    </div>
  );
}

// Thanh bước tiến độ dạng tích chọn.
function OrderProgressStepper({ status, onStepClick }) {
  const currentIndex = STATUS_FLOW.indexOf(status);
  return (
    <div className="flex items-center mb-1">
      {STATUS_FLOW.map((s, i) => {
        const done = i < currentIndex;
        const current = i === currentIndex;
        const clickable = i === currentIndex + 1;
        return (
          <React.Fragment key={s}>
            {i > 0 && <span className={`flex-1 h-0.5 ${i <= currentIndex ? "bg-indigo-600" : "bg-slate-200"}`} />}
            <button
              type="button"
              onClick={() => onStepClick(s, i)}
              title={done ? `${ORDER_STATUS[s].label} — đã hoàn tất, không thể sửa lại` : ORDER_STATUS[s].label}
              className="flex flex-col items-center gap-1 shrink-0 px-1 cursor-pointer"
            >
              <span
                className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-semibold border-2 transition ${
                  done ? "bg-indigo-600 border-indigo-600 text-white"
                  : current ? "bg-white border-indigo-600 text-indigo-700"
                  : clickable ? "bg-indigo-50 border-indigo-600 text-indigo-700 ring-4 ring-indigo-100 hover:bg-indigo-100"
                  : "bg-white border-slate-200 text-slate-300 hover:border-slate-300"
                }`}
              >
                {done ? <CheckCircle2 size={15} /> : i + 1}
              </span>
              <span className={`text-[10px] text-center leading-tight w-16 ${current ? "text-indigo-700 font-medium" : done ? "text-slate-500" : clickable ? "text-indigo-700 font-medium" : "text-slate-400"}`}>
                {ORDER_STATUS[s].label}
              </span>
            </button>
          </React.Fragment>
        );
      })}
    </div>
  );
}

function ConfirmDialog({ title, message, confirmLabel = "Xác nhận", danger, onConfirm, onCancel }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
        <div className="flex items-center gap-2 mb-2">
          <AlertTriangle size={18} className={danger ? "text-rose-600" : "text-amber-600"} />
          <p className="font-semibold text-slate-800">{title}</p>
        </div>
        <p className="text-sm text-slate-600 mb-5">{message}</p>
        <div className="flex justify-end gap-2">
          <GhostButton onClick={onCancel}>Huỷ, kiểm tra lại</GhostButton>
          <button
            onClick={onConfirm}
            className={`inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-white text-sm font-medium transition ${danger ? "bg-rose-600 hover:bg-rose-700" : "bg-indigo-700 hover:bg-indigo-800"}`}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

// Danh sách nhân sự công ty được uỷ quyền nộp hồ sơ HKD thay khách hàng —
// mỗi người ứng với 1 tài khoản nhân viên. Khi tạo Giấy uỷ quyền, hệ thống tự
// chọn đúng người theo tên tài khoản đang đăng nhập (có thể đổi lại thủ công).
const UY_QUYEN_PROFILES = [
  {
    hoTen: "TRẦN VIỆT HƯNG", gioiTinh: "Nam", ngaySinh: "14/04/1990",
    cccd: "008090011162", diaChi: "Số 175 Phố Huế, Phường Hai Bà Trưng, Thành phố Hà Nội",
    dienThoai: "0911799111", email: "hungviettq@gmail.com",
  },
  {
    hoTen: "NGUYỄN MINH ĐỨC", gioiTinh: "Nam", ngaySinh: "09/09/1999",
    cccd: "001099011128", diaChi: "136 Hàng Cỏ, Phường Cửa Nam, Hà Nội",
    dienThoai: "0966448150", email: "ngminhduc59@gmail.com",
  },
  {
    hoTen: "NGUYỄN MINH TUỆ", gioiTinh: "Nam", ngaySinh: "22/03/2001",
    cccd: "008201003889", diaChi: "Số 175 Phố Huế, Phường Hai Bà Trưng, Thành phố Hà Nội",
    dienThoai: "0394379676", email: "",
  },
];

function stripDiacriticsUQ(str) {
  return (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase().trim();
}

function findUyQuyenProfile(name) {
  const target = stripDiacriticsUQ(name);
  return UY_QUYEN_PROFILES.find((p) => stripDiacriticsUQ(p.hoTen) === target) || null;
}

function isoToDDMMYYYY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

// Form thu thập/điền lại thông tin còn thiếu rồi tự động điền vào mẫu
// "Giấy đề nghị đăng ký doanh nghiệp — Công ty TNHH một thành viên" và tải về.
function BusinessRegModal({ order, customer, onSave, onClose }) {
  const [form, setForm] = useState({
    hoTen: (customer?.name || "").toUpperCase(),
    ngaySinh: order.ownerDob || "",
    gioiTinh: order.ownerGender || "Nam",
    soCccd: customer?.cccd || "",
    dienThoai: customer?.phone || "",
    email: order.ownerEmail || "",
    diaChi1: customer?.address || "",
    xaPhuong1: customer?.ward || "",
    tinhTp1: order.ownerProvince || "Hà Tĩnh",
    tenCongTy: order.companyName || "",
    diaChi2: order.hqAddress || customer?.address || "",
    xaPhuong2: order.hqWard || customer?.ward || "",
    tinhTp2: order.hqProvince || "Hà Tĩnh",
    vonDieuLe: order.capital ?? "",
    industries: order.industries?.length ? order.industries : (customer?.industries || []),
  });
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const vonBangChu = form.vonDieuLe ? soThanhChuTien(form.vonDieuLe) : "";

  const submit = async () => {
    if (!form.hoTen || !form.tenCongTy || !form.vonDieuLe) {
      setError("Vui lòng điền đủ Họ tên, Tên công ty và Vốn điều lệ.");
      return;
    }
    if (form.industries.length === 0) {
      setError("Vui lòng thêm ít nhất 1 ngành nghề kinh doanh.");
      return;
    }
    setError("");
    setGenerating(true);
    try {
      const now = new Date();
      const nganhList = form.industries.map((ind, i) => ({
        stt: String(i + 1),
        ten: ind.detail.trim() ? `${ind.name}\nChi tiết: ${ind.detail.trim()}` : ind.name,
        ma: ind.code,
        chinh: ind.isPrimary ? "X" : "",
      }));
      const data = {
        ho_ten: form.hoTen.toUpperCase(),
        ngay_sinh: isoToDDMMYYYY(form.ngaySinh),
        gioi_tinh: form.gioiTinh,
        so_cccd: form.soCccd,
        dia_chi_1: form.diaChi1,
        xa_phuong_1: form.xaPhuong1,
        tinh_tp_1: form.tinhTp1,
        dien_thoai: form.dienThoai,
        email: form.email,
        ten_cong_ty: form.tenCongTy.toUpperCase(),
        dia_chi_2: form.diaChi2,
        xa_phuong_2: form.xaPhuong2,
        tinh_tp_2: form.tinhTp2,
        von_dieu_le: Number(form.vonDieuLe).toLocaleString("vi-VN"),
        von_bang_chu: vonBangChu,
        ngay_lap: String(now.getDate()).padStart(2, "0"),
        thang_lap: String(now.getMonth() + 1).padStart(2, "0"),
        nam_lap: String(now.getFullYear()),
        tinh_lap: form.tinhTp2 || "Hà Tĩnh",
        nganh_list: nganhList,
      };
      await generateBusinessRegistrationDoc(data, `GiayDeNghiDKDN_${order.orderCode}.docx`);
      await onSave({
        companyName: form.tenCongTy, capital: Number(form.vonDieuLe) || null,
        ownerDob: form.ngaySinh || null, ownerGender: form.gioiTinh, ownerEmail: form.email,
        ownerProvince: form.tinhTp1, hqAddress: form.diaChi2, hqWard: form.xaPhuong2, hqProvince: form.tinhTp2,
        industries: form.industries,
      });
      onClose();
    } catch (err) {
      setError(err.message || "Có lỗi khi tạo file.");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-2xl w-full p-5 max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-1">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><FileText size={18} className="text-indigo-700" /> Tạo Giấy đề nghị đăng ký doanh nghiệp</p>
          <button onClick={onClose}><XCircle size={18} className="text-slate-400" /></button>
        </div>
        <p className="text-xs text-slate-500 mb-4">Kiểm tra/bổ sung thông tin bên dưới — hệ thống sẽ tự động điền vào mẫu chính thức (Công ty TNHH một thành viên) và tải file Word về máy.</p>

        <p className="text-xs font-semibold text-slate-600 mb-2">Chủ sở hữu / Người đại diện theo pháp luật</p>
        <div className="grid sm:grid-cols-2 gap-3 mb-4">
          <TextField label="Họ tên (viết hoa)" value={form.hoTen} onChange={set("hoTen")} className="sm:col-span-2" />
          <TextField label="Ngày sinh" type="date" value={form.ngaySinh} onChange={set("ngaySinh")} />
          <SelectField label="Giới tính" value={form.gioiTinh} onChange={set("gioiTinh")}>
            <option value="Nam">Nam</option>
            <option value="Nữ">Nữ</option>
          </SelectField>
          <TextField label="Số CCCD" value={form.soCccd} onChange={set("soCccd")} />
          <TextField label="Điện thoại" value={form.dienThoai} onChange={set("dienThoai")} />
          <TextField label="Email" value={form.email} onChange={set("email")} className="sm:col-span-2" />
          <TextField label="Địa chỉ liên lạc (số nhà, đường...)" value={form.diaChi1} onChange={set("diaChi1")} className="sm:col-span-2" />
          <TextField label="Xã/Phường" value={form.xaPhuong1} onChange={set("xaPhuong1")} />
          <TextField label="Tỉnh/Thành phố" value={form.tinhTp1} onChange={set("tinhTp1")} />
        </div>

        <p className="text-xs font-semibold text-slate-600 mb-2">Thông tin công ty</p>
        <div className="grid sm:grid-cols-2 gap-3 mb-4">
          <TextField label="Tên công ty (tiếng Việt, viết hoa)" value={form.tenCongTy} onChange={set("tenCongTy")} className="sm:col-span-2" />
          <TextField label="Địa chỉ trụ sở chính" value={form.diaChi2} onChange={set("diaChi2")} className="sm:col-span-2" />
          <TextField label="Xã/Phường (trụ sở)" value={form.xaPhuong2} onChange={set("xaPhuong2")} />
          <TextField label="Tỉnh/Thành phố (trụ sở)" value={form.tinhTp2} onChange={set("tinhTp2")} />
          <TextField label="Vốn điều lệ (VNĐ)" type="number" value={form.vonDieuLe} onChange={set("vonDieuLe")} className="sm:col-span-2" />
          {vonBangChu && <p className="text-xs text-slate-500 sm:col-span-2 -mt-2">Bằng chữ: <span className="italic">{vonBangChu}</span></p>}
        </div>

        <p className="text-xs font-semibold text-slate-600 mb-2">Ngành, nghề kinh doanh</p>
        <div className="mb-4">
          <IndustryListEditor industries={form.industries} onChange={(list) => setForm((f) => ({ ...f, industries: list }))} />
        </div>

        {error && <p className="text-xs text-rose-600 mb-3 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}

        <div className="flex justify-end gap-2">
          <GhostButton onClick={onClose} disabled={generating}>Huỷ</GhostButton>
          <PrimaryButton onClick={submit} disabled={generating}>
            {generating ? <Loader2 size={15} className="animate-spin" /> : <Download size={15} />} Tạo &amp; tải file Word
          </PrimaryButton>
        </div>
      </div>
    </div>
  );
}

// Nút mở form tạo hồ sơ — chỉ hiện với đơn "Mở Công ty"
// Form thu thập/điền lại thông tin cho "Giấy đề nghị đăng ký Hộ kinh doanh +
// Giấy uỷ quyền" (đơn "Mở HKD") và tải về.
function HouseholdRegModal({ order, customer, currentUser, onSave, onClose }) {
  const [form, setForm] = useState({
    hoTen: (customer?.name || "").toUpperCase(),
    ngaySinh: order.ownerDob || "",
    gioiTinh: order.ownerGender || "Nam",
    soCccd: customer?.cccd || "",
    dienThoai: customer?.phone || "",
    tenHoKd: order.companyName || "",
    diaChiTruSo: order.hqAddress || customer?.address || "",
    phuong: order.hqWard || customer?.ward || "",
    tinhTp: order.hqProvince || "Hà Tĩnh",
    industries: order.industries?.length ? order.industries : (customer?.industries || []),
    vonKinhDoanh: order.capital ?? "",
    diaChiCaNhan: customer?.address || "",
    phuongCaNhan: customer?.ward || "",
    tinhTpCaNhan: order.ownerProvince || "Hà Tĩnh",
  });
  const [uqIndex, setUqIndex] = useState(() => {
    const matched = findUyQuyenProfile(currentUser?.name);
    const idx = matched ? UY_QUYEN_PROFILES.indexOf(matched) : 0;
    return idx >= 0 ? idx : 0;
  });
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const uyQuyen = UY_QUYEN_PROFILES[uqIndex];

  const vonBangChu = form.vonKinhDoanh ? soThanhChuTien(form.vonKinhDoanh) : "";

  const submit = async () => {
    if (!form.hoTen || !form.tenHoKd || !form.vonKinhDoanh) {
      setError("Vui lòng điền đủ Họ tên, Tên hộ kinh doanh và Vốn kinh doanh.");
      return;
    }
    if (form.industries.length === 0) {
      setError("Vui lòng thêm ít nhất 1 ngành nghề kinh doanh.");
      return;
    }
    setError("");
    setGenerating(true);
    try {
      const now = new Date();
      const nganhList = form.industries.map((ind, i) => ({
        stt: String(i + 1),
        ten: ind.detail.trim() ? `${ind.name}\nChi tiết: ${ind.detail.trim()}` : ind.name,
        ma: ind.code,
        chinh: ind.isPrimary ? "X" : "",
      }));
      const data = {
        ngay_lap: String(now.getDate()).padStart(2, "0"),
        thang_lap: String(now.getMonth() + 1).padStart(2, "0"),
        nam_lap: String(now.getFullYear()),
        tinh_tp: form.tinhTp || "Hà Tĩnh",
        phuong: form.phuong,
        ho_ten: form.hoTen.toUpperCase(),
        ngay_sinh_ngay: form.ngaySinh ? String(new Date(form.ngaySinh).getDate()).padStart(2, "0") : "",
        ngay_sinh_thang: form.ngaySinh ? String(new Date(form.ngaySinh).getMonth() + 1).padStart(2, "0") : "",
        ngay_sinh_nam: form.ngaySinh ? String(new Date(form.ngaySinh).getFullYear()) : "",
        gioi_tinh: form.gioiTinh,
        so_cccd: form.soCccd,
        dien_thoai: form.dienThoai,
        ten_ho_kd: form.tenHoKd.toUpperCase(),
        dia_chi_tru_so: form.diaChiTruSo,
        nganh_list: nganhList,
        von_kinh_doanh: Number(form.vonKinhDoanh).toLocaleString("vi-VN"),
        von_bang_chu: vonBangChu,
        dia_chi_ca_nhan: form.diaChiCaNhan,
        phuong_ca_nhan: form.phuongCaNhan,
        tinh_tp_ca_nhan: form.tinhTpCaNhan || "Hà Tĩnh",
        uq_ho_ten: uyQuyen.hoTen,
        uq_gioi_tinh: uyQuyen.gioiTinh,
        uq_ngay_sinh: uyQuyen.ngaySinh,
        uq_cccd: uyQuyen.cccd,
        uq_dia_chi: uyQuyen.diaChi,
        uq_dien_thoai: uyQuyen.dienThoai,
        uq_email: uyQuyen.email || "",
      };
      await generateHouseholdBusinessDoc(data, `GiayDeNghiHKD_${order.orderCode}.docx`);
      await onSave({
        companyName: form.tenHoKd, capital: Number(form.vonKinhDoanh) || null,
        ownerDob: form.ngaySinh || null, ownerGender: form.gioiTinh, ownerEmail: order.ownerEmail,
        ownerProvince: form.tinhTpCaNhan, hqAddress: form.diaChiTruSo, hqWard: form.phuong, hqProvince: form.tinhTp,
        industries: form.industries,
      });
      onClose();
    } catch (err) {
      setError(err.message || "Có lỗi khi tạo file.");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-2xl w-full p-5 max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-1">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><FileText size={18} className="text-indigo-700" /> Tạo Giấy đề nghị đăng ký HKD + Giấy uỷ quyền</p>
          <button onClick={onClose}><XCircle size={18} className="text-slate-400" /></button>
        </div>
        <p className="text-xs text-slate-500 mb-4">Kiểm tra/bổ sung thông tin bên dưới — hệ thống sẽ tự động điền vào cả 2 mẫu (Giấy đề nghị đăng ký Hộ kinh doanh + Giấy uỷ quyền) và tải file Word về máy. Phần "Bên nhận uỷ quyền" đã cố định sẵn theo nhân sự công ty.</p>

        <p className="text-xs font-semibold text-slate-600 mb-2">Chủ hộ kinh doanh</p>
        <div className="grid sm:grid-cols-2 gap-3 mb-4">
          <TextField label="Họ tên (viết hoa)" value={form.hoTen} onChange={set("hoTen")} className="sm:col-span-2" />
          <TextField label="Ngày sinh" type="date" value={form.ngaySinh} onChange={set("ngaySinh")} />
          <SelectField label="Giới tính" value={form.gioiTinh} onChange={set("gioiTinh")}>
            <option value="Nam">Nam</option>
            <option value="Nữ">Nữ</option>
          </SelectField>
          <TextField label="Số CCCD" value={form.soCccd} onChange={set("soCccd")} />
          <TextField label="Điện thoại" value={form.dienThoai} onChange={set("dienThoai")} />
          <TextField label="Địa chỉ liên lạc cá nhân (số nhà, đường...)" value={form.diaChiCaNhan} onChange={set("diaChiCaNhan")} className="sm:col-span-2" />
          <TextField label="Xã/Phường (cá nhân)" value={form.phuongCaNhan} onChange={set("phuongCaNhan")} />
          <TextField label="Tỉnh/Thành phố (cá nhân)" value={form.tinhTpCaNhan} onChange={set("tinhTpCaNhan")} />
        </div>

        <p className="text-xs font-semibold text-slate-600 mb-2">Hộ kinh doanh</p>
        <div className="grid sm:grid-cols-2 gap-3 mb-4">
          <TextField label="Tên hộ kinh doanh (không cần ghi 'HỘ KINH DOANH')" value={form.tenHoKd} onChange={set("tenHoKd")} className="sm:col-span-2" />
          <TextField label="Địa chỉ trụ sở" value={form.diaChiTruSo} onChange={set("diaChiTruSo")} className="sm:col-span-2" />
          <TextField label="Xã/Phường (trụ sở, nơi đăng ký)" value={form.phuong} onChange={set("phuong")} />
          <TextField label="Tỉnh/Thành phố (trụ sở)" value={form.tinhTp} onChange={set("tinhTp")} />
          <TextField label="Vốn kinh doanh (VNĐ)" type="number" value={form.vonKinhDoanh} onChange={set("vonKinhDoanh")} className="sm:col-span-2" />
          {vonBangChu && <p className="text-xs text-slate-500 sm:col-span-2 -mt-2">Bằng chữ: <span className="italic">{vonBangChu}</span></p>}
        </div>

        <p className="text-xs font-semibold text-slate-600 mb-2">Ngành, nghề kinh doanh</p>
        <div className="mb-4">
          <IndustryListEditor industries={form.industries} onChange={(list) => setForm((f) => ({ ...f, industries: list }))} />
        </div>

        <p className="text-xs font-semibold text-slate-600 mb-2">Người nhận uỷ quyền (nhân sự công ty nộp hồ sơ)</p>
        <div className="mb-4">
          <SelectField label="Chọn người nhận uỷ quyền" value={uqIndex} onChange={(e) => setUqIndex(Number(e.target.value))}>
            {UY_QUYEN_PROFILES.map((p, i) => <option key={p.cccd} value={i}>{p.hoTen}</option>)}
          </SelectField>
          <p className="text-[11px] text-slate-400 mt-1">Tự động chọn theo tài khoản đang đăng nhập — có thể đổi lại nếu người khác đứng ra nộp hồ sơ.</p>
        </div>

        {error && <p className="text-xs text-rose-600 mb-3 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}

        <div className="flex justify-end gap-2">
          <GhostButton onClick={onClose} disabled={generating}>Huỷ</GhostButton>
          <PrimaryButton onClick={submit} disabled={generating}>
            {generating ? <Loader2 size={15} className="animate-spin" /> : <Download size={15} />} Tạo &amp; tải file Word
          </PrimaryButton>
        </div>
      </div>
    </div>
  );
}

// Nút mở form tạo hồ sơ — chọn đúng mẫu theo loại thủ tục ("Mở Công ty" hoặc
// "Mở HKD"); không hiện với các loại thủ tục khác.
function BusinessRegDocButton({ order, customer, currentUser, onSave, className = "" }) {
  const [open, setOpen] = useState(false);
  if (order.procedureType !== "mo_cty" && order.procedureType !== "mo_hkd") return null;
  const label = order.procedureType === "mo_cty" ? "Tạo Giấy đề nghị ĐKDN" : "Tạo Giấy đề nghị ĐKHKD";
  return (
    <>
      <GhostButton className={`!py-1.5 !text-xs ${className}`} onClick={() => setOpen(true)}>
        <FileText size={13} /> {label}
      </GhostButton>
      {open && order.procedureType === "mo_cty" && (
        <BusinessRegModal order={order} customer={customer} onSave={onSave} onClose={() => setOpen(false)} />
      )}
      {open && order.procedureType === "mo_hkd" && (
        <HouseholdRegModal order={order} customer={customer} currentUser={currentUser} onSave={onSave} onClose={() => setOpen(false)} />
      )}
    </>
  );
}

// ---------------------------------------------------------------------------
// NHÂN VIÊN — PHÂN HỆ 3: TRẠNG THÁI ĐƠN HÀNG
// ---------------------------------------------------------------------------

function OrderStatusRow({ order, customer, currentUser, onAdvance, onMarkRejected, onRetry, onSendLicense, onSaveCompanyInfo, isNew }) {
  const [reason, setReason] = useState(order.overdueReason || "");
  const [confirmStep, setConfirmStep] = useState(null);
  const [stepMsg, setStepMsg] = useState("");
  const [uploadingLicense, setUploadingLicense] = useState(false);
  const [licenseFile, setLicenseFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const currentIndex = STATUS_FLOW.indexOf(order.status);
  const nextStatus = STATUS_FLOW[currentIndex + 1];
  const isRejected = order.status === "chua_duoc_chap_thuan";

  React.useEffect(() => {
    if (!stepMsg) return;
    const t = setTimeout(() => setStepMsg(""), 3000);
    return () => clearTimeout(t);
  }, [stepMsg]);

  const handleStepClick = (step, index) => {
    if (busy) return;
    if (index === currentIndex + 1) {
      if (step === "duoc_chap_thuan") setConfirmStep(step);
      else if (step === "gui_giay_phep") setUploadingLicense(true);
      else runAsync(() => onAdvance(order.id, step));
    } else if (index <= currentIndex) {
      setStepMsg(`Bước "${ORDER_STATUS[step].label}" đã hoàn tất trước đó — không thể sửa lại.`);
    } else {
      setStepMsg(`Cần thực hiện lần lượt từng bước. Bước tiếp theo cần chọn là "${ORDER_STATUS[nextStatus]?.label}".`);
    }
  };

  const runAsync = async (fn) => {
    setBusy(true);
    await fn();
    setBusy(false);
  };

  const handlePickPdf = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLicenseFile({ name: file.name, raw: file });
  };

  const submitRejection = () => {
    if (!reason.trim() || busy) return;
    runAsync(() => onMarkRejected(order.id, reason));
  };

  if (order.status === "hoan_thanh" || order.status === "da_thanh_toan") {
    const paid = order.status === "da_thanh_toan";
    return (
      <div className={`p-4 rounded-xl border ${paid ? "border-emerald-300 bg-emerald-50/40" : "border-indigo-200 bg-indigo-50/40"}`}>
        <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
          <div>
            <p className="text-sm font-semibold text-slate-800">{order.orderCode} · {procedureLabel(order.procedureType)}</p>
            <p className="text-xs text-slate-500">{customer?.name || customer?.phone || "?"} · Tạo ngày {fmtDate(order.createdAt)}</p>
          </div>
          <StatusBadge status={order.status} />
        </div>
        <div className="space-y-1.5 text-xs">
          <p className="flex items-center gap-1.5 text-slate-500"><CheckCircle2 size={13} className="text-emerald-600" /> Đã gửi giấy phép lên nhóm{order.licenseSentAt ? ` (${fmtDate(order.licenseSentAt)})` : ""}</p>
          <p className="flex items-center gap-1.5 text-slate-500"><CheckCircle2 size={13} className="text-emerald-600" /> Quản lý đã xác nhận đơn hàng{order.confirmedAt ? ` (${fmtDate(order.confirmedAt)})` : ""}</p>
          <p className={`flex items-center gap-1.5 ${paid ? "text-slate-500" : "text-indigo-600"}`}>
            {paid ? <CheckCircle2 size={13} className="text-emerald-600" /> : <Clock size={13} />}
            {paid ? `Đã thanh toán${order.completedAt ? ` (${fmtDate(order.completedAt)})` : ""}` : "Đang chờ quản lý cập nhật doanh thu / chi phí / tiền công"}
          </p>
        </div>
        {paid && (
          <div className="mt-3 pt-3 border-t border-emerald-100">
            <p className="text-xs text-slate-500">Tiền công của bạn cho đơn này:</p>
            <p className="text-base font-semibold text-emerald-700">{fmtMoney(order.laborFee)}</p>
          </div>
        )}
      </div>
    );
  }

  if (isRejected) {
    return (
      <div className="p-4 rounded-xl border border-rose-300 bg-rose-50/50">
        <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
          <div>
            <p className="text-sm font-semibold text-slate-800">{order.orderCode} · {procedureLabel(order.procedureType)}</p>
            <p className="text-xs text-slate-500">{customer?.name || customer?.phone || "?"} · Tạo ngày {fmtDate(order.createdAt)}</p>
          </div>
          <StatusBadge status={order.status} />
        </div>
        <div className="p-3 rounded-lg bg-white border border-rose-200 mb-3">
          <p className="text-xs font-medium text-rose-700 flex items-center gap-1 mb-1"><XCircle size={13} /> Lý do chưa được chấp thuận</p>
          <p className="text-sm text-slate-700">{order.overdueReason || "—"}</p>
        </div>
        <p className="text-xs text-slate-500 mb-2">Cảnh báo đã được gửi tới tài khoản quản lý. Nếu vấn đề đã được khắc phục, bạn có thể trình lại lãnh đạo.</p>
        <GhostButton className="!text-xs" disabled={busy} onClick={() => runAsync(() => onRetry(order.id))}>
          {busy ? <Loader2 size={13} className="animate-spin" /> : <ArrowLeft size={13} className="rotate-180" />} Trình lại lãnh đạo
        </GhostButton>
      </div>
    );
  }

  return (
    <div className={`p-4 rounded-xl border transition ${isNew ? "border-emerald-300 bg-emerald-50/50 ring-2 ring-emerald-200" : order.approvalOverdue ? "border-rose-300 bg-rose-50/40" : "border-slate-200"}`}>
      {isNew && (
        <p className="text-xs font-medium text-emerald-700 flex items-center gap-1 mb-2">
          <CheckCircle2 size={13} /> Vừa tạo thành công
        </p>
      )}
      <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
        <div>
          <p className="text-sm font-semibold text-slate-800">{order.orderCode} · {procedureLabel(order.procedureType)}</p>
          <p className="text-xs text-slate-500">{customer?.name || customer?.phone || "?"} · Tạo ngày {fmtDate(order.createdAt)}</p>
        </div>
        <div className="flex items-center gap-2">
          <BusinessRegDocButton order={order} customer={customer} currentUser={currentUser} onSave={(fields) => onSaveCompanyInfo(order.id, fields)} />
          <StatusBadge status={order.status} />
        </div>
      </div>

      <OrderProgressStepper status={order.status} onStepClick={handleStepClick} />

      {stepMsg && (
        <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-2.5 py-1.5 mb-2 flex items-center gap-1.5">
          <AlertTriangle size={12} className="shrink-0" /> {stepMsg}
        </p>
      )}

      {order.receivedAt && (
        <p className="text-xs text-slate-500 mb-2 flex items-center gap-1">
          <Clock size={12} /> Hạn chấp thuận (3 ngày làm việc kể từ khi tiếp nhận): <span className="font-medium text-slate-700">{fmtDate(order.approvalDeadline)}</span>
        </p>
      )}

      {order.status === "trinh_lanh_dao" && (
        <div className={`mb-2 p-3 rounded-lg border ${order.approvalOverdue ? "bg-rose-100/70 border-rose-200" : "bg-slate-50 border-slate-200"}`}>
          <p className="text-xs font-medium text-rose-700 flex items-center gap-1 mb-2">
            <AlertTriangle size={13} />
            {order.isLatePenalty
              ? `Đã tự động ghi nhận CHẬM TIẾN ĐỘ và trừ ${fmtMoney(LATE_PENALTY_AMOUNT)} tiền công do quá 3 ngày làm việc mà chưa được chấp thuận và chưa điền lý do.`
              : "Nếu lãnh đạo không duyệt đơn này, điền lý do bên dưới rồi bấm xác nhận — đơn sẽ chuyển sang trạng thái \"Chưa được chấp thuận\" và gửi cảnh báo cho quản lý."}
          </p>
          <div className="flex gap-2">
            <input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Lý do chưa được chấp thuận..." className="flex-1 px-3 py-1.5 rounded-lg border border-rose-200 text-xs focus:outline-none focus:ring-2 focus:ring-rose-400/40" />
            <GhostButton className="!py-1.5 !text-xs !border-rose-300 !text-rose-700" disabled={!reason.trim() || busy} onClick={submitRejection}>
              <XCircle size={13} /> Chưa được chấp thuận
            </GhostButton>
          </div>
        </div>
      )}

      {order.status === "duoc_chap_thuan" && (
        <p className="text-xs text-slate-500 flex items-center gap-1">
          <Clock size={12} /> Tiếp theo: bấm bước 5 trên thanh tiến độ để gửi giấy phép bản PDF lên nhóm làm việc.
        </p>
      )}

      {order.status === "gui_giay_phep" && (
        <p className="text-xs text-teal-700 flex items-center gap-1">
          <CheckCircle2 size={13} /> Đã gửi giấy phép lên nhóm{order.licensePdfName ? ` (${order.licensePdfName})` : ""} — đang chờ quản lý tải xuống và xác nhận đơn hàng.
        </p>
      )}

      {confirmStep && (
        <ConfirmDialog
          title="Xác nhận chấp thuận đơn hàng"
          message={`Đơn ${order.orderCode} sẽ chuyển sang "Được chấp thuận". Sau khi xác nhận, đơn hàng KHÔNG THỂ sửa lại trạng thái này nữa. Bạn có chắc chắn không?`}
          confirmLabel="Đúng, xác nhận chấp thuận"
          danger
          onCancel={() => setConfirmStep(null)}
          onConfirm={() => { runAsync(() => onAdvance(order.id, "duoc_chap_thuan")); setConfirmStep(null); }}
        />
      )}

      {uploadingLicense && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
            <div className="flex items-center gap-2 mb-2">
              <FileText size={18} className="text-teal-700" />
              <p className="font-semibold text-slate-800">Gửi giấy phép lên nhóm</p>
            </div>
            <p className="text-sm text-slate-600 mb-3">Tải lên file PDF giấy phép của đơn {order.orderCode} để xác nhận đã gửi lên nhóm làm việc.</p>
            <label className="flex items-center gap-3 p-3 rounded-xl border border-dashed border-slate-300 hover:border-teal-400 hover:bg-teal-50/40 transition cursor-pointer mb-4">
              <span className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center shrink-0">
                <FileText size={17} className="text-slate-400" />
              </span>
              <span className="text-xs text-slate-500 truncate">{licenseFile ? licenseFile.name : "Bấm để chọn file PDF..."}</span>
              <input type="file" accept="application/pdf" className="hidden" onChange={handlePickPdf} />
            </label>
            <div className="flex justify-end gap-2">
              <GhostButton disabled={busy} onClick={() => { setUploadingLicense(false); setLicenseFile(null); }}>Huỷ</GhostButton>
              <button
                disabled={!licenseFile || busy}
                onClick={async () => {
                  setBusy(true);
                  await onSendLicense(order.id, licenseFile);
                  setBusy(false);
                  setUploadingLicense(false);
                  setLicenseFile(null);
                }}
                className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-white text-sm font-medium bg-teal-700 hover:bg-teal-800 transition disabled:opacity-40 disabled:cursor-not-allowed"
              >
                {busy ? <Loader2 size={14} className="animate-spin" /> : null} Xác nhận đã gửi
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function OrderStatusModule({ currentUser, orders, customers, highlightOrderId, onAdvance, onMarkRejected, onRetry, onSendLicense, onSaveCompanyInfo }) {
  const derived = useDerivedOrders(orders);
  const mine = derived.filter((o) => o.employeeId === currentUser.id).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const overdueCount = mine.filter((o) => o.approvalOverdue).length;

  return (
    <div className="space-y-5">
      <SectionTitle icon={ClipboardCheck} title="Trạng thái đơn hàng" subtitle="Cập nhật các đơn hàng bạn được giao" />
      {overdueCount > 0 && (
        <Card className="p-3.5 bg-rose-50 border-rose-200">
          <p className="text-sm text-rose-700 flex items-center gap-2"><AlertTriangle size={15} /> Bạn có {overdueCount} đơn hàng quá hạn 3 ngày làm việc chưa được chấp thuận.</p>
        </Card>
      )}
      {mine.length === 0 ? (
        <Card className="p-4"><EmptyState icon={Inbox} text="Chưa có đơn hàng nào." /></Card>
      ) : (
        <div className="space-y-3">
          {mine.map((o) => (
            <OrderStatusRow
              key={o.id}
              order={o}
              customer={customers.find((c) => c.id === o.customerId)}
              currentUser={currentUser}
              onAdvance={onAdvance}
              onMarkRejected={onMarkRejected}
              onRetry={onRetry}
              onSendLicense={onSendLicense}
              onSaveCompanyInfo={onSaveCompanyInfo}
              isNew={o.id === highlightOrderId}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// NHÂN VIÊN — PHÂN HỆ 4: BÁO CÁO
// ---------------------------------------------------------------------------
function EmployeeReportModule({ currentUser, orders }) {
  const derived = useDerivedOrders(orders);
  const mine = derived.filter((o) => o.employeeId === currentUser.id);
  const statusData = [...STATUS_FLOW, "chua_duoc_chap_thuan"].map((s) => ({ name: ORDER_STATUS[s].label, value: mine.filter((o) => o.status === s).length })).filter((d) => d.value > 0);
  const laborRows = mine.filter((o) => o.laborFee != null || o.latePenaltyAmount > 0);
  const totalBase = laborRows.reduce((s, o) => s + (o.laborFee || 0), 0);
  const totalPenalty = mine.reduce((s, o) => s + (o.latePenaltyAmount || 0), 0);
  const totalLabor = totalBase - totalPenalty;

  return (
    <div className="space-y-5">
      <SectionTitle icon={BarChart3} title="Báo cáo" subtitle="Tình trạng đơn hàng & tiền công" />
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <MetricCard label="Tổng số đơn" value={mine.length} icon={ClipboardList} accent="indigo" />
        <MetricCard label="Tiền công thực nhận" value={fmtMoney(totalLabor)} icon={Wallet} accent="amber" />
        <MetricCard label="Bị trừ do chậm tiến độ" value={fmtMoney(totalPenalty)} icon={AlertTriangle} accent="rose" />
      </div>
      <Card className="p-4 sm:p-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Báo cáo tình trạng đơn hàng</p>
        {statusData.length === 0 ? <EmptyState icon={BarChart3} text="Chưa có dữ liệu." /> : (
          <div style={{ width: "100%", height: 240 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={statusData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
                  {statusData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Pie>
                <Tooltip {...CHART_TOOLTIP_STYLE} />
                <Legend verticalAlign="middle" align="right" layout="vertical" wrapperStyle={{ fontSize: 12 }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </Card>
      <Card className="p-4 sm:p-5">
        <p className="font-semibold text-slate-800 text-sm mb-1">Báo cáo tiền công theo đơn hàng</p>
        <p className="text-xs text-slate-500 mb-3">Đơn quá hạn 3 ngày làm việc chưa điền lý do sẽ tự động bị trừ {fmtMoney(LATE_PENALTY_AMOUNT)}.</p>
        {laborRows.length === 0 ? <EmptyState icon={Wallet} text="Chưa có đơn hàng nào được ghi nhận tiền công." /> : (
          <div className="space-y-2">
            {laborRows.map((o) => (
              <div key={o.id} className="flex items-center justify-between text-sm py-1.5 border-b border-slate-50 last:border-0">
                <span className="text-slate-700">{o.orderCode}{o.latePenaltyAmount > 0 && <Badge className="ml-2 bg-rose-50 text-rose-700 border-rose-200 !text-[10px] !py-0.5">Chậm tiến độ</Badge>}</span>
                <span className="text-right">
                  {o.laborFee != null && <span className="font-semibold text-slate-800">{fmtMoney(o.laborFee)}</span>}
                  {o.latePenaltyAmount > 0 && <span className="block text-xs text-rose-600">− {fmtMoney(o.latePenaltyAmount)}</span>}
                </span>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ADMIN — PHÂN HỆ 1: KHÁCH HÀNG (toàn bộ)
// ---------------------------------------------------------------------------
function AdminCustomerModule({ customers, employees }) {
  const [q, setQ] = useState("");
  const filtered = customers.filter((c) => !q || c.phone.includes(q) || (c.cccd || "").includes(q) || (c.name || "").toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="space-y-5">
      <SectionTitle icon={Users} title="Khách hàng" subtitle="Toàn bộ khách hàng đã giao cho nhân viên" />
      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex items-center gap-2">
          <Search size={15} className="text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm theo tên, số điện thoại hoặc CCCD..." className="flex-1 text-sm outline-none" />
        </div>
        {filtered.length === 0 ? <EmptyState icon={Users} text="Chưa có khách hàng." /> : (
          <div className="divide-y divide-slate-100">
            {filtered.map((c) => (
              <div key={c.id} className="p-4 flex items-center gap-3">
                <span className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center shrink-0 overflow-hidden">
                  {c.vnidPhoto ? <img src={c.vnidPhoto} alt="" className="w-full h-full object-cover" /> : <Users size={16} className="text-slate-400" />}
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium text-slate-800 truncate">{c.name || "(Chưa đặt tên)"} · {c.phone}</p>
                  <p className="text-xs text-slate-400 truncate">CCCD: {c.cccd || "—"} · {c.address}, {c.ward} · {c.industry || "—"}</p>
                  {c.referrer && <p className="text-xs text-indigo-500 truncate">Người giới thiệu: {c.referrer}</p>}
                </div>
                <Badge className="bg-indigo-50 text-indigo-700 border-indigo-200 shrink-0">
                  {employees.find((e) => e.id === c.employeeId)?.name || "?"}
                </Badge>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ADMIN — PHÂN HỆ 2: ĐƠN HÀNG (duyệt giấy phép / doanh thu / chi phí / tiền công)
// ---------------------------------------------------------------------------
function ApprovalForm({ order, onSave }) {
  const [revenue, setRevenue] = useState(order.revenue ?? "");
  const [cost, setCost] = useState(order.cost ?? "");
  const [laborFee, setLaborFee] = useState(order.laborFee ?? "");
  const [saving, setSaving] = useState(false);
  const submit = async () => {
    if (saving) return;
    setSaving(true);
    await onSave(order.id, Number(revenue) || 0, Number(cost) || 0, Number(laborFee) || 0);
    setSaving(false);
  };
  return (
    <div className="mt-3 pt-3 border-t border-slate-100">
      <p className="text-xs text-slate-500 mb-2">Điền đầy đủ 3 trường bên dưới — đơn hàng sẽ tự động chuyển sang trạng thái <span className="font-medium text-slate-700">"Đã thanh toán"</span> sau khi ghi nhận.</p>
      <div className="flex flex-wrap items-end gap-3">
        <TextField label="Doanh thu (đ)" type="number" value={revenue} onChange={(e) => setRevenue(e.target.value)} />
        <TextField label="Chi phí (đ)" type="number" value={cost} onChange={(e) => setCost(e.target.value)} />
        <TextField label="Tiền công (đ)" type="number" value={laborFee} onChange={(e) => setLaborFee(e.target.value)} />
        <PrimaryButton className="!py-1.5 !text-xs" disabled={saving} onClick={submit}>
          {saving ? <Loader2 size={13} className="animate-spin" /> : <DollarSign size={13} />} Ghi nhận &amp; đánh dấu đã thanh toán
        </PrimaryButton>
      </div>
    </div>
  );
}

function AdminOrderModule({ orders, customers, employees, currentUser, onConfirmOrder, onSaveFinance, onSaveCompanyInfo }) {
  const derived = useDerivedOrders(orders);
  const sorted = [...derived].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const [confirmingId, setConfirmingId] = useState(null);

  const handleConfirm = async (id) => {
    setConfirmingId(id);
    await onConfirmOrder(id);
    setConfirmingId(null);
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={ClipboardList} title="Đơn hàng" subtitle="Toàn bộ đơn hàng của nhân viên" />
      {sorted.length === 0 ? <Card className="p-4"><EmptyState icon={Inbox} text="Chưa có đơn hàng." /></Card> : (
        <div className="space-y-3">
          {sorted.map((o) => {
            const cust = customers.find((c) => c.id === o.customerId);
            const emp = employees.find((e) => e.id === o.employeeId);
            return (
              <Card key={o.id} className={`p-4 ${o.paymentOverdue ? "border-rose-300 bg-rose-50/30" : ""}`}>
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <p className="text-sm font-semibold text-slate-800">{o.orderCode} · {procedureLabel(o.procedureType)}</p>
                    <p className="text-xs text-slate-500">{cust?.name || cust?.phone} · NV: {emp?.name} · Tạo {fmtDate(o.createdAt)}</p>
                  </div>
                  <div className="flex items-center gap-2">
                    <BusinessRegDocButton order={o} customer={cust} currentUser={currentUser} onSave={(fields) => onSaveCompanyInfo(o.id, fields)} />
                    <StatusBadge status={o.status} />
                  </div>
                </div>

                {o.paymentOverdue && (
                  <p className="mt-2 text-xs font-medium text-rose-700 flex items-center gap-1">
                    <AlertTriangle size={13} /> Quá 2 ngày làm việc kể từ khi được chấp thuận mà chưa hoàn tất thanh toán — cần thu hồi công nợ.
                  </p>
                )}

                {o.status === "chua_duoc_chap_thuan" && (
                  <p className="mt-2 text-xs font-medium text-rose-700 flex items-center gap-1">
                    <AlertTriangle size={13} /> Chưa được chấp thuận — Lý do: {o.overdueReason || "(nhân viên chưa điền lý do)"}
                  </p>
                )}

                {o.status === "gui_giay_phep" && (
                  <div className="mt-3 pt-3 border-t border-slate-100">
                    <p className="text-xs text-teal-700 flex items-center gap-1 mb-2">
                      <FileText size={13} /> Nhân viên đã gửi giấy phép{o.licenseSentAt ? ` (${fmtDate(o.licenseSentAt)})` : ""}. Xem file trước khi xác nhận.
                    </p>
                    <div className="flex flex-wrap items-center gap-2">
                      {o.licensePdfData ? (
                        <button
                          type="button"
                          onClick={() => window.open(o.licensePdfData, "_blank")}
                          className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl border border-teal-300 bg-white text-teal-700 text-xs font-medium hover:bg-teal-50 transition"
                        >
                          <Download size={13} /> Xem / tải {o.licensePdfName || "giấy phép"}
                        </button>
                      ) : (
                        <span className="text-xs text-slate-400">(Không có file — nhân viên chưa tải lên thành công)</span>
                      )}
                      <PrimaryButton className="!py-1.5 !text-xs" disabled={confirmingId === o.id} onClick={() => handleConfirm(o.id)}>
                        {confirmingId === o.id ? <Loader2 size={13} className="animate-spin" /> : <CheckCircle2 size={13} />} Xác nhận đơn hàng
                      </PrimaryButton>
                    </div>
                  </div>
                )}

                {o.status === "hoan_thanh" && (
                  <>
                    <p className="mt-3 text-xs text-indigo-700 flex items-center gap-1">
                      <CheckCircle2 size={13} /> Đã xác nhận đơn hàng{o.confirmedAt ? ` (${fmtDate(o.confirmedAt)})` : ""} — thông báo đã gửi cho nhân viên. Chưa cập nhật doanh thu / chi phí / tiền công.
                    </p>
                    <ApprovalForm order={o} onSave={onSaveFinance} />
                  </>
                )}

                {o.status === "da_thanh_toan" && (
                  <div className="mt-3 pt-3 border-t border-slate-100 flex flex-wrap items-center gap-4">
                    <span className="text-xs text-slate-500">Doanh thu: <span className="font-semibold text-slate-800">{fmtMoney(o.revenue)}</span></span>
                    <span className="text-xs text-slate-500">Chi phí: <span className="font-semibold text-slate-800">{fmtMoney(o.cost)}</span></span>
                    <span className="text-xs text-slate-500">Tiền công: <span className="font-semibold text-slate-800">{fmtMoney(o.laborFee)}</span></span>
                    <Badge className="bg-emerald-50 text-emerald-700 border-emerald-200">Đã thanh toán {fmtDate(o.completedAt)}</Badge>
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// ADMIN — PHÂN HỆ 3: CHI PHÍ PHÁT SINH HÀNG NGÀY
// ---------------------------------------------------------------------------
function AdminExpenseModule({ expenses, onAddExpense }) {
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [desc, setDesc] = useState("");
  const [amount, setAmount] = useState("");
  const [saving, setSaving] = useState(false);
  const submit = async () => {
    if (!desc || !amount || saving) return;
    setSaving(true);
    await onAddExpense({ date, description: desc, amount: Number(amount) });
    setSaving(false);
    setDesc(""); setAmount("");
  };
  const total = expenses.reduce((s, e) => s + e.amount, 0);
  return (
    <div className="space-y-5">
      <SectionTitle icon={Receipt} title="Chi phí phát sinh" subtitle="Cập nhật chi phí hoạt động hàng ngày" />
      <MetricCard label="Tổng chi phí phát sinh đã ghi nhận" value={fmtMoney(total)} icon={Receipt} accent="rose" />
      <Card className="p-4 sm:p-5">
        <div className="flex flex-wrap items-end gap-3">
          <TextField label="Ngày" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          <TextField label="Nội dung" value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="VD: Văn phòng phẩm, xăng xe..." className="flex-1 min-w-[180px]" />
          <TextField label="Số tiền (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
          <PrimaryButton onClick={submit} disabled={saving}>
            {saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Ghi nhận
          </PrimaryButton>
        </div>
      </Card>
      <Card className="p-0 overflow-hidden">
        {expenses.length === 0 ? <EmptyState icon={Receipt} text="Chưa có chi phí nào được ghi nhận." /> : (
          <div className="divide-y divide-slate-100">
            {expenses.map((e) => (
              <div key={e.id} className="p-3.5 flex items-center justify-between text-sm">
                <div>
                  <p className="text-slate-700">{e.description}</p>
                  <p className="text-xs text-slate-400">{fmtDate(e.date)}</p>
                </div>
                <span className="font-semibold text-slate-800">{fmtMoney(e.amount)}</span>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ADMIN — PHÂN HỆ 4: BÁO CÁO TỔNG HỢP
// ---------------------------------------------------------------------------
function AdminReportModule({ orders, expenses }) {
  const completed = orders.filter((o) => o.status === "da_thanh_toan");

  const monthMap = {};
  completed.forEach((o) => {
    const k = monthKey(o.completedAt || o.approvedAt);
    monthMap[k] = monthMap[k] || { month: k, revenue: 0, cost: 0 };
    monthMap[k].revenue += o.revenue || 0;
    monthMap[k].cost += (o.cost || 0) + (o.laborFee || 0);
  });
  expenses.forEach((e) => {
    const k = monthKey(e.date);
    monthMap[k] = monthMap[k] || { month: k, revenue: 0, cost: 0 };
    monthMap[k].cost += e.amount;
  });
  const monthlyData = Object.values(monthMap).sort((a, b) => a.month.localeCompare(b.month)).map((m) => ({ ...m, profit: m.revenue - m.cost }));

  const totalRevenue = completed.reduce((s, o) => s + (o.revenue || 0), 0);
  const totalCost = completed.reduce((s, o) => s + (o.cost || 0) + (o.laborFee || 0), 0) + expenses.reduce((s, e) => s + e.amount, 0);
  const totalProfit = totalRevenue - totalCost;
  const debtOrders = orders.filter((o) => o.approvedAt && o.status !== "da_thanh_toan");
  const totalDebt = debtOrders.length;

  return (
    <div className="space-y-5">
      <SectionTitle icon={BarChart3} title="Báo cáo tổng hợp" subtitle="Doanh thu, chi phí, lợi nhuận & công nợ" />
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(totalRevenue)} icon={TrendingUp} accent="indigo" />
        <MetricCard label="Tổng chi phí" value={fmtMoney(totalCost)} icon={Receipt} accent="amber" />
        <MetricCard label="Lợi nhuận" value={fmtMoney(totalProfit)} icon={Wallet} accent="emerald" />
        <MetricCard label="Đơn chưa hoàn tất thanh toán" value={totalDebt} icon={AlertTriangle} accent="rose" />
      </div>

      <Card className="p-4 sm:p-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Tổng hợp doanh thu &amp; lợi nhuận theo tháng</p>
        {monthlyData.length === 0 ? <EmptyState icon={TrendingUp} text="Chưa có dữ liệu." /> : (
          <div style={{ width: "100%", height: 240 }}>
            <ResponsiveContainer>
              <AreaChart data={monthlyData} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
                <defs>
                  <linearGradient id="rev" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#4338ca" stopOpacity={0.3} />
                    <stop offset="100%" stopColor="#4338ca" stopOpacity={0.02} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} />
                <YAxis tickFormatter={shortMoney} tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} width={48} />
                <Tooltip formatter={(v) => fmtMoney(v)} {...CHART_TOOLTIP_STYLE} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Area type="monotone" dataKey="revenue" name="Doanh thu" stroke="#4338ca" fill="url(#rev)" strokeWidth={2.5} />
                <Area type="monotone" dataKey="profit" name="Lợi nhuận" stroke="#059669" fill="transparent" strokeWidth={2} strokeDasharray="4 3" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        )}
      </Card>

      <div className="grid lg:grid-cols-2 gap-5">
        <Card className="p-4 sm:p-5">
          <p className="font-semibold text-slate-800 text-sm mb-3">Chi phí &amp; lợi nhuận theo đơn hàng</p>
          {completed.length === 0 ? <EmptyState icon={Receipt} text="Chưa có đơn hàng được ghi nhận." /> : (
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {completed.map((o) => (
                <div key={o.id} className="flex items-center justify-between text-xs py-1.5 border-b border-slate-50 last:border-0">
                  <span className="text-slate-600">{o.orderCode}</span>
                  <span className="text-slate-500">CP {fmtMoney((o.cost || 0) + (o.laborFee || 0))}</span>
                  <span className="font-semibold text-emerald-700">LN {fmtMoney((o.revenue || 0) - (o.cost || 0) - (o.laborFee || 0))}</span>
                </div>
              ))}
            </div>
          )}
        </Card>
        <Card className="p-4 sm:p-5">
          <p className="font-semibold text-slate-800 text-sm mb-3">Danh sách đơn hàng chưa hoàn tất thanh toán</p>
          {debtOrders.length === 0 ? <EmptyState icon={ShieldCheck} text="Không có đơn nào tồn đọng." /> : (
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {debtOrders.map((o) => (
                <div key={o.id} className="flex items-center justify-between text-xs py-1.5 border-b border-slate-50 last:border-0">
                  <span className="text-slate-600">{o.orderCode}</span>
                  <StatusBadge status={o.status} />
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ADMIN — PHÂN HỆ 5: CHĂM SÓC SAU ĐƠN HÀNG (đơn đã thanh toán)
// ---------------------------------------------------------------------------
function CareModule({ orders, customers, onToggleCareStep }) {
  const completed = orders.filter((o) => o.status === "da_thanh_toan");
  return (
    <div className="space-y-5">
      <SectionTitle icon={HeartHandshake} title="Chăm sóc sau đơn hàng" subtitle="Theo dõi các bước hỗ trợ sau khi đơn hàng hoàn thành" />
      {completed.length === 0 ? <Card className="p-4"><EmptyState icon={HeartHandshake} text="Chưa có đơn hàng hoàn thành." /></Card> : (
        <div className="space-y-3">
          {completed.map((o) => {
            const steps = CARE_STEPS[o.procedureType] || [];
            const cust = customers.find((c) => c.id === o.customerId);
            const doneCount = steps.filter((_, i) => o.careSteps?.[i]).length;
            return (
              <Card key={o.id} className="p-4">
                <div className="flex items-center justify-between mb-3">
                  <div>
                    <p className="text-sm font-semibold text-slate-800">{o.orderCode} · {procedureLabel(o.procedureType)}</p>
                    <p className="text-xs text-slate-500">{cust?.name || cust?.phone}</p>
                  </div>
                  <Badge className={doneCount === steps.length ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-100 text-slate-600 border-slate-200"}>
                    {doneCount}/{steps.length} bước
                  </Badge>
                </div>
                <div className="space-y-1.5">
                  {steps.map((s, i) => (
                    <button key={i} onClick={() => onToggleCareStep(o.id, i, o.careSteps)} className="w-full flex items-center gap-2.5 p-2 rounded-lg hover:bg-slate-50 transition text-left">
                      {o.careSteps?.[i] ? <CheckCircle2 size={17} className="text-emerald-600 shrink-0" /> : <span className="w-[17px] h-[17px] rounded-full border-2 border-slate-300 shrink-0" />}
                      <span className={`text-sm ${o.careSteps?.[i] ? "text-slate-400 line-through" : "text-slate-700"}`}>{s}</span>
                    </button>
                  ))}
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// CHUÔNG THÔNG BÁO
// ---------------------------------------------------------------------------
const NOTIF_META = {
  cong_no: { label: "Công nợ quá hạn", icon: AlertTriangle, color: "text-rose-700" },
  chua_chap_thuan: { label: "Chưa được chấp thuận", icon: AlertTriangle, color: "text-rose-700" },
  qua_han_chap_thuan: { label: "Quá hạn chấp thuận", icon: AlertTriangle, color: "text-amber-700" },
  xac_nhan: { label: "Đơn hàng được xác nhận", icon: CheckCircle2, color: "text-indigo-700" },
  da_thanh_toan: { label: "Đã thanh toán", icon: CheckCircle2, color: "text-emerald-700" },
};
function NotifBell({ notifications }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="relative">
      <button onClick={() => setOpen((o) => !o)} className="relative w-9 h-9 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition">
        <Bell size={16} className="text-slate-600" />
        {notifications.length > 0 && (
          <span className="absolute -top-1 -right-1 w-4 h-4 rounded-full bg-rose-600 text-white text-[10px] font-semibold flex items-center justify-center">{notifications.length}</span>
        )}
      </button>
      {open && (
        <div className="absolute right-0 mt-2 w-80 max-h-96 overflow-y-auto bg-white rounded-xl border border-slate-200 shadow-lg z-20">
          <div className="p-3 border-b border-slate-100 flex items-center justify-between">
            <p className="text-sm font-semibold text-slate-800">Cảnh báo</p>
            <button onClick={() => setOpen(false)}><XCircle size={15} className="text-slate-400" /></button>
          </div>
          {notifications.length === 0 ? (
            <div className="p-6 text-center text-xs text-slate-400">Không có cảnh báo nào.</div>
          ) : (
            <div className="divide-y divide-slate-100">
              {notifications.map((n) => {
                const meta = NOTIF_META[n.type] || { label: "Cảnh báo", icon: AlertTriangle, color: "text-amber-700" };
                return (
                  <div key={n.id} className="p-3">
                    <p className={`text-xs ${meta.color} font-medium mb-0.5 flex items-center gap-1`}>
                      <meta.icon size={12} /> {meta.label}
                    </p>
                    <p className="text-xs text-slate-600">{n.message}</p>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// APP GỐC
// ---------------------------------------------------------------------------
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [customers, setCustomers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [employees, setEmployees] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("khach_hang");
  const [toast, setToast] = useState(null);
  const [highlightOrderId, setHighlightOrderId] = useState(null);

  const refreshAll = useCallback(async () => {
    const [custRes, ordRes, expRes, empRes] = await Promise.all([
      supabase.from("customers").select("*").order("created_at", { ascending: false }),
      supabase.from("orders").select("*").order("created_at", { ascending: false }),
      supabase.from("expenses").select("*").order("date", { ascending: false }),
      supabase.from("employees").select("id,name,username,role"),
    ]);
    if (custRes.data) setCustomers(custRes.data.map(customerFromRow));
    if (ordRes.data) setOrders(ordRes.data.map(orderFromRow));
    if (expRes.data) setExpenses(expRes.data.map(expenseFromRow));
    if (empRes.data) setEmployees(empRes.data);
  }, []);

  useEffect(() => {
    refreshAll().finally(() => setLoading(false));
  }, [refreshAll]);

  const derivedForNotif = useDerivedOrders(orders);
  const notifications = useNotifications(derivedForNotif, customers);
  const myDerivedOrders = derivedForNotif.filter((o) => currentUser && o.employeeId === currentUser.id);
  const employeeNotifications = useEmployeeNotifications(myDerivedOrders);

  // -------------------- CÁC HÀM GHI DỮ LIỆU (Supabase) --------------------
  const addCustomer = async (form) => {
    let vnidUrl = null;
    try {
      if (form.photoFile) {
        const up = await uploadFile(VNID_BUCKET, form.photoFile);
        vnidUrl = up.url;
      }
      const { error } = await supabase.from("customers").insert({
        name: form.name || null, cccd: form.cccd || null, phone: form.phone, address: form.address,
        ward: form.ward || null, industries: form.industries || [], referrer: form.referrer || null,
        vnid_photo_url: vnidUrl, employee_id: currentUser.id,
      });
      if (error) throw error;
      await refreshAll();
      setToast("Đã lưu khách hàng thành công.");
    } catch (err) {
      setToast("Lỗi khi lưu khách hàng: " + err.message);
    }
  };

  const createOrder = async (customerId, procedureType, orderCode) => {
    try {
      const { data, error } = await supabase.from("orders").insert({
        order_code: orderCode, customer_id: customerId, employee_id: currentUser.id,
        procedure_type: procedureType, status: "cho_xu_ly",
      }).select().single();
      if (error) throw error;
      await refreshAll();
      const newOrder = orderFromRow(data);
      setTab("trang_thai");
      setHighlightOrderId(newOrder.id);
      setToast(`Đã tạo đơn hàng ${newOrder.orderCode} thành công.`);
    } catch (err) {
      setToast("Lỗi khi tạo đơn hàng: " + err.message);
    }
  };

  const advanceOrder = async (orderId, nextStatus) => {
    const patch = { status: nextStatus };
    if (nextStatus === "da_tiep_nhan") patch.received_at = new Date().toISOString();
    if (nextStatus === "trinh_lanh_dao") patch.leader_at = new Date().toISOString();
    if (nextStatus === "duoc_chap_thuan") patch.approved_at = new Date().toISOString();
    const { error } = await supabase.from("orders").update(patch).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
  };

  const markRejected = async (orderId, reason) => {
    const { error } = await supabase.from("orders").update({
      status: "chua_duoc_chap_thuan", overdue_reason: reason, rejected_at: new Date().toISOString(),
    }).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
  };

  const retryOrder = async (orderId) => {
    const { error } = await supabase.from("orders").update({
      status: "trinh_lanh_dao", received_at: new Date().toISOString(), overdue_reason: "", rejected_at: null,
    }).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
  };

  const sendLicense = async (orderId, file) => {
    try {
      let url = null;
      if (file?.raw) {
        const up = await uploadFile(LICENSE_BUCKET, file.raw);
        url = up.url;
      }
      const { error } = await supabase.from("orders").update({
        status: "gui_giay_phep", license_pdf_url: url, license_pdf_name: file?.name || null,
        license_sent_at: new Date().toISOString(),
      }).eq("id", orderId);
      if (error) throw error;
      await refreshAll();
    } catch (err) {
      setToast("Lỗi khi gửi giấy phép: " + err.message);
    }
  };

  const confirmOrder = async (orderId) => {
    const { error } = await supabase.from("orders").update({
      status: "hoan_thanh", confirmed_at: new Date().toISOString(),
    }).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
    setToast("Đã xác nhận đơn hàng — thông báo đã gửi cho nhân viên.");
  };

  const saveFinance = async (orderId, revenue, cost, laborFee) => {
    const { error } = await supabase.from("orders").update({
      revenue, cost, labor_fee: laborFee, status: "da_thanh_toan", completed_at: new Date().toISOString(),
    }).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
    setToast("Đã ghi nhận thanh toán.");
  };

  const saveCompanyInfo = async (orderId, fields) => {
    const { error } = await supabase.from("orders").update({
      company_name: fields.companyName, capital: fields.capital, owner_dob: fields.ownerDob,
      owner_gender: fields.ownerGender, owner_email: fields.ownerEmail, owner_province: fields.ownerProvince,
      hq_address: fields.hqAddress, hq_ward: fields.hqWard, hq_province: fields.hqProvince,
      industries: fields.industries || [],
    }).eq("id", orderId);
    if (error) { setToast("Lưu file thành công nhưng lỗi khi lưu thông tin: " + error.message); return; }
    await refreshAll();
  };

  const addExpense = async (exp) => {
    const { error } = await supabase.from("expenses").insert({
      date: exp.date, description: exp.description, amount: exp.amount,
    });
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
  };

  const toggleCareStep = async (orderId, stepIndex, currentSteps) => {
    const newSteps = { ...currentSteps, [stepIndex]: !currentSteps?.[stepIndex] };
    const { error } = await supabase.from("orders").update({ care_steps: newSteps }).eq("id", orderId);
    if (error) { setToast("Lỗi: " + error.message); return; }
    await refreshAll();
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setTab("khach_hang");
  };

  if (loading) return <FullScreenLoader />;
  if (!currentUser) return <LoginScreen onLogin={(u) => { setCurrentUser(u); setTab(u.role === "admin" ? "kh_admin" : "khach_hang"); }} />;

  const isAdmin = currentUser.role === "admin";
  const NAV_NV = [
    { key: "khach_hang", label: "Khách hàng", icon: Users },
    { key: "tao_don", label: "Tạo đơn hàng", icon: Plus },
    { key: "trang_thai", label: "Trạng thái", icon: ClipboardCheck },
    { key: "bao_cao_nv", label: "Báo cáo", icon: BarChart3 },
  ];
  const NAV_ADMIN = [
    { key: "kh_admin", label: "Khách hàng", icon: Users },
    { key: "don_admin", label: "Đơn hàng", icon: ClipboardList },
    { key: "chi_phi", label: "Chi phí", icon: Receipt },
    { key: "bao_cao_admin", label: "Báo cáo", icon: BarChart3 },
    { key: "cham_soc", label: "Chăm sóc KH", icon: HeartHandshake },
  ];
  const navItems = isAdmin ? NAV_ADMIN : NAV_NV;

  return (
    <div className="min-h-screen bg-slate-50">
      <div className="bg-white/90 backdrop-blur border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-5xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-indigo-700 to-indigo-900 text-white flex items-center justify-center shrink-0 shadow-sm">
              <FileText size={18} />
            </div>
            <div className="min-w-0">
              <p className="text-sm font-semibold text-slate-800 leading-tight tracking-tight">Quản lý đơn hàng NMT</p>
              <p className="text-[11px] text-slate-400 leading-tight truncate">Theo dõi khách hàng &amp; đơn hàng</p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {isAdmin ? <NotifBell notifications={notifications} /> : <NotifBell notifications={employeeNotifications} />}
            <div className="flex items-center gap-2 pl-2 border-l border-slate-200">
              <div className="w-8 h-8 rounded-full bg-gradient-to-br from-indigo-700 to-indigo-900 text-white flex items-center justify-center text-xs font-semibold shrink-0">
                {currentUser.name[0]}
              </div>
              <div className="hidden md:block text-right leading-tight">
                <p className="text-xs font-medium text-slate-800">{currentUser.name}</p>
                <p className="text-[11px] text-slate-400">{isAdmin ? "Quản lý" : "Nhân viên"}</p>
              </div>
              <button onClick={handleLogout} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition">
                <LogOut size={14} className="text-slate-500" />
              </button>
            </div>
          </div>
        </div>
        <div className="max-w-5xl mx-auto px-4 pb-2 flex items-center gap-1 overflow-x-auto sm:flex-wrap sm:overflow-visible">
          {navItems.map((n) => (
            <button
              key={n.key}
              onClick={() => setTab(n.key)}
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-medium whitespace-nowrap transition ${
                tab === n.key ? "bg-indigo-800 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100 hover:text-slate-700"
              }`}
            >
              <n.icon size={15} /> {n.label}
            </button>
          ))}
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-4 py-6">
        {!isAdmin && tab === "khach_hang" && <CustomerModule currentUser={currentUser} customers={customers} onAddCustomer={addCustomer} />}
        {!isAdmin && tab === "tao_don" && <OrderCreateModule currentUser={currentUser} customers={customers} orders={orders} onCreateOrder={createOrder} />}
        {!isAdmin && tab === "trang_thai" && (
          <OrderStatusModule
            currentUser={currentUser} orders={orders} customers={customers} highlightOrderId={highlightOrderId}
            onAdvance={advanceOrder} onMarkRejected={markRejected} onRetry={retryOrder} onSendLicense={sendLicense}
            onSaveCompanyInfo={saveCompanyInfo}
          />
        )}
        {!isAdmin && tab === "bao_cao_nv" && <EmployeeReportModule currentUser={currentUser} orders={orders} />}

        {isAdmin && tab === "kh_admin" && <AdminCustomerModule customers={customers} employees={employees} />}
        {isAdmin && tab === "don_admin" && (
          <AdminOrderModule orders={orders} customers={customers} employees={employees} currentUser={currentUser} onConfirmOrder={confirmOrder} onSaveFinance={saveFinance} onSaveCompanyInfo={saveCompanyInfo} />
        )}
        {isAdmin && tab === "chi_phi" && <AdminExpenseModule expenses={expenses} onAddExpense={addExpense} />}
        {isAdmin && tab === "bao_cao_admin" && <AdminReportModule orders={orders} expenses={expenses} />}
        {isAdmin && tab === "cham_soc" && <CareModule orders={orders} customers={customers} onToggleCareStep={toggleCareStep} />}
      </div>
      {toast && <Toast message={toast} onClose={() => setToast(null)} />}
    </div>
  );
}
