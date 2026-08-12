import React, { useState, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  Package, Truck, ClipboardList, BarChart3, Bell, LogOut, CheckCircle2, XCircle,
  Plus, Search, ChevronRight, Inbox, Warehouse, TrendingUp, TrendingDown, Wallet,
  AlertTriangle, Clock, Loader2, Lock, Unlock, User, X, Pencil, Trash2, Download, Upload, Users,
  ShieldCheck, ArrowDownCircle, ArrowUpCircle, Boxes, Receipt, FileText, ChevronDown, ArrowUpDown, Filter, Coins,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, Cell,
} from "recharts";

/* ============================================================================
   QUẢN LÝ KHO NVL & THÀNH PHẨM — React + Supabase.
   Danh mục gốc (NCC, SP/NVL/TP, Mã doanh thu, Mã xuất) → Nhập hàng → Xuất hàng
   → Báo cáo nhập/xuất → Tồn kho (bình quân gia quyền).
   ========================================================================= */

// ---------------------------------------------------------------------------
// HẰNG SỐ
// ---------------------------------------------------------------------------
const ROLE_META = {
  nhan_vien_kho: { label: "Nhân viên kho", color: "bg-sky-50 text-sky-700 border-sky-200" },
  quan_ly: { label: "Quản lý", color: "bg-slate-800 text-white border-slate-800" },
  bao_cao: { label: "Quản lý (Báo cáo)", color: "bg-indigo-50 text-indigo-700 border-indigo-200" },
  thu_ngan: { label: "Thu ngân", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
};
const PAYMENT_TYPE_META = {
  tien_mat: { label: "Tiền mặt", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
  cong_no: { label: "Công nợ", color: "bg-amber-50 text-amber-700 border-amber-200" },
  noi_bo: { label: "Nội bộ", color: "bg-indigo-50 text-indigo-700 border-indigo-200" },
};
const CLASSIFICATION_META = {
  NL: { label: "Nguyên vật liệu", color: "bg-sky-50 text-sky-700 border-sky-200" },
  TP: { label: "Thành phẩm", color: "bg-purple-50 text-purple-700 border-purple-200" },
};
const SESSION_KEY = "kho_session_employee_id";

// ---------------------------------------------------------------------------
// HÀM TIỆN ÍCH
// ---------------------------------------------------------------------------
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("vi-VN") + " đ";
}
function fmtNumber(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("vi-VN", { maximumFractionDigits: 3 });
}
function fmtDate(d) {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("vi-VN");
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function daysAgoISO(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().slice(0, 10);
}
// Ngày bắt đầu theo dõi quỹ — mặc định cho các bộ lọc trong tab Quỹ.
const FUND_START_DATE = "2026-08-01";
// Liệt kê danh sách ngày (yyyy-mm-dd) từ "from" đến "to", tăng dần. Giới hạn 62 ngày để tránh render quá nặng.
function enumerateDatesISO(from, to) {
  if (!from || !to) return [];
  const start = new Date(from + "T00:00:00");
  const end = new Date(to + "T00:00:00");
  if (isNaN(start) || isNaN(end) || start > end) return [];
  const dates = [];
  const cur = new Date(start);
  while (cur <= end && dates.length < 62) {
    dates.push(cur.toISOString().slice(0, 10));
    cur.setDate(cur.getDate() + 1);
  }
  return dates;
}
function stripDiacritics(str) {
  return (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase();
}
// Chuẩn hoá chuỗi để SO KHỚP tên món/tên NVL giữa 2 nguồn khác nhau (Cost món ăn
// gõ tay vs. tên trích từ file Excel POS) — ngoài bỏ dấu còn loại các ký tự
// "vô hình" hay gặp trong file Excel xuất từ phần mềm POS (khoảng trắng không
// ngắt dòng, zero-width space...) và gộp nhiều khoảng trắng liền nhau thành 1.
function normalizeForMatch(str) {
  return stripDiacritics(str)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, "")
    .replace(/\s+/g, "")
    .trim();
}

// Đọc 1 file Excel (.xlsx/.xls/.csv) thành mảng-các-mảng thô (không giả định dòng 1
// là tiêu đề, vì nhiều file xuất từ phần mềm POS có vài dòng mô tả phía trên bảng thật).
function readExcelRaw(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Không đọc được file."));
    reader.readAsArrayBuffer(file);
  });
}
// Tự dò đúng dòng tiêu đề thật trong file — bỏ qua các dòng mô tả/tên báo cáo phía
// trên (thường gặp ở file xuất từ phần mềm POS), dựa vào các từ khoá bắt buộc phải có.
function detectHeaderRow(rawRows, hints) {
  const normHints = hints.map((h) => stripDiacritics(h).replace(/\s+/g, ""));
  for (let i = 0; i < rawRows.length; i++) {
    const rowText = (rawRows[i] || []).map((c) => stripDiacritics(String(c ?? "")).replace(/\s+/g, "")).join("|");
    if (normHints.every((h) => rowText.includes(h))) return i;
  }
  return -1;
}
// Chuyển các dòng thô (từ đúng dòng tiêu đề trở xuống) thành mảng object theo tên cột.
// Tự bỏ qua dòng trống hoàn toàn (kể cả dòng tổng phụ theo hoá đơn không có đủ cột).
function rowsToObjects(rawRows, headerRowIdx) {
  const headers = (rawRows[headerRowIdx] || []).map((h) => String(h ?? "").trim());
  const out = [];
  for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
    const row = rawRows[i] || [];
    if (row.every((c) => c === "" || c === undefined || c === null)) continue;
    const obj = {};
    headers.forEach((h, idx) => { if (h) obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}
// Lấy giá trị 1 cột trong dòng Excel, thử vài tên cột khác nhau (không phân biệt hoa/thường, dấu).
function pickCol(row, ...names) {
  const keys = Object.keys(row);
  for (const name of names) {
    const target = stripDiacritics(name).replace(/\s+/g, "");
    const found = keys.find((k) => stripDiacritics(k).replace(/\s+/g, "") === target);
    if (found !== undefined) return row[found];
  }
  return undefined;
}
// Tìm cột theo kiểu "chứa" (không cần khớp tuyệt đối tên cột) — dùng cho các cột có
// tên hay bị gộp/biến thể giữa các phần mềm POS khác nhau, VD "Ngày" có thể xuất hiện
// dưới dạng "Ngày/Giờ vào-ra", "Ngày giờ vào ra"... miễn có chứa chữ "ngay" là nhận.
function pickColContains(row, substrNorm) {
  const keys = Object.keys(row);
  const found = keys.find((k) => stripDiacritics(k).replace(/\s+/g, "").includes(substrNorm));
  return found !== undefined ? row[found] : undefined;
}
function excelDateToISO(v) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  // Chấp nhận cả trường hợp cột gộp chung ngày+giờ (VD "08/08/2026 12:35:20") —
  // chỉ cần phần đầu chuỗi là ngày hợp lệ, không cần khớp tuyệt đối cả chuỗi.
  const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  return "";
}

// ---------------------------------------------------------------------------
// TỰ ĐỘNG GỢI Ý MÃ MỚI — nhảy tiếp theo mã lớn nhất hiện có
// ---------------------------------------------------------------------------
function nextSupplierCode(suppliers) {
  let max = 0;
  suppliers.forEach((s) => {
    const m = /^cc(\d+)$/i.exec((s.code || "").trim());
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return "cc" + String(max + 1).padStart(3, "0");
}
function nextProductCode(products, classification) {
  let max = classification === "TP" ? 50000 : 10000;
  products
    .filter((p) => p.classification === classification && /^\d+$/.test(p.code))
    .forEach((p) => { max = Math.max(max, parseInt(p.code, 10)); });
  return String(max + 1);
}
function nextSimpleCode(list, prefix) {
  let max = 0;
  list.forEach((x) => {
    const m = new RegExp(`^${prefix}(\\d+)$`, "i").exec((x.code || "").trim());
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + String(max + 1).padStart(2, "0");
}

// ---------------------------------------------------------------------------
// MAPPER: snake_case (Supabase) <-> camelCase (giao diện)
// ---------------------------------------------------------------------------
function mapEmployee(e) {
  return {
    id: e.id, username: e.username, name: e.name, role: e.role,
    mustChangePassword: !!e.must_change_password, passwordChangeDeadline: e.password_change_deadline || null,
  };
}
function mapSupplier(s) {
  return { id: s.id, code: s.code, name: s.name, paymentType: s.payment_type, createdAt: s.created_at };
}
function mapRevenueCode(r) {
  return { id: r.id, code: r.code, name: r.name, createdAt: r.created_at };
}
function mapExportCode(r) {
  return { id: r.id, code: r.code, name: r.name, createdAt: r.created_at };
}
function mapProduct(p) {
  return {
    id: p.id, code: p.code, name: p.name, unit: p.unit,
    groupCode: p.group_code || "", groupName: p.group_name || "",
    classification: p.classification, createdAt: p.created_at,
  };
}
function mapStockOpening(o) {
  return {
    id: o.id, productId: o.product_id, asOfDate: o.as_of_date,
    quantity: Number(o.quantity) || 0, unitPrice: Number(o.unit_price) || 0,
    note: o.note || "", createdAt: o.created_at,
  };
}
function mapImportRecord(r) {
  return {
    id: r.id, orderNumber: r.order_number || "", receiptCode: r.receipt_code,
    supplierId: r.supplier_id, productId: r.product_id,
    quantity: Number(r.quantity) || 0, unitPrice: Number(r.unit_price) || 0, totalAmount: Number(r.total_amount) || 0,
    paymentType: r.payment_type, importDate: r.import_date, createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapExportRecord(r) {
  return {
    id: r.id, orderNumber: r.order_number || "", receiptCode: r.receipt_code,
    revenueCodeId: r.revenue_code_id, exportCodeId: r.export_code_id, productId: r.product_id,
    lineType: r.line_type, quantity: Number(r.quantity) || 0, unitPrice: Number(r.unit_price) || 0,
    totalAmount: Number(r.total_amount) || 0, exportDate: r.export_date, createdBy: r.created_by, createdAt: r.created_at,
  };
}

function mapExpenseRecord(r) {
  return {
    id: r.id, category: r.category, itemName: r.item_name,
    quantity: r.quantity === null ? null : Number(r.quantity),
    unitPrice: r.unit_price === null ? null : Number(r.unit_price),
    amount: Number(r.amount) || 0,
    paymentMethod: r.payment_method || "tien_mat",
    expenseDate: r.expense_date, note: r.note || "", createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapDish(d) {
  return {
    id: d.id, name: d.name, sellingPrice: d.selling_price === null ? null : Number(d.selling_price),
    note: d.note || "", createdBy: d.created_by, createdAt: d.created_at,
  };
}
function mapDishSale(r) {
  return {
    id: r.id, dishId: r.dish_id, quantity: Number(r.quantity) || 0,
    unitPrice: Number(r.unit_price) || 0, totalAmount: Number(r.total_amount) || 0,
    costAmount: Number(r.cost_amount) || 0, invoiceNo: r.invoice_no || "",
    receiptCode: r.receipt_code, saleDate: r.sale_date, createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapCashierReceipt(r) {
  return {
    id: r.id, receiptDate: r.receipt_date, cashAmount: Number(r.cash_amount) || 0,
    bankAmount: Number(r.bank_amount) || 0, note: r.note || "", createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapInvoiceRevenue(r) {
  return {
    id: r.id, invoiceNo: r.invoice_no, invoiceDate: r.invoice_date, amount: Number(r.amount) || 0,
    createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapDishIngredient(i) {
  return {
    id: i.id, dishId: i.dish_id, productId: i.product_id,
    quantity: Number(i.quantity) || 0, costMode: i.cost_mode,
    allocatedCost: i.allocated_cost === null ? null : Number(i.allocated_cost),
    sortOrder: i.sort_order || 0, createdAt: i.created_at,
  };
}
// Cọc — thu cọc (nhận cọc từ khách) hoặc chi cọc (ứng cọc cho NCC/đối tác), theo dõi
// riêng khỏi Chi phí vì không phải chi phí thực tế phát sinh, mà là dòng tiền tạm giữ
// đang chờ đối trừ/hoàn trả.
function mapDeposit(r) {
  return {
    id: r.id, direction: r.direction, partyName: r.party_name,
    amount: Number(r.amount) || 0, depositDate: r.deposit_date,
    paymentMethod: r.payment_method || "tien_mat",
    status: r.status, note: r.note || "", createdBy: r.created_by, createdAt: r.created_at,
  };
}
function mapNotification(r) {
  return {
    id: r.id, message: r.message, type: r.type, targetRole: r.target_role,
    isRead: r.is_read, createdBy: r.created_by, createdByName: r.created_by_name, createdAt: r.created_at,
  };
}
function mapFundDailyBalance(r) {
  return {
    date: r.balance_date, openingBalance: Number(r.opening_balance) || 0,
    openingBalanceBank: Number(r.opening_balance_bank) || 0,
    remittedOwnerCash: Number(r.remitted_owner_cash) || 0,
    remittedOwnerBank: Number(r.remitted_owner_bank) || 0,
    note: r.note || "", updatedBy: r.updated_by, updatedAt: r.updated_at,
  };
}

async function fetchAll() {
  const [emp, sup, rev, exc, prod, open, imp, exp, cost, dish, dishIng, dishSale, cashierRec, invRev, dep, noti, fundBal, settings] = await Promise.all([
    supabase.from("employees").select("id,username,name,role,must_change_password,password_change_deadline"),
    supabase.from("suppliers").select("*").order("code"),
    supabase.from("revenue_codes").select("*").order("code"),
    supabase.from("export_codes").select("*").order("code"),
    supabase.from("products").select("*").order("code"),
    supabase.from("stock_opening").select("*").order("as_of_date", { ascending: false }),
    supabase.from("import_records").select("*").order("import_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("export_records").select("*").order("export_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("expense_records").select("*").order("expense_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("dishes").select("*").order("name"),
    supabase.from("dish_ingredients").select("*").order("sort_order"),
    supabase.from("dish_sales").select("*").order("sale_date", { ascending: false }),
    supabase.from("cashier_receipts").select("*").order("receipt_date", { ascending: false }),
    supabase.from("invoice_revenue").select("*").order("invoice_date", { ascending: false }),
    supabase.from("deposits").select("*").order("deposit_date", { ascending: false }).order("created_at", { ascending: false }),
    supabase.from("notifications").select("*").order("created_at", { ascending: false }).limit(100),
    supabase.from("fund_daily_balance").select("*").order("balance_date", { ascending: false }),
    supabase.from("app_settings").select("*"),
  ]);
  [emp, sup, rev, exc, prod, open, imp, exp, cost, dish, dishIng, dishSale, cashierRec, invRev, dep, noti, fundBal, settings].forEach((r) => { if (r.error) console.error(r.error); });
  return {
    employees: (emp.data || []).map(mapEmployee),
    suppliers: (sup.data || []).map(mapSupplier),
    revenueCodes: (rev.data || []).map(mapRevenueCode),
    exportCodes: (exc.data || []).map(mapExportCode),
    products: (prod.data || []).map(mapProduct),
    stockOpenings: (open.data || []).map(mapStockOpening),
    importRecords: (imp.data || []).map(mapImportRecord),
    exportRecords: (exp.data || []).map(mapExportRecord),
    expenseRecords: (cost.data || []).map(mapExpenseRecord),
    dishes: (dish.data || []).map(mapDish),
    dishIngredients: (dishIng.data || []).map(mapDishIngredient),
    dishSales: (dishSale.data || []).map(mapDishSale),
    cashierReceipts: (cashierRec.data || []).map(mapCashierReceipt),
    invoiceRevenue: (invRev.data || []).map(mapInvoiceRevenue),
    deposits: (dep.data || []).map(mapDeposit),
    notifications: (noti.data || []).map(mapNotification),
    fundDailyBalances: (fundBal.data || []).map(mapFundDailyBalance),
    settings: Object.fromEntries((settings.data || []).map((r) => [r.key, r.value])),
  };
}

// ---------------------------------------------------------------------------
// TÍNH TỒN KHO & GIÁ BÌNH QUÂN GIA QUYỀN
// ---------------------------------------------------------------------------
function latestOpening(productId, stockOpenings) {
  const list = stockOpenings.filter((o) => o.productId === productId).sort((a, b) => new Date(b.asOfDate) - new Date(a.asOfDate));
  return list[0] || null;
}

// Giá bình quân gia quyền HIỆN TẠI của 1 sản phẩm: (giá trị tồn đầu + giá trị
// nhập từ mốc tồn đầu tới nay) / (số lượng tồn đầu + số lượng nhập tương ứng).
function computeAvgPrice(productId, { stockOpenings, importRecords }) {
  const opening = latestOpening(productId, stockOpenings);
  const baseQty = opening?.quantity || 0;
  const baseValue = opening ? opening.quantity * opening.unitPrice : 0;
  const baseDate = opening?.asOfDate || "1970-01-01";
  const imports = importRecords.filter((r) => r.productId === productId && r.importDate > baseDate);
  const importQty = imports.reduce((s, r) => s + r.quantity, 0);
  const importValue = imports.reduce((s, r) => s + r.totalAmount, 0);
  const totalQty = baseQty + importQty;
  const totalValue = baseValue + importValue;
  if (totalQty <= 0) return 0;
  return totalValue / totalQty;
}

// Chi phí của 1 dòng nguyên liệu trong công thức món ăn:
// - "phan_bo" (dòng bôi xanh trong file Excel gốc): lấy thẳng giá trị phân bổ đã gán cho dòng đó.
// - "xuat_kho" (còn lại): giá xuất kho hiện tại (bình quân gia quyền) × khối lượng xuất dùng trong món.
function dishIngredientCost(ing, data) {
  if (ing.costMode === "phan_bo") return ing.allocatedCost || 0;
  const avgPrice = computeAvgPrice(ing.productId, data);
  return avgPrice * ing.quantity;
}
function dishTotalCost(dishId, data) {
  return data.dishIngredients.filter((i) => i.dishId === dishId).reduce((s, i) => s + dishIngredientCost(i, data), 0);
}

// Tồn kho (số lượng) của 1 sản phẩm tính đến hết 1 ngày cụ thể.
function stockAsOf(productId, asOfDate, { stockOpenings, importRecords, exportRecords }) {
  const opening = latestOpening(productId, stockOpenings);
  const baseQty = opening?.quantity || 0;
  const baseDate = opening?.asOfDate || "1970-01-01";
  const imports = importRecords.filter((r) => r.productId === productId && r.importDate > baseDate && r.importDate <= asOfDate);
  const exports = exportRecords.filter((r) => r.productId === productId && r.exportDate > baseDate && r.exportDate <= asOfDate);
  return baseQty + imports.reduce((s, r) => s + r.quantity, 0) - exports.reduce((s, r) => s + r.quantity, 0);
}

// Báo cáo Nhập-Xuất-Tồn cho 1 sản phẩm trong khoảng [from, to] (Tác vụ 6).
function nktForProduct(productId, from, to, data) {
  const { stockOpenings, importRecords, exportRecords } = data;
  const dayBeforeFrom = new Date(from);
  dayBeforeFrom.setDate(dayBeforeFrom.getDate() - 1);
  const openingQty = stockAsOf(productId, dayBeforeFrom.toISOString().slice(0, 10), data);
  const avgPrice = computeAvgPrice(productId, data);
  const importsInRange = importRecords.filter((r) => r.productId === productId && r.importDate >= from && r.importDate <= to);
  const exportsInRange = exportRecords.filter((r) => r.productId === productId && r.exportDate >= from && r.exportDate <= to);
  const importQty = importsInRange.reduce((s, r) => s + r.quantity, 0);
  const importValue = importsInRange.reduce((s, r) => s + r.totalAmount, 0);
  const exportQty = exportsInRange.reduce((s, r) => s + r.quantity, 0);
  const exportValue = exportsInRange.reduce((s, r) => s + r.totalAmount, 0);
  const closingQty = openingQty + importQty - exportQty;
  return {
    openingQty, openingValue: openingQty * avgPrice,
    importQty, importValue,
    exportQty, exportValue,
    closingQty, closingValue: closingQty * avgPrice,
    avgPrice,
  };
}

// Sinh mã phiếu tự động — VD NK-20260803-0001 / XK-20260803-0001
// Mã phiếu = tiền tố + ngày + giờ:phút:giây:mili giây — không dựa vào số lượng bản ghi hiện có,
// nên không bao giờ bị trùng dù dữ liệu vừa bị xoá sạch trước đó (tránh trùng mã phiếu).
function genReceiptCode(prefix) {
  const d = new Date();
  const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
  const hms = `${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}${String(d.getSeconds()).padStart(2, "0")}${String(d.getMilliseconds()).padStart(3, "0")}`;
  return `${prefix}-${ymd}-${hms}`;
}

// ---------------------------------------------------------------------------
// UI DÙNG CHUNG
// ---------------------------------------------------------------------------
function Card({ children, className = "" }) {
  return <div className={`bg-white/90 backdrop-blur-sm rounded-3xl border border-slate-200/80 shadow-sm hover:shadow-md transition-shadow duration-200 ${className}`}>{children}</div>;
}
function SectionTitle({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex items-start gap-3 mb-4">
      <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center shrink-0 shadow-sm shadow-sky-900/20">
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
function MetricCard({ label, value, icon: Icon, accent = "teal" }) {
  const a = {
    teal: { text: "text-sky-700", bg: "bg-sky-50", bar: "bg-sky-600" },
    amber: { text: "text-amber-700", bg: "bg-amber-50", bar: "bg-amber-600" },
    emerald: { text: "text-emerald-700", bg: "bg-emerald-50", bar: "bg-emerald-600" },
    rose: { text: "text-rose-700", bg: "bg-rose-50", bar: "bg-rose-600" },
    indigo: { text: "text-indigo-700", bg: "bg-indigo-50", bar: "bg-indigo-600" },
  }[accent];
  return (
    <div className="relative bg-white/90 backdrop-blur-sm rounded-3xl border border-slate-200/80 shadow-sm p-4 flex items-center gap-3 overflow-hidden">
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
function PrimaryButton({ children, className = "", ...props }) {
  return (
    <button {...props} className={`inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-gradient-to-b from-sky-700 to-sky-800 text-white text-sm font-medium shadow-sm hover:from-sky-800 hover:to-sky-900 active:scale-[0.98] transition disabled:opacity-40 disabled:cursor-not-allowed ${className}`}>
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
function DangerButton({ children, className = "", ...props }) {
  return (
    <button {...props} className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl border border-rose-300 bg-white text-rose-700 text-sm font-medium hover:bg-rose-50 active:scale-[0.98] transition disabled:opacity-40 ${className}`}>
      {children}
    </button>
  );
}
function TextField({ label, hint, ...props }) {
  return (
    <label className="block">
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <input {...props} className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600" />
      {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}
// Ô nhập số tiền — tự hiện dấu phẩy ngăn cách hàng nghìn trong lúc gõ (VD: 1,500,000).
// Giá trị thật (không dấu phẩy) trả về qua onChange(rawDigitsString), cho phần thập phân
// (số lượng dạng kg lẻ) dùng allowDecimal để giữ được dấu chấm thập phân.
function MoneyField({ label, value, onChange, placeholder = "0", hint, className = "", allowDecimal = false, disabled }) {
  const formatDisplay = (v) => {
    if (v === "" || v === null || v === undefined) return "";
    const s = String(v);
    if (allowDecimal) {
      const [intPart, decPart] = s.split(".");
      const intFmt = intPart === "" ? "" : Number(intPart || 0).toLocaleString("en-US");
      return decPart !== undefined ? `${intFmt}.${decPart}` : intFmt;
    }
    return s === "" ? "" : Number(s).toLocaleString("en-US");
  };
  const handleChange = (e) => {
    const raw = allowDecimal
      ? e.target.value.replace(/[^\d.]/g, "").replace(/(\..*)\./g, "$1")
      : e.target.value.replace(/[^\d]/g, "");
    onChange(raw);
  };
  return (
    <label className="block">
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <input
        type="text"
        inputMode="decimal"
        value={formatDisplay(value)}
        onChange={handleChange}
        placeholder={placeholder}
        disabled={disabled}
        className={`w-full px-3 py-2 rounded-xl border border-slate-300 text-sm transition focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600 disabled:bg-slate-50 disabled:text-slate-400 ${className}`}
      />
      {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}
function SelectField({ label, children, hint, ...props }) {
  return (
    <label className="block">
      {label && <span className="block text-xs font-medium text-slate-600 mb-1">{label}</span>}
      <select {...props} className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white transition focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600">
        {children}
      </select>
      {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}
function Toast({ toast }) {
  if (!toast) return null;
  return (
    <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-sm px-4 py-2.5 rounded-xl shadow-lg shadow-slate-900/30 flex items-center gap-2 animate-[fadeIn_.2s_ease-out]">
      <CheckCircle2 size={15} className="text-emerald-400 shrink-0" /> {toast}
    </div>
  );
}
function FullScreenLoader() {
  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50">
      <Loader2 size={28} className="text-sky-600 animate-spin" />
    </div>
  );
}
function fmtCountdown(ms) {
  if (ms <= 0) return "00:00:00";
  const totalSec = Math.floor(ms / 1000);
  const h = String(Math.floor(totalSec / 3600)).padStart(2, "0");
  const m = String(Math.floor((totalSec % 3600) / 60)).padStart(2, "0");
  const s = String(totalSec % 60).padStart(2, "0");
  return `${h}:${m}:${s}`;
}

// ---------------------------------------------------------------------------
// ĐĂNG NHẬP — bcrypt qua RPC, chặn brute-force
// ---------------------------------------------------------------------------
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const LOCK_KEY_PREFIX = "kho_login_fail_";
  const MAX_ATTEMPTS = 5;
  const LOCK_MINUTES = 5;
  const getFailState = (u) => { try { return JSON.parse(localStorage.getItem(LOCK_KEY_PREFIX + u) || "null"); } catch { return null; } };
  const setFailState = (u, s) => { try { localStorage.setItem(LOCK_KEY_PREFIX + u, JSON.stringify(s)); } catch {} };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password) { setError("Vui lòng nhập tên đăng nhập và mật khẩu."); return; }
    const u = username.trim();
    const fail = getFailState(u);
    if (fail && fail.count >= MAX_ATTEMPTS) {
      const remainMs = fail.lockedAt + LOCK_MINUTES * 60_000 - Date.now();
      if (remainMs > 0) { setError(`Bạn đã nhập sai quá ${MAX_ATTEMPTS} lần. Vui lòng thử lại sau ${Math.ceil(remainMs / 60000)} phút.`); return; }
      setFailState(u, null);
    }
    setError(""); setLoading(true);
    try {
      const { data, error: qErr } = await supabase.rpc("verify_employee_login", { p_username: u, p_password: password });
      if (qErr) throw qErr;
      const employee = Array.isArray(data) ? data[0] : data;
      if (!employee) {
        const prev = getFailState(u) || { count: 0 };
        const nextCount = prev.count + 1;
        setFailState(u, { count: nextCount, lockedAt: nextCount >= MAX_ATTEMPTS ? Date.now() : null });
        setError(nextCount >= MAX_ATTEMPTS ? `Sai mật khẩu quá ${MAX_ATTEMPTS} lần. Tài khoản bị khoá tạm ${LOCK_MINUTES} phút.` : "Tên đăng nhập hoặc mật khẩu không đúng.");
        return;
      }
      if (employee.locked) { setError("Tài khoản đã bị khoá do không đổi mật khẩu đúng hạn. Vui lòng liên hệ Quản lý."); return; }
      setFailState(u, null);
      let mapped = mapEmployee(employee);
      if (mapped.mustChangePassword && !mapped.passwordChangeDeadline) {
        const { data: deadline } = await supabase.rpc("start_password_deadline", { p_employee_id: mapped.id });
        if (deadline) mapped = { ...mapped, passwordChangeDeadline: deadline };
      }
      onLogin(mapped);
    } catch (err) {
      console.error(err);
      setError("Không kết nối được máy chủ, vui lòng thử lại.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-b from-sky-50 via-slate-50 to-slate-50 px-4 py-10">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-sky-700 to-sky-900 mx-auto mb-3 flex items-center justify-center shadow-lg shadow-sky-900/25">
            <Warehouse size={26} className="text-white" />
          </div>
          <h1 className="text-2xl font-semibold text-slate-800 tracking-tight">P&amp;L The Eros 143</h1>
          <p className="text-sm text-slate-500 mt-1">Đăng nhập bằng tài khoản nhân sự</p>
        </div>
        <form onSubmit={handleSubmit} className="bg-white rounded-2xl border border-slate-200/80 shadow-xl shadow-slate-900/5 p-5 space-y-4">
          <label className="block">
            <span className="block text-xs font-medium text-slate-600 mb-1">Tên đăng nhập</span>
            <div className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-300 focus-within:ring-2 focus-within:ring-sky-500/40 focus-within:border-sky-600">
              <User size={15} className="text-slate-400 shrink-0" />
              <input value={username} onChange={(e) => setUsername(e.target.value)} className="flex-1 text-sm outline-none" placeholder="vd: nvkho1" autoFocus />
            </div>
          </label>
          <label className="block">
            <span className="block text-xs font-medium text-slate-600 mb-1">Mật khẩu</span>
            <div className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-300 focus-within:ring-2 focus-within:ring-sky-500/40 focus-within:border-sky-600">
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

function ChangePasswordForm({ currentUser, onSuccess, onCancel, mandatory }) {
  const [oldPassword, setOldPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async () => {
    if (!oldPassword || !newPassword || !confirmPassword) { setError("Vui lòng điền đủ các ô."); return; }
    if (newPassword.length < 6) { setError("Mật khẩu mới cần ít nhất 6 ký tự."); return; }
    if (newPassword !== confirmPassword) { setError("Mật khẩu xác nhận không khớp."); return; }
    setError(""); setSaving(true);
    try {
      const { data, error: qErr } = await supabase.rpc("change_own_password", {
        p_employee_id: currentUser.id, p_old_password: oldPassword, p_new_password: newPassword,
      });
      if (qErr) throw qErr;
      if (!data) { setError("Mật khẩu hiện tại không đúng."); return; }
      onSuccess();
    } catch (err) {
      console.error(err);
      setError("Không đổi được mật khẩu, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-3">
      <TextField label="Mật khẩu hiện tại" type="password" value={oldPassword} onChange={(e) => setOldPassword(e.target.value)} />
      <TextField label="Mật khẩu mới (tối thiểu 6 ký tự)" type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} />
      <TextField label="Xác nhận mật khẩu mới" type="password" value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} />
      {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertTriangle size={14} /> {error}</p>}
      <div className="flex gap-2">
        <PrimaryButton type="button" onClick={submit} disabled={saving} className={mandatory ? "w-full justify-center" : ""}>
          {saving ? "Đang lưu..." : "Đổi mật khẩu"}
        </PrimaryButton>
        {!mandatory && onCancel && <GhostButton type="button" onClick={onCancel}>Hủy</GhostButton>}
      </div>
    </div>
  );
}

function ForcePasswordChangeGate({ currentUser, onChanged, onLogout }) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, []);
  const deadline = currentUser.passwordChangeDeadline ? new Date(currentUser.passwordChangeDeadline).getTime() : null;
  const remainMs = deadline ? deadline - now : null;
  const expired = remainMs !== null && remainMs <= 0;

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-b from-rose-50 via-slate-50 to-slate-50 px-4 py-10">
      <div className="w-full max-w-sm">
        <div className="text-center mb-6">
          <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-rose-600 to-rose-700 mx-auto mb-3 flex items-center justify-center shadow-lg shadow-rose-900/25">
            <Lock size={26} className="text-white" />
          </div>
          <h1 className="text-xl font-semibold text-slate-800">Yêu cầu đổi mật khẩu</h1>
          <p className="text-sm text-slate-500 mt-1">Vì lý do bảo mật, bạn cần đặt mật khẩu mới trước khi tiếp tục sử dụng.</p>
          {!expired && deadline && <p className="text-sm text-rose-600 font-semibold mt-2">Thời gian còn lại: {fmtCountdown(remainMs)}</p>}
          {expired && <p className="text-sm text-rose-600 font-semibold mt-2">Đã hết hạn 24 giờ — tài khoản đã bị khoá. Vui lòng liên hệ Quản lý.</p>}
        </div>
        {!expired ? (
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-xl shadow-slate-900/5 p-5">
            <ChangePasswordForm currentUser={currentUser} mandatory onSuccess={onChanged} />
          </div>
        ) : (
          <GhostButton className="w-full justify-center" onClick={onLogout}>Đăng xuất</GhostButton>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// DANH MỤC GỐC — Nhà cung cấp / NVL & TP / Mã doanh thu / Mã xuất
// ---------------------------------------------------------------------------
function AddSupplierForm({ suppliers, onAdd }) {
  const [code, setCode] = useState(() => nextSupplierCode(suppliers));
  const [name, setName] = useState("");
  const [paymentType, setPaymentType] = useState("cong_no");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!code.trim() || !name.trim()) { setError("Vui lòng nhập đủ Mã và Tên nhà cung cấp."); return; }
    setError(""); setSaving(true);
    try {
      await onAdd({ code: code.trim(), name: name.trim(), paymentType });
      setCode(nextSupplierCode([...suppliers, { code }]));
      setName("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <p className="font-semibold text-slate-800 text-sm mb-3">Thêm nhà cung cấp mới</p>
      <div className="grid sm:grid-cols-3 gap-3 mb-3">
        <TextField label="Mã NCC (gợi ý tự động)" value={code} onChange={(e) => setCode(e.target.value)} />
        <TextField label="Tên nhà cung cấp" value={name} onChange={(e) => setName(e.target.value)} className="sm:col-span-2" />
        <SelectField label="Hình thức thanh toán mặc định" value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="sm:col-span-3">
          <option value="cong_no">Công nợ</option>
          <option value="tien_mat">Tiền mặt</option>
          <option value="noi_bo">Nội bộ</option>
        </SelectField>
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Thêm nhà cung cấp</PrimaryButton>
    </Card>
  );
}

function SupplierList({ suppliers }) {
  const [q, setQ] = useState("");
  const filtered = q ? suppliers.filter((s) => stripDiacritics(s.name).includes(stripDiacritics(q)) || s.code.toLowerCase().includes(q.toLowerCase())) : suppliers;
  return (
    <Card className="p-0 overflow-hidden">
      <div className="p-4 border-b border-slate-100 flex items-center gap-2">
        <Search size={15} className="text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm theo mã hoặc tên..." className="flex-1 text-sm outline-none" />
      </div>
      {filtered.length === 0 ? <EmptyState icon={Truck} text="Chưa có nhà cung cấp nào." /> : (
        <div className="divide-y divide-slate-100 max-h-96 overflow-y-auto">
          {filtered.map((s) => (
            <div key={s.id} className="p-3.5 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <p className="text-sm font-medium text-slate-800">{s.name}</p>
                <p className="text-xs text-slate-400">Mã: {s.code}</p>
              </div>
              <Badge className={PAYMENT_TYPE_META[s.paymentType]?.color}>{PAYMENT_TYPE_META[s.paymentType]?.label}</Badge>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function AddProductForm({ products, onAdd }) {
  const [classification, setClassification] = useState("NL");
  const [code, setCode] = useState(() => nextProductCode(products, "NL"));
  const [name, setName] = useState("");
  const [unit, setUnit] = useState("");
  const [groupCode, setGroupCode] = useState("N10");
  const [groupName, setGroupName] = useState("FOOD");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => { setCode(nextProductCode(products, classification)); }, [classification]);

  const submit = async () => {
    if (!code.trim() || !name.trim() || !unit.trim()) { setError("Vui lòng nhập đủ Mã, Tên và Đơn vị tính."); return; }
    setError(""); setSaving(true);
    try {
      await onAdd({ code: code.trim(), name: name.trim(), unit: unit.trim(), groupCode: groupCode.trim(), groupName: groupName.trim(), classification });
      setCode(nextProductCode([...products, { code, classification }], classification));
      setName(""); setUnit("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <p className="font-semibold text-slate-800 text-sm mb-3">Thêm sản phẩm mới (Nguyên vật liệu / Thành phẩm)</p>
      <div className="grid sm:grid-cols-3 gap-3 mb-3">
        <SelectField label="Phân loại" value={classification} onChange={(e) => setClassification(e.target.value)}>
          <option value="NL">Nguyên vật liệu (NL)</option>
          <option value="TP">Thành phẩm (TP)</option>
        </SelectField>
        <TextField label="Mã SP (gợi ý tự động)" value={code} onChange={(e) => setCode(e.target.value)} />
        <TextField label="Đơn vị tính" value={unit} onChange={(e) => setUnit(e.target.value)} placeholder="kg, cái, bó, đĩa..." />
        <TextField label="Tên sản phẩm" value={name} onChange={(e) => setName(e.target.value)} className="sm:col-span-3" />
        <TextField label="Mã nhóm" value={groupCode} onChange={(e) => setGroupCode(e.target.value)} />
        <TextField label="Mã nhóm lớn" value={groupName} onChange={(e) => setGroupName(e.target.value)} className="sm:col-span-2" />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Thêm sản phẩm</PrimaryButton>
    </Card>
  );
}

function ProductList({ products }) {
  const [q, setQ] = useState("");
  const [filter, setFilter] = useState("all");
  let filtered = filter === "all" ? products : products.filter((p) => p.classification === filter);
  if (q) filtered = filtered.filter((p) => stripDiacritics(p.name).includes(stripDiacritics(q)) || p.code.includes(q));
  return (
    <Card className="p-0 overflow-hidden">
      <div className="p-4 border-b border-slate-100 flex flex-wrap items-center gap-2">
        <Search size={15} className="text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm theo mã hoặc tên..." className="flex-1 min-w-[140px] text-sm outline-none" />
        <select value={filter} onChange={(e) => setFilter(e.target.value)} className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
          <option value="all">Tất cả</option>
          <option value="NL">Nguyên vật liệu</option>
          <option value="TP">Thành phẩm</option>
        </select>
      </div>
      {filtered.length === 0 ? <EmptyState icon={Package} text="Chưa có sản phẩm nào." /> : (
        <div className="divide-y divide-slate-100 max-h-96 overflow-y-auto">
          {filtered.map((p) => (
            <div key={p.id} className="p-3.5 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <p className="text-sm font-medium text-slate-800 truncate">{p.name}</p>
                <p className="text-xs text-slate-400">Mã: {p.code} · ĐVT: {p.unit} · {p.groupCode}/{p.groupName}</p>
              </div>
              <Badge className={CLASSIFICATION_META[p.classification]?.color}>{p.classification}</Badge>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function AddSimpleCodeForm({ list, prefix, label, onAdd }) {
  const [code, setCode] = useState(() => nextSimpleCode(list, prefix));
  const [name, setName] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!code.trim() || !name.trim()) { setError("Vui lòng nhập đủ Mã và Tên."); return; }
    setError(""); setSaving(true);
    try {
      await onAdd({ code: code.trim(), name: name.trim() });
      setCode(nextSimpleCode([...list, { code }], prefix));
      setName("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <p className="font-semibold text-slate-800 text-sm mb-3">Thêm {label.toLowerCase()} mới</p>
      <div className="grid sm:grid-cols-3 gap-3 mb-3">
        <TextField label="Mã (gợi ý tự động)" value={code} onChange={(e) => setCode(e.target.value)} />
        <TextField label={`Tên ${label.toLowerCase()}`} value={name} onChange={(e) => setName(e.target.value)} className="sm:col-span-2" />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Thêm {label.toLowerCase()}</PrimaryButton>
    </Card>
  );
}

function SimpleCodeList({ list, icon: Icon, emptyText }) {
  const [q, setQ] = useState("");
  const filtered = q ? list.filter((x) => stripDiacritics(x.name).includes(stripDiacritics(q)) || x.code.toLowerCase().includes(q.toLowerCase())) : list;
  return (
    <Card className="p-0 overflow-hidden">
      <div className="p-4 border-b border-slate-100 flex items-center gap-2">
        <Search size={15} className="text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm theo mã hoặc tên..." className="flex-1 text-sm outline-none" />
      </div>
      {filtered.length === 0 ? <EmptyState icon={Icon} text={emptyText} /> : (
        <div className="divide-y divide-slate-100 max-h-96 overflow-y-auto">
          {filtered.map((x) => (
            <div key={x.id} className="p-3.5">
              <p className="text-sm font-medium text-slate-800">{x.name}</p>
              <p className="text-xs text-slate-400">Mã: {x.code}</p>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function DanhMucModule({ data, onAddSupplier, onAddProduct, onAddRevenueCode, onAddExportCode }) {
  const [tab, setTab] = useState("ncc");
  const TABS = [
    { key: "ncc", label: "Nhà cung cấp", icon: Truck },
    { key: "sp", label: "NVL & Thành phẩm", icon: Package },
    { key: "dt", label: "Mã doanh thu", icon: TrendingUp },
    { key: "xuat", label: "Mã xuất", icon: ArrowUpCircle },
  ];
  return (
    <div>
      <SectionTitle icon={Boxes} title="Danh mục" subtitle="Dữ liệu gốc — thêm mới bất cứ lúc nào, mã tự động gợi ý" />
      <div className="flex gap-1 mb-4 overflow-x-auto">
        {TABS.map((t) => (
          <button key={t.key} onClick={() => setTab(t.key)} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-medium whitespace-nowrap transition ${tab === t.key ? "bg-sky-800 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100"}`}>
            <t.icon size={15} /> {t.label}
          </button>
        ))}
      </div>
      {tab === "ncc" && (<><AddSupplierForm suppliers={data.suppliers} onAdd={onAddSupplier} /><SupplierList suppliers={data.suppliers} /></>)}
      {tab === "sp" && (<><AddProductForm products={data.products} onAdd={onAddProduct} /><ProductList products={data.products} /></>)}
      {tab === "dt" && (<><AddSimpleCodeForm list={data.revenueCodes} prefix="DT" label="Mã doanh thu" onAdd={onAddRevenueCode} /><SimpleCodeList list={data.revenueCodes} icon={TrendingUp} emptyText="Chưa có mã doanh thu nào." /></>)}
      {tab === "xuat" && (<><AddSimpleCodeForm list={data.exportCodes} prefix="MX" label="Mã xuất" onAdd={onAddExportCode} /><SimpleCodeList list={data.exportCodes} icon={ArrowUpCircle} emptyText="Chưa có mã xuất nào." /></>)}
    </div>
  );
}

// ---------------------------------------------------------------------------
// NHẬP HÀNG — Tác vụ 2 (nhập liệu) + Tác vụ 3 (báo cáo nhập)
// ---------------------------------------------------------------------------
// Cặp ô "Mã" ↔ "Tên" liên kết 2 chiều: gõ đúng mã thì tên tự nhảy ra, gõ đúng tên
// thì mã tự nhảy ra; gõ 1 phần thì hiện gợi ý để bấm chọn (khớp cả mã lẫn tên).
function ProductCodeNameFields({ products, productId, onSelectProduct, codeLabel = "Mã sản phẩm", nameLabel = "Tên sản phẩm" }) {
  const selected = products.find((p) => p.id === productId);
  const [codeText, setCodeText] = useState(selected?.code || "");
  const [nameText, setNameText] = useState(selected?.name || "");
  const [showCodeList, setShowCodeList] = useState(false);
  const [showNameList, setShowNameList] = useState(false);

  useEffect(() => {
    const p = products.find((x) => x.id === productId);
    setCodeText(p?.code || "");
    setNameText(p?.name || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [productId]);

  const codeMatches = codeText
    ? products.filter((p) => p.code.toLowerCase().includes(codeText.trim().toLowerCase())).slice(0, 8)
    : [];
  const nameMatches = nameText
    ? products.filter((p) => stripDiacritics(p.name).includes(stripDiacritics(nameText.trim()))).slice(0, 8)
    : [];

  const handleCodeChange = (v) => {
    setCodeText(v);
    const exact = products.find((p) => p.code.toLowerCase() === v.trim().toLowerCase());
    if (exact) { onSelectProduct(exact.id); setNameText(exact.name); }
    else if (productId) onSelectProduct("");
  };
  const handleNameChange = (v) => {
    setNameText(v);
    const exact = products.find((p) => p.name.toLowerCase() === v.trim().toLowerCase());
    if (exact) { onSelectProduct(exact.id); setCodeText(exact.code); }
    else if (productId) onSelectProduct("");
  };
  const pick = (p) => {
    onSelectProduct(p.id);
    setCodeText(p.code); setNameText(p.name);
    setShowCodeList(false); setShowNameList(false);
  };

  return (
    <div className="grid sm:grid-cols-2 gap-2">
      <label className="block relative">
        <span className="block text-[11px] text-slate-500 mb-1">{codeLabel}</span>
        <input
          value={codeText}
          onChange={(e) => { handleCodeChange(e.target.value); setShowCodeList(true); }}
          onFocus={() => setShowCodeList(true)}
          onBlur={() => setTimeout(() => setShowCodeList(false), 150)}
          placeholder="Gõ mã..."
          className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600"
        />
        {showCodeList && codeMatches.length > 0 && (
          <div className="absolute z-20 left-0 right-0 mt-1 bg-white rounded-xl border border-slate-200 shadow-lg max-h-56 overflow-y-auto">
            {codeMatches.map((p) => (
              <button type="button" key={p.id} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(p)} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <span className="text-sky-600 font-medium">{p.code}</span> {p.name} <span className="text-slate-400 text-xs">({p.unit})</span>
              </button>
            ))}
          </div>
        )}
      </label>
      <label className="block relative">
        <span className="block text-[11px] text-slate-500 mb-1">{nameLabel}</span>
        <input
          value={nameText}
          onChange={(e) => { handleNameChange(e.target.value); setShowNameList(true); }}
          onFocus={() => setShowNameList(true)}
          onBlur={() => setTimeout(() => setShowNameList(false), 150)}
          placeholder="Gõ tên..."
          className="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600"
        />
        {showNameList && nameMatches.length > 0 && (
          <div className="absolute z-20 left-0 right-0 mt-1 bg-white rounded-xl border border-slate-200 shadow-lg max-h-56 overflow-y-auto">
            {nameMatches.map((p) => (
              <button type="button" key={p.id} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(p)} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <span className="text-sky-600 font-medium">{p.code}</span> {p.name} <span className="text-slate-400 text-xs">({p.unit})</span>
              </button>
            ))}
          </div>
        )}
      </label>
    </div>
  );
}

// Cặp ô "Mã" ↔ "Tên" gọn, dùng bên trong 1 dòng của bảng kiểu Excel (không có
// nhãn/khung riêng như ProductCodeNameFields, chỉ 2 ô input cạnh nhau).
function ProductCodeNameCells({ products, productId, onSelectProduct, codePlaceholder = "Mã...", namePlaceholder = "Tên...", codeInputId, nameInputId, onCellKeyDown }) {
  const selected = products.find((p) => p.id === productId);
  const [codeText, setCodeText] = useState(selected?.code || "");
  const [nameText, setNameText] = useState(selected?.name || "");
  const [showCodeList, setShowCodeList] = useState(false);
  const [showNameList, setShowNameList] = useState(false);

  useEffect(() => {
    const p = products.find((x) => x.id === productId);
    setCodeText(p?.code || "");
    setNameText(p?.name || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [productId]);

  const codeMatches = codeText
    ? products.filter((p) => p.code.toLowerCase().includes(codeText.trim().toLowerCase())).slice(0, 8)
    : [];
  const nameMatches = nameText
    ? products.filter((p) => stripDiacritics(p.name).includes(stripDiacritics(nameText.trim()))).slice(0, 8)
    : [];

  const handleCodeChange = (v) => {
    setCodeText(v);
    const exact = products.find((p) => p.code.toLowerCase() === v.trim().toLowerCase());
    if (exact) { onSelectProduct(exact.id); setNameText(exact.name); }
    else if (productId) onSelectProduct("");
  };
  const handleNameChange = (v) => {
    setNameText(v);
    const exact = products.find((p) => p.name.toLowerCase() === v.trim().toLowerCase());
    if (exact) { onSelectProduct(exact.id); setCodeText(exact.code); }
    else if (productId) onSelectProduct("");
  };
  const pick = (p) => {
    onSelectProduct(p.id);
    setCodeText(p.code); setNameText(p.name);
    setShowCodeList(false); setShowNameList(false);
  };

  return (
    <>
      <td className="px-2 py-1.5 relative">
        <input
          id={codeInputId}
          value={codeText}
          onChange={(e) => { handleCodeChange(e.target.value); setShowCodeList(true); }}
          onFocus={() => setShowCodeList(true)}
          onBlur={() => setTimeout(() => setShowCodeList(false), 150)}
          onKeyDown={(e) => onCellKeyDown?.(e, 0)}
          placeholder={codePlaceholder}
          className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600"
        />
        {showCodeList && codeMatches.length > 0 && (
          <div className="absolute z-20 left-2 right-2 mt-1 bg-white rounded-xl border border-slate-200 shadow-lg max-h-56 overflow-y-auto">
            {codeMatches.map((p) => (
              <button type="button" key={p.id} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(p)} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <span className="text-sky-600 font-medium">{p.code}</span> {p.name} <span className="text-slate-400 text-xs">({p.unit})</span>
              </button>
            ))}
          </div>
        )}
      </td>
      <td className="px-2 py-1.5 relative">
        <input
          id={nameInputId}
          value={nameText}
          onChange={(e) => { handleNameChange(e.target.value); setShowNameList(true); }}
          onFocus={() => setShowNameList(true)}
          onBlur={() => setTimeout(() => setShowNameList(false), 150)}
          onKeyDown={(e) => onCellKeyDown?.(e, 1)}
          placeholder={namePlaceholder}
          className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40 focus:border-sky-600"
        />
        {showNameList && nameMatches.length > 0 && (
          <div className="absolute z-20 left-2 right-2 mt-1 bg-white rounded-xl border border-slate-200 shadow-lg max-h-56 overflow-y-auto">
            {nameMatches.map((p) => (
              <button type="button" key={p.id} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(p)} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <span className="text-sky-600 font-medium">{p.code}</span> {p.name} <span className="text-slate-400 text-xs">({p.unit})</span>
              </button>
            ))}
          </div>
        )}
      </td>
    </>
  );
}

function NhapHangForm({ data, currentUser, onSubmit }) {
  const [orderNumber, setOrderNumber] = useState("");
  const [supplierId, setSupplierId] = useState(data.suppliers[0]?.id || "");
  const [lines, setLines] = useState([{ key: Math.random().toString(36).slice(2), productId: "", quantity: "", unitPrice: "" }]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [pendingFocusIndex, setPendingFocusIndex] = useState(null);

  const supplier = data.suppliers.find((s) => s.id === supplierId);

  const updateLine = (key, patch) => setLines((prev) => prev.map((l) => (l.key === key ? { ...l, ...patch } : l)));
  const addRow = () => {
    const newIndex = lines.length;
    setLines((prev) => [...prev, { key: Math.random().toString(36).slice(2), productId: "", quantity: "", unitPrice: "" }]);
    setPendingFocusIndex(newIndex);
  };
  const removeRow = (key) => setLines((prev) => (prev.length > 1 ? prev.filter((l) => l.key !== key) : prev));

  // Điều hướng ô bằng phím mũi tên kiểu Excel — không bị "khoá" trong 1 ô,
  // quay lại sửa dữ liệu dòng trước đó mà không cần bấm chuột.
  // col: 0=Mã SP, 1=Tên SP, 2=Số lượng, 3=Đơn giá
  const focusCell = (rowIndex, col) => {
    const el = document.getElementById(`nhap-cell-${rowIndex}-${col}`);
    if (el) {
      el.focus();
      try { el.select(); } catch (_) { /* 1 số trình duyệt không hỗ trợ select() trên input number, bỏ qua */ }
    }
  };

  // Sau khi thêm dòng mới (bấm nút hoặc Tab ở ô Đơn giá của dòng cuối), tự đưa
  // con trỏ vào ô "Mã SP" của dòng vừa thêm để gõ tiếp luôn không cần bấm chuột.
  useEffect(() => {
    if (pendingFocusIndex !== null) {
      focusCell(pendingFocusIndex, 0);
      setPendingFocusIndex(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lines, pendingFocusIndex]);

  // Bấm Tab ở ô Đơn giá của dòng CUỐI CÙNG → tự thêm dòng mới thay vì nhảy ra
  // khỏi bảng (Shift+Tab hoặc không phải dòng cuối thì vẫn Tab bình thường).
  const handlePriceKeyDown = (e, isLastRow) => {
    if (e.key === "Tab" && !e.shiftKey && isLastRow) {
      e.preventDefault();
      addRow();
    }
  };

  const handleCellKeyDown = (e, rowIndex, col) => {
    if (e.key === "ArrowUp") {
      e.preventDefault();
      if (rowIndex > 0) focusCell(rowIndex - 1, col);
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      if (rowIndex < lines.length - 1) focusCell(rowIndex + 1, col);
    } else if (e.key === "ArrowLeft") {
      // Chỉ nhảy sang ô bên trái khi con trỏ đang ở đầu chữ (không phá việc di chuyển con trỏ trong ô)
      const atStart = e.target.selectionStart === 0 && e.target.selectionEnd === 0;
      if (atStart && col > 0) { e.preventDefault(); focusCell(rowIndex, col - 1); }
    } else if (e.key === "ArrowRight") {
      const len = e.target.value?.length ?? 0;
      const atEnd = e.target.selectionStart === len && e.target.selectionEnd === len;
      if (atEnd && col < 3) { e.preventDefault(); focusCell(rowIndex, col + 1); }
    }
  };

  const validLines = lines.filter((l) => l.productId && Number(l.quantity) > 0 && Number(l.unitPrice) >= 0);
  const grandTotal = validLines.reduce((s, l) => s + Number(l.quantity) * Number(l.unitPrice), 0);

  const submit = async () => {
    if (!supplierId) { setError("Vui lòng chọn Nhà cung cấp."); return; }
    if (validLines.length === 0) { setError("Cần ít nhất 1 dòng đủ Mã SP, Số lượng, Đơn giá."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({
        orderNumber: orderNumber.trim(), supplierId,
        lines: validLines.map((l) => ({ productId: l.productId, quantity: Number(l.quantity), unitPrice: Number(l.unitPrice), totalAmount: Number(l.quantity) * Number(l.unitPrice) })),
        paymentType: supplier?.paymentType || "cong_no",
      });
      setOrderNumber("");
      setLines([{ key: Math.random().toString(36).slice(2), productId: "", quantity: "", unitPrice: "" }]);
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Nhập hàng mới</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Đơn số (tự đặt, không bắt buộc)" value={orderNumber} onChange={(e) => setOrderNumber(e.target.value)} />
        <SelectField label="Nhà cung cấp" value={supplierId} onChange={(e) => setSupplierId(e.target.value)}>
          {data.suppliers.map((s) => <option key={s.id} value={s.id}>{s.code} — {s.name}</option>)}
        </SelectField>
        {supplier && (
          <p className="text-xs text-slate-400 sm:col-span-2 -mt-1">Tình trạng thanh toán sẽ tự ghi cho cả phiếu: <Badge className={PAYMENT_TYPE_META[supplier.paymentType]?.color}>{PAYMENT_TYPE_META[supplier.paymentType]?.label}</Badge></p>
        )}
      </div>

      <div className="overflow-x-auto rounded-xl border border-slate-200 mb-3">
        <table className="w-full text-sm min-w-[720px]">
          <thead>
            <tr className="text-left text-xs text-slate-500 bg-slate-50 border-b border-slate-200">
              <th className="px-2 py-2 w-40">Mã SP</th>
              <th className="px-2 py-2">Tên SP</th>
              <th className="px-2 py-2 w-16">ĐVT</th>
              <th className="px-2 py-2 w-28">Số lượng</th>
              <th className="px-2 py-2 w-32">Đơn giá</th>
              <th className="px-2 py-2 w-32 text-right">Thành tiền</th>
              <th className="px-2 py-2 w-10"></th>
            </tr>
          </thead>
          <tbody>
            {lines.map((l, idx) => {
              const p = data.products.find((x) => x.id === l.productId);
              const rowTotal = (Number(l.quantity) || 0) * (Number(l.unitPrice) || 0);
              const isLastRow = idx === lines.length - 1;
              return (
                <tr key={l.key} className="border-b border-slate-100 last:border-0">
                  <ProductCodeNameCells
                    products={data.products}
                    productId={l.productId}
                    onSelectProduct={(id) => updateLine(l.key, { productId: id })}
                    codePlaceholder="Mã SP"
                    namePlaceholder="Tên SP"
                    codeInputId={`nhap-cell-${idx}-0`}
                    nameInputId={`nhap-cell-${idx}-1`}
                    onCellKeyDown={(e, col) => handleCellKeyDown(e, idx, col)}
                  />
                  <td className="px-2 py-1.5 text-slate-500">{p?.unit || "—"}</td>
                  <td className="px-2 py-1.5">
                    <input
                      id={`nhap-cell-${idx}-2`}
                      type="number"
                      value={l.quantity}
                      onChange={(e) => updateLine(l.key, { quantity: e.target.value })}
                      onKeyDown={(e) => handleCellKeyDown(e, idx, 2)}
                      placeholder="0"
                      className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40"
                    />
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      id={`nhap-cell-${idx}-3`}
                      type="number"
                      value={l.unitPrice}
                      onChange={(e) => updateLine(l.key, { unitPrice: e.target.value })}
                      onKeyDown={(e) => { handleCellKeyDown(e, idx, 3); handlePriceKeyDown(e, isLastRow); }}
                      placeholder="0"
                      className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40"
                    />
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium text-slate-700">{rowTotal > 0 ? fmtMoney(rowTotal) : "—"}</td>
                  <td className="px-2 py-1.5 text-center">
                    <button
                      type="button"
                      onClick={() => removeRow(l.key)}
                      onKeyDown={(e) => handlePriceKeyDown(e, isLastRow)}
                      className="text-slate-400 hover:text-rose-600"
                    ><X size={15} /></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="bg-sky-50/60">
              <td colSpan={5} className="px-2 py-2 text-right text-sky-700 font-medium">Tổng thành tiền cả phiếu</td>
              <td className="px-2 py-2 text-right font-semibold text-sky-800">{fmtMoney(grandTotal)}</td>
              <td></td>
            </tr>
          </tfoot>
        </table>
      </div>
      <GhostButton type="button" onClick={addRow} className="mb-3"><Plus size={14} /> Thêm dòng</GhostButton>

      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <ArrowDownCircle size={15} />} Lưu phiếu nhập</PrimaryButton>
    </Card>
  );
}

function NhapHangList({ data, rows, showFilterHint }) {
  return (
    <Card className="p-0 overflow-hidden">
      {showFilterHint !== false && (
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Lịch sử nhập hàng gần đây</p></div>
      )}
      {rows.length === 0 ? <EmptyState icon={Inbox} text="Không có phiếu nhập nào khớp." /> : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Mã phiếu</th><th className="px-3 py-2">Ngày</th><th className="px-3 py-2">NCC</th>
              <th className="px-3 py-2">Sản phẩm</th><th className="px-3 py-2 text-right">SL</th><th className="px-3 py-2 text-right">Đơn giá</th>
              <th className="px-3 py-2 text-right">Thành tiền</th><th className="px-3 py-2">TT thanh toán</th>
            </tr></thead>
            <tbody>
              {rows.map((r) => {
                const s = data.suppliers.find((x) => x.id === r.supplierId);
                const p = data.products.find((x) => x.id === r.productId);
                return (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2 text-slate-500">{r.receiptCode}</td>
                    <td className="px-3 py-2 text-slate-500">{fmtDate(r.importDate)}</td>
                    <td className="px-3 py-2">{s?.name || "—"}</td>
                    <td className="px-3 py-2">{p?.name || "—"}</td>
                    <td className="px-3 py-2 text-right">{fmtNumber(r.quantity)} {p?.unit}</td>
                    <td className="px-3 py-2 text-right">{fmtMoney(r.unitPrice)}</td>
                    <td className="px-3 py-2 text-right font-medium">{fmtMoney(r.totalAmount)}</td>
                    <td className="px-3 py-2"><Badge className={PAYMENT_TYPE_META[r.paymentType]?.color}>{PAYMENT_TYPE_META[r.paymentType]?.label}</Badge></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

// Ô tiêu đề cột kiểu Excel: bấm để sắp xếp, bấm icon phễu để mở danh sách check-box lọc theo giá trị.
function ExcelHeaderFilter({ label, align = "left", options, selected, onChangeSelected, sortDir, onSort, className = "" }) {
  const [open, setOpen] = useState(false);
  const allChecked = selected === null;
  const isChecked = (v) => allChecked || selected.has(v);

  const toggleValue = (v) => {
    const next = allChecked ? new Set(options.map((o) => o.value)) : new Set(selected);
    if (next.has(v)) next.delete(v); else next.add(v);
    onChangeSelected(next.size === options.length ? null : next);
  };
  const selectAll = () => onChangeSelected(null);
  const clearAll = () => onChangeSelected(new Set());

  return (
    <th className={`px-3 py-2 relative select-none ${align === "right" ? "text-right" : "text-left"} ${className}`}>
      <div className={`flex items-center gap-1 ${align === "right" ? "justify-end" : ""}`}>
        <button type="button" onClick={onSort} className="flex items-center gap-1 hover:text-sky-700">
          {label}
          <ArrowUpDown size={11} className={sortDir ? "text-sky-600" : "text-slate-300"} />
        </button>
        {options && (
          <button type="button" onClick={() => setOpen((o) => !o)} className={`p-0.5 rounded hover:bg-slate-100 ${!allChecked ? "text-sky-600" : "text-slate-300"}`}>
            <Filter size={11} />
          </button>
        )}
      </div>
      {open && options && (
        <div className="absolute z-30 top-full mt-1 left-0 w-56 bg-white rounded-xl border border-slate-200 shadow-lg text-xs font-normal normal-case text-slate-700 max-h-64 overflow-y-auto">
          <div className="flex items-center justify-between px-3 py-2 border-b border-slate-100 sticky top-0 bg-white">
            <button type="button" onClick={selectAll} className="text-sky-700 hover:underline">Chọn tất cả</button>
            <button type="button" onClick={clearAll} className="text-slate-400 hover:underline">Bỏ chọn</button>
          </div>
          {options.map((o) => (
            <label key={o.value} className="flex items-center gap-2 px-3 py-1.5 hover:bg-slate-50 cursor-pointer">
              <input type="checkbox" checked={isChecked(o.value)} onChange={() => toggleValue(o.value)} className="rounded border-slate-300 text-sky-600 focus:ring-sky-500" />
              <span className="truncate">{o.label}</span>
            </label>
          ))}
          <div className="p-2 border-t border-slate-100 sticky bottom-0 bg-white">
            <button type="button" onClick={() => setOpen(false)} className="w-full text-center text-sky-700 text-xs py-1 hover:bg-sky-50 rounded-lg">Đóng</button>
          </div>
        </div>
      )}
    </th>
  );
}

// ---------------------------------------------------------------------------
// LỊCH SỬ NHẬP HÀNG — màn riêng, có bộ lọc chi tiết theo ngày/giá/mã phiếu/SL/thành tiền
// ---------------------------------------------------------------------------
// Xoá gọn cả 1 phiếu (nhập hoặc xuất) theo Mã phiếu — không cần lọc/tìm từng dòng.
function DeleteByReceiptCard({ records, onDeleteMany, label = "phiếu" }) {
  const receipts = useMemo(() => {
    const map = new Map();
    records.forEach((r) => {
      if (!r.receiptCode) return;
      const d = r.importDate || r.exportDate;
      const cur = map.get(r.receiptCode) || { receiptCode: r.receiptCode, lineCount: 0, totalAmount: 0, minDate: d, maxDate: d };
      cur.lineCount += 1;
      cur.totalAmount += r.totalAmount;
      if (d && (!cur.minDate || d < cur.minDate)) cur.minDate = d;
      if (d && (!cur.maxDate || d > cur.maxDate)) cur.maxDate = d;
      map.set(r.receiptCode, cur);
    });
    return Array.from(map.values()).sort((a, b) => (b.maxDate || "").localeCompare(a.maxDate || ""));
  }, [records]);

  const [selected, setSelected] = useState("");
  const [deleting, setDeleting] = useState(false);
  const [deletingAll, setDeletingAll] = useState(false);
  const [error, setError] = useState("");
  const chosen = receipts.find((r) => r.receiptCode === selected);

  const totalAllLines = records.length;
  const totalAllAmount = records.reduce((s, r) => s + r.totalAmount, 0);

  const handleDelete = async () => {
    if (!chosen) return;
    if (!window.confirm(`Xoá toàn bộ ${label} "${chosen.receiptCode}" (${chosen.lineCount} dòng, tổng ${fmtMoney(chosen.totalAmount)})? Không thể hoàn tác.`)) return;
    setDeleting(true); setError("");
    try {
      const ids = records.filter((r) => r.receiptCode === chosen.receiptCode).map((r) => r.id);
      await onDeleteMany(ids);
      setSelected("");
    } catch (e) {
      setError(e.message || "Không xoá được, vui lòng thử lại.");
    } finally {
      setDeleting(false);
    }
  };

  const handleDeleteAll = async () => {
    if (records.length === 0) return;
    if (!window.confirm(`Xoá SẠCH TOÀN BỘ lịch sử ${label} — tất cả ${receipts.length} mã phiếu, ${totalAllLines} dòng, tổng ${fmtMoney(totalAllAmount)}? Hành động này xoá hết mọi phiếu (không chỉ 1 phiếu), không thể hoàn tác.`)) return;
    setDeletingAll(true); setError("");
    try {
      const ids = records.map((r) => r.id);
      await onDeleteMany(ids);
      setSelected("");
    } catch (e) {
      setError(e.message || "Không xoá được, vui lòng thử lại.");
    } finally {
      setDeletingAll(false);
    }
  };

  if (receipts.length === 0) return null;

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Xoá nhanh theo Mã phiếu</p>
      <div className="flex flex-wrap items-end gap-3">
        <div className="flex-1 min-w-[220px]">
          <SelectField label="Chọn phiếu" value={selected} onChange={(e) => setSelected(e.target.value)}>
            <option value="">— Chọn mã phiếu —</option>
            {receipts.map((r) => (
              <option key={r.receiptCode} value={r.receiptCode}>
                {r.receiptCode} · {r.minDate === r.maxDate ? fmtDate(r.minDate) : `${fmtDate(r.minDate)}–${fmtDate(r.maxDate)}`} · {r.lineCount} dòng · {fmtMoney(r.totalAmount)}
              </option>
            ))}
          </SelectField>
        </div>
        <button
          type="button"
          onClick={handleDelete}
          disabled={!chosen || deleting || deletingAll}
          className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-sm font-medium text-rose-600 border border-rose-200 bg-rose-50 hover:bg-rose-100 disabled:opacity-50"
        >
          {deleting ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />} Xoá phiếu này
        </button>
      </div>
      {receipts.length > 1 && (
        <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between flex-wrap gap-2">
          <p className="text-xs text-slate-400">Có {receipts.length} mã phiếu, tổng {totalAllLines} dòng — nếu chỉ muốn làm sạch hết để import lại từ đầu, xoá 1 lần luôn:</p>
          <button
            type="button"
            onClick={handleDeleteAll}
            disabled={deleting || deletingAll}
            className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-sm font-medium text-white bg-rose-600 hover:bg-rose-700 disabled:opacity-50 shrink-0"
          >
            {deletingAll ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />} Xoá SẠCH toàn bộ ({receipts.length} phiếu)
          </button>
        </div>
      )}
      {error && <p className="text-xs text-rose-600 mt-2 flex items-center gap-1"><AlertTriangle size={12} className="shrink-0" /> {error}</p>}
    </Card>
  );
}

function LichSuNhapModule({ data, onDelete, onDeleteMany }) {
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());
  const [receiptCode, setReceiptCode] = useState("");
  const [productQuery, setProductQuery] = useState("");
  const [supplierId, setSupplierId] = useState("");
  const [priceFrom, setPriceFrom] = useState("");
  const [priceTo, setPriceTo] = useState("");
  const [qtyFrom, setQtyFrom] = useState("");
  const [qtyTo, setQtyTo] = useState("");
  const [amountFrom, setAmountFrom] = useState("");
  const [amountTo, setAmountTo] = useState("");

  // Bộ lọc kiểu Excel theo từng cột (null = chọn tất cả)
  const [colFilters, setColFilters] = useState({ receiptCode: null, supplierId: null, productId: null, paymentType: null });
  const [sortKey, setSortKey] = useState("importDate");
  const [sortDir, setSortDir] = useState("desc");

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("asc"); }
  };

  const preFiltered = data.importRecords.filter((r) => {
    if (from && r.importDate < from) return false;
    if (to && r.importDate > to) return false;
    if (receiptCode && !r.receiptCode?.toLowerCase().includes(receiptCode.trim().toLowerCase())) return false;
    if (supplierId && r.supplierId !== supplierId) return false;
    if (productQuery) {
      const p = data.products.find((x) => x.id === r.productId);
      const q = stripDiacritics(productQuery.trim());
      if (!p || (!stripDiacritics(p.name).includes(q) && !p.code.toLowerCase().includes(productQuery.trim().toLowerCase()))) return false;
    }
    if (priceFrom !== "" && r.unitPrice < Number(priceFrom)) return false;
    if (priceTo !== "" && r.unitPrice > Number(priceTo)) return false;
    if (qtyFrom !== "" && r.quantity < Number(qtyFrom)) return false;
    if (qtyTo !== "" && r.quantity > Number(qtyTo)) return false;
    if (amountFrom !== "" && r.totalAmount < Number(amountFrom)) return false;
    if (amountTo !== "" && r.totalAmount > Number(amountTo)) return false;
    return true;
  });

  // Danh sách lựa chọn cho từng cột kiểu Excel — lấy từ tập đã lọc phía trên
  const receiptOptions = useMemo(() => {
    const codes = Array.from(new Set(preFiltered.map((r) => r.receiptCode).filter(Boolean)));
    return codes.sort().map((c) => ({ value: c, label: c }));
  }, [preFiltered]);
  const supplierOptions = useMemo(() => {
    const ids = Array.from(new Set(preFiltered.map((r) => r.supplierId).filter(Boolean)));
    return ids.map((id) => ({ value: id, label: data.suppliers.find((s) => s.id === id)?.name || "—" })).sort((a, b) => a.label.localeCompare(b.label));
  }, [preFiltered, data.suppliers]);
  const productOptions = useMemo(() => {
    const ids = Array.from(new Set(preFiltered.map((r) => r.productId).filter(Boolean)));
    return ids.map((id) => ({ value: id, label: data.products.find((p) => p.id === id)?.name || "—" })).sort((a, b) => a.label.localeCompare(b.label));
  }, [preFiltered, data.products]);
  const paymentOptions = useMemo(() => {
    const types = Array.from(new Set(preFiltered.map((r) => r.paymentType).filter(Boolean)));
    return types.map((t) => ({ value: t, label: PAYMENT_TYPE_META[t]?.label || t }));
  }, [preFiltered]);

  let filtered = preFiltered.filter((r) => {
    if (colFilters.receiptCode && !colFilters.receiptCode.has(r.receiptCode)) return false;
    if (colFilters.supplierId && !colFilters.supplierId.has(r.supplierId)) return false;
    if (colFilters.productId && !colFilters.productId.has(r.productId)) return false;
    if (colFilters.paymentType && !colFilters.paymentType.has(r.paymentType)) return false;
    return true;
  });

  filtered = [...filtered].sort((a, b) => {
    let av, bv;
    if (sortKey === "receiptCode") { av = a.receiptCode || ""; bv = b.receiptCode || ""; }
    else if (sortKey === "importDate") { av = a.importDate || ""; bv = b.importDate || ""; }
    else if (sortKey === "supplier") { av = data.suppliers.find((s) => s.id === a.supplierId)?.name || ""; bv = data.suppliers.find((s) => s.id === b.supplierId)?.name || ""; }
    else if (sortKey === "product") { av = data.products.find((p) => p.id === a.productId)?.name || ""; bv = data.products.find((p) => p.id === b.productId)?.name || ""; }
    else { av = a[sortKey]; bv = b[sortKey]; }
    if (av < bv) return sortDir === "asc" ? -1 : 1;
    if (av > bv) return sortDir === "asc" ? 1 : -1;
    return 0;
  });

  const totalAmount = filtered.reduce((s, r) => s + r.totalAmount, 0);
  const [deleting, setDeleting] = useState(null);
  const [bulkDeleting, setBulkDeleting] = useState(false);

  const handleDeleteRow = async (id) => {
    if (!window.confirm("Xoá dòng nhập hàng này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };
  const handleDeleteAllFiltered = async () => {
    if (filtered.length === 0) return;
    if (!window.confirm(`Xoá toàn bộ ${filtered.length} dòng đang hiển thị (theo bộ lọc hiện tại)? Không thể hoàn tác.`)) return;
    setBulkDeleting(true);
    try { await onDeleteMany(filtered.map((r) => r.id)); } finally { setBulkDeleting(false); }
  };

  const resetFilters = () => {
    setFrom(daysAgoISO(30)); setTo(todayISO()); setReceiptCode(""); setProductQuery("");
    setSupplierId(""); setPriceFrom(""); setPriceTo(""); setQtyFrom(""); setQtyTo(""); setAmountFrom(""); setAmountTo("");
    setColFilters({ receiptCode: null, supplierId: null, productId: null, paymentType: null });
  };

  return (
    <div>
      <SectionTitle icon={Inbox} title="Lịch sử nhập hàng" subtitle="Tra cứu chi tiết các phiếu đã nhập theo ngày, giá, mã phiếu, số lượng, thành tiền" />

      <DeleteByReceiptCard records={data.importRecords} onDeleteMany={onDeleteMany} label="phiếu nhập" />

      <Card className="p-4 sm:p-5 mb-5">
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
        <div className="flex items-center justify-between mt-3">
          <div className="flex items-center gap-2">
            <GhostButton onClick={resetFilters}><X size={14} /> Xoá bộ lọc</GhostButton>
            {filtered.length > 0 && (
              <button type="button" onClick={handleDeleteAllFiltered} disabled={bulkDeleting} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-medium text-rose-600 border border-rose-200 bg-rose-50 hover:bg-rose-100 disabled:opacity-50">
                {bulkDeleting ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />} Xoá {filtered.length} dòng đang hiển thị
              </button>
            )}
          </div>
          <p className="text-sm text-slate-500">
            <span className="font-medium text-slate-700">{filtered.length}</span> dòng · Tổng thành tiền{" "}
            <span className="font-semibold text-sky-700">{fmtMoney(totalAmount)}</span>
          </p>
        </div>
      </Card>

      <Card className="p-0 overflow-visible">
        {filtered.length === 0 ? <EmptyState icon={Inbox} text="Không có phiếu nhập nào khớp." /> : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-slate-400 border-b border-slate-100 bg-slate-50/60">
                  <ExcelHeaderFilter label="Mã phiếu" options={receiptOptions} selected={colFilters.receiptCode} onChangeSelected={(v) => setColFilters((f) => ({ ...f, receiptCode: v }))} sortDir={sortKey === "receiptCode" ? sortDir : null} onSort={() => toggleSort("receiptCode")} />
                  <ExcelHeaderFilter label="Ngày" sortDir={sortKey === "importDate" ? sortDir : null} onSort={() => toggleSort("importDate")} />
                  <ExcelHeaderFilter label="NCC" options={supplierOptions} selected={colFilters.supplierId} onChangeSelected={(v) => setColFilters((f) => ({ ...f, supplierId: v }))} sortDir={sortKey === "supplier" ? sortDir : null} onSort={() => toggleSort("supplier")} />
                  <ExcelHeaderFilter label="Sản phẩm" options={productOptions} selected={colFilters.productId} onChangeSelected={(v) => setColFilters((f) => ({ ...f, productId: v }))} sortDir={sortKey === "product" ? sortDir : null} onSort={() => toggleSort("product")} />
                  <ExcelHeaderFilter label="SL" align="right" sortDir={sortKey === "quantity" ? sortDir : null} onSort={() => toggleSort("quantity")} />
                  <ExcelHeaderFilter label="Đơn giá" align="right" sortDir={sortKey === "unitPrice" ? sortDir : null} onSort={() => toggleSort("unitPrice")} />
                  <ExcelHeaderFilter label="Thành tiền" align="right" sortDir={sortKey === "totalAmount" ? sortDir : null} onSort={() => toggleSort("totalAmount")} />
                  <ExcelHeaderFilter label="TT thanh toán" options={paymentOptions} selected={colFilters.paymentType} onChangeSelected={(v) => setColFilters((f) => ({ ...f, paymentType: v }))} />
                  <th className="px-3 py-2 w-10"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((r) => {
                  const s = data.suppliers.find((x) => x.id === r.supplierId);
                  const p = data.products.find((x) => x.id === r.productId);
                  return (
                    <tr key={r.id} className="border-b border-slate-50 last:border-0">
                      <td className="px-3 py-2 text-slate-500">{r.receiptCode}</td>
                      <td className="px-3 py-2 text-slate-500">{fmtDate(r.importDate)}</td>
                      <td className="px-3 py-2">{s?.name || "—"}</td>
                      <td className="px-3 py-2">{p?.name || "—"}</td>
                      <td className="px-3 py-2 text-right">{fmtNumber(r.quantity)} {p?.unit}</td>
                      <td className="px-3 py-2 text-right">{fmtMoney(r.unitPrice)}</td>
                      <td className="px-3 py-2 text-right font-medium">{fmtMoney(r.totalAmount)}</td>
                      <td className="px-3 py-2"><Badge className={PAYMENT_TYPE_META[r.paymentType]?.color}>{PAYMENT_TYPE_META[r.paymentType]?.label}</Badge></td>
                      <td className="px-3 py-2 text-center">
                        <button type="button" onClick={() => handleDeleteRow(r.id)} disabled={deleting === r.id} className="text-slate-400 hover:text-rose-600 p-1">
                          {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

// Import Nhập hàng hàng loạt từ file Excel — mỗi dòng: Mã NCC, Mã SP, Số lượng, Đơn giá
// (tuỳ chọn: Đơn số, Ngày nhập). Không khớp được NCC/SP nào thì báo lỗi dòng đó, các dòng
// hợp lệ khác vẫn nhập được bình thường.
function NhapExcelImportForm({ data, onImport }) {
  const [fileName, setFileName] = useState("");
  const [preview, setPreview] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setDone(null);
    try {
      const rawRows = await readExcelRaw(file);
      const headerIdx = detectHeaderRow(rawRows, ["Mã NCC", "Mã SP"]);
      if (headerIdx === -1) { setErrors(['Không tìm thấy dòng tiêu đề (cần có cột "Mã NCC" và "Mã SP") trong file.']); setPreview([]); return; }
      const rawObjRows = rowsToObjects(rawRows, headerIdx);
      const valid = [];
      const errs = [];
      rawObjRows.forEach((row, idx) => {
        const supplierCode = String(pickCol(row, "Mã NCC", "Ma NCC") ?? "").trim();
        const productCode = String(pickCol(row, "Mã SP", "Ma SP", "Mã NVL") ?? "").trim();
        const quantity = Number(pickCol(row, "Số lượng", "So luong")) || 0;
        const unitPrice = Number(pickCol(row, "Đơn giá", "Don gia")) || 0;
        const orderNumber = String(pickCol(row, "Đơn số", "Don so") ?? "").trim();
        const importDate = excelDateToISO(pickCol(row, "Ngày nhập", "Ngay nhap"));
        if (!supplierCode && !productCode) return; // dòng trống, bỏ qua âm thầm
        const supplier = data.suppliers.find((s) => s.code.trim().toLowerCase() === supplierCode.toLowerCase());
        const product = data.products.find((p) => p.code.trim().toLowerCase() === productCode.toLowerCase());
        const excelRowNo = headerIdx + idx + 3; // +1 header 0-based→1-based, +1 dòng dữ liệu đầu tiên, +1 idx 0-based
        if (!supplier) { errs.push(`Dòng ${excelRowNo}: không tìm thấy Mã NCC "${supplierCode}"`); return; }
        if (!product) { errs.push(`Dòng ${excelRowNo}: không tìm thấy Mã SP "${productCode}"`); return; }
        if (quantity <= 0) { errs.push(`Dòng ${excelRowNo}: Số lượng không hợp lệ`); return; }
        valid.push({
          supplierId: supplier.id, supplierCode: supplier.code, productId: product.id, productCode: product.code, productName: product.name,
          quantity, unitPrice, orderNumber, importDate, paymentType: supplier.paymentType,
        });
      });
      setPreview(valid);
      setErrors(errs);
    } catch (err) {
      setErrors([err.message || "Không đọc được file, kiểm tra lại định dạng .xlsx."]);
      setPreview([]);
    }
  };

  const submit = async () => {
    if (preview.length === 0) return;
    setImporting(true);
    try {
      await onImport(preview);
      setDone(preview.length);
      setPreview([]); setFileName(""); setErrors([]);
    } catch (e) {
      setErrors([e.message || "Không nhập được, vui lòng thử lại."]);
    } finally {
      setImporting(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Nhập hàng loạt từ Excel</p>
      <p className="text-xs text-slate-500 mb-3">File cần có các cột: <b>Mã NCC</b>, <b>Mã SP</b>, <b>Số lượng</b>, <b>Đơn giá</b> (tuỳ chọn: Đơn số, Ngày nhập).</p>
      <div className="flex items-center gap-2 mb-3">
        <label className="inline-flex items-center gap-2 cursor-pointer text-sm font-medium text-sky-700 border border-sky-300 bg-sky-50 hover:bg-sky-100 rounded-xl px-4 py-2">
          <Upload size={15} /> {fileName || "Chọn file Excel..."}
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
        </label>
        {fileName && (
          <button type="button" onClick={() => { setFileName(""); setPreview([]); setErrors([]); }} className="text-slate-400 hover:text-rose-600 p-1.5" title="Bỏ file đã chọn">
            <X size={16} />
          </button>
        )}
      </div>

      {errors.length > 0 && (
        <div className="mb-3 text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-xl p-3 space-y-1 max-h-40 overflow-y-auto">
          {errors.map((e, i) => <p key={i} className="flex items-center gap-1"><AlertTriangle size={12} className="shrink-0" /> {e}</p>)}
        </div>
      )}

      {preview.length > 0 && (
        <>
          <div className="overflow-x-auto rounded-xl border border-slate-200 mb-3 max-h-72 overflow-y-auto">
            <table className="w-full text-sm min-w-[560px]">
              <thead className="sticky top-0 bg-slate-50">
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="px-2 py-2">NCC</th><th className="px-2 py-2">SP</th>
                  <th className="px-2 py-2 text-right">SL</th><th className="px-2 py-2 text-right">Đơn giá</th>
                  <th className="px-2 py-2 text-right">Thành tiền</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((r, i) => (
                  <tr key={i} className="border-b border-slate-100 last:border-0">
                    <td className="px-2 py-1.5">{r.supplierCode}</td>
                    <td className="px-2 py-1.5">{r.productCode} — {r.productName}</td>
                    <td className="px-2 py-1.5 text-right">{fmtNumber ? fmtNumber(r.quantity) : r.quantity}</td>
                    <td className="px-2 py-1.5 text-right">{fmtMoney(r.unitPrice)}</td>
                    <td className="px-2 py-1.5 text-right font-medium">{fmtMoney(r.quantity * r.unitPrice)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <PrimaryButton onClick={submit} disabled={importing}>{importing ? <Loader2 size={15} className="animate-spin" /> : <Upload size={15} />} Nhập {preview.length} dòng vào kho</PrimaryButton>
        </>
      )}
      {done !== null && <p className="text-xs text-emerald-600 mt-2">Đã nhập thành công {done} dòng.</p>}
    </Card>
  );
}

// Import Xuất kho NVL tự động từ báo cáo doanh thu chi tiết theo hoá đơn & món ăn —
// mỗi dòng: Ngày, Tên món, Số lượng bán (tuỳ chọn: Số hoá đơn). Hệ thống tự nổ theo công thức
// Cost món ăn để trừ đúng NVL tiêu hao. Ngày lấy trực tiếp từ cột "Ngày" của từng dòng trong
// file (không dùng 1 ngày chung cho cả file) — vì 1 file có thể gộp doanh thu nhiều ngày.
function XuatExcelImportForm({ data, onImport }) {
  const [fileName, setFileName] = useState("");
  const [preview, setPreview] = useState([]);
  const [unmatched, setUnmatched] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setDone(null);
    try {
      const rawRows = await readExcelRaw(file);
      const headerIdx = detectHeaderRow(rawRows, ["Tên món", "SL bán"]);
      if (headerIdx === -1) { setErrors(['Không tìm thấy dòng tiêu đề (cần có cột "Tên món" và "SL bán") trong file.']); setPreview([]); setUnmatched([]); return; }
      const rawObjRows = rowsToObjects(rawRows, headerIdx);
      const valid = [];
      const unmatchedNames = new Set();
      rawObjRows.forEach((row) => {
        const dishName = String(pickCol(row, "Tên món", "Ten mon", "Món ăn", "Mon an") ?? "").trim();
        const quantitySold = Number(pickCol(row, "SL bán", "SL ban", "Số lượng bán", "Số lượng", "So luong")) || 0;
        const unitPriceInFile = Number(pickCol(row, "Đơn giá", "Don gia")) || 0;
        // Doanh thu thực tế của dòng — ưu tiên lấy thẳng cột "Doanh thu" của file thay vì tự nhân
        // SL bán × Đơn giá, vì có trường hợp Đơn giá niêm yết không đổi nhưng thực thu khác
        // (khuyến mãi/tặng kèm/giảm giá riêng dòng) khiến SL × Đơn giá bị sai lệch so với thực tế.
        const revenueInFile = Number(pickCol(row, "Doanh thu", "Doanh thu bán hàng")) || 0;
        const invoiceNo = String(pickCol(row, "Số hóa đơn", "Số hoá đơn", "So hoa don", "Mã hoá đơn") ?? "").trim();
        const saleDate = excelDateToISO(pickCol(row, "Ngày", "Ngay") ?? pickColContains(row, "ngay")) || todayISO();
        if (!dishName || quantitySold <= 0) return;
        if (unitPriceInFile <= 0) return; // bỏ dòng rác của POS (VD "Mở két"...) — đơn giá 0đ không phải bán hàng thật
        const dish = data.dishes.find((d) => normalizeForMatch(d.name) === normalizeForMatch(dishName));
        if (!dish) { unmatchedNames.add(dishName); return; }
        valid.push({ dishName, dishId: dish.id, quantitySold, unitPriceInFile, revenueInFile, invoiceNo, saleDate });
      });
      setPreview(valid);
      setUnmatched(Array.from(unmatchedNames));
      setErrors([]);
    } catch (err) {
      setErrors([err.message || "Không đọc được file, kiểm tra lại định dạng .xlsx."]);
      setPreview([]); setUnmatched([]);
    }
  };

  const submit = async () => {
    if (preview.length === 0) return;
    setImporting(true);
    try {
      await onImport({ rows: preview });
      setDone(preview.length);
      setPreview([]); setUnmatched([]); setFileName("");
    } catch (e) {
      setErrors([e.message || "Không nhập được, vui lòng thử lại."]);
    } finally {
      setImporting(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Xuất kho NVL tự động từ báo cáo doanh thu</p>
      <p className="text-xs text-slate-500 mb-3">File cần có các cột: <b>Ngày</b>, <b>Tên món</b>, <b>SL bán</b>, <b>Đơn giá</b>, <b>Doanh thu</b> (tuỳ chọn: Số hoá đơn). Tên món phải khớp đúng tên đã tạo trong "Cost món ăn"; <b>doanh thu ghi nhận dùng đúng cột "Doanh thu" của file</b> (không tự nhân SL × Đơn giá, không dùng giá cấu hình sẵn trong Cost món ăn) — vì có trường hợp Đơn giá niêm yết không đổi nhưng thực thu khác do khuyến mãi/giảm giá/tặng kèm riêng từng dòng. App tự nhận diện đúng dòng tiêu đề dù file có vài dòng mô tả phía trên (kiểu file xuất từ phần mềm POS), tự bỏ qua các dòng tổng phụ theo hoá đơn và các dòng có Đơn giá = 0đ (dữ liệu rác của máy POS như "Mở két"...). Ngày xuất của từng dòng lấy trực tiếp từ cột "Ngày" trong file (không cần chọn ngày thủ công) — file gộp nhiều ngày vẫn tách đúng theo từng ngày.</p>

      <div className="flex items-center gap-2 mb-3">
        <label className="inline-flex items-center gap-2 cursor-pointer text-sm font-medium text-sky-700 border border-sky-300 bg-sky-50 hover:bg-sky-100 rounded-xl px-4 py-2">
          <Upload size={15} /> {fileName || "Chọn file Excel..."}
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
        </label>
        {fileName && (
          <button type="button" onClick={() => { setFileName(""); setPreview([]); setUnmatched([]); setErrors([]); }} className="text-slate-400 hover:text-rose-600 p-1.5" title="Bỏ file đã chọn">
            <X size={16} />
          </button>
        )}
      </div>

      {errors.length > 0 && (
        <div className="mb-3 text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-xl p-3 space-y-1">
          {errors.map((e, i) => <p key={i} className="flex items-center gap-1"><AlertTriangle size={12} className="shrink-0" /> {e}</p>)}
        </div>
      )}
      {unmatched.length > 0 && (
        <div className="mb-3 text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-xl p-3">
          <p className="font-medium mb-1 flex items-center gap-1"><AlertTriangle size={12} /> {unmatched.length} tên món không khớp với "Cost món ăn" (sẽ bị bỏ qua):</p>
          <p>{unmatched.join(", ")}</p>
        </div>
      )}

      {preview.length > 0 && (
        <>
          <div className="overflow-x-auto rounded-xl border border-slate-200 mb-3 max-h-72 overflow-y-auto">
            <table className="w-full text-sm min-w-[640px]">
              <thead className="sticky top-0 bg-slate-50">
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="px-2 py-2">Ngày</th><th className="px-2 py-2">Món</th><th className="px-2 py-2">Hoá đơn</th>
                  <th className="px-2 py-2 text-right">SL bán</th><th className="px-2 py-2 text-right">Đơn giá</th>
                  <th className="px-2 py-2 text-right">Doanh thu</th><th className="px-2 py-2 text-right">Giá vốn/suất</th><th className="px-2 py-2 text-right">Cost %</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((r, i) => {
                  const costPerUnit = dishTotalCost(r.dishId, data);
                  const totalCost = costPerUnit * r.quantitySold;
                  const costPct = r.revenueInFile > 0 ? (totalCost / r.revenueInFile) * 100 : 0;
                  return (
                    <tr key={i} className="border-b border-slate-100 last:border-0">
                      <td className="px-2 py-1.5 text-slate-400">{fmtDate(r.saleDate)}</td>
                      <td className="px-2 py-1.5">{r.dishName}</td>
                      <td className="px-2 py-1.5 text-slate-400">{r.invoiceNo || "—"}</td>
                      <td className="px-2 py-1.5 text-right">{r.quantitySold}</td>
                      <td className="px-2 py-1.5 text-right text-slate-400">{fmtMoney(r.unitPriceInFile)}</td>
                      <td className="px-2 py-1.5 text-right">{fmtMoney(r.revenueInFile)}</td>
                      <td className="px-2 py-1.5 text-right text-amber-700">{fmtMoney(totalCost)}</td>
                      <td className="px-2 py-1.5 text-right text-slate-500">{costPct.toFixed(1)}%</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <PrimaryButton onClick={submit} disabled={importing}>{importing ? <Loader2 size={15} className="animate-spin" /> : <Upload size={15} />} Xuất kho theo {preview.length} dòng món bán</PrimaryButton>
        </>
      )}
      {done !== null && <p className="text-xs text-emerald-600 mt-2">Đã xuất kho thành công theo {done} dòng món bán.</p>}
    </Card>
  );
}

// Thẻ tổng quan hiện ngay sau khi lưu 1 phiếu (nhập hoặc xuất) — tách biệt với
// phần "Lịch sử" chi tiết từng dòng.
function ReceiptSummaryCard({ summary, icon: Icon, actionLabel }) {
  if (!summary) return null;
  return (
    <Card className="p-4 sm:p-5 mb-5 border-sky-200 bg-sky-50/40">
      <div className="flex items-center gap-3">
        <div className="w-11 h-11 rounded-2xl bg-sky-100 text-sky-700 flex items-center justify-center shrink-0"><Icon size={20} /></div>
        <div className="min-w-0 flex-1">
          <p className="text-sm font-semibold text-sky-800">{actionLabel} thành công — Phiếu {summary.receiptCode}</p>
          <p className="text-xs text-sky-700/80 mt-0.5">
            {summary.lineCount} dòng
            {summary.supplierName ? ` · NCC: ${summary.supplierName}` : ""}
          </p>
        </div>
        <div className="text-right shrink-0">
          <p className="text-xs text-sky-700/80">Tổng giá trị phiếu</p>
          <p className="text-lg font-bold text-sky-800">{fmtMoney(summary.totalAmount)}</p>
        </div>
      </div>
    </Card>
  );
}

function NhapHangModule({ data, currentUser, onSubmit, onBulkImport }) {
  const [lastReceipt, setLastReceipt] = useState(null);

  const handleSubmit = async (payload) => {
    const summary = await onSubmit(payload);
    if (summary) setLastReceipt(summary);
  };
  const handleBulkImport = async (rows) => {
    const summary = await onBulkImport(rows);
    if (summary) setLastReceipt(summary);
  };

  return (
    <div>
      <SectionTitle icon={ArrowDownCircle} title="Nhập hàng" subtitle="Ghi nhận nhập hàng từ nhà cung cấp" />
      <ReceiptSummaryCard summary={lastReceipt} icon={ArrowDownCircle} actionLabel="Nhập hàng" />
      <NhapExcelImportForm data={data} onImport={handleBulkImport} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// BÁO CÁO NHẬP — Tác vụ 3: theo ngày, NCC, nhóm sản phẩm, tình trạng thanh toán
// ---------------------------------------------------------------------------
function BaoCaoNhapModule({ data }) {
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());
  const [supplierId, setSupplierId] = useState("");
  const [groupCode, setGroupCode] = useState("");
  const [paymentType, setPaymentType] = useState("");

  const filtered = data.importRecords.filter((r) => {
    if (r.importDate < from || r.importDate > to) return false;
    if (supplierId && r.supplierId !== supplierId) return false;
    if (paymentType && r.paymentType !== paymentType) return false;
    if (groupCode) {
      const p = data.products.find((x) => x.id === r.productId);
      if (!p || p.groupCode !== groupCode) return false;
    }
    return true;
  });

  const totalAmount = filtered.reduce((s, r) => s + r.totalAmount, 0);
  const totalByPayment = { tien_mat: 0, cong_no: 0, noi_bo: 0 };
  filtered.forEach((r) => { totalByPayment[r.paymentType] = (totalByPayment[r.paymentType] || 0) + r.totalAmount; });

  const bySupplier = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      const s = data.suppliers.find((x) => x.id === r.supplierId);
      const key = s?.name || "Không rõ";
      map.set(key, (map.get(key) || 0) + r.totalAmount);
    });
    return [...map.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, 10);
  }, [filtered]);

  // Báo cáo mua hàng theo Nhà cung cấp × Tình trạng thanh toán — mỗi NCC 1 dòng,
  // tách rõ đã trả (tiền mặt/nội bộ) và còn nợ (công nợ) để dễ đối chiếu công nợ NCC.
  const bySupplierPayment = useMemo(() => {
    const map = new Map();
    filtered.forEach((r) => {
      const s = data.suppliers.find((x) => x.id === r.supplierId);
      const key = s?.name || "Không rõ";
      const cur = map.get(key) || { name: key, tien_mat: 0, cong_no: 0, noi_bo: 0, total: 0 };
      cur[r.paymentType] = (cur[r.paymentType] || 0) + r.totalAmount;
      cur.total += r.totalAmount;
      map.set(key, cur);
    });
    return Array.from(map.values()).sort((a, b) => b.total - a.total);
  }, [filtered]);

  const groupCodes = [...new Set(data.products.map((p) => p.groupCode).filter(Boolean))];

  return (
    <div>
      <SectionTitle icon={BarChart3} title="Báo cáo nhập hàng" subtitle="Lọc theo ngày, nhà cung cấp, nhóm sản phẩm, tình trạng thanh toán" />
      <Card className="p-4 mb-5">
        <div className="grid sm:grid-cols-4 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
          <SelectField label="Nhà cung cấp" value={supplierId} onChange={(e) => setSupplierId(e.target.value)}>
            <option value="">Tất cả</option>
            {data.suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </SelectField>
          <SelectField label="Nhóm sản phẩm" value={groupCode} onChange={(e) => setGroupCode(e.target.value)}>
            <option value="">Tất cả</option>
            {groupCodes.map((g) => <option key={g} value={g}>{g}</option>)}
          </SelectField>
          <SelectField label="Tình trạng thanh toán" value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="sm:col-span-2">
            <option value="">Tất cả</option>
            <option value="tien_mat">Tiền mặt</option>
            <option value="cong_no">Công nợ</option>
            <option value="noi_bo">Nội bộ</option>
          </SelectField>
        </div>
      </Card>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-5">
        <MetricCard label="Tổng giá trị nhập" value={fmtMoney(totalAmount)} icon={ArrowDownCircle} accent="teal" />
        <MetricCard label="Tiền mặt" value={fmtMoney(totalByPayment.tien_mat)} icon={Wallet} accent="emerald" />
        <MetricCard label="Công nợ" value={fmtMoney(totalByPayment.cong_no)} icon={Receipt} accent="amber" />
        <MetricCard label="Nội bộ" value={fmtMoney(totalByPayment.noi_bo)} icon={Boxes} accent="indigo" />
      </div>
      <Card className="p-4 sm:p-5 mb-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Top nhà cung cấp theo giá trị nhập</p>
        {bySupplier.length === 0 ? <EmptyState icon={BarChart3} text="Chưa có dữ liệu." /> : (
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <BarChart data={bySupplier} layout="vertical" margin={{ left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                <XAxis type="number" tickFormatter={(v) => fmtNumber(v)} tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="name" width={140} tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v) => fmtMoney(v)} />
                <Bar dataKey="value" fill="#0369a1" radius={[0, 6, 6, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </Card>
      <Card className="p-0 overflow-hidden mb-5">
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Báo cáo mua hàng theo Nhà cung cấp &amp; Tình trạng thanh toán</p></div>
        {bySupplierPayment.length === 0 ? <EmptyState icon={BarChart3} text="Chưa có dữ liệu." /> : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Nhà cung cấp</th>
                <th className="px-3 py-2 text-right">Tiền mặt</th>
                <th className="px-3 py-2 text-right">Công nợ</th>
                <th className="px-3 py-2 text-right">Nội bộ</th>
                <th className="px-3 py-2 text-right">Tổng cộng</th>
              </tr></thead>
              <tbody>
                {bySupplierPayment.map((r) => (
                  <tr key={r.name} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2 font-medium text-slate-700">{r.name}</td>
                    <td className="px-3 py-2 text-right text-emerald-700">{r.tien_mat ? fmtMoney(r.tien_mat) : "—"}</td>
                    <td className="px-3 py-2 text-right text-amber-700">{r.cong_no ? fmtMoney(r.cong_no) : "—"}</td>
                    <td className="px-3 py-2 text-right text-indigo-700">{r.noi_bo ? fmtMoney(r.noi_bo) : "—"}</td>
                    <td className="px-3 py-2 text-right font-semibold text-slate-800">{fmtMoney(r.total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Chi tiết ({filtered.length} dòng)</p></div>
        {filtered.length === 0 ? <EmptyState icon={Inbox} text="Không có dữ liệu phù hợp bộ lọc." /> : (
          <div className="overflow-x-auto max-h-96 overflow-y-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-white"><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Ngày</th><th className="px-3 py-2">Mã phiếu</th><th className="px-3 py-2">NCC</th>
                <th className="px-3 py-2">Sản phẩm</th><th className="px-3 py-2 text-right">SL</th><th className="px-3 py-2 text-right">Thành tiền</th><th className="px-3 py-2">TT</th>
              </tr></thead>
              <tbody>
                {filtered.map((r) => {
                  const s = data.suppliers.find((x) => x.id === r.supplierId);
                  const p = data.products.find((x) => x.id === r.productId);
                  return (
                    <tr key={r.id} className="border-b border-slate-50 last:border-0">
                      <td className="px-3 py-2 text-slate-500">{fmtDate(r.importDate)}</td>
                      <td className="px-3 py-2 text-slate-500">{r.receiptCode}</td>
                      <td className="px-3 py-2">{s?.name}</td>
                      <td className="px-3 py-2">{p?.name}</td>
                      <td className="px-3 py-2 text-right">{fmtNumber(r.quantity)} {p?.unit}</td>
                      <td className="px-3 py-2 text-right font-medium">{fmtMoney(r.totalAmount)}</td>
                      <td className="px-3 py-2"><Badge className={PAYMENT_TYPE_META[r.paymentType]?.color}>{PAYMENT_TYPE_META[r.paymentType]?.label}</Badge></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// XUẤT HÀNG — Tác vụ 4: 1 phiếu xuất gồm nhiều dòng NVL (tiêu hao) và/hoặc
// TP (bán ra), dùng chung Đơn số / Mã doanh thu / Mã xuất / Ngày.
// ---------------------------------------------------------------------------
function XuatHangForm({ data, currentUser, onSubmit }) {
  const [orderNumber, setOrderNumber] = useState("");
  const [revenueCodeId, setRevenueCodeId] = useState(data.revenueCodes[0]?.id || "");
  const [exportCodeId, setExportCodeId] = useState(data.exportCodes[0]?.id || "");
  const [lines, setLines] = useState([]);
  const [lineType, setLineType] = useState("NL");
  const [productId, setProductId] = useState("");
  const [quantity, setQuantity] = useState("");
  const [unitPrice, setUnitPrice] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const productPool = data.products.filter((p) => p.classification === lineType);
  const selectedProduct = data.products.find((p) => p.id === productId);
  const avgPrice = selectedProduct && lineType === "NL" ? computeAvgPrice(selectedProduct.id, data) : 0;
  const currentStock = selectedProduct ? stockAsOf(selectedProduct.id, todayISO(), data) : 0;

  const addLine = () => {
    if (!productId || !quantity) { setError("Chọn sản phẩm và nhập số lượng trước khi thêm dòng."); return; }
    const price = lineType === "NL" ? avgPrice : Number(unitPrice) || 0;
    if (lineType === "TP" && !unitPrice) { setError("Dòng thành phẩm cần nhập đơn giá bán."); return; }
    setError("");
    setLines((prev) => [...prev, {
      key: Math.random().toString(36).slice(2),
      lineType, productId, quantity: Number(quantity), unitPrice: price,
      totalAmount: Number(quantity) * price,
    }]);
    setProductId(""); setQuantity(""); setUnitPrice("");
  };

  const removeLine = (key) => setLines((prev) => prev.filter((l) => l.key !== key));

  const totalNVL = lines.filter((l) => l.lineType === "NL").reduce((s, l) => s + l.totalAmount, 0);
  const totalTP = lines.filter((l) => l.lineType === "TP").reduce((s, l) => s + l.totalAmount, 0);

  const submit = async () => {
    if (lines.length === 0) { setError("Cần thêm ít nhất 1 dòng NVL hoặc TP."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ orderNumber: orderNumber.trim(), revenueCodeId: revenueCodeId || null, exportCodeId: exportCodeId || null, lines });
      setOrderNumber(""); setLines([]);
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Xuất hàng mới</p>
      <div className="grid sm:grid-cols-3 gap-3 mb-4">
        <TextField label="Đơn số (tự đặt, không bắt buộc)" value={orderNumber} onChange={(e) => setOrderNumber(e.target.value)} />
        <SelectField label="Mã doanh thu" value={revenueCodeId} onChange={(e) => setRevenueCodeId(e.target.value)}>
          <option value="">— Không chọn —</option>
          {data.revenueCodes.map((r) => <option key={r.id} value={r.id}>{r.code} — {r.name}</option>)}
        </SelectField>
        <SelectField label="Mã xuất" value={exportCodeId} onChange={(e) => setExportCodeId(e.target.value)}>
          <option value="">— Không chọn —</option>
          {data.exportCodes.map((r) => <option key={r.id} value={r.id}>{r.code} — {r.name}</option>)}
        </SelectField>
      </div>

      <div className="bg-slate-50 rounded-xl p-3 mb-4">
        <p className="text-xs font-medium text-slate-600 mb-2">Thêm dòng nguyên liệu / thành phẩm</p>
        <div className="grid sm:grid-cols-3 gap-2 mb-2">
          <SelectField value={lineType} onChange={(e) => { setLineType(e.target.value); setProductId(""); }}>
            <option value="NL">Nguyên vật liệu (NVL)</option>
            <option value="TP">Thành phẩm (TP)</option>
          </SelectField>
          <div className="sm:col-span-2">
            <ProductCodeNameFields
              products={productPool}
              productId={productId}
              onSelectProduct={(id) => setProductId(id)}
              codeLabel={lineType === "NL" ? "Mã NVL" : "Mã TP"}
              nameLabel={lineType === "NL" ? "Tên NVL" : "Tên TP"}
            />
          </div>
        </div>
        {selectedProduct && (
          <p className="text-xs text-slate-500 mb-2">
            Tồn hiện tại: <span className="font-medium">{fmtNumber(currentStock)} {selectedProduct.unit}</span>
            {lineType === "NL" && <> · Giá bình quân gia quyền: <span className="font-medium">{fmtMoney(avgPrice)}</span>/{selectedProduct.unit}</>}
          </p>
        )}
        <div className="grid sm:grid-cols-3 gap-2">
          <TextField placeholder={`Số lượng${selectedProduct ? ` (${selectedProduct.unit})` : ""}`} type="number" value={quantity} onChange={(e) => setQuantity(e.target.value)} />
          {lineType === "TP" ? (
            <TextField placeholder="Đơn giá bán (đ)" type="number" value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} />
          ) : (
            <div className="px-3 py-2 rounded-xl border border-dashed border-slate-300 text-xs text-slate-400 flex items-center">Đơn giá tự tính (bình quân gia quyền)</div>
          )}
          <GhostButton type="button" onClick={addLine}><Plus size={14} /> Thêm dòng</GhostButton>
        </div>
      </div>

      {lines.length > 0 && (
        <div className="mb-4 space-y-1.5">
          {lines.map((l) => {
            const p = data.products.find((x) => x.id === l.productId);
            return (
              <div key={l.key} className="flex items-center justify-between text-sm bg-white border border-slate-200 rounded-lg px-3 py-2">
                <div className="flex items-center gap-2 min-w-0">
                  <Badge className={CLASSIFICATION_META[l.lineType]?.color}>{l.lineType}</Badge>
                  <span className="truncate">{p?.name} — {fmtNumber(l.quantity)} {p?.unit} × {fmtMoney(l.unitPrice)}</span>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <span className="font-medium">{fmtMoney(l.totalAmount)}</span>
                  <button onClick={() => removeLine(l.key)} className="text-slate-400 hover:text-rose-600"><X size={14} /></button>
                </div>
              </div>
            );
          })}
          <div className="flex items-center justify-between text-xs text-slate-500 pt-1">
            <span>Tổng giá vốn NVL: <span className="font-medium text-slate-700">{fmtMoney(totalNVL)}</span></span>
            <span>Tổng doanh thu TP: <span className="font-medium text-slate-700">{fmtMoney(totalTP)}</span></span>
          </div>
        </div>
      )}

      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <ArrowUpCircle size={15} />} Lưu phiếu xuất</PrimaryButton>
    </Card>
  );
}

function XuatHangList({ data, onDelete }) {
  const rows = data.exportRecords.slice(0, 40);
  const [deleting, setDeleting] = useState(null);
  const handleDeleteRow = async (id) => {
    if (!window.confirm("Xoá dòng xuất hàng này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };
  return (
    <Card className="p-0 overflow-hidden">
      <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Lịch sử xuất hàng gần đây</p></div>
      {rows.length === 0 ? <EmptyState icon={Inbox} text="Chưa có phiếu xuất nào." /> : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Mã phiếu</th><th className="px-3 py-2">Ngày</th><th className="px-3 py-2">Loại</th>
              <th className="px-3 py-2">Sản phẩm</th><th className="px-3 py-2 text-right">SL</th>
              <th className="px-3 py-2 text-right">Đơn giá</th><th className="px-3 py-2 text-right">Thành tiền</th>
              <th className="px-3 py-2 w-10"></th>
            </tr></thead>
            <tbody>
              {rows.map((r) => {
                const p = data.products.find((x) => x.id === r.productId);
                return (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2 text-slate-500">{r.receiptCode}</td>
                    <td className="px-3 py-2 text-slate-500">{fmtDate(r.exportDate)}</td>
                    <td className="px-3 py-2"><Badge className={CLASSIFICATION_META[r.lineType]?.color}>{r.lineType}</Badge></td>
                    <td className="px-3 py-2">{p?.name || "—"}</td>
                    <td className="px-3 py-2 text-right">{fmtNumber(r.quantity)} {p?.unit}</td>
                    <td className="px-3 py-2 text-right text-slate-500">{fmtMoney(r.unitPrice)}</td>
                    <td className="px-3 py-2 text-right font-medium">{fmtMoney(r.totalAmount)}</td>
                    <td className="px-3 py-2 text-center">
                      <button type="button" onClick={() => handleDeleteRow(r.id)} disabled={deleting === r.id} className="text-slate-400 hover:text-rose-600 p-1">
                        {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

function XuatHangModule({ data, currentUser, onSubmit, onBulkImportFromBills, onDelete, onDeleteMany }) {
  const [lastReceipt, setLastReceipt] = useState(null);
  const [resetKey, setResetKey] = useState(0);

  const handleSubmit = async (payload) => {
    const summary = await onSubmit(payload);
    if (summary) setLastReceipt(summary);
  };
  const handleBulkImport = async (payload) => {
    const result = await onBulkImportFromBills(payload);
    if (result) setLastReceipt(result);
    return result;
  };
  // Sau khi xoá phiếu, làm mới toàn bộ màn hình về trạng thái sạch như trước khi import —
  // xoá luôn thẻ "Tổng quan phiếu vừa xuất" và dòng thông báo "Đã xuất kho thành công..." còn đọng lại.
  const handleDeleteMany = async (ids) => {
    await onDeleteMany(ids);
    setLastReceipt(null);
    setResetKey((k) => k + 1);
  };

  return (
    <div>
      <SectionTitle icon={ArrowUpCircle} title="Xuất hàng" subtitle="Ghi nhận tiêu hao nguyên liệu và bán thành phẩm" />
      <ReceiptSummaryCard summary={lastReceipt} icon={ArrowUpCircle} actionLabel="Xuất hàng" />
      <DeleteByReceiptCard records={data.exportRecords} onDeleteMany={handleDeleteMany} label="phiếu xuất" />
      <XuatExcelImportForm key={resetKey} data={data} onImport={handleBulkImport} />
      <XuatHangList data={data} onDelete={onDelete} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// BÁO CÁO XUẤT — Tác vụ 5: theo loại hình doanh thu & theo nguồn (mã xuất)
// Giá vốn = tổng Thành tiền các dòng NVL cùng mã; Doanh thu = tổng Thành tiền
// các dòng TP cùng mã. Lợi nhuận = Doanh thu - Giá vốn.
// ---------------------------------------------------------------------------
function profitReportBy(records, codeList, codeField, from, to) {
  const filtered = records.filter((r) => r.exportDate >= from && r.exportDate <= to);
  const totalRevenueAll = filtered.filter((r) => r.lineType === "TP").reduce((s, r) => s + r.totalAmount, 0);
  const rows = codeList.map((c) => {
    const lines = filtered.filter((r) => r[codeField] === c.id);
    const giaVon = lines.filter((r) => r.lineType === "NL").reduce((s, r) => s + r.totalAmount, 0);
    const doanhThu = lines.filter((r) => r.lineType === "TP").reduce((s, r) => s + r.totalAmount, 0);
    const loiNhuan = doanhThu - giaVon;
    return {
      name: c.name, giaVon, doanhThu, loiNhuan,
      tiTrong: totalRevenueAll > 0 ? (doanhThu / totalRevenueAll) * 100 : 0,
      tiSuat: doanhThu > 0 ? (loiNhuan / doanhThu) * 100 : 0,
    };
  }).filter((r) => r.giaVon > 0 || r.doanhThu > 0);
  const totals = rows.reduce((acc, r) => ({ giaVon: acc.giaVon + r.giaVon, doanhThu: acc.doanhThu + r.doanhThu, loiNhuan: acc.loiNhuan + r.loiNhuan }), { giaVon: 0, doanhThu: 0, loiNhuan: 0 });
  return { rows: rows.sort((a, b) => b.doanhThu - a.doanhThu), totals };
}

function ProfitTable({ title, rows, totals }) {
  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">{title}</p>
      {rows.length === 0 ? <EmptyState icon={BarChart3} text="Chưa có dữ liệu trong khoảng thời gian này." /> : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Tên</th><th className="px-3 py-2 text-right">Giá vốn</th><th className="px-3 py-2 text-right">Doanh thu</th>
              <th className="px-3 py-2 text-right">Tỉ trọng DT</th><th className="px-3 py-2 text-right">Lợi nhuận</th><th className="px-3 py-2 text-right">Tỉ suất LN</th>
            </tr></thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i} className="border-b border-slate-50 last:border-0">
                  <td className="px-3 py-2 font-medium text-slate-700">{r.name}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{fmtMoney(r.giaVon)}</td>
                  <td className="px-3 py-2 text-right">{fmtMoney(r.doanhThu)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{r.tiTrong.toFixed(1)}%</td>
                  <td className={`px-3 py-2 text-right font-medium ${r.loiNhuan >= 0 ? "text-emerald-700" : "text-rose-600"}`}>{fmtMoney(r.loiNhuan)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{r.tiSuat.toFixed(1)}%</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-200 font-semibold text-slate-800">
                <td className="px-3 py-2">TỔNG</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totals.giaVon)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totals.doanhThu)}</td>
                <td className="px-3 py-2 text-right">—</td>
                <td className={`px-3 py-2 text-right ${totals.loiNhuan >= 0 ? "text-emerald-700" : "text-rose-600"}`}>{fmtMoney(totals.loiNhuan)}</td>
                <td className="px-3 py-2 text-right">{totals.doanhThu > 0 ? ((totals.loiNhuan / totals.doanhThu) * 100).toFixed(1) : "0.0"}%</td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </Card>
  );
}

// Gộp dish_sales theo ngày để ra 5 chỉ số: số hoá đơn, doanh số, giá vốn, số món bán, tỉ lệ cost.
function dailySalesReport(dishSales, from, to) {
  const filtered = dishSales.filter((s) => s.saleDate >= from && s.saleDate <= to);
  const byDate = new Map();
  filtered.forEach((s) => {
    const cur = byDate.get(s.saleDate) || { date: s.saleDate, invoices: new Set(), revenue: 0, cost: 0, qty: 0 };
    if (s.invoiceNo) cur.invoices.add(s.invoiceNo);
    cur.revenue += s.totalAmount;
    cur.cost += s.costAmount;
    cur.qty += s.quantity;
    byDate.set(s.saleDate, cur);
  });
  const rows = Array.from(byDate.values()).map((r) => ({
    date: r.date, invoiceCount: r.invoices.size, revenue: r.revenue, cost: r.cost, qty: r.qty,
    costRatio: r.revenue > 0 ? (r.cost / r.revenue) * 100 : 0,
  })).sort((a, b) => b.date.localeCompare(a.date));
  const totals = rows.reduce((acc, r) => ({
    invoiceCount: acc.invoiceCount + r.invoiceCount, revenue: acc.revenue + r.revenue,
    cost: acc.cost + r.cost, qty: acc.qty + r.qty,
  }), { invoiceCount: 0, revenue: 0, cost: 0, qty: 0 });
  totals.costRatio = totals.revenue > 0 ? (totals.cost / totals.revenue) * 100 : 0;
  return { rows, totals };
}

function BaoCaoDoanhThuNgayModule({ data }) {
  const [from, setFrom] = useState(daysAgoISO(7));
  const [to, setTo] = useState(todayISO());
  const report = dailySalesReport(data.dishSales, from, to);

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Báo cáo doanh thu theo ngày</p>
      <p className="text-xs text-slate-500 mb-3">Số liệu lấy từ các lần "Xuất kho tự động từ báo cáo doanh thu" — chỉ có dữ liệu từ thời điểm tính năng này được bật.</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
        <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-5 gap-2 mb-4">
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Số hoá đơn</p>
          <p className="text-sm font-semibold text-slate-700">{fmtNumber(report.totals.invoiceCount)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Số món bán</p>
          <p className="text-sm font-semibold text-slate-700">{fmtNumber(report.totals.qty)}</p>
        </div>
        <div className="bg-sky-50 rounded-xl px-3 py-2">
          <p className="text-xs text-sky-700">Tổng doanh số</p>
          <p className="text-sm font-semibold text-sky-800">{fmtMoney(report.totals.revenue)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Tổng giá cost</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(report.totals.cost)}</p>
        </div>
        <div className={`rounded-xl px-3 py-2 ${report.totals.costRatio <= 35 ? "bg-emerald-50" : report.totals.costRatio <= 45 ? "bg-amber-50" : "bg-rose-50"}`}>
          <p className={`text-xs ${report.totals.costRatio <= 35 ? "text-emerald-700" : report.totals.costRatio <= 45 ? "text-amber-700" : "text-rose-700"}`}>Tỉ lệ cost bình quân</p>
          <p className={`text-sm font-semibold ${report.totals.costRatio <= 35 ? "text-emerald-800" : report.totals.costRatio <= 45 ? "text-amber-800" : "text-rose-800"}`}>{report.totals.costRatio.toFixed(1)}%</p>
        </div>
      </div>

      {report.rows.length === 0 ? (
        <EmptyState icon={BarChart3} text="Chưa có dữ liệu bán hàng trong khoảng thời gian này." />
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Ngày</th><th className="px-3 py-2 text-right">Số hoá đơn</th>
              <th className="px-3 py-2 text-right">Số món bán</th><th className="px-3 py-2 text-right">Doanh số</th>
              <th className="px-3 py-2 text-right">Giá cost</th><th className="px-3 py-2 text-right">Tỉ lệ cost</th>
            </tr></thead>
            <tbody>
              {report.rows.map((r) => (
                <tr key={r.date} className="border-b border-slate-50 last:border-0">
                  <td className="px-3 py-2 font-medium text-slate-700">{fmtDate(r.date)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{fmtNumber(r.invoiceCount)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{fmtNumber(r.qty)}</td>
                  <td className="px-3 py-2 text-right">{fmtMoney(r.revenue)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{fmtMoney(r.cost)}</td>
                  <td className="px-3 py-2 text-right text-slate-500">{r.costRatio.toFixed(1)}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

// Tổng chi phí vận hành theo khoảng ngày, gộp theo từng nhóm — KHÔNG gồm nhóm "nvl"
// (Chi phí nguyên vật liệu) vì đã được tính trong Giá vốn (dish_sales.costAmount) rồi,
// tính thêm vào đây sẽ bị trùng lặp.
function operatingExpenseReport(expenseRecords, from, to) {
  const filtered = expenseRecords.filter((r) => r.category !== "nvl" && r.expenseDate >= from && r.expenseDate <= to);
  const byCategory = {};
  filtered.forEach((r) => { byCategory[r.category] = (byCategory[r.category] || 0) + r.amount; });
  const total = filtered.reduce((s, r) => s + r.amount, 0);
  return { total, byCategory };
}

// P&L tổng hợp: Doanh thu & Giá vốn lấy từ dish_sales (đã snapshot đúng giá bán/giá vốn
// tại thời điểm bán), trừ tiếp Chi phí vận hành (expense_records, trừ nhóm nvl) để ra
// Lợi nhuận ròng — khác với "Lợi nhuận gộp" (chỉ mới trừ giá vốn nguyên liệu).
function PLTongHopModule({ data }) {
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());

  const sales = data.dishSales.filter((s) => s.saleDate >= from && s.saleDate <= to);
  const revenue = sales.reduce((s, r) => s + r.totalAmount, 0);
  const cogs = sales.reduce((s, r) => s + r.costAmount, 0);
  const grossProfit = revenue - cogs;
  const opex = operatingExpenseReport(data.expenseRecords, from, to);
  const netProfit = grossProfit - opex.total;
  const netMargin = revenue > 0 ? (netProfit / revenue) * 100 : 0;

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Lợi nhuận ròng (P&L tổng hợp)</p>
      <p className="text-xs text-slate-500 mb-3">
        Doanh thu & giá vốn lấy từ dữ liệu bán món (mục "Báo cáo doanh thu theo ngày" ở trên) — trừ tiếp toàn bộ
        Chi phí vận hành (tab "Chi phí", không gồm nhóm Nguyên vật liệu vì đã nằm trong giá vốn) để ra Lợi nhuận ròng thực tế.
      </p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
        <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-3 gap-2 mb-4">
        <div className="bg-sky-50 rounded-xl px-3 py-2">
          <p className="text-xs text-sky-700">Doanh thu</p>
          <p className="text-sm font-semibold text-sky-800">{fmtMoney(revenue)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Giá vốn</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(cogs)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Lợi nhuận gộp</p>
          <p className="text-sm font-semibold text-slate-700">{fmtMoney(grossProfit)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl px-3 py-2">
          <p className="text-xs text-slate-500">Chi phí vận hành</p>
          <p className="text-sm font-semibold text-slate-700">{fmtMoney(opex.total)}</p>
        </div>
        <div className={`rounded-xl px-3 py-2 ${netProfit >= 0 ? "bg-emerald-50" : "bg-rose-50"}`}>
          <p className={`text-xs ${netProfit >= 0 ? "text-emerald-700" : "text-rose-700"}`}>Lợi nhuận ròng</p>
          <p className={`text-sm font-semibold ${netProfit >= 0 ? "text-emerald-800" : "text-rose-800"}`}>{fmtMoney(netProfit)}</p>
        </div>
        <div className={`rounded-xl px-3 py-2 ${netProfit >= 0 ? "bg-emerald-50" : "bg-rose-50"}`}>
          <p className={`text-xs ${netProfit >= 0 ? "text-emerald-700" : "text-rose-700"}`}>Tỉ suất LN ròng</p>
          <p className={`text-sm font-semibold ${netProfit >= 0 ? "text-emerald-800" : "text-rose-800"}`}>{netMargin.toFixed(1)}%</p>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
            <th className="px-3 py-2">Nhóm chi phí vận hành</th><th className="px-3 py-2 text-right">Số tiền</th>
          </tr></thead>
          <tbody>
            {EXPENSE_CATEGORIES.filter((c) => c.key !== "nvl").map((c) => (
              <tr key={c.key} className="border-b border-slate-50 last:border-0">
                <td className="px-3 py-2 text-slate-700">{c.label}</td>
                <td className="px-3 py-2 text-right text-slate-600">{fmtMoney(opex.byCategory[c.key] || 0)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function BaoCaoXuatModule({ data }) {
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());

  const byRevenue = profitReportBy(data.exportRecords, data.revenueCodes, "revenueCodeId", from, to);
  const byExportCode = profitReportBy(data.exportRecords, data.exportCodes, "exportCodeId", from, to);

  return (
    <div>
      <SectionTitle icon={BarChart3} title="Báo cáo xuất hàng" subtitle="Doanh thu, giá vốn, lợi nhuận theo loại hình & nguồn doanh thu" />
      <BaoCaoDoanhThuNgayModule data={data} />
      <PLTongHopModule data={data} />
      <Card className="p-4 mb-5">
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
      </Card>
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3 mb-5">
        <MetricCard label="Tổng doanh thu" value={fmtMoney(byRevenue.totals.doanhThu)} icon={TrendingUp} accent="teal" />
        <MetricCard label="Tổng giá vốn" value={fmtMoney(byRevenue.totals.giaVon)} icon={Receipt} accent="amber" />
        <MetricCard label="Tổng lợi nhuận" value={fmtMoney(byRevenue.totals.loiNhuan)} icon={Wallet} accent="emerald" />
      </div>
      <ProfitTable title="I. Báo cáo theo loại hình doanh thu (Mã doanh thu)" rows={byRevenue.rows} totals={byRevenue.totals} />
      <ProfitTable title="II. Báo cáo theo nguồn doanh thu (Mã xuất)" rows={byExportCode.rows} totals={byExportCode.totals} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// TỒN KHO — Tác vụ 6: Nhập - Xuất - Tồn theo mã SP/NVL/TP + chỉnh tồn đầu
// ---------------------------------------------------------------------------
function StockOpeningForm({ data, currentUser, onSubmit }) {
  const [productId, setProductId] = useState("");
  const [asOfDate, setAsOfDate] = useState(todayISO());
  const [quantity, setQuantity] = useState("");
  const [unitPrice, setUnitPrice] = useState("");
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const selectedProduct = data.products.find((p) => p.id === productId);

  const submit = async () => {
    if (!productId || !asOfDate || quantity === "") { setError("Vui lòng chọn sản phẩm, ngày chốt và số lượng."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ productId, asOfDate, quantity: Number(quantity), unitPrice: Number(unitPrice) || 0, note: note.trim() });
      setProductId(""); setQuantity(""); setUnitPrice(""); setNote("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Chốt / điều chỉnh tồn đầu kỳ</p>
      <p className="text-xs text-slate-500 mb-3">Dùng khi kiểm kho thực tế — mốc này sẽ là gốc để tính tồn kho & giá bình quân gia quyền cho các lần nhập/xuất sau đó.</p>
      <div className="mb-3">
        <ProductCodeNameFields products={data.products} productId={productId} onSelectProduct={(id) => setProductId(id)} codeLabel="Mã sản phẩm" nameLabel="Tên sản phẩm" />
      </div>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <TextField label="Ngày chốt" type="date" value={asOfDate} onChange={(e) => setAsOfDate(e.target.value)} />
        <MoneyField label={`Số lượng thực tế${selectedProduct ? ` (${selectedProduct.unit})` : ""}`} allowDecimal value={quantity} onChange={setQuantity} />
        <MoneyField label="Đơn giá (đ)" value={unitPrice} onChange={setUnitPrice} />
        <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} className="sm:col-span-2" />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <CheckCircle2 size={15} />} Lưu mốc tồn đầu</PrimaryButton>
    </Card>
  );
}

function TonKhoModule({ data, currentUser, onSaveOpening }) {
  const [from, setFrom] = useState(daysAgoISO(30));
  const [to, setTo] = useState(todayISO());
  const [classification, setClassification] = useState("all");
  const [q, setQ] = useState("");
  const isQuanLy = currentUser.role === "quan_ly";

  let products = classification === "all" ? data.products : data.products.filter((p) => p.classification === classification);
  if (q) products = products.filter((p) => stripDiacritics(p.name).includes(stripDiacritics(q)) || p.code.includes(q));

  const rows = products.map((p) => ({ product: p, ...nktForProduct(p.id, from, to, data) }));
  const totals = rows.reduce((acc, r) => ({
    openingValue: acc.openingValue + r.openingValue, importValue: acc.importValue + r.importValue,
    exportValue: acc.exportValue + r.exportValue, closingValue: acc.closingValue + r.closingValue,
  }), { openingValue: 0, importValue: 0, exportValue: 0, closingValue: 0 });

  return (
    <div>
      <SectionTitle icon={Warehouse} title="Tồn kho — Nhập - Xuất - Tồn" subtitle="Theo dõi tồn đầu, nhập, xuất, tồn cuối theo từng sản phẩm" />
      {isQuanLy && <StockOpeningForm data={data} currentUser={currentUser} onSubmit={onSaveOpening} />}
      <Card className="p-4 mb-5">
        <div className="grid sm:grid-cols-4 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
          <SelectField label="Phân loại" value={classification} onChange={(e) => setClassification(e.target.value)}>
            <option value="all">Tất cả</option>
            <option value="NL">Nguyên vật liệu</option>
            <option value="TP">Thành phẩm</option>
          </SelectField>
          <TextField label="Tìm sản phẩm" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Mã hoặc tên..." />
        </div>
      </Card>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-5">
        <MetricCard label="Giá trị tồn đầu" value={fmtMoney(totals.openingValue)} icon={Boxes} accent="indigo" />
        <MetricCard label="Giá trị nhập" value={fmtMoney(totals.importValue)} icon={ArrowDownCircle} accent="teal" />
        <MetricCard label="Giá trị xuất" value={fmtMoney(totals.exportValue)} icon={ArrowUpCircle} accent="amber" />
        <MetricCard label="Giá trị tồn cuối" value={fmtMoney(totals.closingValue)} icon={Warehouse} accent="emerald" />
      </div>
      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Chi tiết theo sản phẩm ({rows.length})</p></div>
        {rows.length === 0 ? <EmptyState icon={Package} text="Chưa có sản phẩm nào." /> : (
          <div className="overflow-x-auto max-h-[32rem] overflow-y-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-white"><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Sản phẩm</th><th className="px-3 py-2">Loại</th>
                <th className="px-3 py-2 text-right">Tồn đầu</th><th className="px-3 py-2 text-right">Nhập</th>
                <th className="px-3 py-2 text-right">Xuất</th><th className="px-3 py-2 text-right">Tồn cuối</th>
                <th className="px-3 py-2 text-right">Giá trị tồn cuối</th>
              </tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.product.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2">
                      <p className="font-medium text-slate-700">{r.product.name}</p>
                      <p className="text-xs text-slate-400">{r.product.code}</p>
                    </td>
                    <td className="px-3 py-2"><Badge className={CLASSIFICATION_META[r.product.classification]?.color}>{r.product.classification}</Badge></td>
                    <td className="px-3 py-2 text-right">{fmtNumber(r.openingQty)} {r.product.unit}</td>
                    <td className="px-3 py-2 text-right text-sky-700">+{fmtNumber(r.importQty)}</td>
                    <td className="px-3 py-2 text-right text-rose-600">-{fmtNumber(r.exportQty)}</td>
                    <td className="px-3 py-2 text-right font-medium">{fmtNumber(r.closingQty)} {r.product.unit}</td>
                    <td className="px-3 py-2 text-right font-medium">{fmtMoney(r.closingValue)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      <p className="text-xs text-slate-400 mt-3">* Giá trị tính theo giá bình quân gia quyền hiện tại của từng sản phẩm (tồn đầu kỳ gần nhất + toàn bộ nhập kể từ mốc đó).</p>
    </div>
  );
}

// ---------------------------------------------------------------------------
// TÀI KHOẢN (Quản lý) — tra cứu + đặt lại mật khẩu
// ---------------------------------------------------------------------------
function ResetPasswordModal({ currentUser, employee, onClose, onSuccess }) {
  const [adminPassword, setAdminPassword] = useState("");
  const [newPassword, setNewPassword] = useState("123456");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async () => {
    if (!adminPassword) { setError("Vui lòng nhập mật khẩu của chính bạn để xác nhận."); return; }
    if (!newPassword || newPassword.length < 6) { setError("Mật khẩu mới cần ít nhất 6 ký tự."); return; }
    setError(""); setSaving(true);
    try {
      const { data, error: qErr } = await supabase.rpc("admin_reset_password", {
        p_admin_id: currentUser.id, p_admin_password: adminPassword,
        p_target_employee_id: employee.id, p_new_password: newPassword,
      });
      if (qErr) throw qErr;
      if (!data) { setError("Mật khẩu của bạn không đúng."); return; }
      onSuccess();
    } catch (e) {
      console.error(e);
      setError("Không đặt lại được mật khẩu, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
        <div className="flex items-center justify-between mb-1">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><Lock size={17} className="text-sky-700" /> Đặt lại mật khẩu</p>
          <button onClick={onClose}><X size={18} className="text-slate-400" /></button>
        </div>
        <p className="text-sm text-slate-500 mb-4">Cho tài khoản: <span className="font-medium text-slate-700">{employee.name}</span> ({employee.username})</p>
        <div className="space-y-3">
          <TextField label="Mật khẩu tạm thời mới" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} />
          <p className="text-[11px] text-slate-400 -mt-2">Nhân viên sẽ bị bắt buộc đổi mật khẩu này trong 24h ở lần đăng nhập tiếp theo.</p>
          <TextField label="Xác nhận: mật khẩu của chính bạn" type="password" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} />
          {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertTriangle size={14} /> {error}</p>}
          <div className="flex gap-2">
            <PrimaryButton type="button" onClick={submit} disabled={saving}>{saving ? "Đang lưu..." : "Đặt lại mật khẩu"}</PrimaryButton>
            <GhostButton type="button" onClick={onClose}>Hủy</GhostButton>
          </div>
        </div>
      </div>
    </div>
  );
}

function AddEmployeeForm({ onAdd }) {
  const [username, setUsername] = useState("");
  const [name, setName] = useState("");
  const [role, setRole] = useState("nhan_vien_kho");
  const [adminPassword, setAdminPassword] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!username.trim() || !name.trim()) { setError("Vui lòng nhập đủ tên đăng nhập và họ tên."); return; }
    if (!adminPassword) { setError("Vui lòng nhập mật khẩu của chính bạn để xác nhận."); return; }
    setError(""); setSaving(true);
    try {
      await onAdd({ username: username.trim(), name: name.trim(), role, adminPassword });
      setUsername(""); setName(""); setAdminPassword("");
    } catch (e) {
      setError(e.message || "Không tạo được tài khoản.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Tạo tài khoản mới</p>
      <div className="grid sm:grid-cols-3 gap-3 mb-3">
        <TextField label="Tên đăng nhập" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="vd: nvkho2" />
        <TextField label="Họ tên" value={name} onChange={(e) => setName(e.target.value)} />
        <SelectField label="Vai trò" value={role} onChange={(e) => setRole(e.target.value)}>
          <option value="nhan_vien_kho">Nhân viên kho</option>
          <option value="quan_ly">Quản lý</option>
          <option value="bao_cao">Quản lý (Báo cáo)</option>
          <option value="thu_ngan">Thu ngân</option>
        </SelectField>
        <TextField label="Xác nhận: mật khẩu của chính bạn" type="password" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} className="sm:col-span-3" />
      </div>
      <p className="text-xs text-slate-400 mb-3">Mật khẩu mặc định ban đầu: <span className="font-mono font-medium">123456</span> — tài khoản sẽ bị bắt buộc đổi mật khẩu trong 24h ở lần đăng nhập đầu tiên.</p>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Tạo tài khoản</PrimaryButton>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// CHI PHÍ (Vận hành / Marketing & bán hàng / Bảo trì & vật tư / Khác)
// ---------------------------------------------------------------------------
const EXPENSE_CATEGORIES = [
  { key: "nvl", label: "Chi phí nguyên vật liệu" },
  { key: "van_hanh", label: "Chi phí vận hành", presetItems: ["Nhân công", "Mặt bằng", "Điện", "Nước", "Gas", "Internet", "Điện thoại"] },
  { key: "marketing", label: "Chi phí Marketing & bán hàng", presetItems: ["Quảng cáo"] },
  { key: "bao_tri_vat_tu", label: "Chi phí bảo trì và vật tư" },
  { key: "khac", label: "Chi phí khác" },
];
const EXPENSE_CATEGORY_META = Object.fromEntries(EXPENSE_CATEGORIES.map((c) => [c.key, c]));
const EXPENSE_PAYMENT_METHOD_META = {
  tien_mat: { label: "Tiền mặt", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
  chuyen_khoan: { label: "Chuyển khoản", color: "bg-sky-50 text-sky-700 border-sky-200" },
};

// ---------------------------------------------------------------------------
// QUỸ (Sổ quỹ thu-chi) — Phiếu thu lấy trực tiếp từ dữ liệu bán hàng (dish_sales,
// nguồn "Xuất kho tự động từ báo cáo doanh thu"), không nhập tay. Phiếu chi nhập tay
// theo ngày, chọn đúng loại chi phí — khi lưu sẽ ghi thẳng vào expense_records nên
// tự động xuất hiện trong "Chi phí" / báo cáo chi phí, không cần đồng bộ thủ công.
// ---------------------------------------------------------------------------
// Sổ quỹ theo ngày — tách riêng Tiền mặt và Ngân hàng, đi qua đủ 4 dòng tiền: Thu ngân
// (cashier_receipts), Thu/Chi cọc (deposits, lọc theo payment_method), Chi phí (expense_records,
// lọc theo payment_method). Tồn đầu ngày lấy override trong overridesMap nếu có (dạng
// {cash, bank}), không thì lấy Tồn cuối của ngày liền trước trong danh sách.
function buildFundLedger(dates, cashierReceipts, expenseRecords, deposits, overridesMap, remittedMap = {}) {
  let prevCash = null;
  let prevBank = null;
  return dates.map((d) => {
    const thuNganCash = cashierReceipts.filter((r) => r.receiptDate === d).reduce((s, r) => s + r.cashAmount, 0);
    const thuNganBank = cashierReceipts.filter((r) => r.receiptDate === d).reduce((s, r) => s + r.bankAmount, 0);
    const thuCocCash = deposits.filter((r) => r.direction === "thu" && r.depositDate === d && r.paymentMethod === "tien_mat").reduce((s, r) => s + r.amount, 0);
    const thuCocBank = deposits.filter((r) => r.direction === "thu" && r.depositDate === d && r.paymentMethod === "ngan_hang").reduce((s, r) => s + r.amount, 0);
    const chiCocCash = deposits.filter((r) => r.direction === "chi" && r.depositDate === d && r.paymentMethod === "tien_mat").reduce((s, r) => s + r.amount, 0);
    const chiCocBank = deposits.filter((r) => r.direction === "chi" && r.depositDate === d && r.paymentMethod === "ngan_hang").reduce((s, r) => s + r.amount, 0);
    const chiPhiCash = expenseRecords.filter((r) => r.expenseDate === d && r.paymentMethod === "tien_mat").reduce((s, r) => s + r.amount, 0);
    const chiPhiBank = expenseRecords.filter((r) => r.expenseDate === d && r.paymentMethod === "chuyen_khoan").reduce((s, r) => s + r.amount, 0);
    const remittedCash = remittedMap[d]?.cash || 0;
    const remittedBank = remittedMap[d]?.bank || 0;

    const ov = overridesMap[d];
    const hasOverride = !!ov;
    const openingCash = hasOverride ? ov.cash : (prevCash !== null ? prevCash : 0);
    const openingBank = hasOverride ? ov.bank : (prevBank !== null ? prevBank : 0);

    const closingCash = openingCash + thuNganCash + thuCocCash - chiPhiCash - chiCocCash - remittedCash;
    const closingBank = openingBank + thuNganBank + thuCocBank - chiPhiBank - chiCocBank - remittedBank;
    prevCash = closingCash;
    prevBank = closingBank;

    return {
      date: d, hasOverride,
      openingCash, openingBank,
      thuNganCash, thuNganBank, thuCocCash, thuCocBank, chiPhiCash, chiPhiBank, chiCocCash, chiCocBank,
      remittedCash, remittedBank,
      closingCash, closingBank,
    };
  });
}

function dailyReceiptsFromSales(dishSales, from, to) {
  const filtered = dishSales.filter((s) => s.saleDate >= from && s.saleDate <= to);
  const map = new Map();
  filtered.forEach((s) => {
    const cur = map.get(s.saleDate) || { date: s.saleDate, amount: 0, qty: 0 };
    cur.amount += s.totalAmount;
    cur.qty += s.quantity;
    map.set(s.saleDate, cur);
  });
  return Array.from(map.values()).sort((a, b) => b.date.localeCompare(a.date));
}

function PhieuChiForm({ onSubmit }) {
  const [expenseDate, setExpenseDate] = useState(todayISO());
  const [category, setCategory] = useState("van_hanh");
  const [itemName, setItemName] = useState("");
  const [amount, setAmount] = useState("");
  const [paymentMethod, setPaymentMethod] = useState("tien_mat");
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const chiCategories = EXPENSE_CATEGORIES; // NVL cũng chọn được ở đây — dùng cho các khoản mua nguyên liệu lẻ (thịt/cá/tôm...) không qua phiếu nhập Excel

  const submit = async () => {
    if (!itemName.trim()) { setError("Vui lòng nhập tên khoản chi."); return; }
    if (!amount || Number(amount) <= 0) { setError("Vui lòng nhập số tiền hợp lệ."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ category, expenseDate, paymentMethod, lines: [{ itemName: itemName.trim(), amount: Number(amount), note: note.trim() || null }] });
      setItemName(""); setAmount(""); setNote("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Tạo phiếu chi</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <TextField label="Ngày chi" type="date" value={expenseDate} onChange={(e) => setExpenseDate(e.target.value)} />
        <SelectField label="Loại chi phí" value={category} onChange={(e) => setCategory(e.target.value)}>
          {chiCategories.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
        </SelectField>
        <TextField label="Tên khoản chi" value={itemName} onChange={(e) => setItemName(e.target.value)} placeholder="VD: Tiền điện tháng 8" />
        <MoneyField label="Số tiền" value={amount} onChange={setAmount} />
        <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
          <option value="tien_mat">Tiền mặt</option>
          <option value="chuyen_khoan">Chuyển khoản</option>
        </SelectField>
        <TextField label="Ghi chú (tuỳ chọn)" value={note} onChange={(e) => setNote(e.target.value)} className="sm:col-span-2" />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Wallet size={15} />} Lưu phiếu chi</PrimaryButton>
      <p className="text-xs text-slate-400 mt-2">Phiếu chi sau khi lưu sẽ tự động xuất hiện trong tab "Chi phí" đúng nhóm "{EXPENSE_CATEGORY_META[category]?.label}".</p>
    </Card>
  );
}

function CashierReceiptForm({ onSubmit }) {
  const [receiptDate, setReceiptDate] = useState(todayISO());
  const [cashAmount, setCashAmount] = useState("");
  const [bankAmount, setBankAmount] = useState("");
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const total = (Number(cashAmount) || 0) + (Number(bankAmount) || 0);

  const submit = async () => {
    if (total <= 0) { setError("Vui lòng nhập ít nhất 1 trong 2 khoản: tiền mặt hoặc tiền ngân hàng."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ receiptDate, cashAmount, bankAmount, note: note.trim() });
      setCashAmount(""); setBankAmount(""); setNote("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Ghi nhận Thu ngân</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <TextField label="Ngày thu" type="date" value={receiptDate} onChange={(e) => setReceiptDate(e.target.value)} />
        <div />
        <MoneyField label="Thu tiền mặt" value={cashAmount} onChange={setCashAmount} />
        <MoneyField label="Thu tiền ngân hàng" value={bankAmount} onChange={setBankAmount} />
        <TextField label="Ghi chú (tuỳ chọn)" value={note} onChange={(e) => setNote(e.target.value)} className="sm:col-span-2" />
      </div>
      <div className="bg-emerald-50 rounded-xl px-3 py-2 flex items-center justify-between text-sm mb-4">
        <span className="text-emerald-700">Tổng thu</span>
        <span className="font-semibold text-emerald-800">{fmtMoney(total)}</span>
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <TrendingUp size={15} />} Lưu phiếu thu</PrimaryButton>
      <p className="text-xs text-slate-400 mt-2">Số liệu này độc lập với doanh thu tính từ báo cáo bán hàng — dùng để Quản lý đối chiếu quỹ thực tế.</p>
    </Card>
  );
}

function EditCashierReceiptModal({ receipt, onSave, onClose }) {
  const [receiptDate, setReceiptDate] = useState(receipt.receiptDate);
  const [cashAmount, setCashAmount] = useState(String(receipt.cashAmount ?? ""));
  const [bankAmount, setBankAmount] = useState(String(receipt.bankAmount ?? ""));
  const [note, setNote] = useState(receipt.note || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!(Number(cashAmount) > 0 || Number(bankAmount) > 0)) { setError("Vui lòng nhập ít nhất 1 số tiền hợp lệ."); return; }
    setError(""); setSaving(true);
    try {
      await onSave(receipt.id, { receiptDate, cashAmount: Number(cashAmount) || 0, bankAmount: Number(bankAmount) || 0, note: note.trim() });
      onClose();
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><Pencil size={17} className="text-sky-700" /> Sửa phiếu thu</p>
          <button onClick={onClose}><X size={18} className="text-slate-400" /></button>
        </div>
        <div className="space-y-3">
          <TextField label="Ngày" type="date" value={receiptDate} onChange={(e) => setReceiptDate(e.target.value)} />
          <MoneyField label="Tiền mặt" value={cashAmount} onChange={setCashAmount} />
          <MoneyField label="Tiền ngân hàng" value={bankAmount} onChange={setBankAmount} />
          <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />
          {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertTriangle size={14} /> {error}</p>}
          <div className="flex gap-2">
            <PrimaryButton type="button" onClick={submit} disabled={saving}>{saving ? "Đang lưu..." : "Lưu thay đổi"}</PrimaryButton>
            <GhostButton type="button" onClick={onClose}>Huỷ</GhostButton>
          </div>
        </div>
      </div>
    </div>
  );
}

function ThuModule({ data, currentUser, onSubmit, onUpdate, onDelete, editEnabled }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const myReceipts = data.cashierReceipts
    .filter((r) => r.createdBy === currentUser.id)
    .filter((r) => (!from || r.receiptDate >= from) && (!to || r.receiptDate <= to))
    .slice(0, 50);
  const todayReceipts = data.cashierReceipts.filter((r) => r.receiptDate === todayISO());
  const todayTotal = todayReceipts.reduce((s, r) => s + r.cashAmount + r.bankAmount, 0);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);

  const handleDelete = async (id) => {
    if (!window.confirm("Xoá phiếu thu này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };

  return (
    <div>
      <SectionTitle icon={Wallet} title="Thu ngân" subtitle="Ghi nhận phiếu thu tiền mặt/ngân hàng trong ngày" />

      {!editEnabled && (
        <div className="flex items-center gap-2 text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2 mb-5">
          <Lock size={15} /> Quản lý đã khoá quyền tự sửa/xoá phiếu thu — bạn chỉ ghi nhận mới được, cần chỉnh sửa thì báo Quản lý.
        </div>
      )}

      <div className="grid grid-cols-2 gap-3 mb-5">
        <MetricCard label="Tổng thu hôm nay" value={fmtMoney(todayTotal)} icon={TrendingUp} accent="emerald" />
        <MetricCard label="Số phiếu thu hôm nay" value={fmtNumber(todayReceipts.length)} icon={Receipt} accent="teal" />
      </div>

      <CashierReceiptForm onSubmit={onSubmit} />

      <Card className="p-4 sm:p-5 mb-5">
        <p className="text-xs font-medium text-slate-500 mb-2">Lọc theo ngày</p>
        <div className="grid grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
        {(from || to) && (
          <button type="button" onClick={() => { setFrom(""); setTo(""); }} className="text-xs text-sky-700 hover:underline mt-2">Bỏ lọc ngày</button>
        )}
      </Card>

      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Phiếu thu {(from || to) ? "" : "gần đây "}(của bạn)</p></div>
        {myReceipts.length === 0 ? <EmptyState icon={TrendingUp} text="Chưa có phiếu thu nào." /> : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Ngày</th><th className="px-3 py-2 text-right">Tiền mặt</th><th className="px-3 py-2 text-right">Ngân hàng</th><th className="px-3 py-2 text-right">Tổng</th><th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {myReceipts.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2 text-slate-500">{fmtDate(r.receiptDate)}</td>
                    <td className="px-3 py-2 text-right">{fmtMoney(r.cashAmount)}</td>
                    <td className="px-3 py-2 text-right">{fmtMoney(r.bankAmount)}</td>
                    <td className="px-3 py-2 text-right font-medium text-emerald-700">{fmtMoney(r.cashAmount + r.bankAmount)}</td>
                    <td className="px-3 py-2">
                      {editEnabled && (
                        <div className="flex items-center justify-end gap-1">
                          <button type="button" onClick={() => setEditing(r)} className="text-slate-400 hover:text-sky-700 p-1" title="Sửa phiếu thu này"><Pencil size={14} /></button>
                          <button type="button" onClick={() => handleDelete(r.id)} disabled={deleting === r.id} className="text-slate-400 hover:text-rose-600 p-1" title="Xoá phiếu thu này">
                            {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                          </button>
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      {editing && editEnabled && (
        <EditCashierReceiptModal
          receipt={editing}
          onSave={onUpdate}
          onClose={() => setEditing(null)}
        />
      )}
    </div>
  );
}

function ChiPhieuModule({ data, currentUser, onSubmit, onUpdate, onDelete, editEnabled }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [category, setCategory] = useState("all");
  const filteredExpenses = data.expenseRecords
    .filter((r) => r.createdBy === currentUser.id)
    .filter((r) => (!from || r.expenseDate >= from) && (!to || r.expenseDate <= to))
    .filter((r) => category === "all" || r.category === category);
  const totalFilteredAmount = filteredExpenses.reduce((s, r) => s + r.amount, 0);
  const myExpenses = filteredExpenses.slice(0, 50);
  const [editing, setEditing] = useState(null);
  const [deleting, setDeleting] = useState(null);

  const handleDelete = async (id) => {
    if (!window.confirm("Xoá phiếu chi này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };

  return (
    <div>
      <SectionTitle icon={Receipt} title="Phiếu chi" subtitle="Ghi nhận phiếu chi phát sinh trong ngày" />

      {!editEnabled && (
        <div className="flex items-center gap-2 text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2 mb-5">
          <Lock size={15} /> Quản lý đã khoá quyền tự sửa/xoá phiếu chi — bạn chỉ ghi nhận mới được, cần chỉnh sửa thì báo Quản lý.
        </div>
      )}

      <PhieuChiForm onSubmit={onSubmit} />

      <Card className="p-4 sm:p-5 mb-5">
        <p className="text-xs font-medium text-slate-500 mb-2">Lọc theo ngày &amp; loại chi phí</p>
        <div className="grid sm:grid-cols-3 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
          <SelectField label="Loại chi phí" value={category} onChange={(e) => setCategory(e.target.value)}>
            <option value="all">Tất cả</option>
            {EXPENSE_CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
          </SelectField>
        </div>
        {(from || to || category !== "all") && (
          <button type="button" onClick={() => { setFrom(""); setTo(""); setCategory("all"); }} className="text-xs text-sky-700 hover:underline mt-2">Bỏ lọc</button>
        )}
      </Card>

      <div className="grid grid-cols-2 gap-3 mb-5">
        <MetricCard label="Số phiếu chi (theo bộ lọc)" value={fmtNumber(filteredExpenses.length)} icon={Receipt} accent="rose" />
        <MetricCard label="Tổng chi phí (theo bộ lọc)" value={fmtMoney(totalFilteredAmount)} icon={TrendingDown} accent="rose" />
      </div>

      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100"><p className="font-semibold text-slate-800 text-sm">Phiếu chi {(from || to || category !== "all") ? "" : "gần đây "}(của bạn)</p></div>
        {myExpenses.length === 0 ? <EmptyState icon={Receipt} text="Chưa có phiếu chi nào." /> : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Ngày</th><th className="px-3 py-2">Khoản chi</th><th className="px-3 py-2">Nhóm</th><th className="px-3 py-2">Hình thức</th><th className="px-3 py-2 text-right">Số tiền</th><th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {myExpenses.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2 text-slate-500">{fmtDate(r.expenseDate)}</td>
                    <td className="px-3 py-2 text-slate-700">{r.itemName}</td>
                    <td className="px-3 py-2 text-slate-500">{EXPENSE_CATEGORY_META[r.category]?.label || r.category}</td>
                    <td className="px-3 py-2"><span className={`text-xs px-1.5 py-0.5 rounded-lg border ${EXPENSE_PAYMENT_METHOD_META[r.paymentMethod]?.color}`}>{EXPENSE_PAYMENT_METHOD_META[r.paymentMethod]?.label || r.paymentMethod}</span></td>
                    <td className="px-3 py-2 text-right font-medium text-rose-700">{fmtMoney(r.amount)}</td>
                    <td className="px-3 py-2">
                      {editEnabled && (
                        <div className="flex items-center justify-end gap-1">
                          <button type="button" onClick={() => setEditing(r)} className="text-slate-400 hover:text-sky-700 p-1" title="Sửa phiếu chi này"><Pencil size={14} /></button>
                          <button type="button" onClick={() => handleDelete(r.id)} disabled={deleting === r.id} className="text-slate-400 hover:text-rose-600 p-1" title="Xoá phiếu chi này">
                            {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                          </button>
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      {editing && editEnabled && (
        <EditExpenseModal
          expense={editing}
          onSave={onUpdate}
          onClose={() => setEditing(null)}
        />
      )}
    </div>
  );
}

// Import "Bảng kê hoá đơn" (POS) — Doanh thu bán hàng theo hoá đơn, cột: Ngày, Số hoá đơn, Doanh thu.
function InvoiceRevenueImportForm({ onImport }) {
  const [fileName, setFileName] = useState("");
  const [preview, setPreview] = useState([]);
  const [errors, setErrors] = useState([]);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [cancelled, setCancelled] = useState(0);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setDone(null);
    try {
      const rawRows = await readExcelRaw(file);
      const headerIdx = detectHeaderRow(rawRows, ["Số hóa đơn", "Doanh thu"]);
      if (headerIdx === -1) { setErrors(['Không tìm thấy dòng tiêu đề (cần có cột "Số hóa đơn" và "Doanh thu") trong file.']); setPreview([]); return; }
      const rawObjRows = rowsToObjects(rawRows, headerIdx);
      const valid = [];
      let cancelledCount = 0;
      rawObjRows.forEach((row) => {
        const invoiceNo = String(pickCol(row, "Số hóa đơn", "So hoa don") ?? "").trim();
        const amount = Number(pickCol(row, "Doanh thu", "Doanh thu bán hàng")) || 0;
        const invoiceDate = excelDateToISO(pickCol(row, "Ngày", "Ngay") ?? pickColContains(row, "ngay")) || todayISO();
        const note = String(pickCol(row, "Ghi chú", "Ghi chu") ?? "").trim();
        if (!invoiceNo) return;
        if (stripDiacritics(note).includes("huy")) { cancelledCount += 1; return; } // bỏ hoá đơn đã huỷ
        valid.push({ invoiceNo, invoiceDate, amount });
      });
      setPreview(valid);
      setCancelled(cancelledCount);
      setErrors([]);
    } catch (err) {
      setErrors([err.message || "Không đọc được file, kiểm tra lại định dạng .xlsx/.xls."]);
      setPreview([]);
    }
  };

  const submit = async () => {
    if (preview.length === 0) return;
    setImporting(true);
    try {
      await onImport(preview);
      setDone(preview.length);
      setPreview([]); setFileName("");
    } catch (e) {
      setErrors([e.message || "Không import được, vui lòng thử lại."]);
    } finally {
      setImporting(false);
    }
  };

  const totalPreview = preview.reduce((s, r) => s + r.amount, 0);

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-1">Import Doanh thu bán hàng theo hoá đơn</p>
      <p className="text-xs text-slate-500 mb-3">Dùng file <b>"Bảng kê hoá đơn"</b> xuất từ POS (khác với file "Chi tiết doanh thu theo hoá đơn và mặt hàng"). Cần có cột <b>Ngày</b>, <b>Số hoá đơn</b>, <b>Doanh thu</b>. Import lại cùng số hoá đơn sẽ tự ghi đè, không bị trùng. Hoá đơn có ghi chú "đã huỷ" sẽ tự động bị loại bỏ, không tính vào doanh thu.</p>
      <div className="flex items-center gap-2 mb-3">
        <label className="inline-flex items-center gap-2 cursor-pointer text-sm font-medium text-sky-700 border border-sky-300 bg-sky-50 hover:bg-sky-100 rounded-xl px-4 py-2">
          <Upload size={15} /> {fileName || "Chọn file Excel..."}
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
        </label>
        {fileName && (
          <button type="button" onClick={() => { setFileName(""); setPreview([]); setErrors([]); setCancelled(0); }} className="text-slate-400 hover:text-rose-600 p-1.5" title="Bỏ file đã chọn">
            <X size={16} />
          </button>
        )}
      </div>
      {errors.length > 0 && (
        <div className="mb-3 text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-xl p-3 space-y-1">
          {errors.map((e, i) => <p key={i} className="flex items-center gap-1"><AlertTriangle size={12} className="shrink-0" /> {e}</p>)}
        </div>
      )}
      {cancelled > 0 && (
        <div className="mb-3 text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-xl p-3 flex items-center gap-1">
          <AlertTriangle size={12} className="shrink-0" /> Đã loại bỏ {cancelled} hoá đơn có ghi chú "đã huỷ" — không tính vào doanh thu.
        </div>
      )}
      {preview.length > 0 && (
        <>
          <p className="text-xs text-slate-500 mb-3">{preview.length} hoá đơn · Tổng doanh thu <span className="font-semibold text-sky-700">{fmtMoney(totalPreview)}</span></p>
          <PrimaryButton onClick={submit} disabled={importing}>{importing ? <Loader2 size={15} className="animate-spin" /> : <Upload size={15} />} Import {preview.length} hoá đơn</PrimaryButton>
        </>
      )}
      {done !== null && <p className="text-xs text-emerald-600 mt-2">Đã import thành công {done} hoá đơn.</p>}
    </Card>
  );
}

const DEPOSIT_PAYMENT_METHOD_META = {
  tien_mat: { label: "Tiền mặt", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
  ngan_hang: { label: "Ngân hàng", color: "bg-sky-50 text-sky-700 border-sky-200" },
};
const DEPOSIT_STATUS_META = {
  dang_giu: { label: "Đang giữ", color: "bg-amber-50 text-amber-700 border-amber-200" },
  da_doi_tru: { label: "Đã đối trừ", color: "bg-sky-50 text-sky-700 border-sky-200" },
  da_hoan_tra: { label: "Đã hoàn trả", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
};
const DEPOSIT_STATUSES = Object.keys(DEPOSIT_STATUS_META);

function DepositForm({ onSubmit }) {
  const [direction, setDirection] = useState("thu");
  const [partyName, setPartyName] = useState("");
  const [amount, setAmount] = useState("");
  const [paymentMethod, setPaymentMethod] = useState("tien_mat");
  const [depositDate, setDepositDate] = useState(todayISO());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!partyName.trim()) { setError(direction === "thu" ? "Vui lòng nhập tên khách hàng." : "Vui lòng nhập tên NCC/đối tác."); return; }
    if (!amount || Number(amount) <= 0) { setError("Vui lòng nhập số tiền hợp lệ."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ direction, partyName, amount: Number(amount), paymentMethod, depositDate, note: note.trim() });
      setPartyName(""); setAmount(""); setNote("");
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Ghi nhận cọc</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <SelectField label="Loại" value={direction} onChange={(e) => setDirection(e.target.value)}>
          <option value="thu">Thu cọc (nhận cọc từ khách)</option>
          <option value="chi">Chi cọc (ứng cọc cho NCC/đối tác)</option>
        </SelectField>
        <TextField label="Ngày" type="date" value={depositDate} onChange={(e) => setDepositDate(e.target.value)} />
        <TextField label={direction === "thu" ? "Tên khách hàng" : "Tên NCC/đối tác"} value={partyName} onChange={(e) => setPartyName(e.target.value)} placeholder={direction === "thu" ? "VD: Anh Tuấn - đặt bàn 20/8" : "VD: NCC hải sản An Phát"} />
        <MoneyField label="Số tiền" value={amount} onChange={setAmount} />
        <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
          <option value="tien_mat">Tiền mặt</option>
          <option value="ngan_hang">Ngân hàng</option>
        </SelectField>
        <TextField label="Ghi chú (tuỳ chọn)" value={note} onChange={(e) => setNote(e.target.value)} className="sm:col-span-2" />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Coins size={15} />} Lưu khoản cọc</PrimaryButton>
    </Card>
  );
}

function DepositList({ data, currentUser, from, to, onDelete, onUpdate }) {
  const isQuanLy = currentUser?.role === "quan_ly" || currentUser?.role === "bao_cao";
  const isThuNgan = currentUser?.role === "thu_ngan";
  const [direction, setDirection] = useState("all");
  const [deleting, setDeleting] = useState(null);
  const [editing, setEditing] = useState(null);
  const [changingStatus, setChangingStatus] = useState(null);

  const rows = data.deposits
    .filter((r) => direction === "all" || r.direction === direction)
    .filter((r) => (!from || r.depositDate >= from) && (!to || r.depositDate <= to));

  const totalThu = data.deposits.filter((r) => r.direction === "thu" && (!from || r.depositDate >= from) && (!to || r.depositDate <= to)).reduce((s, r) => s + r.amount, 0);
  const totalChi = data.deposits.filter((r) => r.direction === "chi" && (!from || r.depositDate >= from) && (!to || r.depositDate <= to)).reduce((s, r) => s + r.amount, 0);
  const dangGiuThu = data.deposits.filter((r) => r.direction === "thu" && r.status === "dang_giu").reduce((s, r) => s + r.amount, 0);
  const dangGiuChi = data.deposits.filter((r) => r.direction === "chi" && r.status === "dang_giu").reduce((s, r) => s + r.amount, 0);

  const handleDelete = async (id) => {
    if (!window.confirm("Xoá khoản cọc này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };
  const handleStatusChange = async (r, newStatus) => {
    if (newStatus === r.status) return;
    setChangingStatus(r.id);
    try { await onUpdate(r.id, { direction: r.direction, partyName: r.partyName, amount: r.amount, paymentMethod: r.paymentMethod, depositDate: r.depositDate, status: newStatus, note: r.note }); }
    finally { setChangingStatus(null); }
  };

  return (
    <>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-5">
        <div className="bg-emerald-50 rounded-xl px-3 py-2">
          <p className="text-xs text-emerald-700">Tổng thu cọc</p>
          <p className="text-sm font-semibold text-emerald-800">{fmtMoney(totalThu)}</p>
        </div>
        <div className="bg-rose-50 rounded-xl px-3 py-2">
          <p className="text-xs text-rose-700">Tổng chi cọc</p>
          <p className="text-sm font-semibold text-rose-800">{fmtMoney(totalChi)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Đang giữ (thu cọc)</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(dangGiuThu)}</p>
        </div>
        <div className="bg-amber-50 rounded-xl px-3 py-2">
          <p className="text-xs text-amber-700">Đang giữ (chi cọc)</p>
          <p className="text-sm font-semibold text-amber-800">{fmtMoney(dangGiuChi)}</p>
        </div>
      </div>

      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-2">
          <p className="font-semibold text-slate-800 text-sm">Danh sách cọc</p>
          <div className="flex gap-1">
            {[{ key: "all", label: "Tất cả" }, { key: "thu", label: "Thu cọc" }, { key: "chi", label: "Chi cọc" }].map((t) => (
              <button
                key={t.key}
                type="button"
                onClick={() => setDirection(t.key)}
                className={`text-xs px-2.5 py-1 rounded-lg border ${direction === t.key ? "bg-sky-700 text-white border-sky-700" : "text-slate-500 border-slate-200 hover:bg-slate-50"}`}
              >
                {t.label}
              </button>
            ))}
          </div>
        </div>
        {rows.length === 0 ? <EmptyState icon={Coins} text="Chưa có khoản cọc nào." /> : (
          <div className="divide-y divide-slate-100">
            {rows.map((r) => (
              <div key={r.id} className="px-4 py-3 flex items-center justify-between gap-3">
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium text-slate-800 truncate">
                    {r.partyName} <span className={`ml-1 text-[10px] px-1.5 py-0.5 rounded-full border ${r.direction === "thu" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-rose-50 text-rose-700 border-rose-200"}`}>{r.direction === "thu" ? "Thu cọc" : "Chi cọc"}</span>
                  </p>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    {isQuanLy ? (
                      <select
                        value={r.status}
                        onChange={(e) => handleStatusChange(r, e.target.value)}
                        disabled={changingStatus === r.id}
                        className="text-xs border border-slate-200 rounded-lg px-1.5 py-0.5 text-slate-600 bg-slate-50 focus:outline-none focus:ring-1 focus:ring-sky-400 disabled:opacity-50"
                      >
                        {DEPOSIT_STATUSES.map((s) => <option key={s} value={s}>{DEPOSIT_STATUS_META[s].label}</option>)}
                      </select>
                    ) : (
                      <span className={`text-xs px-1.5 py-0.5 rounded-lg border ${DEPOSIT_STATUS_META[r.status]?.color}`}>{DEPOSIT_STATUS_META[r.status]?.label || r.status}</span>
                    )}
                    <p className="text-xs text-slate-400">· {fmtDate(r.depositDate)}</p>
                    <span className={`text-xs px-1.5 py-0.5 rounded-lg border ${DEPOSIT_PAYMENT_METHOD_META[r.paymentMethod]?.color}`}>{DEPOSIT_PAYMENT_METHOD_META[r.paymentMethod]?.label || r.paymentMethod}</span>
                    {r.note && <p className="text-xs text-slate-400 truncate">· {r.note}</p>}
                    {changingStatus === r.id && <Loader2 size={11} className="animate-spin text-slate-400" />}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <p className={`text-sm font-semibold ${r.direction === "thu" ? "text-emerald-700" : "text-rose-700"}`}>{fmtMoney(r.amount)}</p>
                  {isQuanLy && (
                    <>
                      <button type="button" onClick={() => setEditing(r)} className="text-slate-400 hover:text-sky-700 p-1" title="Sửa khoản cọc này"><Pencil size={14} /></button>
                      <button type="button" onClick={() => handleDelete(r.id)} disabled={deleting === r.id} className="text-slate-400 hover:text-rose-600 p-1" title="Xoá khoản cọc này">
                        {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                      </button>
                    </>
                  )}
                  {!isQuanLy && isThuNgan && r.createdBy === currentUser.id && (
                    <button type="button" onClick={() => setEditing(r)} className="text-slate-400 hover:text-sky-700 p-1" title="Sửa khoản cọc này"><Pencil size={14} /></button>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>
      {editing && (
        <EditDepositModal
          deposit={editing}
          onSave={async (id, patch) => { await onUpdate(id, patch); }}
          onClose={() => setEditing(null)}
        />
      )}
    </>
  );
}

function EditDepositModal({ deposit, onSave, onClose }) {
  const [direction, setDirection] = useState(deposit.direction);
  const [partyName, setPartyName] = useState(deposit.partyName);
  const [amount, setAmount] = useState(String(deposit.amount ?? ""));
  const [paymentMethod, setPaymentMethod] = useState(deposit.paymentMethod || "tien_mat");
  const [depositDate, setDepositDate] = useState(deposit.depositDate);
  const [status, setStatus] = useState(deposit.status);
  const [note, setNote] = useState(deposit.note || "");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!partyName.trim()) { setError("Vui lòng nhập tên."); return; }
    if (!(Number(amount) > 0)) { setError("Vui lòng nhập số tiền hợp lệ."); return; }
    setError(""); setSaving(true);
    try {
      await onSave(deposit.id, { direction, partyName: partyName.trim(), amount: Number(amount), paymentMethod, depositDate, status, note: note.trim() });
      onClose();
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><Pencil size={17} className="text-sky-700" /> Sửa khoản cọc</p>
          <button onClick={onClose}><X size={18} className="text-slate-400" /></button>
        </div>
        <div className="space-y-3">
          <SelectField label="Loại" value={direction} onChange={(e) => setDirection(e.target.value)}>
            <option value="thu">Thu cọc</option>
            <option value="chi">Chi cọc</option>
          </SelectField>
          <TextField label="Tên khách/NCC" value={partyName} onChange={(e) => setPartyName(e.target.value)} />
          <TextField label="Ngày" type="date" value={depositDate} onChange={(e) => setDepositDate(e.target.value)} />
          <SelectField label="Trạng thái" value={status} onChange={(e) => setStatus(e.target.value)}>
            {DEPOSIT_STATUSES.map((s) => <option key={s} value={s}>{DEPOSIT_STATUS_META[s].label}</option>)}
          </SelectField>
          <MoneyField label="Số tiền" value={amount} onChange={setAmount} />
          <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
            <option value="tien_mat">Tiền mặt</option>
            <option value="ngan_hang">Ngân hàng</option>
          </SelectField>
          <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />
          {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertTriangle size={14} /> {error}</p>}
          <div className="flex gap-2">
            <PrimaryButton type="button" onClick={submit} disabled={saving}>{saving ? "Đang lưu..." : "Lưu thay đổi"}</PrimaryButton>
            <GhostButton type="button" onClick={onClose}>Huỷ</GhostButton>
          </div>
        </div>
      </div>
    </div>
  );
}

function DepositModule({ data, currentUser, onSubmit, onUpdate, onDelete }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");

  return (
    <div>
      <SectionTitle icon={Coins} title="Cọc" subtitle="Theo dõi thu cọc (nhận từ khách) và chi cọc (ứng cho NCC/đối tác), tách riêng khỏi Chi phí" />
      <Card className="p-4 sm:p-5 mb-5">
        <p className="text-xs font-medium text-slate-500 mb-2">Lọc theo ngày</p>
        <div className="grid grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
        {(from || to) && (
          <button type="button" onClick={() => { setFrom(""); setTo(""); }} className="text-xs text-sky-700 hover:underline mt-2">Bỏ lọc ngày</button>
        )}
      </Card>
      {currentUser?.role !== "bao_cao" && <DepositForm onSubmit={onSubmit} />}
      <DepositList data={data} currentUser={currentUser} from={from} to={to} onDelete={onDelete} onUpdate={onUpdate} />
    </div>
  );
}

function timeAgo(iso) {
  if (!iso) return "";
  const diffMs = Date.now() - new Date(iso).getTime();
  const mins = Math.floor(diffMs / 60000);
  if (mins < 1) return "vừa xong";
  if (mins < 60) return `${mins} phút trước`;
  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours} giờ trước`;
  const days = Math.floor(hours / 24);
  return `${days} ngày trước`;
}

const NOTI_TYPE_ICON = { chi_phi: Receipt, thu_ngan: Wallet, coc: Coins };

function NotificationBell({ notifications, onMarkRead, onMarkAllRead }) {
  const [open, setOpen] = useState(false);
  const unread = notifications.filter((n) => !n.isRead);
  const rows = notifications.slice(0, 20);

  return (
    <div className="relative">
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className="relative w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition shrink-0"
        title="Thông báo"
      >
        <Bell size={14} className="text-slate-500" />
        {unread.length > 0 && (
          <span className="absolute -top-1 -right-1 min-w-[16px] h-4 px-1 rounded-full bg-rose-600 text-white text-[10px] font-semibold flex items-center justify-center">
            {unread.length > 9 ? "9+" : unread.length}
          </span>
        )}
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-40" onClick={() => setOpen(false)} />
          <div className="absolute right-0 mt-2 w-80 max-w-[90vw] bg-white rounded-2xl shadow-xl border border-slate-100 z-50 overflow-hidden" style={{ animation: "fadeIn .15s ease-out" }}>
            <div className="p-3 border-b border-slate-100 flex items-center justify-between">
              <p className="text-sm font-semibold text-slate-800">Thông báo</p>
              {unread.length > 0 && (
                <button type="button" onClick={onMarkAllRead} className="text-xs text-sky-700 hover:underline">Đánh dấu đã đọc hết</button>
              )}
            </div>
            <div className="max-h-96 overflow-y-auto divide-y divide-slate-50">
              {rows.length === 0 ? (
                <div className="p-6 text-center text-sm text-slate-400">Chưa có thông báo nào.</div>
              ) : rows.map((n) => {
                const Icon = NOTI_TYPE_ICON[n.type] || Bell;
                return (
                  <button
                    key={n.id}
                    type="button"
                    onClick={() => !n.isRead && onMarkRead(n.id)}
                    className={`w-full text-left px-3 py-2.5 flex items-start gap-2.5 hover:bg-slate-50 transition ${!n.isRead ? "bg-sky-50/60" : ""}`}
                  >
                    <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 mt-0.5 ${!n.isRead ? "bg-sky-100 text-sky-700" : "bg-slate-100 text-slate-400"}`}>
                      <Icon size={13} />
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className={`text-xs leading-snug ${!n.isRead ? "text-slate-800 font-medium" : "text-slate-500"}`}>{n.message}</p>
                      <p className="text-[11px] text-slate-400 mt-0.5">{timeAgo(n.createdAt)}</p>
                    </div>
                    {!n.isRead && <span className="w-2 h-2 rounded-full bg-sky-600 shrink-0 mt-1.5" />}
                  </button>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// Sổ quỹ tiền mặt theo ngày — Tồn đầu ngày tự nhảy từ Tồn cuối ngày liền trước, sửa được từng ngày
// (lưu vào fund_daily_balance); khi sửa 1 ngày, các ngày sau đó trong bảng tự tính lại theo.
function FundLedgerTable({ ledger, onSaveOpening, onSaveRemitted }) {
  const [currency, setCurrency] = useState("cash"); // "cash" | "bank"
  const [editingCell, setEditingCell] = useState(null); // { date, field: "opening" | "remitted" }
  const [draftValue, setDraftValue] = useState("");
  const [saving, setSaving] = useState(false);

  const openingKey = currency === "cash" ? "openingCash" : "openingBank";
  const thuNganKey = currency === "cash" ? "thuNganCash" : "thuNganBank";
  const thuCocKey = currency === "cash" ? "thuCocCash" : "thuCocBank";
  const chiPhiKey = currency === "cash" ? "chiPhiCash" : "chiPhiBank";
  const chiCocKey = currency === "cash" ? "chiCocCash" : "chiCocBank";
  const remittedKey = currency === "cash" ? "remittedCash" : "remittedBank";
  const closingKey = currency === "cash" ? "closingCash" : "closingBank";

  const startEdit = (row, field) => { setEditingCell({ date: row.date, field }); setDraftValue(String(field === "opening" ? row[openingKey] : row[remittedKey])); };
  const cancelEdit = () => { setEditingCell(null); setDraftValue(""); };
  const save = async (date, field) => {
    setSaving(true);
    try {
      if (field === "opening") await onSaveOpening(date, currency, Number(draftValue) || 0);
      else await onSaveRemitted(date, currency, Number(draftValue) || 0);
      setEditingCell(null);
      setDraftValue("");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-0 overflow-hidden mb-5">
      <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-2 flex-wrap">
        <div>
          <p className="font-semibold text-slate-800 text-sm">Sổ quỹ theo ngày</p>
          <p className="text-xs text-slate-400 mt-0.5">Tồn đầu ngày tự động lấy Tồn cuối ngày liền trước — bấm vào số để sửa tay, các ngày sau đó tự tính lại. Đã cộng cả Thu/Chi cọc và trừ tiền nộp.</p>
        </div>
        <div className="flex gap-1">
          {[{ key: "cash", label: "Tiền mặt" }, { key: "bank", label: "Ngân hàng" }].map((t) => (
            <button
              key={t.key}
              type="button"
              onClick={() => setCurrency(t.key)}
              className={`text-xs px-2.5 py-1 rounded-lg border ${currency === t.key ? "bg-sky-700 text-white border-sky-700" : "text-slate-500 border-slate-200 hover:bg-slate-50"}`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>
      {ledger.length === 0 ? <EmptyState icon={Wallet} text="Chọn khoảng ngày để xem sổ quỹ." /> : (
        <div className="overflow-x-auto max-h-96 overflow-y-auto">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-white"><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Ngày</th>
              <th className="px-3 py-2 text-right">Tồn đầu ngày</th>
              <th className="px-3 py-2 text-right">Thu ngân</th>
              <th className="px-3 py-2 text-right">Thu cọc</th>
              <th className="px-3 py-2 text-right">Chi phí</th>
              <th className="px-3 py-2 text-right">Chi cọc</th>
              <th className="px-3 py-2 text-right">Nộp cho cô</th>
              <th className="px-3 py-2 text-right">Tồn cuối ngày</th>
            </tr></thead>
            <tbody>
              {ledger.map((row) => (
                <tr key={row.date} className="border-b border-slate-50 last:border-0">
                  <td className="px-3 py-2 text-slate-500">{fmtDate(row.date)}</td>
                  <td className="px-3 py-2 text-right">
                    {editingCell?.date === row.date && editingCell?.field === "opening" ? (
                      <div className="flex items-center justify-end gap-1">
                        <MoneyField value={draftValue} onChange={setDraftValue} className="w-32 !py-1" />
                        <button type="button" onClick={() => save(row.date, "opening")} disabled={saving} className="text-emerald-600 hover:text-emerald-700 p-1">
                          {saving ? <Loader2 size={14} className="animate-spin" /> : <CheckCircle2 size={14} />}
                        </button>
                        <button type="button" onClick={cancelEdit} className="text-slate-400 hover:text-rose-600 p-1"><X size={14} /></button>
                      </div>
                    ) : (
                      <button type="button" onClick={() => startEdit(row, "opening")} className="inline-flex items-center gap-1 font-medium text-slate-700 hover:text-sky-700" title="Bấm để sửa tay">
                        {fmtMoney(row[openingKey])}
                        <Pencil size={11} className="text-slate-300" />
                      </button>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right text-emerald-700">{fmtMoney(row[thuNganKey])}</td>
                  <td className="px-3 py-2 text-right text-teal-700">{fmtMoney(row[thuCocKey])}</td>
                  <td className="px-3 py-2 text-right text-rose-700">{fmtMoney(row[chiPhiKey])}</td>
                  <td className="px-3 py-2 text-right text-amber-700">{fmtMoney(row[chiCocKey])}</td>
                  <td className="px-3 py-2 text-right">
                    {editingCell?.date === row.date && editingCell?.field === "remitted" ? (
                      <div className="flex items-center justify-end gap-1">
                        <MoneyField value={draftValue} onChange={setDraftValue} className="w-32 !py-1" />
                        <button type="button" onClick={() => save(row.date, "remitted")} disabled={saving} className="text-emerald-600 hover:text-emerald-700 p-1">
                          {saving ? <Loader2 size={14} className="animate-spin" /> : <CheckCircle2 size={14} />}
                        </button>
                        <button type="button" onClick={cancelEdit} className="text-slate-400 hover:text-rose-600 p-1"><X size={14} /></button>
                      </div>
                    ) : (
                      <button type="button" onClick={() => startEdit(row, "remitted")} className="inline-flex items-center gap-1 font-medium text-purple-700 hover:text-purple-800" title="Bấm để nhập số tiền đã nộp">
                        {fmtMoney(row[remittedKey])}
                        <Pencil size={11} className="text-slate-300" />
                      </button>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right font-semibold text-slate-800">{fmtMoney(row[closingKey])}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

// Đối chiếu theo ngày — so sánh Thu ngân báo cáo (tiền mặt+ngân hàng) với Doanh thu theo
// hoá đơn Excel cho TỪNG NGÀY, cảnh báo ngay dòng nào bị lệch để dễ truy đúng ngày phát sinh.
// Đối chiếu theo ngày — so sánh 2 vế: (1) Tồn quỹ dự kiến theo Doanh thu hoá đơn Excel
// (Tồn đầu + Doanh thu hoá đơn + Thu cọc - Chi phí - Chi cọc) và (2) Tồn quỹ thực tế theo
// Thu ngân báo cáo (lấy thẳng từ fundLedger, dùng đúng Thu ngân thay vì Doanh thu hoá đơn).
// Chi phí/Thu cọc/Chi cọc/Tồn đầu dùng chung 1 nguồn dữ liệu cho cả 2 vế nên không lệch.
function buildDailyReconciliation(dates, fundLedger, invoiceRevenue) {
  return dates.map((d, i) => {
    const row = fundLedger[i];
    const thuNganCash = row.thuNganCash;
    const thuNganBank = row.thuNganBank;
    const thuNgan = thuNganCash + thuNganBank;
    const hoaDon = invoiceRevenue.filter((r) => r.invoiceDate === d).reduce((s, r) => s + r.amount, 0);
    const diff = thuNgan - hoaDon;

    return { date: d, thuNganCash, thuNganBank, thuNgan, hoaDon, diff };
  });
}

function DailyReconciliationTable({ rows }) {
  const mismatchCount = rows.filter((r) => r.diff !== 0 && (r.thuNgan > 0 || r.hoaDon > 0)).length;

  return (
    <Card className="p-0 overflow-hidden mb-5">
      <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-2 flex-wrap">
        <div>
          <p className="font-semibold text-slate-800 text-sm">Đối chiếu Doanh thu theo ngày</p>
          <p className="text-xs text-slate-400 mt-0.5">So sánh Doanh thu Thu ngân báo cáo (tách Tiền mặt/Ngân hàng) với Doanh thu theo hoá đơn Excel cho từng ngày.</p>
        </div>
        {mismatchCount > 0 && (
          <span className="text-xs px-2 py-1 rounded-lg bg-rose-50 text-rose-700 border border-rose-200 font-medium flex items-center gap-1">
            <AlertTriangle size={13} /> {mismatchCount} ngày lệch số liệu
          </span>
        )}
      </div>
      {rows.length === 0 ? <EmptyState icon={AlertTriangle} text="Chọn khoảng ngày để đối chiếu." /> : (
        <div className="overflow-x-auto max-h-96 overflow-y-auto">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-white"><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Ngày</th>
              <th className="px-3 py-2 text-right">Doanh thu hoá đơn</th>
              <th className="px-3 py-2 text-right">Thu ngân - Tiền mặt</th>
              <th className="px-3 py-2 text-right">Thu ngân - Ngân hàng</th>
              <th className="px-3 py-2 text-right">Tổng Thu ngân</th>
              <th className="px-3 py-2 text-right">Chênh lệch</th>
              <th className="px-3 py-2">Trạng thái</th>
            </tr></thead>
            <tbody>
              {rows.map((r) => {
                const noData = r.thuNgan === 0 && r.hoaDon === 0;
                const rowClass = noData ? "" : r.diff === 0 ? "" : r.diff > 0 ? "bg-amber-50/60" : "bg-rose-50/60";
                return (
                  <tr key={r.date} className={`border-b border-slate-50 last:border-0 ${rowClass}`}>
                    <td className="px-3 py-2 text-slate-500">{fmtDate(r.date)}</td>
                    <td className="px-3 py-2 text-right text-slate-700">{fmtMoney(r.hoaDon)}</td>
                    <td className="px-3 py-2 text-right text-slate-700">{fmtMoney(r.thuNganCash)}</td>
                    <td className="px-3 py-2 text-right text-slate-700">{fmtMoney(r.thuNganBank)}</td>
                    <td className="px-3 py-2 text-right font-medium text-slate-800">{fmtMoney(r.thuNgan)}</td>
                    <td className={`px-3 py-2 text-right font-semibold ${r.diff === 0 ? "text-slate-400" : r.diff > 0 ? "text-amber-700" : "text-rose-700"}`}>
                      {r.diff > 0 ? "+" : ""}{fmtMoney(r.diff)}
                    </td>
                    <td className="px-3 py-2">
                      {noData ? (
                        <span className="text-xs text-slate-300">Chưa có dữ liệu</span>
                      ) : r.diff === 0 ? (
                        <span className="inline-flex items-center gap-1 text-xs font-medium text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-lg">
                          <CheckCircle2 size={12} /> Khớp
                        </span>
                      ) : (
                        <span className={`inline-flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-lg border ${r.diff > 0 ? "text-amber-700 bg-amber-50 border-amber-200" : "text-rose-700 bg-rose-50 border-rose-200"}`}>
                          <AlertTriangle size={12} /> Chênh lệch
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

function QuyModule({ data, currentUser, onBulkImportInvoiceRevenue, onUpsertOpeningBalance, onUpsertRemittedAmount, onSetThuNganEditEnabled }) {
  const [from, setFrom] = useState(FUND_START_DATE);
  const [to, setTo] = useState(todayISO());
  const [viewMode, setViewMode] = useState("ngay"); // "ngay" | "tongthe"
  const thuNganEditEnabled = data.settings?.thu_ngan_edit_enabled !== "false";
  const [togglingEdit, setTogglingEdit] = useState(false);
  const handleToggleEdit = async () => {
    setTogglingEdit(true);
    try { await onSetThuNganEditEnabled(!thuNganEditEnabled); } finally { setTogglingEdit(false); }
  };

  // Doanh thu bán hàng theo hoá đơn — import trực tiếp từ file "Bảng kê hoá đơn" POS.
  const invoiceRevenueInRange = data.invoiceRevenue.filter((r) => r.invoiceDate >= from && r.invoiceDate <= to);
  const totalInvoiceRevenue = invoiceRevenueInRange.reduce((s, r) => s + r.amount, 0);

  // Sổ quỹ theo ngày (Tiền mặt & Ngân hàng riêng): Tồn đầu ngày tự động = Tồn cuối ngày liền
  // trước, có thể ghi đè tay từng loại tiền (lưu vào bảng fund_daily_balance). Đi qua đủ 4 dòng
  // tiền mỗi ngày: Thu ngân, Thu cọc, Chi phí, Chi cọc (deposits/expense_records).
  const ledgerDates = enumerateDatesISO(from, to);
  const overridesMap = Object.fromEntries(data.fundDailyBalances.map((b) => [b.date, { cash: b.openingBalance, bank: b.openingBalanceBank }]));
  const remittedMap = Object.fromEntries(data.fundDailyBalances.map((b) => [b.date, { cash: b.remittedOwnerCash, bank: b.remittedOwnerBank }]));
  const fundLedger = buildFundLedger(ledgerDates, data.cashierReceipts, data.expenseRecords, data.deposits, overridesMap, remittedMap);
  const openingBalanceNum = fundLedger.length > 0 ? fundLedger[0].openingCash + fundLedger[0].openingBank : 0;
  const dailyReconciliation = buildDailyReconciliation(ledgerDates, fundLedger, data.invoiceRevenue);

  // Tổng hợp theo khoảng đang lọc — tách riêng Tiền mặt / Ngân hàng cho từng chỉ tiêu.
  const sumField = (key) => fundLedger.reduce((s, r) => s + r[key], 0);
  const totalThuNganCash = sumField("thuNganCash");
  const totalThuNganBank = sumField("thuNganBank");
  const totalThuNgan = totalThuNganCash + totalThuNganBank;
  const totalThuCocCash = sumField("thuCocCash");
  const totalThuCocBank = sumField("thuCocBank");
  const totalChiPhiCash = sumField("chiPhiCash");
  const totalChiPhiBank = sumField("chiPhiBank");
  const totalChiCocCash = sumField("chiCocCash");
  const totalChiCocBank = sumField("chiCocBank");
  const totalRemittedCash = sumField("remittedCash");
  const totalRemittedBank = sumField("remittedBank");
  const closingCash = fundLedger.length > 0 ? fundLedger[fundLedger.length - 1].closingCash : 0;
  const closingBank = fundLedger.length > 0 ? fundLedger[fundLedger.length - 1].closingBank : 0;

  // Đối chiếu tổng thể cả khoảng: Tồn quỹ theo hoá đơn Excel vs theo Thu ngân báo cáo.
  // Trừ cả "Nộp cho cô" ở vế hoá đơn để nhất quán với closingCash/closingBank (đã trừ khoản này).
  const tonQuyHoaDonTongThe = openingBalanceNum + totalInvoiceRevenue + totalThuCocCash + totalThuCocBank - totalChiPhiCash - totalChiPhiBank - totalChiCocCash - totalChiCocBank - totalRemittedCash - totalRemittedBank;
  const tonQuyThuNganTongThe = closingCash + closingBank;
  const diffTongThe = tonQuyThuNganTongThe - tonQuyHoaDonTongThe;

  return (
    <div>
      <SectionTitle icon={Wallet} title="Quỹ" subtitle="Đối chiếu Doanh thu (hoá đơn Excel & Thu ngân báo cáo) và Tồn quỹ mỗi ngày" />

      <Card className="p-4 sm:p-5 mb-5">
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div>
            <p className="font-semibold text-slate-800 text-sm flex items-center gap-2">
              {thuNganEditEnabled ? <Unlock size={16} className="text-emerald-600" /> : <Lock size={16} className="text-rose-600" />}
              Quyền sửa/xoá phiếu thu, phiếu chi của Thu ngân
            </p>
            <p className="text-xs text-slate-400 mt-0.5">
              {thuNganEditEnabled
                ? "Đang MỞ — tài khoản Thu ngân tự sửa/xoá được phiếu thu & phiếu chi do chính mình tạo."
                : "Đang KHOÁ — tài khoản Thu ngân chỉ được thêm mới, không tự sửa/xoá được nữa."}
            </p>
          </div>
          <button
            type="button"
            onClick={handleToggleEdit}
            disabled={togglingEdit}
            className={`relative inline-flex items-center h-7 w-12 rounded-full transition-colors shrink-0 disabled:opacity-50 ${thuNganEditEnabled ? "bg-emerald-500" : "bg-slate-300"}`}
            title={thuNganEditEnabled ? "Bấm để khoá" : "Bấm để mở khoá"}
          >
            <span className={`inline-block h-5 w-5 transform rounded-full bg-white shadow transition-transform ${thuNganEditEnabled ? "translate-x-6" : "translate-x-1"}`} />
          </button>
        </div>
      </Card>

      <Card className="p-4 sm:p-5 mb-5">
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
      </Card>

      <div className="grid grid-cols-2 gap-3 mb-5">
        <MetricCard label="Doanh thu bán hàng theo hoá đơn" value={fmtMoney(totalInvoiceRevenue)} icon={FileText} accent="indigo" />
        <MetricCard label="Tồn quỹ cuối kỳ (theo Thu ngân)" value={fmtMoney(closingCash + closingBank)} icon={Wallet} accent={closingCash + closingBank >= 0 ? "teal" : "rose"} />
      </div>

      {/* Phân loại Tổng thu / Tổng chi / Cọc theo Tiền mặt & Ngân hàng */}
      <Card className="p-4 sm:p-5 mb-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Tổng thu, tổng chi &amp; cọc theo hình thức</p>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Chỉ tiêu</th>
              <th className="px-3 py-2 text-right">Tiền mặt</th>
              <th className="px-3 py-2 text-right">Ngân hàng</th>
              <th className="px-3 py-2 text-right">Tổng</th>
            </tr></thead>
            <tbody>
              <tr className="border-b border-slate-50">
                <td className="px-3 py-2 text-emerald-700 font-medium">Thu ngân báo cáo</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalThuNganCash)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalThuNganBank)}</td>
                <td className="px-3 py-2 text-right font-semibold">{fmtMoney(totalThuNganCash + totalThuNganBank)}</td>
              </tr>
              <tr className="border-b border-slate-50">
                <td className="px-3 py-2 text-teal-700 font-medium">Thu cọc</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalThuCocCash)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalThuCocBank)}</td>
                <td className="px-3 py-2 text-right font-semibold">{fmtMoney(totalThuCocCash + totalThuCocBank)}</td>
              </tr>
              <tr className="border-b border-slate-50">
                <td className="px-3 py-2 text-rose-700 font-medium">Chi phí</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalChiPhiCash)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalChiPhiBank)}</td>
                <td className="px-3 py-2 text-right font-semibold">{fmtMoney(totalChiPhiCash + totalChiPhiBank)}</td>
              </tr>
              <tr className="border-b border-slate-50">
                <td className="px-3 py-2 text-amber-700 font-medium">Chi cọc</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalChiCocCash)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalChiCocBank)}</td>
                <td className="px-3 py-2 text-right font-semibold">{fmtMoney(totalChiCocCash + totalChiCocBank)}</td>
              </tr>
              <tr>
                <td className="px-3 py-2 text-purple-700 font-medium">Nộp cho cô</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalRemittedCash)}</td>
                <td className="px-3 py-2 text-right">{fmtMoney(totalRemittedBank)}</td>
                <td className="px-3 py-2 text-right font-semibold">{fmtMoney(totalRemittedCash + totalRemittedBank)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Card>

      <InvoiceRevenueImportForm onImport={onBulkImportInvoiceRevenue} />

      <FundLedgerTable ledger={fundLedger} onSaveOpening={onUpsertOpeningBalance} onSaveRemitted={onUpsertRemittedAmount} />

      {/* Đối chiếu Doanh thu & Tồn quỹ — xem theo từng ngày hoặc tổng thể cả khoảng đang lọc */}
      <Card className="p-4 sm:p-5 mb-2">
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <p className="font-semibold text-slate-800 text-sm">Đối chiếu Doanh thu &amp; Tồn quỹ</p>
          <div className="flex gap-1">
            {[{ key: "ngay", label: "Theo ngày" }, { key: "tongthe", label: "Tổng thể" }].map((t) => (
              <button
                key={t.key}
                type="button"
                onClick={() => setViewMode(t.key)}
                className={`text-xs px-2.5 py-1 rounded-lg border ${viewMode === t.key ? "bg-sky-700 text-white border-sky-700" : "text-slate-500 border-slate-200 hover:bg-slate-50"}`}
              >
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </Card>

      {viewMode === "ngay" ? (
        <DailyReconciliationTable rows={dailyReconciliation} />
      ) : (
        <Card className="p-4 sm:p-5 mb-5">
          <p className="text-xs text-slate-400 mb-3">Tổng thể từ {fmtDate(from)} đến {fmtDate(to)} — Tồn đầu kỳ {fmtMoney(openingBalanceNum)}.</p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mb-3">
            <div className="bg-indigo-50 border border-indigo-100 rounded-xl px-3 py-2">
              <p className="text-xs text-indigo-700">Tồn quỹ theo hoá đơn Excel</p>
              <p className="text-lg font-semibold text-indigo-800">{fmtMoney(tonQuyHoaDonTongThe)}</p>
            </div>
            <div className="bg-teal-50 border border-teal-100 rounded-xl px-3 py-2">
              <p className="text-xs text-teal-700">Tồn quỹ theo Thu ngân báo cáo</p>
              <p className="text-lg font-semibold text-teal-800">{fmtMoney(tonQuyThuNganTongThe)}</p>
            </div>
          </div>
          {totalInvoiceRevenue === 0 && totalThuNgan === 0 ? (
            <p className="text-xs text-slate-400">Chưa có đủ dữ liệu để đối chiếu trong khoảng thời gian này.</p>
          ) : diffTongThe === 0 ? (
            <div className="flex items-center gap-2 text-sm text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-2 font-medium">
              <CheckCircle2 size={16} /> Khớp — Tồn quỹ theo hoá đơn Excel và theo Thu ngân báo cáo bằng nhau.
            </div>
          ) : (
            <div className="flex items-center gap-2 text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded-xl px-3 py-2 font-medium">
              <AlertTriangle size={16} /> Chênh lệch {fmtMoney(Math.abs(diffTongThe))} — {diffTongThe > 0 ? "Thu ngân báo cao hơn hoá đơn Excel" : "Thu ngân báo thấp hơn hoá đơn Excel"}.
            </div>
          )}
        </Card>
      )}
    </div>
  );
}

// Chi phí "Bảo trì & vật tư" (mua sắm CCDC, vật tư, sửa chữa, vật dụng tiêu hao) —
// bảng nhiều dòng tự do (không gắn danh mục sản phẩm), giống kiểu Nhập hàng.
function BaoTriVatTuForm({ currentUser, onSubmit }) {
  const [lines, setLines] = useState([{ key: Math.random().toString(36).slice(2), itemName: "", quantity: "", unitPrice: "" }]);
  const [expenseDate, setExpenseDate] = useState(todayISO());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [paymentMethod, setPaymentMethod] = useState("tien_mat");

  const updateLine = (key, patch) => setLines((prev) => prev.map((l) => (l.key === key ? { ...l, ...patch } : l)));
  const addRow = () => setLines((prev) => [...prev, { key: Math.random().toString(36).slice(2), itemName: "", quantity: "", unitPrice: "" }]);
  const removeRow = (key) => setLines((prev) => (prev.length > 1 ? prev.filter((l) => l.key !== key) : prev));

  const validLines = lines.filter((l) => l.itemName.trim() && Number(l.quantity) > 0 && Number(l.unitPrice) >= 0);
  const grandTotal = validLines.reduce((s, l) => s + Number(l.quantity) * Number(l.unitPrice), 0);

  const submit = async () => {
    if (validLines.length === 0) { setError("Cần ít nhất 1 dòng đủ Tên khoản chi, Số lượng, Đơn giá."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({
        category: "bao_tri_vat_tu", expenseDate, paymentMethod,
        lines: validLines.map((l) => ({ itemName: l.itemName.trim(), quantity: Number(l.quantity), unitPrice: Number(l.unitPrice), amount: Number(l.quantity) * Number(l.unitPrice) })),
      });
      setLines([{ key: Math.random().toString(36).slice(2), itemName: "", quantity: "", unitPrice: "" }]);
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Chi phí bảo trì & vật tư (CCDC, vật tư, sửa chữa, vật dụng tiêu hao...)</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Ngày phát sinh" type="date" value={expenseDate} onChange={(e) => setExpenseDate(e.target.value)} />
        <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
          <option value="tien_mat">Tiền mặt</option>
          <option value="chuyen_khoan">Chuyển khoản</option>
        </SelectField>
      </div>
      <div className="overflow-x-auto rounded-xl border border-slate-200 mb-3">
        <table className="w-full text-sm min-w-[640px]">
          <thead>
            <tr className="text-left text-xs text-slate-500 bg-slate-50 border-b border-slate-200">
              <th className="px-2 py-2">Tên khoản chi (CCDC / vật tư / sửa chữa...)</th>
              <th className="px-2 py-2 w-28">Số lượng</th>
              <th className="px-2 py-2 w-32">Đơn giá</th>
              <th className="px-2 py-2 w-32 text-right">Thành tiền</th>
              <th className="px-2 py-2 w-10"></th>
            </tr>
          </thead>
          <tbody>
            {lines.map((l) => {
              const rowTotal = (Number(l.quantity) || 0) * (Number(l.unitPrice) || 0);
              return (
                <tr key={l.key} className="border-b border-slate-100 last:border-0">
                  <td className="px-2 py-1.5">
                    <input value={l.itemName} onChange={(e) => updateLine(l.key, { itemName: e.target.value })} placeholder="VD: Sửa máy hút mùi, mua khay inox..." className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40" />
                  </td>
                  <td className="px-2 py-1.5">
                    <MoneyField allowDecimal value={l.quantity} onChange={(v) => updateLine(l.key, { quantity: v })} className="!py-1.5" />
                  </td>
                  <td className="px-2 py-1.5">
                    <MoneyField value={l.unitPrice} onChange={(v) => updateLine(l.key, { unitPrice: v })} className="!py-1.5" />
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium text-slate-700">{rowTotal > 0 ? fmtMoney(rowTotal) : "—"}</td>
                  <td className="px-2 py-1.5 text-center">
                    <button type="button" onClick={() => removeRow(l.key)} className="text-slate-400 hover:text-rose-600"><X size={15} /></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="bg-sky-50/60">
              <td colSpan={3} className="px-2 py-2 text-right text-sky-700 font-medium">Tổng chi phí</td>
              <td className="px-2 py-2 text-right font-semibold text-sky-800">{fmtMoney(grandTotal)}</td>
              <td></td>
            </tr>
          </tfoot>
        </table>
      </div>
      <GhostButton type="button" onClick={addRow} className="mb-3"><Plus size={14} /> Thêm dòng</GhostButton>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Receipt size={15} />} Lưu chi phí</PrimaryButton>
    </Card>
  );
}

// Chi phí "Vận hành" / "Marketing & bán hàng" — danh sách khoản chi cố định, mỗi
// khoản chỉ cần điền số tiền (không cần số lượng/đơn giá).
function PresetExpenseForm({ category, onSubmit }) {
  const meta = EXPENSE_CATEGORY_META[category];
  const [amounts, setAmounts] = useState(() => Object.fromEntries(meta.presetItems.map((n) => [n, ""])));
  const [expenseDate, setExpenseDate] = useState(todayISO());
  const [paymentMethod, setPaymentMethod] = useState("tien_mat");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const filled = meta.presetItems.filter((n) => Number(amounts[n]) > 0);
  const grandTotal = filled.reduce((s, n) => s + Number(amounts[n]), 0);

  const submit = async () => {
    if (filled.length === 0) { setError("Vui lòng nhập ít nhất 1 khoản chi phí."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({
        category, expenseDate, paymentMethod,
        lines: filled.map((n) => ({ itemName: n, amount: Number(amounts[n]) })),
      });
      setAmounts(Object.fromEntries(meta.presetItems.map((n) => [n, ""])));
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">{meta.label}</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Ngày phát sinh" type="date" value={expenseDate} onChange={(e) => setExpenseDate(e.target.value)} />
        <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
          <option value="tien_mat">Tiền mặt</option>
          <option value="chuyen_khoan">Chuyển khoản</option>
        </SelectField>
      </div>
      <div className="space-y-2 mb-4">
        {meta.presetItems.map((n) => (
          <div key={n} className="flex items-center gap-3">
            <span className="text-sm text-slate-600 w-32 shrink-0">{n}</span>
            <MoneyField value={amounts[n]} onChange={(v) => setAmounts((prev) => ({ ...prev, [n]: v }))} className="flex-1" />
            <span className="text-xs text-slate-400 w-8">đ</span>
          </div>
        ))}
      </div>
      <div className="bg-sky-50 rounded-xl px-3 py-2 flex items-center justify-between text-sm mb-4">
        <span className="text-sky-700">Tổng chi phí</span>
        <span className="font-semibold text-sky-800">{fmtMoney(grandTotal)}</span>
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Receipt size={15} />} Lưu chi phí</PrimaryButton>
    </Card>
  );
}

// Chi phí "Khác" — các khoản không thuộc 4 nhóm còn lại, tự đặt tên khoản chi.
function OtherExpenseForm({ onSubmit }) {
  const [lines, setLines] = useState([{ key: Math.random().toString(36).slice(2), itemName: "", amount: "" }]);
  const [expenseDate, setExpenseDate] = useState(todayISO());
  const [paymentMethod, setPaymentMethod] = useState("tien_mat");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const updateLine = (key, patch) => setLines((prev) => prev.map((l) => (l.key === key ? { ...l, ...patch } : l)));
  const addRow = () => setLines((prev) => [...prev, { key: Math.random().toString(36).slice(2), itemName: "", amount: "" }]);
  const removeRow = (key) => setLines((prev) => (prev.length > 1 ? prev.filter((l) => l.key !== key) : prev));

  const validLines = lines.filter((l) => l.itemName.trim() && Number(l.amount) > 0);
  const grandTotal = validLines.reduce((s, l) => s + Number(l.amount), 0);

  const submit = async () => {
    if (validLines.length === 0) { setError("Cần ít nhất 1 dòng đủ Tên khoản chi và Số tiền."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({
        category: "khac", expenseDate, paymentMethod,
        lines: validLines.map((l) => ({ itemName: l.itemName.trim(), amount: Number(l.amount) })),
      });
      setLines([{ key: Math.random().toString(36).slice(2), itemName: "", amount: "" }]);
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Chi phí khác</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <TextField label="Ngày phát sinh" type="date" value={expenseDate} onChange={(e) => setExpenseDate(e.target.value)} />
        <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
          <option value="tien_mat">Tiền mặt</option>
          <option value="chuyen_khoan">Chuyển khoản</option>
        </SelectField>
      </div>
      <div className="space-y-2 mb-4">
        {lines.map((l) => (
          <div key={l.key} className="flex items-center gap-2">
            <input value={l.itemName} onChange={(e) => updateLine(l.key, { itemName: e.target.value })} placeholder="Tên khoản chi..." className="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40" />
            <MoneyField value={l.amount} onChange={(v) => updateLine(l.key, { amount: v })} placeholder="Số tiền" className="w-40" />
            <button type="button" onClick={() => removeRow(l.key)} className="text-slate-400 hover:text-rose-600 shrink-0"><X size={16} /></button>
          </div>
        ))}
      </div>
      <GhostButton type="button" onClick={addRow} className="mb-3"><Plus size={14} /> Thêm dòng</GhostButton>
      <div className="bg-sky-50 rounded-xl px-3 py-2 flex items-center justify-between text-sm mb-4">
        <span className="text-sky-700">Tổng chi phí</span>
        <span className="font-semibold text-sky-800">{fmtMoney(grandTotal)}</span>
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Receipt size={15} />} Lưu chi phí</PrimaryButton>
    </Card>
  );
}

// Sửa 1 khoản chi phí — dùng chung cho mọi nhóm chi phí (kể cả khoản do
// Thu ngân gửi lên, vì tất cả đều nằm chung bảng expense_records).
function EditExpenseModal({ expense, onSave, onClose }) {
  const [category, setCategory] = useState(expense.category);
  const [itemName, setItemName] = useState(expense.itemName);
  const [expenseDate, setExpenseDate] = useState(expense.expenseDate);
  const [amount, setAmount] = useState(String(expense.amount ?? ""));
  const [quantity, setQuantity] = useState(expense.quantity === null || expense.quantity === undefined ? "" : String(expense.quantity));
  const [unitPrice, setUnitPrice] = useState(expense.unitPrice === null || expense.unitPrice === undefined ? "" : String(expense.unitPrice));
  const [paymentMethod, setPaymentMethod] = useState(expense.paymentMethod || "tien_mat");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!itemName.trim()) { setError("Vui lòng nhập tên khoản chi."); return; }
    if (!(Number(amount) > 0)) { setError("Vui lòng nhập số tiền hợp lệ."); return; }
    setError(""); setSaving(true);
    try {
      await onSave(expense.id, {
        category, itemName: itemName.trim(), expenseDate,
        amount: Number(amount),
        quantity: quantity === "" ? null : Number(quantity),
        unitPrice: unitPrice === "" ? null : Number(unitPrice),
        paymentMethod,
      });
      onClose();
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 flex items-center gap-2"><Pencil size={17} className="text-sky-700" /> Sửa khoản chi phí</p>
          <button onClick={onClose}><X size={18} className="text-slate-400" /></button>
        </div>
        <div className="space-y-3">
          <SelectField label="Nhóm chi phí" value={category} onChange={(e) => setCategory(e.target.value)}>
            {EXPENSE_CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
          </SelectField>
          <TextField label="Tên khoản chi" value={itemName} onChange={(e) => setItemName(e.target.value)} />
          <TextField label="Ngày phát sinh" type="date" value={expenseDate} onChange={(e) => setExpenseDate(e.target.value)} />
          <div className="grid grid-cols-2 gap-3">
            <MoneyField label="Số lượng (nếu có)" value={quantity} onChange={setQuantity} />
            <MoneyField label="Đơn giá (nếu có)" value={unitPrice} onChange={setUnitPrice} />
          </div>
          <MoneyField label="Số tiền" value={amount} onChange={setAmount} />
          <SelectField label="Hình thức" value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
            <option value="tien_mat">Tiền mặt</option>
            <option value="chuyen_khoan">Chuyển khoản</option>
          </SelectField>
          {error && <p className="text-sm text-rose-600 flex items-center gap-1.5"><AlertTriangle size={14} /> {error}</p>}
          <div className="flex gap-2">
            <PrimaryButton type="button" onClick={submit} disabled={saving}>{saving ? "Đang lưu..." : "Lưu thay đổi"}</PrimaryButton>
            <GhostButton type="button" onClick={onClose}>Huỷ</GhostButton>
          </div>
        </div>
      </div>
    </div>
  );
}

function ExpenseList({ data, currentUser, onDelete, onUpdate, filterCategory, from, to }) {
  const [paymentFilter, setPaymentFilter] = useState("all");
  const rows = data.expenseRecords
    .filter((r) => !filterCategory || r.category === filterCategory)
    .filter((r) => (!from || r.expenseDate >= from) && (!to || r.expenseDate <= to))
    .filter((r) => paymentFilter === "all" || r.paymentMethod === paymentFilter)
    .slice(0, 30);
  const isQuanLy = currentUser?.role === "quan_ly" || currentUser?.role === "bao_cao";
  const [deleting, setDeleting] = useState(null);
  const [editing, setEditing] = useState(null);
  const [changingCategory, setChangingCategory] = useState(null);

  const handleDelete = async (id) => {
    if (!window.confirm("Xoá khoản chi phí này? Không thể hoàn tác.")) return;
    setDeleting(id);
    try { await onDelete(id); } finally { setDeleting(null); }
  };

  const handleCategoryChange = async (r, newCategory) => {
    if (newCategory === r.category) return;
    setChangingCategory(r.id);
    try {
      await onUpdate(r.id, {
        category: newCategory, itemName: r.itemName, expenseDate: r.expenseDate,
        amount: r.amount, quantity: r.quantity, unitPrice: r.unitPrice, paymentMethod: r.paymentMethod,
      });
    } finally {
      setChangingCategory(null);
    }
  };

  return (
    <Card className="p-0 overflow-hidden">
      <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-2 flex-wrap">
        <div>
          <p className="font-semibold text-slate-800 text-sm">
            {filterCategory ? `Chi phí nhóm "${EXPENSE_CATEGORY_META[filterCategory]?.label || filterCategory}"` : "Chi phí ghi nhận gần đây"}
          </p>
          {isQuanLy && <p className="text-xs text-slate-400 mt-0.5">Có thể đổi nhanh "Loại chi phí" ngay tại đây nếu tự động phân nhóm bị sai.</p>}
        </div>
        <div className="flex gap-1">
          {[{ key: "all", label: "Tất cả" }, { key: "tien_mat", label: "Tiền mặt" }, { key: "chuyen_khoan", label: "Chuyển khoản" }].map((t) => (
            <button
              key={t.key}
              type="button"
              onClick={() => setPaymentFilter(t.key)}
              className={`text-xs px-2.5 py-1 rounded-lg border ${paymentFilter === t.key ? "bg-sky-700 text-white border-sky-700" : "text-slate-500 border-slate-200 hover:bg-slate-50"}`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>
      {rows.length === 0 ? <EmptyState icon={Receipt} text="Chưa có khoản chi phí nào ở nhóm này." /> : (
        <div className="divide-y divide-slate-100">
          {rows.map((r) => (
            <div key={r.id} className="px-4 py-3 flex items-center justify-between gap-3">
              <div className="min-w-0 flex-1">
                <p className="text-sm font-medium text-slate-800 truncate">{r.itemName}</p>
                <div className="flex items-center gap-2 mt-1 flex-wrap">
                  {isQuanLy ? (
                    <select
                      value={r.category}
                      onChange={(e) => handleCategoryChange(r, e.target.value)}
                      disabled={changingCategory === r.id}
                      className="text-xs border border-slate-200 rounded-lg px-1.5 py-0.5 text-slate-600 bg-slate-50 focus:outline-none focus:ring-1 focus:ring-sky-400 disabled:opacity-50"
                      title="Đổi nhanh loại chi phí"
                    >
                      {EXPENSE_CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
                    </select>
                  ) : (
                    <span className="text-xs text-slate-400">{EXPENSE_CATEGORY_META[r.category]?.label || r.category}</span>
                  )}
                  <p className="text-xs text-slate-400">· {fmtDate(r.expenseDate)}</p>
                  <span className={`text-xs px-1.5 py-0.5 rounded-lg border ${EXPENSE_PAYMENT_METHOD_META[r.paymentMethod]?.color}`}>{EXPENSE_PAYMENT_METHOD_META[r.paymentMethod]?.label || r.paymentMethod}</span>
                  {changingCategory === r.id && <Loader2 size={11} className="animate-spin text-slate-400" />}
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <p className="text-sm font-semibold text-slate-700">{fmtMoney(r.amount)}</p>
                {isQuanLy && (
                  <>
                    <button
                      type="button"
                      onClick={() => setEditing(r)}
                      className="text-slate-400 hover:text-sky-700 p-1"
                      title="Sửa khoản chi phí này"
                    >
                      <Pencil size={14} />
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDelete(r.id)}
                      disabled={deleting === r.id}
                      className="text-slate-400 hover:text-rose-600 p-1"
                      title="Xoá khoản chi phí này"
                    >
                      {deleting === r.id ? <Loader2 size={14} className="animate-spin" /> : <Trash2 size={14} />}
                    </button>
                  </>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
      {editing && (
        <EditExpenseModal
          expense={editing}
          onSave={async (id, patch) => { await onUpdate(id, patch); }}
          onClose={() => setEditing(null)}
        />
      )}
    </Card>
  );
}

// Bảng tổng hợp Chi phí: tổng theo từng loại (nvl gộp cả nhập kho Excel + phiếu chi lẻ) +
// chi tiết theo từng khoản mục (gộp các dòng cùng tên, sắp xếp theo tổng tiền giảm dần).
function ExpenseSummaryTable({ data, from, to }) {
  const allExpense = data.expenseRecords.filter((r) => (!from || r.expenseDate >= from) && (!to || r.expenseDate <= to));
  const importInRange = data.importRecords.filter((r) => (!from || r.importDate >= from) && (!to || r.importDate <= to));

  const totalsByCategory = EXPENSE_CATEGORIES.map((c) => {
    const rowsInCat = allExpense.filter((r) => r.category === c.key);
    let total = rowsInCat.reduce((s, r) => s + r.amount, 0);
    let count = rowsInCat.length;
    if (c.key === "nvl") {
      total += importInRange.reduce((s, r) => s + r.totalAmount, 0);
      count += importInRange.length;
    }
    return { key: c.key, label: c.label, total, count };
  });
  const grandTotal = totalsByCategory.reduce((s, c) => s + c.total, 0);

  return (
    <Card className="p-0 overflow-hidden mb-5">
      <div className="p-4 border-b border-slate-100">
        <p className="font-semibold text-slate-800 text-sm">Tổng hợp chi phí theo loại</p>
        <p className="text-xs text-slate-400 mt-0.5">{from || to ? `Từ ${from ? fmtDate(from) : "…"} đến ${to ? fmtDate(to) : "…"}` : "Toàn bộ dữ liệu"}</p>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-4 py-2">Loại chi phí</th>
              <th className="px-3 py-2 text-right">Số khoản</th>
              <th className="px-3 py-2 text-right">Tổng tiền</th>
              <th className="px-4 py-2 text-right">Tỷ trọng</th>
            </tr>
          </thead>
          <tbody>
            {totalsByCategory.map((c) => (
              <tr key={c.key} className="border-b border-slate-50 last:border-0">
                <td className="px-4 py-2 text-slate-700">{c.label}</td>
                <td className="px-3 py-2 text-right text-slate-500">{fmtNumber(c.count)}</td>
                <td className="px-3 py-2 text-right font-medium text-slate-800">{fmtMoney(c.total)}</td>
                <td className="px-4 py-2 text-right text-slate-500">{grandTotal > 0 ? `${((c.total / grandTotal) * 100).toFixed(1)}%` : "—"}</td>
              </tr>
            ))}
            <tr className="bg-slate-50">
              <td className="px-4 py-2 font-semibold text-slate-800">Tổng cộng</td>
              <td className="px-3 py-2 text-right font-semibold text-slate-800">{fmtNumber(totalsByCategory.reduce((s, c) => s + c.count, 0))}</td>
              <td className="px-3 py-2 text-right font-semibold text-slate-800">{fmtMoney(grandTotal)}</td>
              <td className="px-4 py-2 text-right font-semibold text-slate-800">100%</td>
            </tr>
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function ChiPhiModule({ data, currentUser, onSubmitExpense, onSubmitImport, onDeleteExpense, onUpdateExpense }) {
  const [category, setCategory] = useState("van_hanh");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const meta = EXPENSE_CATEGORY_META[category];

  // Công nợ nguyên vật liệu: các phiếu nhập NVL có tình trạng thanh toán = Công nợ.
  const nvlRows = data.importRecords.filter((r) => (!from || r.importDate >= from) && (!to || r.importDate <= to)).slice(0, 30);
  // NVL ghi qua Phiếu chi (Thu ngân mua lẻ thịt/cá/tôm...), luôn coi là đã thanh toán (không có khái niệm công nợ ở Phiếu chi).
  const nvlExpenseRows = data.expenseRecords.filter((r) => r.category === "nvl" && (!from || r.expenseDate >= from) && (!to || r.expenseDate <= to)).slice(0, 30);
  const nvlExpenseTotal = nvlExpenseRows.reduce((s, r) => s + r.amount, 0);
  const nvlTotal = nvlRows.reduce((s, r) => s + r.totalAmount, 0) + nvlExpenseTotal;
  const nvlDebt = nvlRows.filter((r) => r.paymentType === "cong_no").reduce((s, r) => s + r.totalAmount, 0);

  return (
    <div>
      <SectionTitle icon={Receipt} title="Chi phí" subtitle="Ghi nhận toàn bộ chi phí phát sinh hàng tháng" />
      <Card className="p-4 sm:p-5 mb-5">
        <p className="text-xs font-medium text-slate-500 mb-2">Lọc theo ngày (áp dụng cho bảng tổng hợp và danh sách chi phí bên dưới)</p>
        <div className="grid grid-cols-2 gap-3">
          <TextField label="Từ ngày" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </div>
        {(from || to) && (
          <button type="button" onClick={() => { setFrom(""); setTo(""); }} className="text-xs text-sky-700 hover:underline mt-2">
            Bỏ lọc ngày
          </button>
        )}
      </Card>
      <ExpenseSummaryTable data={data} from={from} to={to} />
      <Card className="p-4 sm:p-5 mb-5">
        <SelectField label="Loại chi phí" value={category} onChange={(e) => setCategory(e.target.value)}>
          {EXPENSE_CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
        </SelectField>
      </Card>

      {category === "nvl" && (
        <>
          <div className="mb-4 flex items-center gap-2 text-sm text-sky-700 bg-sky-50 border border-sky-200 rounded-xl px-3 py-2">
            <ArrowDownCircle size={15} /> Chi phí nguyên vật liệu chủ yếu ghi nhận qua màn "Nhập hàng" (file Excel). Ngoài ra các khoản mua lẻ (VD Thu ngân mua thịt/cá/tôm ngoài chợ) có thể ghi qua "Phiếu chi" chọn nhóm NVL — số tiền sẽ tự cộng vào Tổng giá trị NVL bên dưới.
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
            <div className="bg-slate-50 rounded-xl px-3 py-2">
              <p className="text-xs text-slate-500">Tổng giá trị NVL (nhập kho + phiếu chi)</p>
              <p className="text-sm font-semibold text-slate-700">{fmtMoney(nvlTotal)}</p>
            </div>
            <div className="bg-amber-50 rounded-xl px-3 py-2">
              <p className="text-xs text-amber-700">Công nợ</p>
              <p className="text-sm font-semibold text-amber-800">{fmtMoney(nvlDebt)}</p>
            </div>
            <div className="bg-emerald-50 rounded-xl px-3 py-2">
              <p className="text-xs text-emerald-700">Đã thanh toán</p>
              <p className="text-sm font-semibold text-emerald-800">{fmtMoney(nvlTotal - nvlDebt)}</p>
            </div>
          </div>
          {nvlRows.length > 0 && <NhapHangList data={data} rows={nvlRows} />}
        </>
      )}
      {category === "bao_tri_vat_tu" && currentUser?.role !== "bao_cao" && <BaoTriVatTuForm currentUser={currentUser} onSubmit={onSubmitExpense} />}
      {(category === "van_hanh" || category === "marketing") && currentUser?.role !== "bao_cao" && <PresetExpenseForm category={category} onSubmit={onSubmitExpense} />}
      {category === "khac" && currentUser?.role !== "bao_cao" && <OtherExpenseForm onSubmit={onSubmitExpense} />}
      {currentUser?.role === "bao_cao" && category !== "nvl" && (
        <div className="mb-4 flex items-center gap-2 text-sm text-slate-500 bg-slate-50 border border-slate-200 rounded-xl px-3 py-2">
          <Receipt size={15} /> Tài khoản Báo cáo chỉ xem, đổi nhóm và sửa/xoá chi phí — không nhập khoản chi mới.
        </div>
      )}

      <ExpenseList data={data} currentUser={currentUser} onDelete={onDeleteExpense} onUpdate={onUpdateExpense} filterCategory={category} from={from} to={to} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// COST MÓN ĂN (công thức + chi phí nguyên liệu + lợi nhuận theo giá bán)
// ---------------------------------------------------------------------------
function DishCreateForm({ onSubmit }) {
  const [name, setName] = useState("");
  const [sellingPrice, setSellingPrice] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const submit = async () => {
    if (!name.trim()) { setError("Vui lòng nhập tên món."); return; }
    setError(""); setSaving(true);
    try {
      await onSubmit({ name, sellingPrice });
      setName(""); setSellingPrice("");
    } catch (e) {
      setError(e.message || "Không tạo được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <p className="font-semibold text-slate-800 text-sm mb-3">Thêm món ăn mới</p>
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <TextField label="Tên món" value={name} onChange={(e) => setName(e.target.value)} placeholder="VD: Gỏi cá trê xù" />
        <MoneyField label="Giá bán / suất (đ, tuỳ chọn)" value={sellingPrice} onChange={setSellingPrice} />
      </div>
      {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
      <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Thêm món</PrimaryButton>
    </Card>
  );
}

function DishRecipeEditor({ dish, data, onSave, onClose }) {
  const existing = data.dishIngredients.filter((i) => i.dishId === dish.id).sort((a, b) => a.sortOrder - b.sortOrder);
  const [lines, setLines] = useState(
    existing.length > 0
      ? existing.map((i) => ({ key: i.id, productId: i.productId, quantity: String(i.quantity), costMode: i.costMode, allocatedCost: i.allocatedCost === null ? "" : String(i.allocatedCost) }))
      : [{ key: Math.random().toString(36).slice(2), productId: "", quantity: "", costMode: "xuat_kho", allocatedCost: "" }]
  );
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const updateLine = (key, patch) => setLines((prev) => prev.map((l) => (l.key === key ? { ...l, ...patch } : l)));
  const addRow = () => setLines((prev) => [...prev, { key: Math.random().toString(36).slice(2), productId: "", quantity: "", costMode: "xuat_kho", allocatedCost: "" }]);
  const removeRow = (key) => setLines((prev) => (prev.length > 1 ? prev.filter((l) => l.key !== key) : prev));

  const lineCost = (l) => {
    if (l.costMode === "phan_bo") return Number(l.allocatedCost) || 0;
    const p = data.products.find((x) => x.id === l.productId);
    if (!p) return 0;
    return computeAvgPrice(p.id, data) * (Number(l.quantity) || 0);
  };
  const totalCost = lines.reduce((s, l) => s + lineCost(l), 0);
  const sellingPrice = dish.sellingPrice || 0;
  const profit = sellingPrice - totalCost;
  const profitPct = sellingPrice > 0 ? (profit / sellingPrice) * 100 : null;

  const submit = async () => {
    const validLines = lines.filter((l) => l.productId && Number(l.quantity) > 0);
    if (validLines.length === 0) { setError("Cần ít nhất 1 dòng nguyên liệu đủ Sản phẩm + Định lượng."); return; }
    if (validLines.some((l) => l.costMode === "phan_bo" && !(Number(l.allocatedCost) > 0))) {
      setError("Dòng tính theo giá phân bổ cần nhập Giá phân bổ > 0."); return;
    }
    setError(""); setSaving(true);
    try {
      await onSave(dish.id, validLines.map((l) => ({
        productId: l.productId, quantity: Number(l.quantity), costMode: l.costMode,
        allocatedCost: l.costMode === "phan_bo" ? Number(l.allocatedCost) : null,
      })));
      onClose();
    } catch (e) {
      setError(e.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-4xl w-full max-h-[90vh] overflow-y-auto p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-base">Công thức: {dish.name}</p>
          <button onClick={onClose}><X size={18} className="text-slate-400" /></button>
        </div>

        <div className="overflow-x-auto rounded-xl border border-slate-200 mb-3">
          <table className="w-full text-sm min-w-[820px]">
            <thead>
              <tr className="text-left text-xs text-slate-500 bg-slate-50 border-b border-slate-200">
                <th className="px-2 py-2 w-40">Mã NVL</th>
                <th className="px-2 py-2">Tên NVL</th>
                <th className="px-2 py-2 w-24">Định lượng</th>
                <th className="px-2 py-2 w-40">Cách tính chi phí</th>
                <th className="px-2 py-2 w-32">Giá phân bổ / Giá xuất kho</th>
                <th className="px-2 py-2 w-32 text-right">Chi phí dòng</th>
                <th className="px-2 py-2 w-10"></th>
              </tr>
            </thead>
            <tbody>
              {lines.map((l) => {
                const p = data.products.find((x) => x.id === l.productId);
                const avgPrice = p ? computeAvgPrice(p.id, data) : 0;
                return (
                  <tr key={l.key} className="border-b border-slate-100 last:border-0">
                    <ProductCodeNameCells
                      products={data.products}
                      productId={l.productId}
                      onSelectProduct={(id) => updateLine(l.key, { productId: id })}
                      codePlaceholder="Mã NVL"
                      namePlaceholder="Tên NVL"
                    />
                    <td className="px-2 py-1.5">
                      <MoneyField allowDecimal value={l.quantity} onChange={(v) => updateLine(l.key, { quantity: v })} placeholder={p ? p.unit : "0"} className="!py-1.5" />
                    </td>
                    <td className="px-2 py-1.5">
                      <select value={l.costMode} onChange={(e) => updateLine(l.key, { costMode: e.target.value })} className="w-full px-2 py-1.5 rounded-lg border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/40">
                        <option value="xuat_kho">Giá xuất kho</option>
                        <option value="phan_bo">Phân bổ (giá cố định)</option>
                      </select>
                    </td>
                    <td className="px-2 py-1.5">
                      {l.costMode === "phan_bo" ? (
                        <MoneyField value={l.allocatedCost} onChange={(v) => updateLine(l.key, { allocatedCost: v })} placeholder="Giá phân bổ (đ)" className="!py-1.5" />
                      ) : (
                        <p className="px-2 py-1.5 text-slate-400 text-xs">{p ? `${fmtMoney(avgPrice)}/${p.unit}` : "—"}</p>
                      )}
                    </td>
                    <td className="px-2 py-1.5 text-right font-medium text-slate-700">{fmtMoney(lineCost(l))}</td>
                    <td className="px-2 py-1.5 text-center">
                      <button type="button" onClick={() => removeRow(l.key)} className="text-slate-400 hover:text-rose-600"><X size={15} /></button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <GhostButton type="button" onClick={addRow} className="mb-4"><Plus size={14} /> Thêm dòng</GhostButton>

        <div className="grid sm:grid-cols-3 gap-2 mb-4">
          <div className="bg-slate-50 rounded-xl px-3 py-2">
            <p className="text-xs text-slate-500">Tổng chi phí nguyên liệu / suất</p>
            <p className="text-sm font-semibold text-slate-700">{fmtMoney(totalCost)}</p>
          </div>
          <div className="bg-slate-50 rounded-xl px-3 py-2">
            <p className="text-xs text-slate-500">Giá bán / suất</p>
            <p className="text-sm font-semibold text-slate-700">{sellingPrice > 0 ? fmtMoney(sellingPrice) : "Chưa đặt giá bán"}</p>
          </div>
          <div className={`rounded-xl px-3 py-2 ${sellingPrice > 0 ? (profit >= 0 ? "bg-emerald-50" : "bg-rose-50") : "bg-slate-50"}`}>
            <p className={`text-xs ${sellingPrice > 0 ? (profit >= 0 ? "text-emerald-700" : "text-rose-700") : "text-slate-500"}`}>Lợi nhuận / suất</p>
            <p className={`text-sm font-semibold ${sellingPrice > 0 ? (profit >= 0 ? "text-emerald-800" : "text-rose-800") : "text-slate-700"}`}>
              {sellingPrice > 0 ? `${fmtMoney(profit)} (${profitPct.toFixed(1)}%)` : "—"}
            </p>
          </div>
        </div>

        {error && <p className="text-xs text-rose-600 mb-2 flex items-center gap-1"><AlertTriangle size={12} /> {error}</p>}
        <div className="flex gap-2">
          <PrimaryButton onClick={submit} disabled={saving}>{saving ? <Loader2 size={15} className="animate-spin" /> : <CheckCircle2 size={15} />} Lưu công thức</PrimaryButton>
          <GhostButton onClick={onClose}>Đóng</GhostButton>
        </div>
      </div>
    </div>
  );
}

function MonAnModule({ data, onAddDish, onSaveRecipe, onDeleteDish }) {
  const [editingDish, setEditingDish] = useState(null);
  const [deleting, setDeleting] = useState(null);
  const [search, setSearch] = useState("");

  const handleDelete = async (dishId) => {
    setDeleting(dishId);
    try { await onDeleteDish(dishId); } finally { setDeleting(null); }
  };

  // Tỷ lệ cost bình quân: tổng chi phí nguyên liệu / tổng giá bán, tính trên các
  // món ĐÃ có giá bán (món chưa đặt giá bán không có ý nghĩa để đưa vào tỷ lệ này).
  const costRatioStats = useMemo(() => {
    let totalCost = 0, totalRevenue = 0, countPriced = 0;
    data.dishes.forEach((d) => {
      if (!d.sellingPrice || d.sellingPrice <= 0) return;
      totalCost += dishTotalCost(d.id, data);
      totalRevenue += d.sellingPrice;
      countPriced += 1;
    });
    const ratio = totalRevenue > 0 ? (totalCost / totalRevenue) * 100 : null;
    return { ratio, totalCost, totalRevenue, countPriced };
  }, [data]);

  // Lọc theo tên món — bỏ dấu + không phân biệt hoa/thường, để gõ nhanh không cần gõ đúng dấu.
  const filteredDishes = useMemo(() => {
    const q = normalizeForMatch(search);
    if (!q) return data.dishes;
    return data.dishes.filter((d) => normalizeForMatch(d.name).includes(q));
  }, [data.dishes, search]);

  return (
    <div>
      <SectionTitle icon={Package} title="Cost món ăn" subtitle="Công thức, chi phí nguyên liệu và lợi nhuận theo từng món" />

      <Card className="p-4 sm:p-5 mb-5">
        <p className="font-semibold text-slate-800 text-sm mb-3">Tỷ lệ cost bình quân</p>
        {costRatioStats.ratio === null ? (
          <p className="text-xs text-slate-400">Chưa có món nào có giá bán để tính tỷ lệ cost.</p>
        ) : (
          <div className="grid sm:grid-cols-3 gap-2">
            <div className={`rounded-xl px-3 py-2 ${costRatioStats.ratio <= 35 ? "bg-emerald-50" : costRatioStats.ratio <= 45 ? "bg-amber-50" : "bg-rose-50"}`}>
              <p className="text-xs text-slate-500">Tỷ lệ cost bình quân</p>
              <p className={`text-lg font-semibold ${costRatioStats.ratio <= 35 ? "text-emerald-800" : costRatioStats.ratio <= 45 ? "text-amber-800" : "text-rose-800"}`}>
                {costRatioStats.ratio.toFixed(1)}%
              </p>
            </div>
            <div className="bg-slate-50 rounded-xl px-3 py-2">
              <p className="text-xs text-slate-500">Tổng chi phí nguyên liệu</p>
              <p className="text-sm font-semibold text-slate-700">{fmtMoney(costRatioStats.totalCost)}</p>
            </div>
            <div className="bg-slate-50 rounded-xl px-3 py-2">
              <p className="text-xs text-slate-500">Tổng giá bán ({costRatioStats.countPriced} món)</p>
              <p className="text-sm font-semibold text-slate-700">{fmtMoney(costRatioStats.totalRevenue)}</p>
            </div>
          </div>
        )}
      </Card>

      <DishCreateForm onSubmit={onAddDish} />

      <Card className="p-0 overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex items-center justify-between gap-3 flex-wrap">
          <p className="font-semibold text-slate-800 text-sm">
            Danh sách món ăn ({filteredDishes.length}{filteredDishes.length !== data.dishes.length ? `/${data.dishes.length}` : ""})
          </p>
          <div className="relative w-full sm:w-64">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm món ăn..."
              className="w-full pl-8 pr-8 py-1.5 rounded-lg border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-500/40"
            />
            {search && (
              <button type="button" onClick={() => setSearch("")} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-rose-600">
                <X size={14} />
              </button>
            )}
          </div>
        </div>
        {data.dishes.length === 0 ? (
          <EmptyState icon={Package} text="Chưa có món ăn nào." />
        ) : filteredDishes.length === 0 ? (
          <EmptyState icon={Search} text={`Không tìm thấy món nào khớp với "${search}".`} />
        ) : (
          <div className="divide-y divide-slate-100">
            {filteredDishes.map((d) => {
              const cost = dishTotalCost(d.id, data);
              const ingCount = data.dishIngredients.filter((i) => i.dishId === d.id).length;
              const profit = d.sellingPrice ? d.sellingPrice - cost : null;
              const costRatio = d.sellingPrice > 0 ? (cost / d.sellingPrice) * 100 : null;
              return (
                <div key={d.id} className="px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-slate-800 truncate">{d.name}</p>
                    <p className="text-xs text-slate-400">
                      {ingCount} nguyên liệu · Chi phí {fmtMoney(cost)}{d.sellingPrice ? ` · Giá bán ${fmtMoney(d.sellingPrice)}` : ""}
                      {costRatio !== null && ` · Cost ${costRatio.toFixed(1)}%`}
                    </p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    {profit !== null && (
                      <span className={`text-xs font-medium px-2 py-1 rounded-full ${profit >= 0 ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"}`}>
                        {profit >= 0 ? "Lãi" : "Lỗ"} {fmtMoney(Math.abs(profit))}
                      </span>
                    )}
                    <GhostButton onClick={() => setEditingDish(d)}><Pencil size={13} /> Công thức</GhostButton>
                    <button onClick={() => handleDelete(d.id)} disabled={deleting === d.id} className="text-slate-400 hover:text-rose-600 p-1.5"><Trash2 size={15} /></button>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </Card>

      {editingDish && (
        <DishRecipeEditor dish={editingDish} data={data} onSave={onSaveRecipe} onClose={() => setEditingDish(null)} />
      )}
    </div>
  );
}

function TaiKhoanModule({ currentUser, employees, onAddEmployee }) {
  const [query, setQuery] = useState("");
  const [resetting, setResetting] = useState(null);
  const q = query.trim().toLowerCase();
  const filtered = q ? employees.filter((e) => e.name?.toLowerCase().includes(q) || e.username?.toLowerCase().includes(q)) : employees;

  return (
    <div>
      <SectionTitle icon={ShieldCheck} title="Quản lý tài khoản" subtitle={`${employees.length} tài khoản nhân sự`} />
      <AddEmployeeForm onAdd={onAddEmployee} />
      <Card className="p-3 mb-4 flex items-center gap-2">
        <Search size={15} className="text-slate-400 shrink-0" />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo tên hoặc tên đăng nhập..." className="flex-1 text-sm outline-none" />
      </Card>
      {filtered.length === 0 ? (
        <EmptyState icon={Search} text="Không tìm thấy tài khoản phù hợp." />
      ) : (
        <div className="space-y-2">
          {filtered.map((e) => (
            <Card key={e.id} className="p-3.5 flex items-center justify-between gap-3">
              <div className="min-w-0 flex items-center gap-3">
                <div className="w-9 h-9 rounded-full bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center text-xs font-semibold shrink-0">
                  {e.name?.trim()?.split(" ").slice(-1)[0]?.[0] || "?"}
                </div>
                <div className="min-w-0">
                  <p className="text-sm font-medium text-slate-800 truncate">{e.name} <span className="text-slate-400 font-normal">({e.username})</span></p>
                  <p className="text-xs text-slate-400 truncate">{ROLE_META[e.role]?.label}{e.mustChangePassword ? " · Đang chờ đổi mật khẩu" : ""}</p>
                </div>
              </div>
              <GhostButton className="!text-xs shrink-0" onClick={() => setResetting(e)}><Lock size={13} /> Đặt lại mật khẩu</GhostButton>
            </Card>
          ))}
        </div>
      )}
      {resetting && <ResetPasswordModal currentUser={currentUser} employee={resetting} onClose={() => setResetting(null)} onSuccess={() => setResetting(null)} />}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Chặn lỗi phát sinh trong 1 tab lan ra làm sập toàn bộ giao diện.
// ---------------------------------------------------------------------------
class TabErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError() { return { hasError: true }; }
  componentDidCatch(error, info) { console.error("Lỗi hiển thị tab:", error, info); }
  componentDidUpdate(prevProps) {
    if (prevProps.resetKey !== this.props.resetKey && this.state.hasError) this.setState({ hasError: false });
  }
  render() {
    if (this.state.hasError) {
      return (
        <div className="bg-white rounded-2xl border border-rose-200 shadow-sm p-6 text-center">
          <AlertTriangle size={28} className="text-rose-500 mx-auto mb-2" />
          <p className="font-medium text-slate-800 mb-1">Không hiển thị được mục này</p>
          <p className="text-sm text-slate-500 mb-4">Đã có lỗi xảy ra. Bạn có thể chuyển sang mục khác từ thanh menu, hoặc tải lại trang.</p>
          <GhostButton onClick={() => window.location.reload()}>Tải lại trang</GhostButton>
        </div>
      );
    }
    return this.props.children;
  }
}

// ---------------------------------------------------------------------------
// APP GỐC
// ---------------------------------------------------------------------------
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("nhap");
  const [toast, setToast] = useState(null);
  const [showChangePassword, setShowChangePassword] = useState(false);

  const [data, setData] = useState({
    employees: [], suppliers: [], revenueCodes: [], exportCodes: [], products: [],
    stockOpenings: [], importRecords: [], exportRecords: [], expenseRecords: [],
    dishes: [], dishIngredients: [], deposits: [], notifications: [], fundDailyBalances: [], settings: {},
  });

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 3000); };

  const refreshAll = useCallback(async () => {
    const fresh = await fetchAll();
    setData(fresh);
    return fresh;
  }, []);

  useEffect(() => {
    (async () => {
      await refreshAll();
      const savedId = localStorage.getItem(SESSION_KEY);
      if (savedId) {
        const fresh = await fetchAll();
        const found = fresh.employees.find((e) => e.id === savedId);
        if (found) setCurrentUser(found);
        else localStorage.removeItem(SESSION_KEY);
      }
      setLoading(false);
    })();
  }, [refreshAll]);

  useEffect(() => {
    if (!currentUser) return;
    const t = setInterval(() => { refreshAll(); }, 20000);
    return () => clearInterval(t);
  }, [currentUser, refreshAll]);

  const handleLogin = (employee) => {
    setCurrentUser(employee);
    localStorage.setItem(SESSION_KEY, employee.id);
    setTab(employee.role === "bao_cao" ? "lich_su_nhap" : employee.role === "thu_ngan" ? "thu" : "nhap");
  };
  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem(SESSION_KEY);
  };

  // ---------------- Danh mục ----------------
  const addSupplier = async ({ code, name, paymentType }) => {
    const { error } = await supabase.from("suppliers").insert({ code, name, payment_type: paymentType });
    if (error) throw error;
    await refreshAll();
    showToast("Đã thêm nhà cung cấp");
  };
  const addProduct = async ({ code, name, unit, groupCode, groupName, classification }) => {
    const { error } = await supabase.from("products").insert({ code, name, unit, group_code: groupCode, group_name: groupName, classification });
    if (error) throw error;
    await refreshAll();
    showToast("Đã thêm sản phẩm");
  };
  const addRevenueCode = async ({ code, name }) => {
    const { error } = await supabase.from("revenue_codes").insert({ code, name });
    if (error) throw error;
    await refreshAll();
    showToast("Đã thêm mã doanh thu");
  };
  const addExportCode = async ({ code, name }) => {
    const { error } = await supabase.from("export_codes").insert({ code, name });
    if (error) throw error;
    await refreshAll();
    showToast("Đã thêm mã xuất");
  };

  // ---------------- Nhập hàng ----------------
  const submitImport = async ({ orderNumber, supplierId, lines, paymentType }) => {
    const receiptCode = genReceiptCode("NK");
    const rows = lines.map((l) => ({
      order_number: orderNumber || null, receipt_code: receiptCode, supplier_id: supplierId, product_id: l.productId,
      quantity: l.quantity, unit_price: l.unitPrice, total_amount: l.totalAmount, payment_type: paymentType,
      import_date: todayISO(), created_by: currentUser.id,
    }));
    const { error } = await supabase.from("import_records").insert(rows);
    if (error) throw error;
    await refreshAll();
    const totalAmount = rows.reduce((s, r) => s + r.total_amount, 0);
    showToast(`Đã lưu phiếu nhập ${receiptCode} (${rows.length} dòng)`);
    return { receiptCode, lineCount: rows.length, totalAmount, supplierName: data.suppliers.find((s) => s.id === supplierId)?.name };
  };

  // Xoá theo từng đợt nhỏ (thay vì xoá hàng nghìn id cùng lúc trong 1 câu lệnh)
  // — tránh request quá dài bị từ chối khi phiếu có rất nhiều dòng.
  const CHUNK_SIZE = 100;
  const deleteInChunks = async (table, ids) => {
    let totalDeleted = 0;
    for (let i = 0; i < ids.length; i += CHUNK_SIZE) {
      const chunk = ids.slice(i, i + CHUNK_SIZE);
      // .select("id") để biết CHÍNH XÁC bao nhiêu dòng thực sự bị xoá — nếu RLS chặn ngầm,
      // Supabase trả về thành công nhưng 0 dòng bị xoá (không báo lỗi), nên phải tự kiểm tra.
      const { data: deletedRows, error } = await supabase.from(table).delete().in("id", chunk).select("id");
      if (error) throw error;
      totalDeleted += (deletedRows || []).length;
    }
    if (totalDeleted < ids.length) {
      throw new Error(`Chỉ xoá được ${totalDeleted}/${ids.length} dòng — có thể bị chặn quyền (RLS) trên Supabase. Vui lòng chạy DELETE trực tiếp trong Supabase SQL Editor hoặc kiểm tra lại policy.`);
    }
  };

  // Xoá 1 dòng phiếu nhập, hoặc cả phiếu (theo receiptCode) nếu truyền receiptCode thay vì id.
  const deleteImportRecord = async (id) => {
    const { data: deletedRows, error } = await supabase.from("import_records").delete().eq("id", id).select("id");
    if (error) throw error;
    if ((deletedRows || []).length === 0) throw new Error("Không xoá được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã xoá dòng nhập hàng");
  };
  const deleteImportRecordsByIds = async (ids) => {
    if (ids.length === 0) return;
    await deleteInChunks("import_records", ids);
    await refreshAll();
    showToast(`Đã xoá ${ids.length} dòng nhập hàng`);
  };

  // Xoá 1 dòng phiếu xuất, hoặc nhiều dòng cùng lúc theo danh sách id.
  const deleteExportRecord = async (id) => {
    const { data: deletedRows, error } = await supabase.from("export_records").delete().eq("id", id).select("id");
    if (error) throw error;
    if ((deletedRows || []).length === 0) throw new Error("Không xoá được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã xoá dòng xuất hàng");
  };
  const deleteExportRecordsByIds = async (ids) => {
    if (ids.length === 0) return;
    // Xoá kèm dish_sales cùng phiếu (nếu có) để "Báo cáo doanh thu theo ngày" không còn dữ liệu ma.
    const receiptCodes = Array.from(new Set(
      ids.map((id) => data.exportRecords.find((r) => r.id === id)?.receiptCode).filter(Boolean)
    ));
    if (receiptCodes.length > 0) {
      const { error: saleErr } = await supabase.from("dish_sales").delete().in("receipt_code", receiptCodes);
      if (saleErr) console.error(saleErr);
    }
    await deleteInChunks("export_records", ids);
    await refreshAll();
    showToast(`Đã xoá ${ids.length} dòng xuất hàng`);
  };

  // Nhập hàng hàng loạt từ file Excel — mỗi dòng Excel tự mang đúng NCC/SP/SL/Đơn giá riêng.
  const bulkImportNhap = async (rows) => {
    const receiptCode = genReceiptCode("NK");
    const dbRows = rows.map((r) => ({
      order_number: r.orderNumber || null, receipt_code: receiptCode, supplier_id: r.supplierId, product_id: r.productId,
      quantity: r.quantity, unit_price: r.unitPrice, total_amount: r.quantity * r.unitPrice,
      payment_type: r.paymentType, import_date: r.importDate || todayISO(), created_by: currentUser.id,
    }));
    const { error } = await supabase.from("import_records").insert(dbRows);
    if (error) throw error;
    await refreshAll();
    const totalAmount = dbRows.reduce((s, r) => s + r.total_amount, 0);
    showToast(`Đã nhập ${dbRows.length} dòng từ file Excel (phiếu ${receiptCode})`);
    return { receiptCode, lineCount: dbRows.length, totalAmount };
  };

  // ---------------- Xuất hàng ----------------
  const submitExport = async ({ orderNumber, revenueCodeId, exportCodeId, lines }) => {
    const receiptCode = genReceiptCode("XK");
    const rows = lines.map((l) => ({
      order_number: orderNumber || null, receipt_code: receiptCode,
      revenue_code_id: revenueCodeId, export_code_id: exportCodeId,
      product_id: l.productId, line_type: l.lineType, quantity: l.quantity, unit_price: l.unitPrice,
      total_amount: l.totalAmount, export_date: todayISO(), created_by: currentUser.id,
    }));
    const { error } = await supabase.from("export_records").insert(rows);
    if (error) throw error;
    await refreshAll();
    const totalAmount = rows.reduce((s, r) => s + r.total_amount, 0);
    showToast(`Đã lưu phiếu xuất ${receiptCode}`);
    return { receiptCode, lineCount: rows.length, totalAmount };
  };

  // Xuất kho NVL tự động từ báo cáo doanh thu chi tiết theo hoá đơn & món ăn:
  // mỗi dòng "món X bán N suất" được nổ ra thành các dòng NVL theo đúng công thức
  // (Cost món ăn) — số lượng NVL tiêu hao = định lượng trong công thức × N.
  // Ngày của từng dòng lấy từ cột "Ngày" trong file (r.saleDate), không dùng 1 ngày chung
  // cho cả file — vì 1 file có thể gộp doanh thu nhiều ngày.
  // Đồng thời lưu lại chi tiết từng lượt bán món vào dish_sales — phục vụ
  // "Báo cáo doanh thu theo ngày" (số hoá đơn, doanh số, giá vốn, số món bán, tỉ lệ cost).
  const bulkImportXuatFromBills = async ({ rows }) => {
    const exportRows = [];
    const saleRows = [];
    const notFoundDishes = new Set();
    rows.forEach((r) => {
      const dish = data.dishes.find((d) => normalizeForMatch(d.name) === normalizeForMatch(r.dishName));
      if (!dish) { notFoundDishes.add(r.dishName); return; }
      const rowDate = r.saleDate || todayISO();
      const ingredients = data.dishIngredients.filter((i) => i.dishId === dish.id);
      ingredients.forEach((ing) => {
        const product = data.products.find((p) => p.id === ing.productId);
        if (!product) return;
        const unitPrice = ing.costMode === "phan_bo"
          ? (ing.quantity > 0 ? (ing.allocatedCost || 0) / ing.quantity : 0)
          : computeAvgPrice(ing.productId, data);
        const quantity = ing.quantity * r.quantitySold;
        exportRows.push({
          order_number: r.invoiceNo || null, receipt_code: null,
          revenue_code_id: null, export_code_id: null,
          product_id: ing.productId, line_type: product.classification, quantity, unit_price: unitPrice,
          total_amount: quantity * unitPrice, export_date: rowDate, created_by: currentUser.id,
        });
      });
      const dishCostPerUnit = dishTotalCost(dish.id, data);
      // Doanh thu (dish_sales) lấy thẳng cột "Doanh thu" của file (r.revenueInFile) — không tự nhân
      // SL bán × Đơn giá, không dùng giá cấu hình cứng trong Cost món ăn — vì có trường hợp Đơn giá
      // niêm yết không đổi nhưng thực thu khác (khuyến mãi/giảm giá/tặng kèm riêng từng dòng hoá đơn).
      // Nhờ đó "Doanh thu từ phiếu xuất kho" luôn khớp đúng với "Doanh thu bán hàng theo hoá đơn".
      const lineRevenue = r.revenueInFile || 0;
      const dishSellingPrice = r.quantitySold > 0 ? lineRevenue / r.quantitySold : 0;
      saleRows.push({
        dish_id: dish.id, quantity: r.quantitySold, unit_price: dishSellingPrice,
        total_amount: lineRevenue, cost_amount: dishCostPerUnit * r.quantitySold,
        invoice_no: r.invoiceNo || null, receipt_code: null,
        sale_date: rowDate, created_by: currentUser.id,
      });
    });
    if (exportRows.length === 0) {
      throw new Error("Không có dòng nào khớp được với danh sách món ăn trong Cost món ăn.");
    }
    const receiptCode = genReceiptCode("XK");
    exportRows.forEach((r) => { r.receipt_code = receiptCode; });
    saleRows.forEach((r) => { r.receipt_code = receiptCode; });
    const { error } = await supabase.from("export_records").insert(exportRows);
    if (error) throw error;
    const { error: saleError } = await supabase.from("dish_sales").insert(saleRows);
    if (saleError) console.error(saleError); // không chặn luồng chính nếu lỗi phần thống kê doanh thu
    await refreshAll();
    const matchedBills = rows.length - notFoundDishes.size;
    const totalAmount = exportRows.reduce((s, r) => s + r.total_amount, 0);
    showToast(`Đã xuất kho ${exportRows.length} dòng NVL từ ${matchedBills} dòng báo cáo doanh thu (phiếu ${receiptCode})` + (notFoundDishes.size ? ` — bỏ qua ${notFoundDishes.size} tên món không khớp` : ""));
    return { notFoundDishes: Array.from(notFoundDishes), receiptCode, lineCount: exportRows.length, totalAmount };
  };

  // ---------------- Chi phí (Vận hành / Marketing / Bảo trì & vật tư / Khác) ----------------
  // Gửi thông báo đến tài khoản Báo cáo mỗi khi có cập nhật Thu/Chi (Thu ngân, Chi phí, Cọc).
  // Không chặn luồng chính nếu lỗi (chỉ log ra console) — tránh làm hỏng thao tác chính của người dùng.
  const pushNotification = async ({ message, type }) => {
    try {
      await supabase.from("notifications").insert({
        message, type, target_role: "bao_cao",
        created_by: currentUser.id, created_by_name: currentUser.name, is_read: false,
      });
    } catch (e) {
      console.error("pushNotification lỗi:", e);
    }
  };

  const submitExpense = async ({ category, lines, expenseDate, paymentMethod }) => {
    const rows = lines.map((l) => ({
      category, item_name: l.itemName,
      quantity: l.quantity ?? null, unit_price: l.unitPrice ?? null, amount: l.amount,
      payment_method: paymentMethod || "tien_mat",
      expense_date: expenseDate || todayISO(), note: l.note || null, created_by: currentUser.id,
    }));
    const { error } = await supabase.from("expense_records").insert(rows);
    if (error) throw error;
    const total = rows.reduce((s, r) => s + (r.amount || 0), 0);
    const catLabel = EXPENSE_CATEGORY_META[category]?.label || category;
    await pushNotification({
      type: "chi_phi",
      message: `${currentUser.name} vừa ghi nhận ${rows.length} khoản chi phí "${catLabel}" — tổng ${fmtMoney(total)}`,
    });
    await refreshAll();
    showToast(`Đã ghi nhận ${rows.length} khoản chi phí`);
  };

  // Xoá 1 khoản chi phí — chỉ dành cho tài khoản Quản lý (kiểm soát ở giao diện ExpenseList).
  const deleteExpenseRecord = async (id) => {
    const { data: deletedRows, error } = await supabase.from("expense_records").delete().eq("id", id).select("id");
    if (error) throw error;
    if ((deletedRows || []).length === 0) throw new Error("Không xoá được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã xoá khoản chi phí");
  };

  // Sửa 1 khoản chi phí — áp dụng cho mọi khoản chi trong bảng expense_records, kể
  // cả khoản do tài khoản Thu ngân gửi lên (dùng chung 1 bảng, không phân biệt người tạo).
  const updateExpenseRecord = async (id, { category, itemName, expenseDate, amount, quantity, unitPrice, paymentMethod }) => {
    const { data: updatedRows, error } = await supabase
      .from("expense_records")
      .update({ category, item_name: itemName, expense_date: expenseDate, amount, quantity, unit_price: unitPrice, payment_method: paymentMethod || "tien_mat" })
      .eq("id", id)
      .select("id");
    if (error) throw error;
    if ((updatedRows || []).length === 0) throw new Error("Không sửa được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã cập nhật khoản chi phí");
  };

  // Thu ngân — ghi nhận Thu tiền mặt / Thu tiền ngân hàng theo ngày (độc lập với doanh thu
  // tính từ báo cáo bán hàng/dish_sales) — dùng để đối chiếu quỹ thực tế trong tab Quỹ.
  const submitCashierReceipt = async ({ receiptDate, cashAmount, bankAmount, note }) => {
    const { error } = await supabase.from("cashier_receipts").insert({
      receipt_date: receiptDate || todayISO(), cash_amount: Number(cashAmount) || 0,
      bank_amount: Number(bankAmount) || 0, note: note || null, created_by: currentUser.id,
    });
    if (error) throw error;
    await pushNotification({
      type: "thu_ngan",
      message: `${currentUser.name} vừa ghi nhận phiếu thu ngày ${fmtDate(receiptDate || todayISO())} — Tiền mặt ${fmtMoney(Number(cashAmount) || 0)}, Ngân hàng ${fmtMoney(Number(bankAmount) || 0)}`,
    });
    await refreshAll();
    showToast("Đã ghi nhận phiếu thu ngân");
  };
  const updateCashierReceipt = async (id, { receiptDate, cashAmount, bankAmount, note }) => {
    const { data: updatedRows, error } = await supabase
      .from("cashier_receipts")
      .update({ receipt_date: receiptDate, cash_amount: Number(cashAmount) || 0, bank_amount: Number(bankAmount) || 0, note: note || null })
      .eq("id", id)
      .select("id");
    if (error) throw error;
    if ((updatedRows || []).length === 0) throw new Error("Không sửa được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã cập nhật phiếu thu ngân");
  };
  const deleteCashierReceipt = async (id) => {
    const { data: deletedRows, error } = await supabase.from("cashier_receipts").delete().eq("id", id).select("id");
    if (error) throw error;
    if ((deletedRows || []).length === 0) throw new Error("Không xoá được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã xoá phiếu thu ngân");
  };

  // Cọc — thu cọc (nhận cọc từ khách) / chi cọc (ứng cọc cho NCC, đối tác). Tách riêng khỏi
  // Chi phí vì đây là dòng tiền tạm giữ, chờ đối trừ hoặc hoàn trả, không phải chi phí thực.
  const submitDeposit = async ({ direction, partyName, amount, paymentMethod, depositDate, note }) => {
    const { error } = await supabase.from("deposits").insert({
      direction, party_name: partyName.trim(), amount: Number(amount) || 0,
      payment_method: paymentMethod || "tien_mat",
      deposit_date: depositDate || todayISO(), status: "dang_giu", note: note || null, created_by: currentUser.id,
    });
    if (error) throw error;
    await pushNotification({
      type: "coc",
      message: `${currentUser.name} vừa ghi nhận ${direction === "thu" ? "thu cọc từ" : "chi cọc cho"} "${partyName.trim()}" — ${fmtMoney(Number(amount) || 0)}`,
    });
    await refreshAll();
    showToast(direction === "thu" ? "Đã ghi nhận khoản thu cọc" : "Đã ghi nhận khoản chi cọc");
  };
  const updateDeposit = async (id, { direction, partyName, amount, paymentMethod, depositDate, status, note }) => {
    const { data: updatedRows, error } = await supabase
      .from("deposits")
      .update({ direction, party_name: partyName, amount, payment_method: paymentMethod || "tien_mat", deposit_date: depositDate, status, note: note || null })
      .eq("id", id)
      .select("id");
    if (error) throw error;
    if ((updatedRows || []).length === 0) throw new Error("Không sửa được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã cập nhật khoản cọc");
  };
  const deleteDeposit = async (id) => {
    const { data: deletedRows, error } = await supabase.from("deposits").delete().eq("id", id).select("id");
    if (error) throw error;
    if ((deletedRows || []).length === 0) throw new Error("Không xoá được — có thể bị chặn quyền (RLS) trên Supabase.");
    await refreshAll();
    showToast("Đã xoá khoản cọc");
  };

  // Tồn quỹ tiền mặt đầu ngày — lưu theo từng ngày, mặc định ngày sau tự lấy tồn cuối
  // ngày trước (tính ở QuyModule), nhưng có thể ghi đè tay bất kỳ ngày nào (kiểm đếm thực tế).
  const upsertFundOpeningBalance = async (date, field, amount) => {
    const column = field === "bank" ? "opening_balance_bank" : "opening_balance";
    const { error } = await supabase
      .from("fund_daily_balance")
      .upsert({ balance_date: date, [column]: Number(amount) || 0, updated_by: currentUser.id }, { onConflict: "balance_date" });
    if (error) throw error;
    await refreshAll();
    showToast("Đã cập nhật tồn quỹ đầu ngày");
  };

  // Tiền nộp cho chủ (cô) mỗi ngày — trừ khỏi Tồn cuối ngày trong Sổ quỹ theo ngày.
  const upsertFundRemittedAmount = async (date, field, amount) => {
    const column = field === "bank" ? "remitted_owner_bank" : "remitted_owner_cash";
    const { error } = await supabase
      .from("fund_daily_balance")
      .upsert({ balance_date: date, [column]: Number(amount) || 0, updated_by: currentUser.id }, { onConflict: "balance_date" });
    if (error) throw error;
    await refreshAll();
    showToast("Đã cập nhật tiền nộp cho cô");
  };

  // Mở/khoá quyền tự sửa-xoá phiếu thu & phiếu chi của tài khoản Thu ngân — cấu hình chung áp
  // dụng cho mọi tài khoản thu_ngan, do Quản lý hoặc Quản lý (Báo cáo) bật/tắt trong tab Quỹ.
  const setThuNganEditEnabled = async (enabled) => {
    const { error } = await supabase
      .from("app_settings")
      .upsert({ key: "thu_ngan_edit_enabled", value: enabled ? "true" : "false", updated_by: currentUser.id }, { onConflict: "key" });
    if (error) throw error;
    await refreshAll();
    showToast(enabled ? "Đã mở khoá cho Thu ngân tự sửa/xoá" : "Đã khoá — Thu ngân không tự sửa/xoá được nữa");
  };

  const markNotificationRead = async (id) => {
    const { error } = await supabase.from("notifications").update({ is_read: true }).eq("id", id);
    if (error) { console.error(error); return; }
    await refreshAll();
  };
  const markAllNotificationsRead = async () => {
    const unreadIds = data.notifications.filter((n) => !n.isRead).map((n) => n.id);
    if (unreadIds.length === 0) return;
    const { error } = await supabase.from("notifications").update({ is_read: true }).in("id", unreadIds);
    if (error) { console.error(error); return; }
    await refreshAll();
  };

  // Import "Bảng kê hoá đơn" (POS) — Doanh thu bán hàng theo hoá đơn, độc lập với doanh thu
  // tính từ Xuất kho tự động (dish_sales). Upsert theo invoice_no để import lại không bị trùng.
  const bulkImportInvoiceRevenue = async (rows) => {
    if (rows.length === 0) return;
    const dbRows = rows.map((r) => ({
      invoice_no: r.invoiceNo, invoice_date: r.invoiceDate, amount: r.amount, created_by: currentUser.id,
    }));
    const { error } = await supabase.from("invoice_revenue").upsert(dbRows, { onConflict: "invoice_no" });
    if (error) throw error;
    await refreshAll();
    showToast(`Đã import ${rows.length} hoá đơn doanh thu`);
  };

  // ---------------- Cost món ăn ----------------
  const addDish = async ({ name, sellingPrice, note }) => {
    const { error } = await supabase.from("dishes").insert({
      name: name.trim(), selling_price: sellingPrice === "" || sellingPrice === undefined ? null : Number(sellingPrice),
      note: note?.trim() || null, created_by: currentUser.id,
    });
    if (error) throw error;
    await refreshAll();
    showToast("Đã tạo món ăn");
  };

  const deleteDish = async (dishId) => {
    const { error } = await supabase.from("dishes").delete().eq("id", dishId);
    if (error) throw error;
    await refreshAll();
    showToast("Đã xoá món ăn");
  };

  // Lưu lại toàn bộ công thức của 1 món: xoá hết dòng cũ rồi ghi lại danh sách mới
  // (đơn giản, tránh phải so khớp thêm/sửa/xoá từng dòng riêng lẻ).
  const saveDishRecipe = async (dishId, lines) => {
    const { error: delErr } = await supabase.from("dish_ingredients").delete().eq("dish_id", dishId);
    if (delErr) throw delErr;
    if (lines.length > 0) {
      const rows = lines.map((l, idx) => ({
        dish_id: dishId, product_id: l.productId, quantity: l.quantity,
        cost_mode: l.costMode, allocated_cost: l.costMode === "phan_bo" ? l.allocatedCost : null,
        sort_order: idx,
      }));
      const { error: insErr } = await supabase.from("dish_ingredients").insert(rows);
      if (insErr) throw insErr;
    }
    await refreshAll();
    showToast("Đã lưu công thức món ăn");
  };

  // ---------------- Tồn kho ----------------
  const saveStockOpening = async ({ productId, asOfDate, quantity, unitPrice, note }) => {
    const { error } = await supabase.from("stock_opening").insert({
      product_id: productId, as_of_date: asOfDate, quantity, unit_price: unitPrice, note: note || null, created_by: currentUser.id,
    });
    if (error) throw error;
    await refreshAll();
    showToast("Đã lưu mốc tồn đầu");
  };

  // ---------------- Tài khoản ----------------
  const addEmployee = async ({ username, name, role, adminPassword }) => {
    const { data: newId, error } = await supabase.rpc("admin_create_employee", {
      p_admin_id: currentUser.id, p_admin_password: adminPassword,
      p_username: username, p_name: name, p_role: role, p_initial_password: "123456",
    });
    if (error) {
      if (error.message?.includes("ADMIN_PASSWORD_INCORRECT")) throw new Error("Mật khẩu của bạn không đúng.");
      if (error.message?.includes("duplicate")) throw new Error("Tên đăng nhập này đã tồn tại.");
      throw error;
    }
    await refreshAll();
    showToast("Đã tạo tài khoản mới — mật khẩu ban đầu: 123456");
  };

  const handlePasswordChanged = () => {
    setCurrentUser((prev) => prev && { ...prev, mustChangePassword: false, passwordChangeDeadline: null });
    showToast("Đã đổi mật khẩu thành công");
  };

  if (loading) return <FullScreenLoader />;
  if (!currentUser) return <LoginScreen onLogin={handleLogin} />;
  if (currentUser.mustChangePassword) {
    return <ForcePasswordChangeGate currentUser={currentUser} onLogout={handleLogout} onChanged={handlePasswordChanged} />;
  }

  const isQuanLy = currentUser.role === "quan_ly";
  const isBaoCao = currentUser.role === "bao_cao";
  const isThuNgan = currentUser.role === "thu_ngan";
  const canViewReports = isQuanLy || isBaoCao;
  const NAV_NHAN_VIEN = [
    { key: "nhap", label: "Nhập hàng", icon: ArrowDownCircle },
    { key: "lich_su_nhap", label: "Lịch sử nhập hàng", icon: Inbox },
    { key: "xuat", label: "Xuất hàng", icon: ArrowUpCircle },
    { key: "chi_phi", label: "Chi phí", icon: Receipt },
    { key: "mon_an", label: "Cost món ăn", icon: Package },
    { key: "danh_muc", label: "Danh mục", icon: Boxes },
    { key: "ton_kho", label: "Tồn kho", icon: Warehouse },
  ];
  const NAV_QUAN_LY = [
    { key: "nhap", label: "Nhập hàng", icon: ArrowDownCircle },
    { key: "lich_su_nhap", label: "Lịch sử nhập hàng", icon: Inbox },
    { key: "xuat", label: "Xuất hàng", icon: ArrowUpCircle },
    { key: "chi_phi", label: "Chi phí", icon: Receipt },
    { key: "coc", label: "Cọc", icon: Coins },
    { key: "quy", label: "Quỹ", icon: Wallet },
    { key: "mon_an", label: "Cost món ăn", icon: Package },
    { key: "danh_muc", label: "Danh mục", icon: Boxes },
    { key: "bao_cao_nhap", label: "Báo cáo nhập", icon: BarChart3 },
    { key: "bao_cao_xuat", label: "Báo cáo xuất", icon: BarChart3 },
    { key: "ton_kho", label: "Tồn kho", icon: Warehouse },
    { key: "tai_khoan", label: "Tài khoản", icon: ShieldCheck },
  ];
  // Nhóm 3: Quản lý (Báo cáo) — chỉ xem báo cáo/lịch sử, không có Nhập hàng, Danh mục, Tồn kho, Tài khoản.
  const NAV_BAO_CAO = [
    { key: "lich_su_nhap", label: "Lịch sử nhập hàng", icon: Inbox },
    { key: "xuat", label: "Lịch sử xuất hàng", icon: ArrowUpCircle },
    { key: "chi_phi", label: "Chi phí", icon: Receipt },
    { key: "coc", label: "Cọc", icon: Coins },
    { key: "quy", label: "Quỹ", icon: Wallet },
    { key: "mon_an", label: "Cost món ăn", icon: Package },
    { key: "bao_cao_nhap", label: "Báo cáo nhập", icon: BarChart3 },
    { key: "bao_cao_xuat", label: "Báo cáo xuất", icon: BarChart3 },
  ];
  // Nhóm 4: Thu ngân — chỉ có màn ghi phiếu thu (tiền mặt/ngân hàng) + phiếu chi phát sinh.
  const NAV_THU_NGAN = [
    { key: "thu", label: "Thu ngân", icon: Wallet },
    { key: "chi", label: "Phiếu chi", icon: Receipt },
    { key: "coc", label: "Cọc", icon: Coins },
  ];
  const navItems = isQuanLy ? NAV_QUAN_LY : isBaoCao ? NAV_BAO_CAO : isThuNgan ? NAV_THU_NGAN : NAV_NHAN_VIEN;

  return (
    <div className="min-h-screen bg-gradient-to-b from-sky-50 via-slate-50 to-slate-50 lg:flex">
      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}`}</style>

      {/* Sidebar dọc cố định bên trái — chỉ từ breakpoint lg trở lên */}
      <aside className="hidden lg:flex lg:flex-col lg:w-60 lg:shrink-0 lg:h-screen lg:sticky lg:top-0 bg-white/90 backdrop-blur-sm border-r border-slate-200 px-3 py-4">
        <div className="flex items-center gap-2.5 px-2 mb-6">
          <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center shrink-0 shadow-sm">
            <Warehouse size={18} />
          </div>
          <div className="min-w-0">
            <p className="text-sm font-semibold text-slate-800 leading-tight tracking-tight">P&amp;L The Eros 143</p>
            <p className="text-[11px] text-slate-400 leading-tight truncate">Nhập - Xuất - Tồn kho</p>
          </div>
        </div>
        <nav className="flex-1 space-y-1 overflow-y-auto">
          {navItems.map((n) => (
            <button key={n.key} onClick={() => setTab(n.key)} className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl text-sm font-medium text-left transition ${tab === n.key ? "bg-sky-800 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100 hover:text-slate-700"}`}>
              <n.icon size={16} className="shrink-0" /> {n.label}
            </button>
          ))}
        </nav>
        <div className="pt-3 mt-3 border-t border-slate-200 flex items-center gap-2">
          <div className="w-8 h-8 rounded-full bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center text-xs font-semibold shrink-0">
            {currentUser.name?.trim()?.split(" ").slice(-1)[0]?.[0] || "?"}
          </div>
          <div className="min-w-0 flex-1 leading-tight">
            <p className="text-xs font-medium text-slate-800 truncate">{currentUser.name}</p>
            <p className="text-[11px] text-slate-400 truncate">{ROLE_META[currentUser.role]?.label}</p>
          </div>
          <button onClick={() => setShowChangePassword(true)} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition shrink-0" title="Đổi mật khẩu">
            <Lock size={14} className="text-slate-500" />
          </button>
          {canViewReports && <NotificationBell notifications={data.notifications} onMarkRead={markNotificationRead} onMarkAllRead={markAllNotificationsRead} />}
          <button onClick={handleLogout} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition shrink-0" title="Đăng xuất">
            <LogOut size={14} className="text-slate-500" />
          </button>
        </div>
      </aside>

      <div className="flex-1 min-w-0 pb-20 lg:pb-0">
        {/* Header trên cùng — chỉ hiện khi chưa có sidebar (dưới breakpoint lg) */}
        <div className="lg:hidden bg-white/90 backdrop-blur border-b border-slate-200 sticky top-0 z-30">
          <div className="px-4 py-3 flex items-center justify-between gap-3">
            <div className="flex items-center gap-2.5 min-w-0">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center shrink-0 shadow-sm">
                <Warehouse size={18} />
              </div>
              <div className="min-w-0">
                <p className="text-sm font-semibold text-slate-800 leading-tight tracking-tight">P&amp;L The Eros 143</p>
                <p className="text-[11px] text-slate-400 leading-tight truncate">Nhập - Xuất - Tồn kho</p>
              </div>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <div className="flex items-center gap-2 pl-2 border-l border-slate-200">
                <div className="w-8 h-8 rounded-full bg-gradient-to-br from-sky-700 to-sky-900 text-white flex items-center justify-center text-xs font-semibold shrink-0">
                  {currentUser.name?.trim()?.split(" ").slice(-1)[0]?.[0] || "?"}
                </div>
                <div className="hidden md:block text-right leading-tight">
                  <p className="text-xs font-medium text-slate-800">{currentUser.name}</p>
                  <p className="text-[11px] text-slate-400">{ROLE_META[currentUser.role]?.label}</p>
                </div>
                <button onClick={() => setShowChangePassword(true)} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition" title="Đổi mật khẩu">
                  <Lock size={14} className="text-slate-500" />
                </button>
                {canViewReports && <NotificationBell notifications={data.notifications} onMarkRead={markNotificationRead} onMarkAllRead={markAllNotificationsRead} />}
                <button onClick={handleLogout} className="w-8 h-8 rounded-xl border border-slate-200 flex items-center justify-center hover:bg-slate-50 transition" title="Đăng xuất">
                  <LogOut size={14} className="text-slate-500" />
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="max-w-6xl mx-auto px-4 py-6">
          <TabErrorBoundary resetKey={tab}>
            {tab === "nhap" && !isBaoCao && <NhapHangModule data={data} currentUser={currentUser} onSubmit={submitImport} onBulkImport={bulkImportNhap} />}
            {tab === "lich_su_nhap" && <LichSuNhapModule data={data} onDelete={deleteImportRecord} onDeleteMany={deleteImportRecordsByIds} />}
            {tab === "xuat" && <XuatHangModule data={data} currentUser={currentUser} onSubmit={submitExport} onBulkImportFromBills={bulkImportXuatFromBills} onDelete={deleteExportRecord} onDeleteMany={deleteExportRecordsByIds} />}
            {tab === "chi_phi" && <ChiPhiModule data={data} currentUser={currentUser} onSubmitExpense={submitExpense} onSubmitImport={submitImport} onDeleteExpense={deleteExpenseRecord} onUpdateExpense={updateExpenseRecord} />}
            {tab === "coc" && (canViewReports || isThuNgan) && <DepositModule data={data} currentUser={currentUser} onSubmit={submitDeposit} onUpdate={updateDeposit} onDelete={deleteDeposit} />}
            {tab === "quy" && canViewReports && <QuyModule data={data} currentUser={currentUser} onBulkImportInvoiceRevenue={bulkImportInvoiceRevenue} onUpsertOpeningBalance={upsertFundOpeningBalance} onUpsertRemittedAmount={upsertFundRemittedAmount} onSetThuNganEditEnabled={setThuNganEditEnabled} />}
            {tab === "thu" && isThuNgan && <ThuModule data={data} currentUser={currentUser} onSubmit={submitCashierReceipt} onUpdate={updateCashierReceipt} onDelete={deleteCashierReceipt} editEnabled={data.settings?.thu_ngan_edit_enabled !== "false"} />}
            {tab === "chi" && isThuNgan && <ChiPhieuModule data={data} currentUser={currentUser} onSubmit={submitExpense} onUpdate={updateExpenseRecord} onDelete={deleteExpenseRecord} editEnabled={data.settings?.thu_ngan_edit_enabled !== "false"} />}
            {tab === "mon_an" && <MonAnModule data={data} onAddDish={addDish} onSaveRecipe={saveDishRecipe} onDeleteDish={deleteDish} />}
            {tab === "danh_muc" && !isBaoCao && (
              <DanhMucModule data={data} onAddSupplier={addSupplier} onAddProduct={addProduct} onAddRevenueCode={addRevenueCode} onAddExportCode={addExportCode} />
            )}
            {tab === "bao_cao_nhap" && canViewReports && <BaoCaoNhapModule data={data} />}
            {tab === "bao_cao_xuat" && canViewReports && <BaoCaoXuatModule data={data} />}
            {tab === "ton_kho" && !isBaoCao && <TonKhoModule data={data} currentUser={currentUser} onSaveOpening={saveStockOpening} />}
            {tab === "tai_khoan" && isQuanLy && <TaiKhoanModule currentUser={currentUser} employees={data.employees} onAddEmployee={addEmployee} />}
          </TabErrorBoundary>
        </div>
      </div>

      {/* Thanh điều hướng cố định đáy màn hình, cuộn ngang — chỉ hiện dưới breakpoint lg */}
      <nav className="lg:hidden fixed bottom-0 inset-x-0 z-30 bg-white/95 backdrop-blur border-t border-slate-200 flex items-center gap-1 px-2 py-2 overflow-x-auto">
        {navItems.map((n) => (
          <button key={n.key} onClick={() => setTab(n.key)} className={`flex flex-col items-center gap-0.5 px-3 py-1.5 rounded-xl text-[11px] font-medium whitespace-nowrap shrink-0 transition ${tab === n.key ? "bg-sky-800 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100"}`}>
            <n.icon size={16} /> {n.label}
          </button>
        ))}
      </nav>

      {showChangePassword && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5">
            <div className="flex items-center justify-between mb-3">
              <p className="font-semibold text-slate-800 flex items-center gap-2"><Lock size={17} className="text-sky-700" /> Đổi mật khẩu</p>
              <button onClick={() => setShowChangePassword(false)}><X size={18} className="text-slate-400" /></button>
            </div>
            <ChangePasswordForm currentUser={currentUser} onCancel={() => setShowChangePassword(false)} onSuccess={() => { setShowChangePassword(false); showToast("Đã đổi mật khẩu thành công"); }} />
          </div>
        </div>
      )}
      <Toast toast={toast} />
    </div>
  );
}
