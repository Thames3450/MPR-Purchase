const SUPABASE_URL = "https://pxhskhfieiozbabqoefl.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB4aHNraGZpZWlvemJhYnFvZWZsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1NDEyMDAsImV4cCI6MjA5NDExNzIwMH0.zXUH3Ptj3kAtxJhuMvXBBfLoaqojpysMpASs4CLW9ys";


const USE_SUPABASE = true;
const STORAGE_BUCKET = "purchase-images";
const STORAGE_KEY = "mpr_purchase_requests_v5";

const REQUESTER_STORAGE_KEY = "mpr_purchase_requesters_v1";
const DEPARTMENT_STORAGE_KEY = "mpr_purchase_departments_v1";

const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let selectedImageFile = null;

let state = {
  lang: "th",
  requests: [],
  masters: [],
  currentPage: "dashboard",
};

const statusList = [
  "Pending",
  "Approved",
  "Sent to Purchasing",
  "Qty Adjusted",
  "PO Issued",
  "Partially Ordered",
  "Ordered",
  "Partially Received",
  "Received",
  "Cancelled",
];

const masterConfig = {
  requester: {
    storageKey: REQUESTER_STORAGE_KEY,
    selectId: "requester",
    listId: "requesterSettingList",
    inputId: "newRequesterInput",
    placeholderKey: "selectRequester",
    defaultList: [
      "Apiwut",
      "Chaiphat",
      "Pang Wei",
      "Ratthathummanun",
      "Thammarak",
    ],
  },

  department: {
    storageKey: DEPARTMENT_STORAGE_KEY,
    selectId: "department",
    listId: "departmentSettingList",
    inputId: "newDepartmentInput",
    placeholderKey: "selectDepartment",
    defaultList: [
      "MVR",
      "MSR",
      "Maintenance",
    ],
  },
};

const i18n = {
  th: {
    systemSubtitle: "ระบบขอสั่งซื้ออะไหล่",
    dashboard: "แดชบอร์ด",
    dashboardDesc: "ภาพรวมคำขอสั่งซื้ออะไหล่",
    dashboardHeroTitle: "ระบบควบคุมคำขอสั่งซื้ออะไหล่",
    dashboardHeroDesc: "รวมรายการขอสั่งซื้อ ติดตามสถานะจัดซื้อ และส่งออกข้อมูลให้จัดซื้อได้เป็นระบบ",

    newRequest: "เพิ่มคำขอสั่งซื้อ",
    newRequestDesc: "กรอกเฉพาะข้อมูลสำคัญสำหรับส่งจัดซื้อ",
    requestList: "รายการคำขอ",
    requestListDesc: "ค้นหา กรอง และอัปเดตสถานะคำขอ",
    purchasingFollowup: "ติดตามจัดซื้อ",
    purchasingFollowupDesc: "ติดตามรายการที่ส่งจัดซื้อแล้ว / PO / Ordered / Received",
    purchaseConsolidation: "รวมรายการสั่งซื้อ",
    purchaseConsolidationDesc: "รวมรายการอะไหล่ที่มีชื่อ / รุ่น / ยี่ห้อเดียวกัน เพื่อให้จัดซื้อสั่งซื้อได้ง่ายขึ้น",
    monthlySummary: "สรุปรายเดือน",
    monthlySummaryDesc: "รวมยอดคำขอสั่งซื้อแยกตามเดือน",
    export: "ส่งออกข้อมูล",
    exportDesc: "เลือกข้อมูลที่ต้องการ แล้วส่งออกเป็นไฟล์ Excel พร้อมรูปอะไหล่สำหรับส่งต่อให้จัดซื้อ",
    settings: "ตั้งค่า",
    settingsDesc: "จัดการรายการ Dropdown ที่ใช้ในฟอร์มขอสั่งซื้อ เช่น ผู้ขอ และแผนก",

    totalRequests: "คำขอทั้งหมด",
    pending: "รอดำเนินการ",
    urgentItems: "รายการด่วน",
    received: "ได้รับของแล้ว",
    latestRequests: "รายการล่าสุด",
    latestRequestsDesc: "คำขอสั่งซื้อที่เพิ่มล่าสุด",
    urgentList: "รายการเร่งด่วน",
    urgentListDesc: "อะไหล่ที่ควรเร่งติดตาม",
    viewAll: "ดูทั้งหมด",

    requestNo: "เลขคำขอ",
    requestDate: "วันที่ขอ",
    requester: "ผู้ขอ",
    requesterSettingDesc: "จัดการรายชื่อผู้ขอในฟอร์ม",
    department: "แผนก",
    departmentSettingDesc: "จัดการรายการแผนก",
    machine: "เครื่องจักร",
    partName: "ชื่ออะไหล่",
    model: "รุ่น / Model",
    brand: "ยี่ห้อ",
    qty: "จำนวน",
    requestQty: "จำนวนที่ขอ",
    urgency: "ความเร่งด่วน",
    reason: "เหตุผลการสั่งซื้อ",
    partImage: "รูปภาพอะไหล่",
    addPartImage: "เพิ่มรูปอะไหล่",
    imageHint: "รองรับ JPG / PNG / รูปจากมือถือ",
    removeImage: "ลบรูป",
    clearForm: "ล้างฟอร์ม",
    saveRequest: "บันทึกคำขอ",

    status: "สถานะ",
    allStatus: "ทุกสถานะ",
    reset: "รีเซ็ต",
    sentToPurchasing: "ส่งจัดซื้อแล้ว",
    add: "เพิ่ม",

    selectRequester: "เลือกผู้ขอ",
    selectDepartment: "เลือกแผนก",

    followupSearchPlaceholder: "ค้นหา PR / อะไหล่ / เครื่องจักร / PO",
    requestSearchPlaceholder: "ค้นหา PR / ชื่ออะไหล่ / เครื่องจักร / รุ่น",
    consolidationSearchPlaceholder: "ค้นหา Model / Part Name / Brand / Department",
    activeItemsOnly: "เฉพาะรายการที่ยังไม่จบ",
    allItems: "ทุกรายการ",

    noPurchasingItems: "ยังไม่มีรายการที่ส่งจัดซื้อ",
    noConsolidationItems: "ยังไม่มีรายการสำหรับรวมยอด",

    supportedFiles: "รูปแบบไฟล์",
    exportFilterTitle: "เลือกข้อมูลที่ต้องการ Export",
    exportFilterDesc: "สามารถเลือกทั้งหมด / รายเดือน / แผนก / สถานะ และเลือกรวมรูปภาพได้",
    exportMonth: "เดือนที่ต้องการ",
    exportMonthHint: "ปล่อยว่าง = Export ทุกเดือน",
    includeImageInExcel: "ใส่รูปภาพลงใน Excel",
    includeImageHint: "ถ้ารูปโหลดไม่ได้ ระบบจะใส่ลิงก์รูปแทน",
    exportCountLabel: "จำนวนรายการที่จะ Export",

    exportAll: "Export ทั้งหมด",
    exportAllDesc: "ล้างตัวกรองและส่งออกข้อมูลทั้งหมด",
    exportThisMonth: "Export เดือนนี้",
    exportThisMonthDesc: "ดึงเฉพาะรายการของเดือนปัจจุบัน",
    exportUrgent: "Export รายการด่วน",
    exportUrgentDesc: "ดึงเฉพาะ Urgent / Critical",
    exportPurchasing: "Export ส่งจัดซื้อแล้ว",
    exportPurchasingDesc: "ดึงเฉพาะรายการที่อยู่ในขั้นตอนจัดซื้อ",
    exportConsolidated: "Export รวมตาม Model",
    exportConsolidatedDesc: "รวมยอดอะไหล่ Model เดียวกันสำหรับจัดซื้อ",

    recommendation: "คำแนะนำ",
    exportNote: "ถ้าต้องส่งให้จัดซื้อ แนะนำเลือกสถานะ Approved หรือ Sent to Purchasing และเปิดตัวเลือกใส่รูปภาพลงใน Excel เพื่อให้จัดซื้อเห็นอะไหล่ชัดเจน",

    normal: "ปกติ",
    urgent: "ด่วน",
    critical: "ด่วนมาก",
    noData: "ยังไม่มีข้อมูล",
  },

  zh: {
    systemSubtitle: "维修备件采购申请系统",
    dashboard: "仪表板",
    dashboardDesc: "备件采购申请总览",
    dashboardHeroTitle: "维修备件采购申请管理系统",
    dashboardHeroDesc: "集中管理采购申请、采购进度跟进，并可导出数据给采购部门",

    newRequest: "新增采购申请",
    newRequestDesc: "填写关键采购信息",
    requestList: "申请列表",
    requestListDesc: "搜索、筛选并更新申请状态",
    purchasingFollowup: "采购跟进",
    purchasingFollowupDesc: "跟进已提交采购 / PO / 已下单 / 已收货",
    purchaseConsolidation: "合并采购项目",
    purchaseConsolidationDesc: "将相同名称 / 型号 / 品牌的备件合并，方便采购统一下单",
    monthlySummary: "月度汇总",
    monthlySummaryDesc: "按月份汇总采购申请",
    export: "导出数据",
    exportDesc: "选择需要导出的数据，并导出为包含备件图片的 Excel 文件",
    settings: "设置",
    settingsDesc: "管理采购申请表单中的下拉选项，例如申请人和部门",

    totalRequests: "申请总数",
    pending: "待处理",
    urgentItems: "紧急项目",
    received: "已收货",
    latestRequests: "最新申请",
    latestRequestsDesc: "最近新增的采购申请",
    urgentList: "紧急清单",
    urgentListDesc: "需要优先跟进的备件",
    viewAll: "查看全部",

    requestNo: "申请编号",
    requestDate: "申请日期",
    requester: "申请人",
    requesterSettingDesc: "管理申请人列表",
    department: "部门",
    departmentSettingDesc: "管理部门列表",
    machine: "机器",
    partName: "备件名称",
    model: "型号",
    brand: "品牌",
    qty: "数量",
    requestQty: "申请数量",
    urgency: "紧急程度",
    reason: "采购原因",
    partImage: "备件图片",
    addPartImage: "添加备件图片",
    imageHint: "支持 JPG / PNG / 手机照片",
    removeImage: "删除图片",
    clearForm: "清空表单",
    saveRequest: "保存申请",

    status: "状态",
    allStatus: "全部状态",
    reset: "重置",
    sentToPurchasing: "已提交采购",
    add: "添加",

    selectRequester: "选择申请人",
    selectDepartment: "选择部门",

    followupSearchPlaceholder: "搜索 PR / 备件 / 机器 / PO",
    requestSearchPlaceholder: "搜索 PR / 备件名称 / 机器 / 型号",
    consolidationSearchPlaceholder: "搜索型号 / 备件名称 / 品牌 / 部门",
    activeItemsOnly: "仅显示未完成项目",
    allItems: "全部项目",

    noPurchasingItems: "暂无已提交采购的项目",
    noConsolidationItems: "暂无可合并的项目",

    supportedFiles: "文件格式",
    exportFilterTitle: "选择需要导出的数据",
    exportFilterDesc: "可按月份 / 部门 / 状态筛选，并选择是否包含图片",
    exportMonth: "月份",
    exportMonthHint: "留空 = 导出全部月份",
    includeImageInExcel: "在 Excel 中包含图片",
    includeImageHint: "如果图片无法加载，将保留图片链接",
    exportCountLabel: "将导出的项目数量",

    exportAll: "导出全部",
    exportAllDesc: "清除筛选并导出全部数据",
    exportThisMonth: "导出本月",
    exportThisMonthDesc: "仅导出当前月份数据",
    exportUrgent: "导出紧急项目",
    exportUrgentDesc: "仅导出 Urgent / Critical",
    exportPurchasing: "导出采购跟进",
    exportPurchasingDesc: "仅导出采购流程中的项目",
    exportConsolidated: "按型号合并导出",
    exportConsolidatedDesc: "合并相同型号的备件，方便采购统一处理",

    recommendation: "建议",
    exportNote: "正式提交给采购部门建议选择 Approved 或 Sent to Purchasing，并开启图片导出，便于采购人员确认备件",

    normal: "普通",
    urgent: "紧急",
    critical: "非常紧急",
    noData: "暂无数据",
  },
};

document.addEventListener("DOMContentLoaded", async () => {
  setToday();
  bindEvents();
  updateLanguage();

  await loadMasterData();
  await loadData();

  refreshAll();
});

function bindEvents() {
  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      goPage(btn.dataset.page);
      closeMobileMenu();
    });
  });

  document.querySelectorAll(".lang-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.lang = btn.dataset.lang;
      updateLanguage();
      refreshAll();
    });
  });

  document.getElementById("requestForm").addEventListener("submit", saveRequest);
  document.getElementById("clearFormBtn").addEventListener("click", clearForm);

  document.getElementById("addRequesterSettingBtn").addEventListener("click", () => addMasterItem("requester"));
  document.getElementById("addDepartmentSettingBtn").addEventListener("click", () => addMasterItem("department"));

  document.getElementById("imageUploadBtn").addEventListener("click", () => {
    document.getElementById("imageFile").click();
  });

  document.getElementById("imageFile").addEventListener("change", handleImageUpload);
  document.getElementById("removeImageBtn").addEventListener("click", removeSelectedImage);

  document.getElementById("searchInput").addEventListener("input", renderRequestList);
  document.getElementById("statusFilter").addEventListener("change", renderRequestList);
  document.getElementById("monthFilter").addEventListener("change", renderRequestList);
  document.getElementById("resetFilterBtn").addEventListener("click", resetFilters);

  document.getElementById("followupSearch").addEventListener("input", renderPurchasingFollowup);
  document.getElementById("followupStatusFilter").addEventListener("change", renderPurchasingFollowup);
  document.getElementById("followupMonthFilter").addEventListener("change", renderPurchasingFollowup);

  document.getElementById("resetFollowupBtn").addEventListener("click", () => {
    document.getElementById("followupSearch").value = "";
    document.getElementById("followupStatusFilter").value = "All";
    document.getElementById("followupMonthFilter").value = "";
    renderPurchasingFollowup();
  });

  document.getElementById("consolidationSearch").addEventListener("input", renderConsolidationPage);
  document.getElementById("consolidationStatusFilter").addEventListener("change", renderConsolidationPage);
  document.getElementById("consolidationMonthFilter").addEventListener("change", renderConsolidationPage);

  document.getElementById("resetConsolidationBtn").addEventListener("click", () => {
    document.getElementById("consolidationSearch").value = "";
    document.getElementById("consolidationStatusFilter").value = "Active";
    document.getElementById("consolidationMonthFilter").value = "";
    renderConsolidationPage();
  });

  document.getElementById("summaryMonth").addEventListener("change", renderMonthlySummary);

  document.getElementById("exportExcelBtn").addEventListener("click", exportExcelAdvanced);
  document.getElementById("exportMonthFilter").addEventListener("change", updateExportPreview);
  document.getElementById("exportDepartmentFilter").addEventListener("change", updateExportPreview);
  document.getElementById("exportStatusFilter").addEventListener("change", updateExportPreview);
  document.getElementById("exportIncludeImage").addEventListener("change", updateExportPreview);

  document.getElementById("mobileMenuBtn").addEventListener("click", openMobileMenu);
  document.getElementById("sidebarOverlay").addEventListener("click", closeMobileMenu);

  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("detailModal").addEventListener("click", (e) => {
    if (e.target.id === "detailModal") closeModal();
  });
}

/* =========================
   Supabase Master Dropdown Settings
========================= */

async function loadMasterData() {
  if (!USE_SUPABASE) {
    renderMasterDropdowns();
    renderSettingsLists();
    return;
  }

  const { data, error } = await supabaseClient
    .from("purchase_masters")
    .select("*")
    .eq("is_active", true)
    .order("sort_order", { ascending: true })
    .order("name", { ascending: true });

  if (error) {
    console.error("loadMasterData error:", error);
    toast(`โหลด Dropdown ไม่สำเร็จ: ${error.message}`);
    state.masters = [];
    renderMasterDropdowns();
    renderSettingsLists();
    return;
  }

  state.masters = data || [];

  renderMasterDropdowns();
  renderSettingsLists();
}

function getMasterList(type) {
  if (USE_SUPABASE) {
    return (state.masters || [])
      .filter((item) => item.type === type && item.is_active !== false)
      .map((item) => item.name);
  }

  const config = masterConfig[type];
  if (!config) return [];

  const raw = localStorage.getItem(config.storageKey);

  if (raw) {
    try {
      return JSON.parse(raw);
    } catch {
      return config.defaultList;
    }
  }

  saveMasterList(type, config.defaultList);
  return config.defaultList;
}

function saveMasterList(type, list) {
  if (USE_SUPABASE) return;

  const config = masterConfig[type];
  if (!config) return;

  localStorage.setItem(config.storageKey, JSON.stringify(list));
}

function renderMasterDropdowns() {
  Object.keys(masterConfig).forEach((type) => renderMasterDropdown(type));
}

function renderMasterDropdown(type, selectedValue = "") {
  const config = masterConfig[type];
  if (!config) return;

  const select = document.getElementById(config.selectId);
  if (!select) return;

  const currentValue = selectedValue || select.value;
  const list = getMasterList(type);
  const placeholderText = t(config.placeholderKey);

  select.innerHTML = `
    <option value="">${escapeHTML(placeholderText)}</option>
    ${list.map((item) => `
      <option value="${escapeHTML(item)}">${escapeHTML(item)}</option>
    `).join("")}
  `;

  if (currentValue && list.includes(currentValue)) {
    select.value = currentValue;
  }
}

function renderSettingsLists() {
  Object.keys(masterConfig).forEach((type) => renderSettingsList(type));
}

function renderSettingsList(type) {
  const config = masterConfig[type];
  if (!config) return;

  const listBox = document.getElementById(config.listId);
  if (!listBox) return;

  let items = [];

  if (USE_SUPABASE) {
    items = (state.masters || [])
      .filter((item) => item.type === type && item.is_active !== false);
  } else {
    items = getMasterList(type).map((name) => ({ name }));
  }

  if (!items.length) {
    listBox.innerHTML = `<div class="empty">${t("noData")}</div>`;
    return;
  }

  listBox.innerHTML = items.map((item) => `
    <div class="setting-item">
      <strong>${escapeHTML(item.name)}</strong>
      <button
        type="button"
        class="setting-delete-btn"
        onclick="deleteMasterItem('${type}', '${escapeJS(item.name)}')">
        ${t("delete") || "ลบ"}
      </button>
    </div>
  `).join("");
}

async function addMasterItem(type) {
  const config = masterConfig[type];
  if (!config) return;

  const input = document.getElementById(config.inputId);
  if (!input) return;

  const value = input.value.trim();

  if (!value) {
    toast("กรุณากรอกข้อมูลก่อนเพิ่ม");
    return;
  }

  const list = getMasterList(type);
  const isDuplicate = list.some(
    (item) => item.toLowerCase() === value.toLowerCase()
  );

  if (isDuplicate) {
    toast("มีรายการนี้อยู่แล้ว");
    return;
  }

  if (USE_SUPABASE) {
    const { error } = await supabaseClient
      .from("purchase_masters")
      .insert({
        type,
        name: value,
        is_active: true,
        sort_order: list.length + 1,
      });

    if (error) {
      console.error("addMasterItem error:", error);
      toast(`เพิ่ม Dropdown ไม่สำเร็จ: ${error.message}`);
      return;
    }

    input.value = "";
    await loadMasterData();
    renderExportFilters();
    toast("เพิ่มรายการลงฐานข้อมูลแล้ว");
    return;
  }

  list.push(value);
  list.sort();

  saveMasterList(type, list);

  input.value = "";

  renderMasterDropdown(type, value);
  renderSettingsList(type);
  renderExportFilters();

  toast("เพิ่มรายการแล้ว");
}

async function deleteMasterItem(type, value) {
  const ok = confirm(`ต้องการลบ "${value}" ใช่ไหม?`);
  if (!ok) return;

  if (USE_SUPABASE) {
    const { error } = await supabaseClient
      .from("purchase_masters")
      .update({ is_active: false })
      .eq("type", type)
      .eq("name", value);

    if (error) {
      console.error("deleteMasterItem error:", error);
      toast(`ลบ Dropdown ไม่สำเร็จ: ${error.message}`);
      return;
    }

    await loadMasterData();
    renderExportFilters();
    toast("ลบรายการจากฐานข้อมูลแล้ว");
    return;
  }

  const list = getMasterList(type).filter((item) => item !== value);

  saveMasterList(type, list);

  renderMasterDropdown(type);
  renderSettingsList(type);
  renderExportFilters();

  toast("ลบรายการแล้ว");
}

/* =========================
   Data
========================= */

async function loadData() {
  if (USE_SUPABASE) {
    const { data, error } = await supabaseClient
      .from("purchase_requests")
      .select("*")
      .order("created_at", { ascending: false });

    if (error) {
      console.error("loadData error:", error);
      toast(`โหลดข้อมูลคำขอไม่สำเร็จ: ${error.message}`);
      state.requests = [];
      return;
    }

    state.requests = data.map(mapFromSupabase);
    return;
  }

  const raw = localStorage.getItem(STORAGE_KEY);
  state.requests = raw ? JSON.parse(raw) : [];
}

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.requests));
}

function setToday() {
  const today = getTodayISO();

  const requestDate = document.getElementById("requestDate");
  if (requestDate) requestDate.value = today;

  const summaryMonth = document.getElementById("summaryMonth");
  if (summaryMonth) summaryMonth.value = today.slice(0, 7);
}

function getTodayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function generatePrNo(dateStr, number) {
  const ym = dateStr.slice(0, 7).replace("-", "");
  return `MPR-PR-${ym}-${String(number).padStart(3, "0")}`;
}

function getNextPrNo(dateStr) {
  const ym = dateStr.slice(0, 7);
  const count = state.requests.filter((r) => r.requestDate && r.requestDate.startsWith(ym)).length;
  return generatePrNo(dateStr, count + 1);
}

function updateNextPrNo() {
  const requestDate = document.getElementById("requestDate");
  const nextPrNo = document.getElementById("nextPrNo");

  if (!requestDate || !nextPrNo) return;

  const date = requestDate.value || getTodayISO();
  nextPrNo.textContent = getNextPrNo(date);
}

async function saveRequest(e) {
  e.preventDefault();

  const requestDate = getValue("requestDate") || getTodayISO();
  const prNo = getNextPrNo(requestDate);

  let uploadedImage = {
    imageUrl: "",
    imagePath: "",
  };

  try {
    if (USE_SUPABASE && selectedImageFile) {
      toast("กำลังอัปโหลดรูปภาพ...");
      uploadedImage = await uploadPurchaseImage(selectedImageFile, prNo);
    }

    const request = {
      id: crypto.randomUUID(),
      prNo,
      requestDate,
      requester: getValue("requester"),
      department: getValue("department"),
      machine: getValue("machine"),
      partName: getValue("partName"),
      model: getValue("model"),
      brand: getValue("brand"),

      qty: Number(getValue("qty")) || 1,
      approvedQty: 0,
      orderedQty: 0,
      receivedQty: 0,
      qtyAdjustReason: "",
      unitPrice: 0,
      totalAmount: 0,
      currency: "THB",

      urgency: getValue("urgency"),
      reason: getValue("reason"),
      imageUrl: uploadedImage.imageUrl,
      imagePath: uploadedImage.imagePath,
      status: "Pending",
      poNo: "",
      supplier: "",
      expectedDeliveryDate: "",
      receivedDate: "",
      purchasingNote: "",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    if (USE_SUPABASE) {
      const { data, error } = await supabaseClient
        .from("purchase_requests")
        .insert(mapToSupabase(request))
        .select()
        .single();

      if (error) {
        console.error("saveRequest error:", error);
        toast(`บันทึกคำขอไม่สำเร็จ: ${error.message}`);
        return;
      }

      state.requests.unshift(mapFromSupabase(data));
    } else {
      state.requests.unshift(request);
      saveData();
    }

    clearForm();
    refreshAll();
    toast("บันทึกคำขอสั่งซื้อแล้ว");
    goPage("requestList");
  } catch (err) {
    console.error(err);
    toast(err.message || "เกิดข้อผิดพลาด");
  }
}

function clearForm() {
  document.getElementById("requestForm").reset();

  document.getElementById("requestDate").value = getTodayISO();
  document.getElementById("qty").value = 1;

  document.getElementById("requester").value = "";
  document.getElementById("department").value = "";
  document.getElementById("machine").value = "";

  removeSelectedImage();
  updateNextPrNo();
}

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

/* =========================
   Image Upload
========================= */

function handleImageUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  if (!file.type.startsWith("image/")) {
    toast("กรุณาเลือกไฟล์รูปภาพเท่านั้น");
    return;
  }

  const maxSizeMB = 5;

  if (file.size > maxSizeMB * 1024 * 1024) {
    toast("รูปภาพต้องไม่เกิน 5MB");
    return;
  }

  selectedImageFile = file;

  const previewUrl = URL.createObjectURL(file);
  document.getElementById("imagePreview").src = previewUrl;
  document.getElementById("imagePreviewBox").classList.add("show");
}

async function uploadPurchaseImage(file, prNo) {
  if (!file) {
    return {
      imageUrl: "",
      imagePath: "",
    };
  }

  const ext = file.name.split(".").pop();
  const safePrNo = prNo.replaceAll("/", "-").replaceAll(" ", "-");
  const filePath = `${safePrNo}/${Date.now()}.${ext}`;

  const { error: uploadError } = await supabaseClient.storage
    .from(STORAGE_BUCKET)
    .upload(filePath, file, {
      cacheControl: "3600",
      upsert: true,
    });

  if (uploadError) {
    console.error("uploadPurchaseImage error:", uploadError);
    throw new Error(`อัปโหลดรูปภาพไม่สำเร็จ: ${uploadError.message}`);
  }

  const { data } = supabaseClient.storage
    .from(STORAGE_BUCKET)
    .getPublicUrl(filePath);

  return {
    imageUrl: data.publicUrl,
    imagePath: filePath,
  };
}

function removeSelectedImage() {
  selectedImageFile = null;

  document.getElementById("imageFile").value = "";
  document.getElementById("imageUrl").value = "";
  document.getElementById("imagePreview").src = "";
  document.getElementById("imagePreviewBox").classList.remove("show");
}

/* =========================
   Render
========================= */

function refreshAll() {
  updateNextPrNo();
  renderDashboard();
  renderRequestList();
  renderPurchasingFollowup();
  renderConsolidationPage();
  renderMonthlySummary();
  renderExportFilters();
  updateExportPreview();
  renderSettingsLists();
}

function renderDashboard() {
  const total = state.requests.length;
  const pending = state.requests.filter((r) => r.status === "Pending").length;
  const urgent = state.requests.filter((r) => r.urgency === "Urgent" || r.urgency === "Critical").length;
  const received = state.requests.filter((r) => r.status === "Received").length;

  setText("totalRequests", total);
  setText("pendingCount", pending);
  setText("urgentCount", urgent);
  setText("receivedCount", received);

  renderLatestTable();
  renderUrgentList();
}

function renderLatestTable() {
  const tbody = document.getElementById("latestTable");
  if (!tbody) return;

  const rows = state.requests.slice(0, 6);

  if (!rows.length) {
    tbody.innerHTML = emptyRow(6);
    return;
  }

  tbody.innerHTML = rows.map((r) => `
    <tr>
      <td><strong>${escapeHTML(r.prNo)}</strong></td>
      <td>${escapeHTML(r.partName)}</td>
      <td>${escapeHTML(r.machine)}</td>
      <td>${Number(r.qty || 0)} pcs</td>
      <td>${Number(r.orderedQty || 0)} pcs</td>
      <td>${statusBadge(r.status)}</td>
    </tr>
  `).join("");
}

function renderUrgentList() {
  const box = document.getElementById("urgentList");
  if (!box) return;

  const items = state.requests
    .filter((r) => r.urgency === "Urgent" || r.urgency === "Critical")
    .slice(0, 6);

  if (!items.length) {
    box.innerHTML = `<div class="empty">${t("noData")}</div>`;
    return;
  }

  box.innerHTML = items.map((r) => `
    <div class="urgent-card">
      <div class="request-meta">
        ${urgencyBadge(r.urgency)}
        ${statusBadge(r.status)}
      </div>
      <strong>${escapeHTML(r.partName)}</strong>
      <p>${escapeHTML(r.machine)} · ${escapeHTML(r.model || "-")} · Request ${Number(r.qty || 0)} pcs</p>
      <p>${escapeHTML(r.reason || "-")}</p>
    </div>
  `).join("");
}

function renderRequestList() {
  const list = document.getElementById("requestCardList");
  if (!list) return;

  const search = document.getElementById("searchInput").value.toLowerCase();
  const status = document.getElementById("statusFilter").value;
  const month = document.getElementById("monthFilter").value;

  let rows = [...state.requests];

  if (search) {
    rows = rows.filter((r) => {
      const text = [
        r.prNo,
        r.partName,
        r.machine,
        r.model,
        r.brand,
        r.requester,
        r.department,
        r.reason,
      ].join(" ").toLowerCase();

      return text.includes(search);
    });
  }

  if (status !== "All") {
    rows = rows.filter((r) => r.status === status);
  }

  if (month) {
    rows = rows.filter((r) => r.requestDate && r.requestDate.startsWith(month));
  }

  if (!rows.length) {
    list.innerHTML = `<div class="empty">${t("noData")}</div>`;
    return;
  }

  list.innerHTML = rows.map((r) => {
    const requestQty = Number(r.qty || 0);
    const orderedQty = Number(r.orderedQty || 0);
    const receivedQty = Number(r.receivedQty || 0);
    const remainingQty = Math.max(requestQty - orderedQty, 0);

    return `
      <div class="request-card">
        <div class="request-card-top">
          <div>
            <div class="request-meta">
              <span class="badge Approved">${escapeHTML(r.prNo)}</span>
              ${urgencyBadge(r.urgency)}
              ${statusBadge(r.status)}
            </div>
            <h4>${escapeHTML(r.partName)}</h4>
            <p>
              ${escapeHTML(r.machine)} · ${escapeHTML(r.model || "-")} · ${escapeHTML(r.brand || "-")}
            </p>
            <p>
              Department: ${escapeHTML(r.department || "-")} · Requester: ${escapeHTML(r.requester || "-")}
            </p>
          </div>

          <div class="qty-mini-summary">
            <strong>Request: ${requestQty} pcs</strong>
            <span>Ordered: ${orderedQty} pcs</span>
            <span>Received: ${receivedQty} pcs</span>
            <span>Remaining: ${remainingQty} pcs</span>
          </div>
        </div>

        ${r.imageUrl ? `
          <div class="card-image">
            <img src="${escapeHTML(r.imageUrl)}" alt="Part Image">
          </div>
        ` : ""}

        <p>${escapeHTML(r.reason || "-")}</p>

        <div class="request-actions">
          <select class="status-select" onchange="updateStatus('${r.id}', this.value)">
            ${statusList.map((s) => `<option value="${s}" ${r.status === s ? "selected" : ""}>${s}</option>`).join("")}
          </select>

          <button class="small-btn" onclick="openDetail('${r.id}')">Detail</button>
          <button class="danger-btn" onclick="deleteRequest('${r.id}')">Delete</button>
        </div>
      </div>
    `;
  }).join("");
}

function renderPurchasingFollowup() {
  const list = document.getElementById("followupList");
  if (!list) return;

  const followupStatuses = [
    "Sent to Purchasing",
    "Qty Adjusted",
    "PO Issued",
    "Partially Ordered",
    "Ordered",
    "Partially Received",
    "Received",
  ];

  const search = document.getElementById("followupSearch").value.toLowerCase();
  const status = document.getElementById("followupStatusFilter").value;
  const month = document.getElementById("followupMonthFilter").value;

  let rows = state.requests.filter((r) => followupStatuses.includes(r.status));

  if (search) {
    rows = rows.filter((r) => {
      const text = [
        r.prNo,
        r.partName,
        r.machine,
        r.model,
        r.brand,
        r.poNo,
        r.supplier,
        r.department,
      ].join(" ").toLowerCase();

      return text.includes(search);
    });
  }

  if (status !== "All") {
    rows = rows.filter((r) => r.status === status);
  }

  if (month) {
    rows = rows.filter((r) => r.requestDate && r.requestDate.startsWith(month));
  }

  if (!rows.length) {
    list.innerHTML = `<div class="empty">${t("noPurchasingItems")}</div>`;
    return;
  }

  list.innerHTML = rows.map((r) => {
    const requestQty = Number(r.qty || 0);
    const approvedQty = Number(r.approvedQty || 0);
    const orderedQty = Number(r.orderedQty || 0);
    const receivedQty = Number(r.receivedQty || 0);

    const remainingQty = Math.max(requestQty - orderedQty, 0);
    const pendingReceiveQty = Math.max(orderedQty - receivedQty, 0);

    return `
      <div class="request-card followup-card">
        <div class="request-card-top">
          <div>
            <div class="request-meta">
              <span class="badge Approved">${escapeHTML(r.prNo)}</span>
              ${statusBadge(r.status)}
              ${urgencyBadge(r.urgency)}
              ${orderedQty < requestQty && orderedQty > 0 ? `<span class="badge Partially">Partial Order</span>` : ""}
              ${receivedQty < orderedQty && orderedQty > 0 ? `<span class="badge Waiting">Waiting Receive</span>` : ""}
            </div>

            <h4>${escapeHTML(r.partName)}</h4>
            <p>${escapeHTML(r.machine)} · ${escapeHTML(r.model || "-")} · ${escapeHTML(r.brand || "-")}</p>
          </div>

          <div class="qty-mini-summary">
            <strong>Request: ${requestQty} pcs</strong>
            <span>Ordered: ${orderedQty} pcs</span>
            <span>Received: ${receivedQty} pcs</span>
          </div>
        </div>

        ${r.imageUrl ? `
          <div class="card-image">
            <img src="${escapeHTML(r.imageUrl)}" alt="Part Image">
          </div>
        ` : ""}

        <div class="purchase-qty-box">
          <div><span>Request Qty</span><strong>${requestQty}</strong></div>
          <div><span>Approved Qty</span><strong>${approvedQty}</strong></div>
          <div><span>Ordered Qty</span><strong>${orderedQty}</strong></div>
          <div><span>Received Qty</span><strong>${receivedQty}</strong></div>
          <div><span>Remaining</span><strong>${remainingQty}</strong></div>
          <div><span>Pending Receive</span><strong>${pendingReceiveQty}</strong></div>
        </div>

        <div class="followup-form purchase-form-grid">
          <div class="field">
            <label>Approved Qty</label>
            <input type="number" min="0" value="${approvedQty}" onchange="updatePurchasingField('${r.id}', 'approvedQty', this.value)">
          </div>

          <div class="field">
            <label>Ordered Qty</label>
            <input type="number" min="0" value="${orderedQty}" onchange="updatePurchasingField('${r.id}', 'orderedQty', this.value)">
          </div>

          <div class="field">
            <label>Received Qty</label>
            <input type="number" min="0" value="${receivedQty}" onchange="updatePurchasingField('${r.id}', 'receivedQty', this.value)">
          </div>

          <div class="field">
            <label>Unit Price</label>
            <input type="number" min="0" step="0.01" value="${Number(r.unitPrice || 0)}" onchange="updatePurchasingField('${r.id}', 'unitPrice', this.value)">
          </div>

          <div class="field">
            <label>PO No.</label>
            <input type="text" value="${escapeHTML(r.poNo || "")}" onchange="updatePurchasingField('${r.id}', 'poNo', this.value)" placeholder="ใส่เลข PO">
          </div>

          <div class="field">
            <label>Supplier</label>
            <input type="text" value="${escapeHTML(r.supplier || "")}" onchange="updatePurchasingField('${r.id}', 'supplier', this.value)" placeholder="ชื่อ Supplier">
          </div>

          <div class="field">
            <label>Expected Delivery</label>
            <input type="date" value="${escapeHTML(r.expectedDeliveryDate || "")}" onchange="updatePurchasingField('${r.id}', 'expectedDeliveryDate', this.value)">
          </div>

          <div class="field">
            <label>Received Date</label>
            <input type="date" value="${escapeHTML(r.receivedDate || "")}" onchange="updatePurchasingField('${r.id}', 'receivedDate', this.value)">
          </div>
        </div>

        <div class="field">
          <label>Qty Adjust Reason</label>
          <textarea rows="2" onchange="updatePurchasingField('${r.id}', 'qtyAdjustReason', this.value)" placeholder="กรอกเหตุผล ถ้าอนุมัติ/สั่งซื้อน้อยกว่าจำนวนที่ขอ">${escapeHTML(r.qtyAdjustReason || "")}</textarea>
        </div>

        <div class="field">
          <label>Purchasing Note</label>
          <textarea rows="2" onchange="updatePurchasingField('${r.id}', 'purchasingNote', this.value)" placeholder="หมายเหตุจากจัดซื้อ">${escapeHTML(r.purchasingNote || "")}</textarea>
        </div>

        <div class="request-actions">
          <select class="status-select" onchange="updateStatus('${r.id}', this.value)">
            ${statusList.map((s) => `<option value="${s}" ${r.status === s ? "selected" : ""}>${s}</option>`).join("")}
          </select>

          <button class="small-btn" onclick="autoUpdatePurchaseStatus('${r.id}')">Auto Status</button>
          <button class="small-btn" onclick="openDetail('${r.id}')">Detail</button>
        </div>
      </div>
    `;
  }).join("");
}

function renderMonthlySummary() {
  const summaryMonth = document.getElementById("summaryMonth");
  if (!summaryMonth) return;

  const month = summaryMonth.value || getTodayISO().slice(0, 7);
  const rows = state.requests.filter((r) => r.requestDate && r.requestDate.startsWith(month));

  const total = rows.length;
  const urgent = rows.filter((r) => r.urgency === "Urgent" || r.urgency === "Critical").length;
  const sent = rows.filter((r) =>
    ["Sent to Purchasing", "Qty Adjusted", "PO Issued", "Partially Ordered", "Ordered", "Partially Received", "Received"].includes(r.status)
  ).length;
  const received = rows.filter((r) => r.status === "Received").length;

  setText("monthTotal", total);
  setText("monthUrgent", urgent);
  setText("monthSent", sent);
  setText("monthReceived", received);

  const tbody = document.getElementById("monthlyTable");

  if (!tbody) return;

  if (!rows.length) {
    tbody.innerHTML = emptyRow(9);
    return;
  }

  tbody.innerHTML = rows.map((r) => `
    <tr>
      <td><strong>${escapeHTML(r.prNo)}</strong></td>
      <td>${formatDate(r.requestDate)}</td>
      <td>${escapeHTML(r.machine)}</td>
      <td>${escapeHTML(r.partName)}</td>
      <td>${Number(r.qty || 0)} pcs</td>
      <td>${Number(r.orderedQty || 0)} pcs</td>
      <td>${Number(r.receivedQty || 0)} pcs</td>
      <td>${urgencyBadge(r.urgency)}</td>
      <td>${statusBadge(r.status)}</td>
    </tr>
  `).join("");
}

/* =========================
   Purchase Consolidation
========================= */

function normalizeText(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function buildPartGroupKey(request) {
  const brand = normalizeText(request.brand || "NO BRAND");
  const model = normalizeText(request.model || "NO MODEL");
  const partName = normalizeText(request.partName || "NO PART NAME");

  return `${brand} | ${model} | ${partName}`;
}

function getConsolidatedGroups(sourceRows = state.requests) {
  const groups = new Map();

  sourceRows.forEach((request) => {
    const key = buildPartGroupKey(request);

    if (!groups.has(key)) {
      groups.set(key, {
        groupKey: key,
        partName: request.partName || "-",
        model: request.model || "-",
        brand: request.brand || "-",
        imageUrl: request.imageUrl || "",

        requestQty: 0,
        approvedQty: 0,
        orderedQty: 0,
        receivedQty: 0,
        remainingQty: 0,
        pendingReceiveQty: 0,
        totalAmount: 0,

        departments: {},
        statuses: {},
        requests: [],
      });
    }

    const group = groups.get(key);

    const requestQty = Number(request.qty || 0);
    const approvedQty = Number(request.approvedQty || 0);
    const orderedQty = Number(request.orderedQty || 0);
    const receivedQty = Number(request.receivedQty || 0);
    const totalAmount = Number(request.totalAmount || 0);

    group.requestQty += requestQty;
    group.approvedQty += approvedQty;
    group.orderedQty += orderedQty;
    group.receivedQty += receivedQty;
    group.totalAmount += totalAmount;

    if (!group.imageUrl && request.imageUrl) {
      group.imageUrl = request.imageUrl;
    }

    const department = request.department || "Unknown";

    if (!group.departments[department]) {
      group.departments[department] = {
        requestQty: 0,
        approvedQty: 0,
        orderedQty: 0,
        receivedQty: 0,
        count: 0,
      };
    }

    group.departments[department].requestQty += requestQty;
    group.departments[department].approvedQty += approvedQty;
    group.departments[department].orderedQty += orderedQty;
    group.departments[department].receivedQty += receivedQty;
    group.departments[department].count += 1;

    const status = request.status || "Pending";
    group.statuses[status] = (group.statuses[status] || 0) + 1;

    group.requests.push(request);
  });

  const result = Array.from(groups.values()).map((group) => {
    group.remainingQty = Math.max(group.requestQty - group.orderedQty, 0);
    group.pendingReceiveQty = Math.max(group.orderedQty - group.receivedQty, 0);
    return group;
  });

  result.sort((a, b) => {
    if (b.remainingQty !== a.remainingQty) return b.remainingQty - a.remainingQty;
    if (b.requestQty !== a.requestQty) return b.requestQty - a.requestQty;
    return a.partName.localeCompare(b.partName);
  });

  return result;
}

function getConsolidationFilteredRows() {
  const search = document.getElementById("consolidationSearch")?.value.toLowerCase() || "";
  const status = document.getElementById("consolidationStatusFilter")?.value || "Active";
  const month = document.getElementById("consolidationMonthFilter")?.value || "";

  let rows = [...state.requests];

  if (status === "Active") {
    rows = rows.filter((r) => !["Received", "Cancelled"].includes(r.status));
  } else if (status !== "All") {
    rows = rows.filter((r) => r.status === status);
  }

  if (month) {
    rows = rows.filter((r) => r.requestDate && r.requestDate.startsWith(month));
  }

  if (search) {
    rows = rows.filter((r) => {
      const text = [
        r.prNo,
        r.partName,
        r.model,
        r.brand,
        r.department,
        r.machine,
        r.requester,
        r.status,
      ].join(" ").toLowerCase();

      return text.includes(search);
    });
  }

  return rows;
}

function renderConsolidationPage() {
  const list = document.getElementById("consolidationList");
  if (!list) return;

  const rows = getConsolidationFilteredRows();
  const groups = getConsolidatedGroups(rows);

  const totalRequestQty = groups.reduce((sum, g) => sum + Number(g.requestQty || 0), 0);
  const totalOrderedQty = groups.reduce((sum, g) => sum + Number(g.orderedQty || 0), 0);
  const totalRemainingQty = groups.reduce((sum, g) => sum + Number(g.remainingQty || 0), 0);

  setText("consolidatedGroupCount", groups.length);
  setText("consolidatedRequestQty", totalRequestQty);
  setText("consolidatedOrderedQty", totalOrderedQty);
  setText("consolidatedRemainingQty", totalRemainingQty);

  if (!groups.length) {
    list.innerHTML = `<div class="empty">${t("noConsolidationItems")}</div>`;
    return;
  }

  list.innerHTML = groups.map((group) => {
    const departmentRows = Object.entries(group.departments)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([department, data]) => `
        <div class="department-row">
          <strong>${escapeHTML(department)}</strong>
          <span>Request: ${data.requestQty} pcs</span>
          <span>Ordered: ${data.orderedQty} pcs</span>
          <span>Received: ${data.receivedQty} pcs</span>
        </div>
      `).join("");

    const statusBadges = Object.entries(group.statuses)
      .map(([status, count]) => `
        <span class="badge ${getStatusBadgeClass(status)}">${escapeHTML(status)}: ${count}</span>
      `).join("");

    const requestRows = group.requests.map((r) => {
      const requestQty = Number(r.qty || 0);
      const orderedQty = Number(r.orderedQty || 0);
      const receivedQty = Number(r.receivedQty || 0);
      const remainingQty = Math.max(requestQty - orderedQty, 0);

      return `
        <tr>
          <td><strong>${escapeHTML(r.prNo)}</strong></td>
          <td>${formatDate(r.requestDate)}</td>
          <td>${escapeHTML(r.department || "-")}</td>
          <td>${escapeHTML(r.requester || "-")}</td>
          <td>${escapeHTML(r.machine || "-")}</td>
          <td>${requestQty}</td>
          <td>${orderedQty}</td>
          <td>${receivedQty}</td>
          <td>${remainingQty}</td>
          <td>${statusBadge(r.status)}</td>
        </tr>
      `;
    }).join("");

    return `
      <div class="consolidation-card">
        <div class="consolidation-card-top">
          <div class="consolidation-title">
            <div class="request-meta">
              ${statusBadges}
            </div>
            <h4>${escapeHTML(group.partName)}</h4>
            <p>
              Brand: <strong>${escapeHTML(group.brand)}</strong>
              · Model: <strong>${escapeHTML(group.model)}</strong>
            </p>
          </div>

          <div class="consolidation-total">
            <strong>${group.requestQty} pcs</strong>
            <span>Total Request Qty</span>
          </div>
        </div>

        ${group.imageUrl ? `
          <div class="card-image">
            <img src="${escapeHTML(group.imageUrl)}" alt="Part Image">
          </div>
        ` : ""}

        <div class="consolidation-qty-grid">
          <div><span>Total Request</span><strong>${group.requestQty}</strong></div>
          <div><span>Total Approved</span><strong>${group.approvedQty}</strong></div>
          <div><span>Total Ordered</span><strong>${group.orderedQty}</strong></div>
          <div><span>Total Received</span><strong>${group.receivedQty}</strong></div>
          <div><span>Remaining</span><strong>${group.remainingQty}</strong></div>
          <div><span>Pending Receive</span><strong>${group.pendingReceiveQty}</strong></div>
        </div>

        <div>
          <strong>Department Breakdown</strong>
          <div class="department-breakdown" style="margin-top: 8px;">
            ${departmentRows}
          </div>
        </div>

        <details>
          <summary style="cursor:pointer;font-weight:900;color:var(--primary);">
            ดูรายการคำขอทั้งหมดในกลุ่มนี้ (${group.requests.length} รายการ)
          </summary>

          <div class="consolidated-request-table" style="margin-top: 10px;">
            <table>
              <thead>
                <tr>
                  <th>PR No.</th>
                  <th>Date</th>
                  <th>Department</th>
                  <th>Requester</th>
                  <th>Machine</th>
                  <th>Request Qty</th>
                  <th>Ordered Qty</th>
                  <th>Received Qty</th>
                  <th>Remaining</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                ${requestRows}
              </tbody>
            </table>
          </div>
        </details>

        <div class="consolidation-actions">
          <button class="small-btn" onclick="copyConsolidatedText('${escapeJS(group.groupKey)}')">
            Copy Summary
          </button>
        </div>
      </div>
    `;
  }).join("");
}

function getStatusBadgeClass(status) {
  if (status === "Sent to Purchasing") return "Sent";
  if (status === "PO Issued") return "PO";
  if (status === "Qty Adjusted") return "Adjusted";
  if (status === "Partially Ordered") return "Partially";
  if (status === "Partially Received") return "Waiting";
  return status;
}

function copyConsolidatedText(groupKey) {
  const rows = getConsolidationFilteredRows();
  const groups = getConsolidatedGroups(rows);
  const group = groups.find((g) => g.groupKey === groupKey);

  if (!group) {
    toast("ไม่พบกลุ่มรายการนี้");
    return;
  }

  const departmentText = Object.entries(group.departments)
    .map(([department, data]) => `- ${department}: ${data.requestQty} pcs`)
    .join("\n");

  const text = [
    `Purchase Consolidation`,
    `Part Name: ${group.partName}`,
    `Brand: ${group.brand}`,
    `Model: ${group.model}`,
    `Total Request Qty: ${group.requestQty} pcs`,
    ``,
    `Department Breakdown:`,
    departmentText,
  ].join("\n");

  navigator.clipboard.writeText(text)
    .then(() => toast("คัดลอก Summary แล้ว"))
    .catch(() => toast("คัดลอกไม่สำเร็จ"));
}

/* =========================
   Update / Delete / Detail
========================= */

async function updateStatus(id, status) {
  const item = state.requests.find((r) => r.id === id);
  if (!item) return;

  item.status = status;
  item.updatedAt = new Date().toISOString();

  if (USE_SUPABASE) {
    const { error } = await supabaseClient
      .from("purchase_requests")
      .update({
        status,
        updated_at: new Date().toISOString(),
      })
      .eq("id", id);

    if (error) {
      console.error("updateStatus error:", error);
      toast(`อัปเดตสถานะไม่สำเร็จ: ${error.message}`);
      return;
    }
  } else {
    saveData();
  }

  refreshAll();
  toast("อัปเดตสถานะแล้ว");
}

async function updatePurchasingField(id, field, value) {
  const item = state.requests.find((r) => r.id === id);
  if (!item) return;

  const numberFields = [
    "approvedQty",
    "orderedQty",
    "receivedQty",
    "unitPrice",
  ];

  if (numberFields.includes(field)) {
    item[field] = Number(value || 0);
  } else {
    item[field] = value;
  }

  item.totalAmount = Number(item.orderedQty || 0) * Number(item.unitPrice || 0);
  item.updatedAt = new Date().toISOString();

  const dbFieldMap = {
    approvedQty: "approved_qty",
    orderedQty: "ordered_qty",
    receivedQty: "received_qty",
    qtyAdjustReason: "qty_adjust_reason",
    unitPrice: "unit_price",
    totalAmount: "total_amount",
    currency: "currency",

    poNo: "po_no",
    supplier: "supplier",
    expectedDeliveryDate: "expected_delivery_date",
    receivedDate: "received_date",
    purchasingNote: "purchasing_note",
  };

  const dbField = dbFieldMap[field];

  if (USE_SUPABASE && dbField) {
    const updatePayload = {
      [dbField]: numberFields.includes(field) ? Number(value || 0) : value || null,
      total_amount: item.totalAmount,
      updated_at: new Date().toISOString(),
    };

    const { error } = await supabaseClient
      .from("purchase_requests")
      .update(updatePayload)
      .eq("id", id);

    if (error) {
      console.error("updatePurchasingField error:", error);
      toast(`อัปเดตข้อมูลจัดซื้อไม่สำเร็จ: ${error.message}`);
      return;
    }
  } else {
    saveData();
  }

  toast("บันทึกข้อมูลจัดซื้อแล้ว");
  refreshAll();
}

async function autoUpdatePurchaseStatus(id) {
  const item = state.requests.find((r) => r.id === id);
  if (!item) return;

  const requestQty = Number(item.qty || 0);
  const approvedQty = Number(item.approvedQty || 0);
  const orderedQty = Number(item.orderedQty || 0);
  const receivedQty = Number(item.receivedQty || 0);

  let nextStatus = item.status;

  if (approvedQty > 0 && approvedQty < requestQty) {
    nextStatus = "Qty Adjusted";
  }

  if (orderedQty > 0 && orderedQty < requestQty) {
    nextStatus = "Partially Ordered";
  }

  if (orderedQty >= requestQty && orderedQty > 0) {
    nextStatus = "Ordered";
  }

  if (item.poNo && orderedQty > 0 && nextStatus !== "Received") {
    nextStatus = orderedQty < requestQty ? "Partially Ordered" : "PO Issued";
  }

  if (receivedQty > 0 && receivedQty < orderedQty) {
    nextStatus = "Partially Received";
  }

  if (receivedQty >= orderedQty && orderedQty > 0) {
    nextStatus = "Received";
  }

  await updateStatus(id, nextStatus);
}

async function deleteRequest(id) {
  const ok = confirm("Delete this request?");
  if (!ok) return;

  const item = state.requests.find((r) => r.id === id);

  if (USE_SUPABASE) {
    const { error } = await supabaseClient
      .from("purchase_requests")
      .delete()
      .eq("id", id);

    if (error) {
      console.error("deleteRequest error:", error);
      toast(`ลบข้อมูลไม่สำเร็จ: ${error.message}`);
      return;
    }

    if (item && item.imagePath) {
      await supabaseClient.storage
        .from(STORAGE_BUCKET)
        .remove([item.imagePath]);
    }
  }

  state.requests = state.requests.filter((r) => r.id !== id);

  if (!USE_SUPABASE) {
    saveData();
  }

  refreshAll();
  toast("ลบรายการแล้ว");
}

function openDetail(id) {
  const r = state.requests.find((item) => item.id === id);
  if (!r) return;

  const requestQty = Number(r.qty || 0);
  const approvedQty = Number(r.approvedQty || 0);
  const orderedQty = Number(r.orderedQty || 0);
  const receivedQty = Number(r.receivedQty || 0);
  const remainingQty = Math.max(requestQty - orderedQty, 0);
  const pendingReceiveQty = Math.max(orderedQty - receivedQty, 0);

  document.getElementById("modalTitle").textContent = r.prNo;

  document.getElementById("modalBody").innerHTML = `
    <div class="detail-grid">
      ${detailItem(t("requestDate"), formatDate(r.requestDate))}
      ${detailItem(t("requester"), r.requester)}
      ${detailItem(t("department"), r.department)}
      ${detailItem(t("machine"), r.machine)}
      ${detailItem(t("partName"), r.partName)}
      ${detailItem(t("model"), r.model || "-")}
      ${detailItem(t("brand"), r.brand || "-")}

      ${detailItem("Request Qty", `${requestQty} pcs`)}
      ${detailItem("Approved Qty", `${approvedQty} pcs`)}
      ${detailItem("Ordered Qty", `${orderedQty} pcs`)}
      ${detailItem("Received Qty", `${receivedQty} pcs`)}
      ${detailItem("Remaining Qty", `${remainingQty} pcs`)}
      ${detailItem("Pending Receive", `${pendingReceiveQty} pcs`)}

      ${detailItem(t("urgency"), r.urgency)}
      ${detailItem(t("status"), r.status)}
      ${detailItem("PO No.", r.poNo || "-")}
      ${detailItem("Supplier", r.supplier || "-")}
      ${detailItem("Expected Delivery", r.expectedDeliveryDate || "-")}
      ${detailItem("Received Date", r.receivedDate || "-")}
      ${detailItem("Qty Adjust Reason", r.qtyAdjustReason || "-", true)}
      ${detailItem(t("reason"), r.reason || "-", true)}
      ${detailItem("Purchasing Note", r.purchasingNote || "-", true)}
      ${
        r.imageUrl
          ? `<div class="detail-item full">
              <span>Image</span>
              <img class="detail-image" src="${escapeHTML(r.imageUrl)}" alt="Part Image">
            </div>`
          : ""
      }
    </div>
  `;

  document.getElementById("detailModal").classList.add("show");
}

function detailItem(label, value, full = false) {
  return `
    <div class="detail-item ${full ? "full" : ""}">
      <span>${escapeHTML(label)}</span>
      <strong>${value}</strong>
    </div>
  `;
}

function closeModal() {
  document.getElementById("detailModal").classList.remove("show");
}

/* =========================
   Export Detail Excel
========================= */

function renderExportFilters() {
  const departmentSelect = document.getElementById("exportDepartmentFilter");
  if (!departmentSelect) return;

  const currentValue = departmentSelect.value || "All";

  const fromMaster = getMasterList("department");
  const fromRequests = state.requests
    .map((r) => r.department)
    .filter(Boolean);

  const departments = [...new Set([...fromMaster, ...fromRequests])].sort();

  departmentSelect.innerHTML = `
    <option value="All">${t("allItems")}</option>
    ${departments.map((dep) => `
      <option value="${escapeHTML(dep)}">${escapeHTML(dep)}</option>
    `).join("")}
  `;

  if (departments.includes(currentValue)) {
    departmentSelect.value = currentValue;
  } else {
    departmentSelect.value = "All";
  }
}

function getExportFilteredRows() {
  const month = document.getElementById("exportMonthFilter")?.value || "";
  const department = document.getElementById("exportDepartmentFilter")?.value || "All";
  const status = document.getElementById("exportStatusFilter")?.value || "All";

  let rows = [...state.requests];

  if (month) {
    rows = rows.filter((r) => r.requestDate && r.requestDate.startsWith(month));
  }

  if (department !== "All") {
    rows = rows.filter((r) => r.department === department);
  }

  if (status !== "All") {
    rows = rows.filter((r) => r.status === status);
  }

  return rows;
}

function updateExportPreview() {
  const preview = document.getElementById("exportCountPreview");
  if (!preview) return;

  const rows = getExportFilteredRows();
  preview.textContent = `${rows.length} รายการ`;
}

async function exportExcelAdvanced() {
  const rows = getExportFilteredRows();
  await exportExcelFromRows(rows, "Filtered");
}

function quickExportAll() {
  document.getElementById("exportMonthFilter").value = "";
  document.getElementById("exportDepartmentFilter").value = "All";
  document.getElementById("exportStatusFilter").value = "All";
  updateExportPreview();
  exportExcelAdvanced();
}

function quickExportThisMonth() {
  document.getElementById("exportMonthFilter").value = getTodayISO().slice(0, 7);
  document.getElementById("exportDepartmentFilter").value = "All";
  document.getElementById("exportStatusFilter").value = "All";
  updateExportPreview();
  exportExcelAdvanced();
}

function quickExportUrgent() {
  const rows = state.requests.filter(
    (r) => r.urgency === "Urgent" || r.urgency === "Critical"
  );

  exportExcelFromRows(rows, "Urgent-Critical");
}

function quickExportPurchasing() {
  const rows = state.requests.filter((r) =>
    ["Sent to Purchasing", "Qty Adjusted", "PO Issued", "Partially Ordered", "Ordered", "Partially Received", "Received"].includes(r.status)
  );

  exportExcelFromRows(rows, "Purchasing-Followup");
}

async function exportExcelFromRows(rows, fileTag = "Export") {
  if (!rows.length) {
    toast("ไม่มีข้อมูลสำหรับ Export");
    return;
  }

  const includeImage = document.getElementById("exportIncludeImage")?.checked;

  try {
    toast("กำลังสร้างไฟล์ Excel...");

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "MPR Purchase Request System";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Purchase Requests", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    worksheet.columns = [
      { header: "No.", key: "no", width: 6 },
      { header: "PR No.", key: "prNo", width: 22 },
      { header: "Request Date", key: "requestDate", width: 16 },
      { header: "Requester", key: "requester", width: 22 },
      { header: "Department", key: "department", width: 16 },
      { header: "Machine", key: "machine", width: 26 },
      { header: "Part Name", key: "partName", width: 34 },
      { header: "Model", key: "model", width: 24 },
      { header: "Brand", key: "brand", width: 20 },

      { header: "Request Qty", key: "requestQty", width: 14 },
      { header: "Approved Qty", key: "approvedQty", width: 14 },
      { header: "Ordered Qty", key: "orderedQty", width: 14 },
      { header: "Received Qty", key: "receivedQty", width: 14 },
      { header: "Remaining Qty", key: "remainingQty", width: 16 },
      { header: "Pending Receive Qty", key: "pendingReceiveQty", width: 20 },
      { header: "Qty Adjust Reason", key: "qtyAdjustReason", width: 38 },

      { header: "Unit Price", key: "unitPrice", width: 14 },
      { header: "Total Amount", key: "totalAmount", width: 16 },
      { header: "Currency", key: "currency", width: 12 },

      { header: "Urgency", key: "urgency", width: 14 },
      { header: "Reason", key: "reason", width: 50 },
      { header: "Status", key: "status", width: 22 },
      { header: "PO No.", key: "poNo", width: 20 },
      { header: "Supplier", key: "supplier", width: 24 },
      { header: "Expected Delivery", key: "expectedDeliveryDate", width: 18 },
      { header: "Received Date", key: "receivedDate", width: 18 },
      { header: "Image", key: "image", width: includeImage ? 20 : 12 },
      { header: "Image URL", key: "imageUrl", width: 50 },
      { header: "Purchasing Note", key: "purchasingNote", width: 40 },
    ];

    styleExcelHeader(worksheet);

    rows.forEach((r, index) => {
      const requestQty = Number(r.qty || 0);
      const approvedQty = Number(r.approvedQty || 0);
      const orderedQty = Number(r.orderedQty || 0);
      const receivedQty = Number(r.receivedQty || 0);
      const remainingQty = Math.max(requestQty - orderedQty, 0);
      const pendingReceiveQty = Math.max(orderedQty - receivedQty, 0);
      const unitPrice = Number(r.unitPrice || 0);
      const totalAmount = Number(r.totalAmount || orderedQty * unitPrice || 0);

      const row = worksheet.addRow({
        no: index + 1,
        prNo: r.prNo,
        requestDate: r.requestDate,
        requester: r.requester,
        department: r.department,
        machine: r.machine,
        partName: r.partName,
        model: r.model,
        brand: r.brand,

        requestQty,
        approvedQty,
        orderedQty,
        receivedQty,
        remainingQty,
        pendingReceiveQty,
        qtyAdjustReason: r.qtyAdjustReason || "",

        unitPrice,
        totalAmount,
        currency: r.currency || "THB",

        urgency: r.urgency,
        reason: r.reason,
        status: r.status,
        poNo: r.poNo || "",
        supplier: r.supplier || "",
        expectedDeliveryDate: r.expectedDeliveryDate || "",
        receivedDate: r.receivedDate || "",
        image: r.imageUrl ? "Image" : "-",
        imageUrl: r.imageUrl || "",
        purchasingNote: r.purchasingNote || "",
      });

      row.height = includeImage && r.imageUrl ? 82 : 26;

      row.eachCell((cell) => {
        cell.font = {
          name: "Arial Unicode MS",
          size: 10,
          color: { argb: "FF0F172A" },
        };
        cell.alignment = {
          vertical: "middle",
          wrapText: true,
        };
        cell.border = excelBorder();
      });

      applyStatusExcelStyle(row.getCell("status"), r.status);
      applyUrgencyExcelStyle(row.getCell("urgency"), r.urgency);

      if (r.imageUrl) {
        const imageUrlCell = row.getCell("imageUrl");
        imageUrlCell.value = {
          text: "Open Image",
          hyperlink: r.imageUrl,
        };
        imageUrlCell.font = {
          name: "Arial Unicode MS",
          color: { argb: "FF155EEF" },
          underline: true,
          size: 10,
        };
      }
    });

    if (includeImage) {
      await addImagesToWorksheet(workbook, worksheet, rows, 26.12);
    }

    worksheet.autoFilter = {
      from: "A1",
      to: "AC1",
    };

    const buffer = await workbook.xlsx.writeBuffer();

    downloadBlob(
      buffer,
      buildExportFileName(fileTag),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    toast("Export Excel สำเร็จ");
  } catch (error) {
    console.error(error);
    toast("Export Excel ไม่สำเร็จ");
  }
}

/* =========================
   Export Consolidated Excel
========================= */

async function quickExportConsolidated() {
  const rows = getConsolidationFilteredRows();
  const groups = getConsolidatedGroups(rows);

  if (!groups.length) {
    toast("ไม่มีข้อมูลสำหรับ Export รวมตาม Model");
    return;
  }

  try {
    toast("กำลังสร้างไฟล์ Excel แบบรวม Model...");

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "MPR Purchase Request System";
    workbook.created = new Date();

    const summarySheet = workbook.addWorksheet("Consolidated Summary", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    summarySheet.columns = [
      { header: "No.", key: "no", width: 6 },
      { header: "Part Name", key: "partName", width: 34 },
      { header: "Brand", key: "brand", width: 20 },
      { header: "Model", key: "model", width: 24 },
      { header: "Total Request Qty", key: "requestQty", width: 18 },
      { header: "Total Approved Qty", key: "approvedQty", width: 18 },
      { header: "Total Ordered Qty", key: "orderedQty", width: 18 },
      { header: "Total Received Qty", key: "receivedQty", width: 18 },
      { header: "Remaining Qty", key: "remainingQty", width: 16 },
      { header: "Pending Receive Qty", key: "pendingReceiveQty", width: 22 },
      { header: "Departments", key: "departments", width: 38 },
      { header: "Request Count", key: "requestCount", width: 14 },
      { header: "Image URL", key: "imageUrl", width: 50 },
    ];

    styleExcelHeader(summarySheet);

    groups.forEach((group, index) => {
      const departments = Object.entries(group.departments)
        .map(([department, data]) => `${department}: ${data.requestQty} pcs`)
        .join("\n");

      const row = summarySheet.addRow({
        no: index + 1,
        partName: group.partName,
        brand: group.brand,
        model: group.model,
        requestQty: group.requestQty,
        approvedQty: group.approvedQty,
        orderedQty: group.orderedQty,
        receivedQty: group.receivedQty,
        remainingQty: group.remainingQty,
        pendingReceiveQty: group.pendingReceiveQty,
        departments,
        requestCount: group.requests.length,
        imageUrl: group.imageUrl || "",
      });

      row.eachCell((cell) => {
        cell.font = {
          name: "Arial Unicode MS",
          size: 10,
          color: { argb: "FF0F172A" },
        };
        cell.alignment = {
          vertical: "middle",
          wrapText: true,
        };
        cell.border = excelBorder();
      });

      if (group.imageUrl) {
        const imageUrlCell = row.getCell("imageUrl");
        imageUrlCell.value = {
          text: "Open Image",
          hyperlink: group.imageUrl,
        };
        imageUrlCell.font = {
          name: "Arial Unicode MS",
          color: { argb: "FF155EEF" },
          underline: true,
          size: 10,
        };
      }
    });

    summarySheet.autoFilter = {
      from: "A1",
      to: "M1",
    };

    const detailSheet = workbook.addWorksheet("Request Details", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    detailSheet.columns = [
      { header: "Group Key", key: "groupKey", width: 55 },
      { header: "PR No.", key: "prNo", width: 22 },
      { header: "Request Date", key: "requestDate", width: 16 },
      { header: "Department", key: "department", width: 16 },
      { header: "Requester", key: "requester", width: 22 },
      { header: "Machine", key: "machine", width: 26 },
      { header: "Part Name", key: "partName", width: 34 },
      { header: "Brand", key: "brand", width: 20 },
      { header: "Model", key: "model", width: 24 },
      { header: "Request Qty", key: "requestQty", width: 14 },
      { header: "Approved Qty", key: "approvedQty", width: 14 },
      { header: "Ordered Qty", key: "orderedQty", width: 14 },
      { header: "Received Qty", key: "receivedQty", width: 14 },
      { header: "Remaining Qty", key: "remainingQty", width: 16 },
      { header: "Pending Receive Qty", key: "pendingReceiveQty", width: 20 },
      { header: "Status", key: "status", width: 22 },
      { header: "PO No.", key: "poNo", width: 20 },
      { header: "Supplier", key: "supplier", width: 24 },
      { header: "Qty Adjust Reason", key: "qtyAdjustReason", width: 38 },
      { header: "Reason", key: "reason", width: 50 },
    ];

    styleExcelHeader(detailSheet);

    groups.forEach((group) => {
      group.requests.forEach((r) => {
        const requestQty = Number(r.qty || 0);
        const approvedQty = Number(r.approvedQty || 0);
        const orderedQty = Number(r.orderedQty || 0);
        const receivedQty = Number(r.receivedQty || 0);
        const remainingQty = Math.max(requestQty - orderedQty, 0);
        const pendingReceiveQty = Math.max(orderedQty - receivedQty, 0);

        const row = detailSheet.addRow({
          groupKey: group.groupKey,
          prNo: r.prNo,
          requestDate: r.requestDate,
          department: r.department,
          requester: r.requester,
          machine: r.machine,
          partName: r.partName,
          brand: r.brand,
          model: r.model,
          requestQty,
          approvedQty,
          orderedQty,
          receivedQty,
          remainingQty,
          pendingReceiveQty,
          status: r.status,
          poNo: r.poNo || "",
          supplier: r.supplier || "",
          qtyAdjustReason: r.qtyAdjustReason || "",
          reason: r.reason || "",
        });

        row.eachCell((cell) => {
          cell.font = {
            name: "Arial Unicode MS",
            size: 10,
            color: { argb: "FF0F172A" },
          };
          cell.alignment = {
            vertical: "middle",
            wrapText: true,
          };
          cell.border = excelBorder();
        });

        applyStatusExcelStyle(row.getCell("status"), r.status);
      });
    });

    detailSheet.autoFilter = {
      from: "A1",
      to: "T1",
    };

    const buffer = await workbook.xlsx.writeBuffer();

    downloadBlob(
      buffer,
      `MPR_Consolidated_Model_${getTodayISO()}.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    toast("Export รวมตาม Model สำเร็จ");
  } catch (error) {
    console.error(error);
    toast("Export รวมตาม Model ไม่สำเร็จ");
  }
}

/* =========================
   Excel Helpers
========================= */

function styleExcelHeader(worksheet) {
  const headerRow = worksheet.getRow(1);
  headerRow.height = 26;

  headerRow.eachCell((cell) => {
    cell.font = {
      name: "Arial Unicode MS",
      bold: true,
      color: { argb: "FFFFFFFF" },
      size: 11,
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF0F2B55" },
    };
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    cell.border = excelBorder();
  });
}

async function addImagesToWorksheet(workbook, worksheet, rows, imageColumnIndex = 26.12) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r.imageUrl) continue;

    try {
      const imageBase64 = await imageUrlToBase64(r.imageUrl);
      if (!imageBase64) continue;

      const extension = getImageExtensionFromBase64(imageBase64);

      const imageId = workbook.addImage({
        base64: imageBase64,
        extension,
      });

      const excelRowNumber = i + 2;

      worksheet.addImage(imageId, {
        tl: { col: imageColumnIndex, row: excelRowNumber - 0.82 },
        ext: { width: 110, height: 75 },
      });

      worksheet.getRow(excelRowNumber).height = 82;
    } catch (error) {
      console.warn("Cannot add image:", r.imageUrl, error);
    }
  }
}

async function imageUrlToBase64(url) {
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error("Image fetch failed");
  }

  const blob = await response.blob();

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;

    reader.readAsDataURL(blob);
  });
}

function getImageExtensionFromBase64(base64) {
  if (base64.startsWith("data:image/png")) return "png";
  if (base64.startsWith("data:image/webp")) return "png";
  if (base64.startsWith("data:image/jpeg")) return "jpeg";
  if (base64.startsWith("data:image/jpg")) return "jpeg";
  return "png";
}

function excelBorder() {
  return {
    top: { style: "thin", color: { argb: "FFE2E8F0" } },
    left: { style: "thin", color: { argb: "FFE2E8F0" } },
    bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
    right: { style: "thin", color: { argb: "FFE2E8F0" } },
  };
}

function applyStatusExcelStyle(cell, status) {
  const map = {
    Pending: "FFF59E0B",
    Approved: "FF155EEF",
    "Sent to Purchasing": "FF7C3AED",
    "Qty Adjusted": "FF7C3AED",
    "PO Issued": "FF0369A1",
    "Partially Ordered": "FFB45309",
    Ordered: "FFCA8A04",
    "Partially Received": "FF0369A1",
    Received: "FF16A34A",
    Cancelled: "FFDC2626",
  };

  const color = map[status] || "FF64748B";

  cell.font = {
    name: "Arial Unicode MS",
    bold: true,
    color: { argb: "FFFFFFFF" },
    size: 10,
  };

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };

  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
  };
}

function applyUrgencyExcelStyle(cell, urgency) {
  const map = {
    Normal: "FF64748B",
    Urgent: "FFF97316",
    Critical: "FFDC2626",
  };

  const color = map[urgency] || "FF64748B";

  cell.font = {
    name: "Arial Unicode MS",
    bold: true,
    color: { argb: "FFFFFFFF" },
    size: 10,
  };

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };

  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
  };
}

function buildExportFileName(fileTag = "Export") {
  const month = document.getElementById("exportMonthFilter")?.value || "All-Month";
  const department = document.getElementById("exportDepartmentFilter")?.value || "All-Department";
  const status = document.getElementById("exportStatusFilter")?.value || "All-Status";

  const safeDepartment = department.replaceAll(" ", "-").replaceAll("/", "-");
  const safeStatus = status.replaceAll(" ", "-").replaceAll("/", "-");

  return `MPR_Purchase_${fileTag}_${month}_${safeDepartment}_${safeStatus}_${getTodayISO()}.xlsx`;
}

function downloadBlob(buffer, fileName, mimeType) {
  const blob = new Blob([buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;

  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  URL.revokeObjectURL(url);
}

/* =========================
   Utility
========================= */

function resetFilters() {
  document.getElementById("searchInput").value = "";
  document.getElementById("statusFilter").value = "All";
  document.getElementById("monthFilter").value = "";
  renderRequestList();
}

function goPage(pageId) {
  state.currentPage = pageId;

  document.querySelectorAll(".page").forEach((p) => p.classList.remove("active"));

  const page = document.getElementById(pageId);
  if (page) page.classList.add("active");

  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.page === pageId);
  });

  const titleMap = {
    dashboard: t("dashboard"),
    newRequest: t("newRequest"),
    requestList: t("requestList"),
    purchasingFollowup: t("purchasingFollowup"),
    consolidationPage: t("purchaseConsolidation"),
    monthlySummary: t("monthlySummary"),
    exportPage: t("export"),
    settingsPage: t("settings"),
  };

  const descMap = {
    dashboard: t("dashboardDesc"),
    newRequest: t("newRequestDesc"),
    requestList: t("requestListDesc"),
    purchasingFollowup: t("purchasingFollowupDesc"),
    consolidationPage: t("purchaseConsolidationDesc"),
    monthlySummary: t("monthlySummaryDesc"),
    exportPage: t("exportDesc"),
    settingsPage: t("settingsDesc"),
  };

  document.getElementById("pageTitle").textContent = titleMap[pageId] || "";
  document.getElementById("pageDesc").textContent = descMap[pageId] || "";
}

function updateLanguage() {
  document.querySelectorAll(".lang-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.lang === state.lang);
  });

  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.dataset.i18n;
    if (i18n[state.lang][key]) {
      el.textContent = i18n[state.lang][key];
    }
  });

  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.dataset.i18nPlaceholder;
    if (i18n[state.lang][key]) {
      el.placeholder = i18n[state.lang][key];
    }
  });

  const optionMap = {
    normal: t("normal"),
    urgent: t("urgent"),
    critical: t("critical"),
  };

  document.querySelectorAll("[data-i18n-option]").forEach((el) => {
    const key = el.dataset.i18nOption;
    if (optionMap[key]) el.textContent = optionMap[key];
  });

  renderMasterDropdowns();
  syncDynamicLanguageText();
  goPage(state.currentPage);
}

function syncDynamicLanguageText() {
  const searchInput = document.getElementById("searchInput");
  if (searchInput) searchInput.placeholder = t("requestSearchPlaceholder");

  const followupSearch = document.getElementById("followupSearch");
  if (followupSearch) followupSearch.placeholder = t("followupSearchPlaceholder");

  const consolidationSearch = document.getElementById("consolidationSearch");
  if (consolidationSearch) consolidationSearch.placeholder = t("consolidationSearchPlaceholder");

  const consolidationStatusFilter = document.getElementById("consolidationStatusFilter");
  if (consolidationStatusFilter) {
    const activeOption = consolidationStatusFilter.querySelector('option[value="Active"]');
    const allOption = consolidationStatusFilter.querySelector('option[value="All"]');

    if (activeOption) activeOption.textContent = t("activeItemsOnly");
    if (allOption) allOption.textContent = t("allItems");
  }

  const followupStatusFilter = document.getElementById("followupStatusFilter");
  if (followupStatusFilter) {
    const allOption = followupStatusFilter.querySelector('option[value="All"]');
    if (allOption) allOption.textContent = t("allStatus");
  }

  const statusFilter = document.getElementById("statusFilter");
  if (statusFilter) {
    const allOption = statusFilter.querySelector('option[value="All"]');
    if (allOption) allOption.textContent = t("allStatus");
  }

  const exportStatusFilter = document.getElementById("exportStatusFilter");
  if (exportStatusFilter) {
    const allOption = exportStatusFilter.querySelector('option[value="All"]');
    if (allOption) allOption.textContent = t("allStatus");
  }
}

function mapFromSupabase(row) {
  return {
    id: row.id,
    prNo: row.pr_no,
    requestDate: row.request_date,
    requester: row.requester || "",
    department: row.department || "",
    machine: row.machine || "",
    partName: row.part_name || "",
    model: row.model || "",
    brand: row.brand || "",

    qty: Number(row.qty || 1),
    approvedQty: Number(row.approved_qty || 0),
    orderedQty: Number(row.ordered_qty || 0),
    receivedQty: Number(row.received_qty || 0),
    qtyAdjustReason: row.qty_adjust_reason || "",
    unitPrice: Number(row.unit_price || 0),
    totalAmount: Number(row.total_amount || 0),
    currency: row.currency || "THB",

    urgency: row.urgency || "Normal",
    reason: row.reason || "",
    imageUrl: row.image_url || "",
    imagePath: row.image_path || "",
    status: row.status || "Pending",
    poNo: row.po_no || "",
    supplier: row.supplier || "",
    expectedDeliveryDate: row.expected_delivery_date || "",
    receivedDate: row.received_date || "",
    purchasingNote: row.purchasing_note || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

function mapToSupabase(request) {
  return {
    pr_no: request.prNo,
    request_date: request.requestDate,
    requester: request.requester,
    department: request.department,
    machine: request.machine,
    part_name: request.partName,
    model: request.model,
    brand: request.brand,

    qty: request.qty,
    approved_qty: request.approvedQty || 0,
    ordered_qty: request.orderedQty || 0,
    received_qty: request.receivedQty || 0,
    qty_adjust_reason: request.qtyAdjustReason || "",
    unit_price: request.unitPrice || 0,
    total_amount: request.totalAmount || 0,
    currency: request.currency || "THB",

    urgency: request.urgency,
    reason: request.reason,
    image_url: request.imageUrl || "",
    image_path: request.imagePath || "",
    status: request.status || "Pending",
    po_no: request.poNo || "",
    supplier: request.supplier || "",
    expected_delivery_date: request.expectedDeliveryDate || null,
    received_date: request.receivedDate || null,
    purchasing_note: request.purchasingNote || "",
    updated_at: new Date().toISOString(),
  };
}

function statusBadge(status) {
  let cls = getStatusBadgeClass(status);
  return `<span class="badge ${cls}">${escapeHTML(status)}</span>`;
}

function urgencyBadge(urgency) {
  return `<span class="urgency ${urgency}">${escapeHTML(urgency)}</span>`;
}

function emptyRow(colspan) {
  return `<tr><td colspan="${colspan}" class="empty">${t("noData")}</td></tr>`;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function t(key) {
  return i18n[state.lang][key] || key;
}

function formatDate(dateStr) {
  if (!dateStr) return "-";
  const [y, m, d] = dateStr.split("-");
  return `${d}/${m}/${y}`;
}

function escapeHTML(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeJS(value) {
  return String(value ?? "")
    .replaceAll("\\", "\\\\")
    .replaceAll("'", "\\'")
    .replaceAll('"', '\\"');
}

function toast(message) {
  const el = document.getElementById("toast");
  el.textContent = message;
  el.classList.add("show");

  clearTimeout(window.__toastTimer);

  window.__toastTimer = setTimeout(() => {
    el.classList.remove("show");
  }, 3000);
}

function openMobileMenu() {
  document.getElementById("sidebar").classList.add("open");
  document.getElementById("sidebarOverlay").classList.add("show");
}

function closeMobileMenu() {
  document.getElementById("sidebar").classList.remove("open");
  document.getElementById("sidebarOverlay").classList.remove("show");
}