/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// --- Type definitions for CDN libraries to inform TypeScript ---
declare const firebase: any;
declare const XLSX: any;
declare const jspdf: any;

// --- FIREBASE INITIALIZATION (from index.tsx 1 pattern) ---
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// --- DOM Elements ---
const fileUpload = document.getElementById("file-upload") as HTMLInputElement;
const searchInput = document.getElementById("search-input") as HTMLInputElement;
const lastUpdate = document.getElementById("last-update") as HTMLParagraphElement;
const placeholder = document.getElementById("placeholder") as HTMLDivElement;
const summaryStats = document.getElementById("summary-stats") as HTMLDivElement;
const deliveryDashboard = document.getElementById("delivery-dashboard") as HTMLDivElement;
const deliveryTabs = document.getElementById("delivery-tabs") as HTMLDivElement;
const deliveryContent = document.getElementById("delivery-content") as HTMLDivElement;
const exportExcelBtn = document.getElementById("export-excel-btn") as HTMLButtonElement;
const exportPdfBtn = document.getElementById("export-pdf-btn") as HTMLButtonElement;
const themeToggleBtn = document.getElementById("theme-toggle") as HTMLButtonElement;
const htmlEl = document.documentElement;

// --- Logo Elements ---
const logoUpload = document.getElementById("logo-upload") as HTMLInputElement;
const logoContainer = document.getElementById("logo-container") as HTMLDivElement;
const companyLogo = document.getElementById("company-logo") as HTMLImageElement;

// --- Modal Elements ---
const modalContainer = document.getElementById("confirmation-modal-container") as HTMLDivElement;
const modalEl = document.getElementById("confirmation-modal") as HTMLDivElement;
const modalTitle = document.getElementById("modal-title") as HTMLHeadingElement;
const modalMessage = document.getElementById("modal-message") as HTMLParagraphElement;
const modalConfirmBtn = document.getElementById("modal-confirm-btn") as HTMLButtonElement;
const modalCancelBtn = document.getElementById("modal-cancel-btn") as HTMLButtonElement;

// --- Language Elements ---
const languageSwitcher = document.getElementById("language-switcher") as HTMLDivElement;

// --- i18n ---
const translations = {
  "pt-BR": {
    pageTitle: "Painel de Entregas de Contêineres",
    headerTitle: "Painel de Entregas",
    uploadPrompt: "Carregue sua planilha de agendamento para começar",
    searchInputPlaceholder: "Pesquisar container, BL, navio...",
    uploadLogoTooltip: "Carregar logo da empresa",
    toggleThemeTooltip: "Alternar tema",
    uploadSheetButton: "Carregar",
    uploadSheetTooltip: "Carregar Planilha",
    exportExcelButton: "Exportar Excel",
    exportPdfButton: "Exportar PDF",
    processing: "Processando...",
    placeholderTitle: "Aguardando planilha...",
    placeholderMessage: "Selecione um arquivo .xlsx para visualizar a programação de entregas.",
    imageTooLarge: "O arquivo de imagem é muito grande (máx 2MB).",
    imageReadError: "Não foi possível ler o arquivo de imagem.",
    logoUpdated: "Logo da empresa atualizado!",
    logoUploadError: "Erro ao carregar o logo.",
    sheetLoaded: "Planilha de entregas carregada!",
    fileReadError: "Falha ao ler o arquivo.",
    emptySheetError: "A planilha de agendamento está vazia.",
    fileProcessError: "Erro ao processar arquivo.",
    noDeliverySheet: "Nenhuma planilha de agendamento de entregas foi encontrada.",
    noDataToExport: "Não há dados para exportar.",
    excelGenerated: "Arquivo Excel gerado!",
    pdfGenerated: "Arquivo PDF gerado!",
    statusUpdated: (containerId: string, status: string) => `Container ${containerId} atualizado para ${status}!`,
    fieldUpdated: (field: string) => `Campo "${field.replace(/_/g, " ")}" atualizado.`,
    confirmAction: "Confirmar Ação",
    areYouSure: "Tem certeza que deseja continuar?",
    confirmButton: "Confirmar",
    cancelButton: "Cancelar",
    confirmStatusChangeTitle: "Confirmar Alteração de Status",
    confirmStatusChangeMessage: (containerId: string, status: string) =>
      `Tem certeza que deseja alterar o status de ${containerId} para "${status}"? Esta ação é definitiva.`,
    exportExcelTitle: "Exportar para Excel",
    exportExcelMessage: "Deseja gerar o arquivo .xlsx com os dados atuais?",
    exportPdfTitle: "Exportar para PDF",
    exportPdfMessage: "Deseja gerar o arquivo .pdf com os dados atuais?",
    totalContainers: "Agendados",
    delivered: "Entregues",
    inTransit: "A Caminho",
    postponed: "Adiados",
    pending: "Pendentes",
    canceled: "Cancelados",
    noResultsTitle: "Nenhum resultado encontrado",
    noResultsMessage: "Nenhum resultado encontrado para os filtros aplicados.",
    containersDelivered: (delivered: number, total: number) => `${delivered} de ${total} containers entregues`,
    undefinedDate: "Data não definida",
    dateNotAvailable: "N/D",
    tableHeaderRow: "#",
    tableHeaderContainer: "Container",
    tableHeaderBL: "BL",
    tableHeaderVessel: "Navio",
    tableHeaderCompany: "Transportadora",
    tableHeaderPlate: "Placa",
    tableHeaderWarehouse: "Armazém",
    tableHeaderStatus: "Status",
    tableHeaderLot: "Lote",
    STATUS_PENDENTE: "Pendente",
    STATUS_A_CAMINHO: "A Caminho",
    STATUS_ADIADO: "Adiado",
    STATUS_ENTREGUE: "Entregue",
    STATUS_CANCELADO: "Cancelado",
    detailsVessel: "Navio (Vessel)",
    detailsWarehouse: "Armazém",
    detailsNotes: "Observações",
    detailsMaterial: "Tipo de Material",
    detailsLot: "Lote (LOT)",
    detailsCompany: "Transportadora",
    performanceTitle: "Desempenho por Transportadora",
    badgeBattery: "Bateria",
    pdfTitle: "Programação de Entregas de Contêineres",
    pdfGeneratedOn: (date: string) => `Relatório gerado em: ${date}`,
    pdfPage: (page: number, total: number) => `Página ${page} de ${total}`,
    lastUpdateText: (sheet: string, date: string) => `Dados de "${sheet}" | Carregado em: ${date}`,
    changeStatusFor: (containerId: string) => `Alterar status do container ${containerId || ""}`,
    viewDetailsFor: (containerId: string) => `Ver detalhes do container ${containerId || "sem identificação"}`,
  },
  "en-US": {
    pageTitle: "Container Delivery Dashboard",
    headerTitle: "Delivery Dashboard",
    uploadPrompt: "Upload your schedule spreadsheet to begin",
    searchInputPlaceholder: "Search container, BL, vessel...",
    uploadLogoTooltip: "Upload company logo",
    toggleThemeTooltip: "Toggle theme",
    uploadSheetButton: "Upload",
    uploadSheetTooltip: "Upload Spreadsheet",
    exportExcelButton: "Export Excel",
    exportPdfButton: "Export PDF",
    processing: "Processing...",
    placeholderTitle: "Waiting for spreadsheet...",
    placeholderMessage: "Select an .xlsx file to view the delivery schedule.",
    imageTooLarge: "Image file is too large (max 2MB).",
    imageReadError: "Could not read image file.",
    logoUpdated: "Company logo updated!",
    logoUploadError: "Error uploading logo.",
    sheetLoaded: "Delivery spreadsheet loaded!",
    fileReadError: "Failed to read the file.",
    emptySheetError: "The scheduling spreadsheet is empty.",
    fileProcessError: "Error processing file.",
    noDeliverySheet: "No delivery schedule sheet was found.",
    noDataToExport: "No data to export.",
    excelGenerated: "Excel file generated!",
    pdfGenerated: "PDF file generated!",
    statusUpdated: (containerId: string, status: string) => `Container ${containerId} updated to ${status}!`,
    fieldUpdated: (field: string) => `Field "${field.replace(/_/g, " ")}" updated.`,
    confirmAction: "Confirm Action",
    areYouSure: "Are you sure you want to continue?",
    confirmButton: "Confirm",
    cancelButton: "Cancel",
    confirmStatusChangeTitle: "Confirm Status Change",
    confirmStatusChangeMessage: (containerId: string, status: string) =>
      `Are you sure you want to change the status of ${containerId} to "${status}"? This action is final.`,
    exportExcelTitle: "Export to Excel",
    exportExcelMessage: "Do you want to generate the .xlsx file with the current data?",
    exportPdfTitle: "Export to PDF",
    exportPdfMessage: "Do you want to generate the .pdf file with the current data?",
    totalContainers: "Scheduled",
    delivered: "Delivered",
    inTransit: "In Transit",
    postponed: "Postponed",
    pending: "Pending",
    canceled: "Canceled",
    noResultsTitle: "No results found",
    noResultsMessage: "No results found for the applied filters.",
    containersDelivered: (delivered: number, total: number) => `${delivered} of ${total} containers delivered`,
    undefinedDate: "Date not set",
    dateNotAvailable: "N/A",
    tableHeaderRow: "#",
    tableHeaderContainer: "Container",
    tableHeaderBL: "BL",
    tableHeaderVessel: "Vessel",
    tableHeaderCompany: "Carrier",
    tableHeaderPlate: "Plate",
    tableHeaderWarehouse: "Warehouse",
    tableHeaderStatus: "Status",
    tableHeaderLot: "LOT",
    STATUS_PENDENTE: "Pending",
    STATUS_A_CAMINHO: "In Transit",
    STATUS_ADIADO: "Postponed",
    STATUS_ENTREGUE: "Delivered",
    STATUS_CANCELADO: "Canceled",
    detailsVessel: "Vessel",
    detailsWarehouse: "Warehouse",
    detailsNotes: "Notes",
    detailsMaterial: "Material Type",
    detailsLot: "LOT Number",
    detailsCompany: "Carrier",
    performanceTitle: "Carrier Performance",
    badgeBattery: "Battery",
    pdfTitle: "Container Delivery Schedule",
    pdfGeneratedOn: (date: string) => `Report generated on: ${date}`,
    pdfPage: (page: number, total: number) => `Page ${page} of ${total}`,
    lastUpdateText: (sheet: string, date: string) => `Data from "${sheet}" | Loaded on: ${date}`,
    changeStatusFor: (containerId: string) => `Change status for container ${containerId || ""}`,
    viewDetailsFor: (containerId: string) => `View details for container ${containerId || "unidentified"}`,
  },
  "zh-CN": {
    pageTitle: "集装箱交付仪表板",
    headerTitle: "交付仪表板",
    uploadPrompt: "上传您的排程电子表格以开始",
    searchInputPlaceholder: "搜索集装箱、提单 (BL)、船名...",
    uploadLogoTooltip: "上传公司标志",
    toggleThemeTooltip: "切换主题",
    uploadSheetButton: "上传",
    uploadSheetTooltip: "上传电子表格",
    exportExcelButton: "导出 Excel",
    exportPdfButton: "导出 PDF",
    processing: "处理中...",
    placeholderTitle: "等待电子表格...",
    placeholderMessage: "选择一个 .xlsx 文件以查看交付计划。",
    imageTooLarge: "图片文件太大（最大 2MB）。",
    imageReadError: "无法读取图片文件。",
    logoUpdated: "公司标志已更新！",
    logoUploadError: "上传标志时出错。",
    sheetLoaded: "交付电子表格已加载！",
    fileReadError: "读取文件失败。",
    emptySheetError: "排程电子表格为空。",
    fileProcessError: "处理文件时出错。",
    noDeliverySheet: "未找到交付计划表。",
    noDataToExport: "无数据可导出。",
    excelGenerated: "Excel 文件已生成！",
    pdfGenerated: "PDF 文件已生成！",
    statusUpdated: (containerId: string, status: string) => `集装箱 ${containerId} 已更新为 ${status}！`,
    fieldUpdated: (field: string) => `字段 "${field.replace(/_/g, " ")}" 已更新。`,
    confirmAction: "确认操作",
    areYouSure: "您确定要继续吗？",
    confirmButton: "确认",
    cancelButton: "取消",
    confirmStatusChangeTitle: "确认状态更改",
    confirmStatusChangeMessage: (containerId: string, status: string) =>
      `您确定要将 ${containerId} 的状态更改为 "${status}" 吗？此操作是最终的。`,
    exportExcelTitle: "导出到 Excel",
    exportExcelMessage: "您要使用当前数据生成 .xlsx 文件吗？",
    exportPdfTitle: "导出到 PDF",
    exportPdfMessage: "您要使用当前数据生成 .pdf 文件吗？",
    totalContainers: "总集装箱数",
    delivered: "已交付",
    inTransit: "运输中",
    postponed: "已推迟",
    pending: "待处理",
    canceled: "已取消",
    noResultsTitle: "未找到结果",
    noResultsMessage: "未找到符合所应用筛选条件的结果。",
    containersDelivered: (delivered: number, total: number) => `${total} 个集装箱中已交付 ${delivered} 个`,
    undefinedDate: "未设置日期",
    dateNotAvailable: "不适用",
    tableHeaderRow: "#",
    tableHeaderContainer: "集装箱",
    tableHeaderBL: "提单 (BL)",
    tableHeaderVessel: "船名",
    tableHeaderCompany: "运输公司",
    tableHeaderPlate: "车牌",
    tableHeaderWarehouse: "仓库",
    tableHeaderStatus: "状态",
    tableHeaderLot: "批号 (LOT)",
    STATUS_PENDENTE: "待处理",
    STATUS_A_CAMINHO: "运输中",
    STATUS_ADIADO: "已推迟",
    STATUS_ENTREGUE: "已交付",
    STATUS_CANCELADO: "已取消",
    detailsVessel: "船名",
    detailsWarehouse: "仓库",
    detailsNotes: "备注",
    detailsMaterial: "物料类型",
    detailsLot: "批号",
    detailsCompany: "运输公司",
    performanceTitle: "承运人绩效",
    badgeBattery: "电池",
    pdfTitle: "集装箱交付计划",
    pdfGeneratedOn: (date: string) => `报告生成于：${date}`,
    pdfPage: (page: number, total: number) => `第 ${page} 页，共 ${total} 页`,
    lastUpdateText: (sheet: string, date: string) => `数据来源 "${sheet}" | 加载于：${date}`,
    changeStatusFor: (containerId: string) => `更改集装箱 ${containerId || ""} 的状态`,
    viewDetailsFor: (containerId: string) => `查看集装箱 ${containerId || "未识别"} 的详细信息`,
  },
};

type Language = keyof typeof translations;
let currentLanguage: Language = "pt-BR";
type TranslationKey = keyof typeof translations["pt-BR"];

function t(key: TranslationKey, ...args: any[]): string {
  const translation = (translations[currentLanguage] as any)?.[key] ?? (translations["pt-BR"] as any)[key];
  if (typeof translation === "function") return translation(...args);
  return translation ?? String(key);
}

const statusKeyMap: { [key: string]: TranslationKey } = {
  PENDENTE: "STATUS_PENDENTE",
  "A CAMINHO": "STATUS_A_CAMINHO",
  ADIADO: "STATUS_ADIADO",
  ENTREGUE: "STATUS_ENTREGUE",
  CANCELADO: "STATUS_CANCELADO",
};

// --- APP STATE ---
let deliveryData: any[] = [];
let searchDebounceTimer: number;
let activeStatusFilter: string | null = null;

function updateStaticText() {
  document.title = t("pageTitle");
  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach((el) => {
    const key = el.dataset.i18n as TranslationKey;
    el.textContent = t(key);
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-placeholder]").forEach((el) => {
    const key = el.dataset.i18nPlaceholder as TranslationKey;
    (el as HTMLInputElement).placeholder = t(key);
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-title]").forEach((el) => {
    const key = el.dataset.i18nTitle as TranslationKey;
    el.title = t(key);
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-aria-label]").forEach((el) => {
    const key = (el as any).dataset.i18nArialabel as TranslationKey;
    el.setAttribute("aria-label", t(key));
  });
}

function setLanguage(lang: Language) {
  if (!(translations as any)[lang]) return;
  currentLanguage = lang;
  htmlEl.lang = lang;
  localStorage.setItem("language", lang);

  languageSwitcher?.querySelectorAll("button").forEach((btn) => {
    btn.classList.toggle("active", (btn as HTMLButtonElement).dataset.lang === lang);
  });

  updateStaticText();

  if (deliveryData.length > 0) applyFiltersAndRender();
  else resetUI();
}

languageSwitcher?.addEventListener("click", (event) => {
  const target = event.target as HTMLButtonElement;
  if (target && target.matches("[data-lang]")) setLanguage(target.dataset.lang as Language);
});

// --- THEME ---
const themeIcon = themeToggleBtn?.querySelector("i");

function setTheme(theme: "light" | "dark") {
  if (!themeIcon) return;
  htmlEl.classList.toggle("dark", theme === "dark");
  themeIcon.classList.toggle("fa-sun", theme === "light");
  themeIcon.classList.toggle("fa-moon", theme === "dark");
}

function toggleTheme() {
  const newTheme = htmlEl.classList.contains("dark") ? "light" : "dark";
  localStorage.setItem("theme", newTheme);
  setTheme(newTheme as any);
}

themeToggleBtn?.addEventListener("click", toggleTheme);

// --- TOAST ---
function showToast(message: string, type: "success" | "error" | "warning" = "success") {
  const toastContainer = document.getElementById("toast-container");
  if (!toastContainer) return;

  const toast = document.createElement("div");
  const icons = { success: "fa-check-circle", error: "fa-times-circle", warning: "fa-exclamation-triangle" };
  const colors = { success: "bg-green-500", error: "bg-red-500", warning: "bg-yellow-500" };

  toast.className = `toast ${colors[type]} text-white py-3 px-5 rounded-lg shadow-xl flex items-center mb-2`;
  toast.innerHTML = `<i class="fas ${icons[type]} mr-3" aria-hidden="true"></i> <p>${message}</p>`;
  toastContainer.appendChild(toast);
  setTimeout(() => toast.remove(), 5000);
}

// --- CONFIRM MODAL ---
function showConfirmationDialog(title: string, message: string): Promise<boolean> {
  const previouslyFocusedElement = document.activeElement as HTMLElement;

  return new Promise((resolve) => {
    modalTitle.textContent = title;
    modalMessage.textContent = message;

    modalContainer.classList.remove("hidden");
    setTimeout(() => modalContainer.classList.add("visible"), 10);

    modalConfirmBtn.focus();

    const closeModal = () => {
      modalContainer.classList.remove("visible");
      setTimeout(() => modalContainer.classList.add("hidden"), 200);
      previouslyFocusedElement?.focus();
    };

    const handleConfirm = () => {
      closeModal();
      resolve(true);
    };

    const handleCancel = () => {
      closeModal();
      resolve(false);
    };

    modalConfirmBtn.addEventListener("click", handleConfirm, { once: true });
    modalCancelBtn.addEventListener("click", handleCancel, { once: true });
  });
}

// --- LOGO (local + optional sync through Firebase state) ---
function handleLogoUpload(event: Event) {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) return;

  if (file.size > 2 * 1024 * 1024) {
    showToast(t("imageTooLarge"), "error");
    return;
  }

  const reader = new FileReader();
  reader.onload = async (e) => {
    if (typeof e.target?.result !== "string") {
      showToast(t("imageReadError"), "error");
      return;
    }
    const dataUrl = e.target.result;
    localStorage.setItem("companyLogo", dataUrl);
    companyLogo.src = dataUrl;
    logoContainer.classList.remove("hidden");
    showToast(t("logoUpdated"), "success");

    // Optional: sync logo via Firebase too (so other users see it)
    await saveStateToFirebase({ companyLogo: dataUrl });
  };

  reader.onerror = () => showToast(t("logoUploadError"), "error");
  reader.readAsDataURL(file);
  logoUpload.value = "";
}

function loadLogoFromStorage() {
  const savedLogo = localStorage.getItem("companyLogo");
  if (savedLogo) {
    companyLogo.src = savedLogo;
    logoContainer.classList.remove("hidden");
  }
}

logoUpload?.addEventListener("change", handleLogoUpload);

// --- FIREBASE INTEGRATION (same pattern as index.tsx 1) ---
let isUpdatingFromFirebase = false;

type FirebaseState = {
  deliveryData?: any[];
  lastUpdate?: any; // Firestore Timestamp
  lastUpdateSheetName?: string;
  companyLogo?: string;
};

const FIREBASE_COLLECTION = "delivery_dashboard";
const FIREBASE_DOC = "live_data";

async function saveStateToFirebase(patch: Partial<FirebaseState> = {}) {
  if (isUpdatingFromFirebase) {
    console.log("Firebase update in progress, skipping save to prevent loop.");
    return;
  }

  try {
    const stateToSave: FirebaseState = {
      deliveryData,
      lastUpdate: new Date(),
      lastUpdateSheetName: lastUpdate.dataset.sheetName || "",
      companyLogo: localStorage.getItem("companyLogo") || "",
      ...patch,
    };

    await db.collection(FIREBASE_COLLECTION).doc(FIREBASE_DOC).set(stateToSave, { merge: true });
    console.log("State saved to Firebase.");
  } catch (error) {
    console.error("Error saving state to Firebase:", error);
    showToast("Failed to sync changes with the server.", "error");
  }
}

function listenForRealtimeUpdates() {
  db.collection(FIREBASE_COLLECTION)
    .doc(FIREBASE_DOC)
    .onSnapshot(
      (doc: any) => {
        isUpdatingFromFirebase = true;
        console.log("Received update from Firebase.");

        if (doc.exists) {
          const data: FirebaseState = doc.data() || {};

          deliveryData = Array.isArray(data.deliveryData) ? data.deliveryData : [];
          activeStatusFilter = null;
          searchInput.value = "";

          // Sync logo if present
          if (data.companyLogo && typeof data.companyLogo === "string") {
            localStorage.setItem("companyLogo", data.companyLogo);
            companyLogo.src = data.companyLogo;
            logoContainer.classList.toggle("hidden", !data.companyLogo);
          }

          const lastUpdateDate = data.lastUpdate?.toDate ? data.lastUpdate.toDate() : null;
          const sheetName = data.lastUpdateSheetName || "Sheet";
          if (lastUpdateDate) {
            lastUpdate.dataset.sheetName = sheetName;
            lastUpdate.textContent = t("lastUpdateText", sheetName, lastUpdateDate.toLocaleString(currentLanguage));
          }

          if (deliveryData.length > 0) applyFiltersAndRender();
          else resetUI();
        } else {
          resetUI();
        }

        setTimeout(() => {
          isUpdatingFromFirebase = false;
        }, 250);
      },
      (error: any) => {
        console.error("Firebase listener error:", error);
        showToast("Connection to the server was lost. Please check your internet.", "error");
      }
    );
}

// --- DATA HELPERS ---
function findDeliverySheet(workbook: any): string {
  const keywords = [
    "DELIVERY",
    "SCHEDULE",
    "MONDAY",
    "TUESDAY",
    "WEDNESDAY",
    "THURSDAY",
    "FRIDAY",
    "SATURDAY",
    "SUNDAY",
    "SEGUNDA",
    "TERÇA",
    "QUARTA",
    "QUINTA",
    "SEXTA",
    "SÁBADO",
    "DOMINGO",
  ];
  return (
    workbook.SheetNames.find((name: string) => {
      const upperName = name.toUpperCase();
      return keywords.some((key) => upperName.includes(key));
    }) || workbook.SheetNames[0]
  );
}

function excelDateToJSDate(serial: string | number): Date | null {
  if (!serial) return null;

  if (typeof serial === "string") {
    const cleaned = serial.trim();
    if (cleaned.includes("/") || cleaned.includes("-")) {
      const parts = cleaned.split(/[/\-]/);
      if (parts.length === 3) {
        let d = parseInt(parts[0], 10),
          m = parseInt(parts[1], 10),
          y = parseInt(parts[2], 10);
        if (y < 100) y += 2000;
        if (m > 12) {
          const t = d;
          d = m;
          m = t;
        }
        const date = new Date(y, m - 1, d);
        return isNaN(date.getTime()) ? null : date;
      }
    }
    if (/^\d+$/.test(cleaned)) serial = parseInt(cleaned, 10);
    else return null;
  }

  if (typeof serial !== "number" || serial < 1) return null;
  const dateInfo = new Date((serial - 25569) * 86400 * 1000);
  return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
}

function getStatusDetails(status: string) {
  const upperStatus = (status || "PENDENTE").toUpperCase();
  switch (upperStatus) {
    case "ENTREGUE":
      return { icon: "fa-check-circle", pillBg: "bg-green-100 dark:bg-green-900/50", pillText: "text-green-700 dark:text-green-300" };
    case "A CAMINHO":
      return { icon: "fa-truck", pillBg: "bg-yellow-100 dark:bg-yellow-900/50", pillText: "text-yellow-700 dark:text-yellow-300" };
    case "ADIADO":
      return { icon: "fa-calendar-alt", pillBg: "bg-blue-100 dark:bg-blue-900/50", pillText: "text-blue-700 dark:text-blue-300" };
    case "CANCELADO":
      return { icon: "fa-times-circle", pillBg: "bg-red-100 dark:bg-red-900/50", pillText: "text-red-700 dark:text-red-300" };
    default:
      return { icon: "fa-hourglass-half", pillBg: "bg-slate-200 dark:bg-slate-700", pillText: "text-slate-700 dark:text-slate-200" };
  }
}

function getStatusPill(status: string): string {
  const upperStatus = (status || "PENDENTE").toUpperCase();
  const details = getStatusDetails(upperStatus);
  return `<span class="status-pill ${details.pillBg} ${details.pillText}">
    <i class="fas ${details.icon} fa-fw"></i>
    <span>${t(statusKeyMap[upperStatus] || "STATUS_PENDENTE")}</span>
  </span>`;
}

// --- UI CORE ---
function resetUI() {
  placeholder.classList.remove("hidden");
  deliveryDashboard.classList.add("hidden");
  summaryStats.classList.add("hidden");
  exportExcelBtn.classList.add("hidden");
  exportPdfBtn.classList.add("hidden");
  deliveryTabs.innerHTML = "";
  deliveryContent.innerHTML = "";
  lastUpdate.textContent = t("uploadPrompt");
}

function applyFiltersAndRender(activeTabId: string | null = null) {
  const query = (searchInput.value || "").trim().toLowerCase();
  let filteredData = deliveryData;

  if (activeStatusFilter) {
    if (activeStatusFilter === "PENDENTE") {
      filteredData = filteredData.filter((row) => {
        const status = String(row["STATUS"] || "").toUpperCase();
        return !["ENTREGUE", "A CAMINHO", "ADIADO", "CANCELADO"].includes(status);
      });
    } else {
      filteredData = filteredData.filter((row) => String(row["STATUS"] || "").toUpperCase() === activeStatusFilter);
    }
  }

  if (query) {
    filteredData = filteredData.filter((row) => Object.values(row).some((value) => String(value).toLowerCase().includes(query)));
  }

  renderDeliveryDashboard(filteredData, activeTabId);
  updateStats();
}

function updateStats() {
  const total = deliveryData.length;
  const delivered = deliveryData.filter((d) => String(d["STATUS"] || "").toUpperCase() === "ENTREGUE").length;
  const inTransit = deliveryData.filter((d) => String(d["STATUS"] || "").toUpperCase() === "A CAMINHO").length;
  const postponed = deliveryData.filter((d) => String(d["STATUS"] || "").toUpperCase() === "ADIADO").length;
  const canceled = deliveryData.filter((d) => String(d["STATUS"] || "").toUpperCase() === "CANCELADO").length;
  const pending = total - delivered - inTransit - postponed - canceled;

  const getCardClasses = (cardStatus: string | null) => {
    const isAll = cardStatus === "ALL";
    const isActive = activeStatusFilter === cardStatus || (activeStatusFilter === null && isAll);
    let classes =
      "summary-card bg-white dark:bg-slate-800 p-5 rounded-lg shadow-sm border flex items-center cursor-pointer transition-all duration-200";
    if (isActive) classes += " border-blue-500 ring-2 ring-blue-500/50 scale-[1.02] z-10";
    else classes += " border-slate-200 dark:border-slate-700 hover:border-blue-300";
    return classes;
  };

  summaryStats.innerHTML = `
    <div class="${getCardClasses("ALL")}" data-status="ALL">
      <div class="bg-blue-100 dark:bg-blue-900/50 text-blue-600 dark:text-blue-400 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-box-open text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("totalContainers")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${total}</div>
      </div>
    </div>

    <div class="${getCardClasses("ENTREGUE")}" data-status="ENTREGUE">
      <div class="bg-green-100 dark:bg-green-900/50 text-green-600 dark:text-green-400 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-check-circle text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("delivered")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${delivered}</div>
      </div>
    </div>

    <div class="${getCardClasses("A CAMINHO")}" data-status="A CAMINHO">
      <div class="bg-yellow-100 dark:bg-yellow-900/50 text-yellow-600 dark:text-yellow-400 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-truck text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("inTransit")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${inTransit}</div>
      </div>
    </div>

    <div class="${getCardClasses("PENDENTE")}" data-status="PENDENTE">
      <div class="bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-hourglass-half text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("pending")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${pending}</div>
      </div>
    </div>

    <div class="${getCardClasses("ADIADO")}" data-status="ADIADO">
      <div class="bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-calendar-alt text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("postponed")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${postponed}</div>
      </div>
    </div>

    <div class="${getCardClasses("CANCELADO")}" data-status="CANCELADO">
      <div class="bg-red-100 dark:bg-red-900/50 text-red-600 dark:text-red-400 rounded-full h-12 w-12 flex items-center justify-center mr-4 flex-shrink-0">
        <i class="fas fa-times-circle text-xl"></i>
      </div>
      <div>
        <div class="text-slate-500 dark:text-slate-400 text-sm font-medium">${t("canceled")}</div>
        <div class="text-3xl font-extrabold text-slate-800 dark:text-slate-100">${canceled}</div>
      </div>
    </div>
  `;
}

function renderDeliveryDashboard(data: any[], activeTabId: string | null = null) {
  placeholder.classList.add("hidden");
  deliveryDashboard.classList.remove("hidden");
  summaryStats.classList.remove("hidden");
  exportExcelBtn.classList.remove("hidden");
  exportPdfBtn.classList.remove("hidden");

  deliveryTabs.innerHTML = "";
  deliveryContent.innerHTML = "";

  if (data.length === 0) {
    deliveryTabs.classList.add("hidden");
    deliveryContent.innerHTML = `
      <div class="text-center py-20 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
        <i class="fas fa-search text-6xl text-slate-300 dark:text-slate-600 mb-4"></i>
        <h2 class="text-2xl font-semibold text-slate-700 dark:text-slate-200">${t("noResultsTitle")}</h2>
        <p class="text-slate-500 dark:text-slate-400 mt-2">${t("noResultsMessage")}</p>
      </div>`;
    return;
  }

  deliveryTabs.classList.remove("hidden");

  const groupedByDate = data.reduce((acc, row) => {
    const dateStr = String(row["DELIVERY AT BYD"] || "").trim();
    const jsDate = excelDateToJSDate(dateStr);
    const finalKey = jsDate ? jsDate.toLocaleDateString("en-US") : dateStr || t("undefinedDate");
    if (!acc[finalKey]) acc[finalKey] = [];
    acc[finalKey].push(row);
    return acc;
  }, {} as Record<string, any[]>);

  const sortedDates = Object.keys(groupedByDate).sort((a, b) => {
    if (a === t("undefinedDate")) return 1;
    if (b === t("undefinedDate")) return -1;
    const dA = new Date(a);
    const dB = new Date(b);
    return dA.getTime() - dB.getTime();
  });

  sortedDates.forEach((dateKey, index) => {
    const deliveries = groupedByDate[dateKey];
    const dateObj = new Date(dateKey);
    const formattedDate = isNaN(dateObj.getTime())
      ? dateKey
      : dateObj.toLocaleDateString(currentLanguage, { day: "2-digit", month: "2-digit", year: "2-digit" });

    const contentId = `content-${index}`;
    const isActive = activeTabId ? contentId === activeTabId : index === 0;

    const tabBtn = document.createElement("button");
    tabBtn.className = `tab-btn flex-shrink-0 px-4 py-3 text-sm font-semibold transition-colors duration-200 flex items-center space-x-2 ${
      isActive ? "active" : ""
    }`;
    tabBtn.innerHTML = `<span class="font-bold">${formattedDate}</span> 
      <span class="tab-count-badge bg-slate-200 dark:bg-slate-700 dark:text-slate-200 text-slate-600 font-bold">${deliveries.length}</span>`;
    tabBtn.dataset.target = contentId;
    deliveryTabs.appendChild(tabBtn);

    const card = document.createElement("div");
    card.id = contentId;
    card.className = `date-card bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 ${!isActive ? "hidden" : ""}`;

    const deliveredInCard = deliveries.filter((d) => String(d["STATUS"] || "").toUpperCase() === "ENTREGUE").length;
    const totalInCard = deliveries.length;
    const percentage = totalInCard > 0 ? (deliveredInCard / totalInCard) * 100 : 0;

    // Carrier breakdown
    const carrierStats: Record<string, { total: number; delivered: number }> = {};
    deliveries.forEach((d) => {
      const carrier = String(d["TRANSPORTATION COMPANY"] || "N/A").trim();
      if (!carrierStats[carrier]) carrierStats[carrier] = { total: 0, delivered: 0 };
      carrierStats[carrier].total++;
      if (String(d["STATUS"] || "").toUpperCase() === "ENTREGUE") carrierStats[carrier].delivered++;
    });

    const carrierBreakdownHTML = Object.entries(carrierStats)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([carrier, stats]) => {
        const carrierPercent = stats.total > 0 ? (stats.delivered / stats.total) * 100 : 0;
        return `
          <div class="bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col justify-between transition-all hover:border-blue-300 dark:hover:border-blue-700">
            <div class="flex justify-between items-start mb-2">
              <span class="font-bold text-sm text-slate-700 dark:text-slate-200 truncate pr-2" title="${carrier}">${carrier}</span>
              <span class="text-[10px] font-bold text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/40 px-1.5 py-0.5 rounded">${carrierPercent.toFixed(
                0
              )}%</span>
            </div>
            <div class="flex items-center justify-between mb-2">
              <span class="text-xs text-slate-500 dark:text-slate-400">${t("totalContainers")}: <strong class="text-slate-700 dark:text-slate-200">${
                stats.total
              }</strong></span>
              <span class="text-xs text-slate-500 dark:text-slate-400">${t("delivered")}: <strong class="text-green-600 dark:text-green-400">${
                stats.delivered
              }</strong></span>
            </div>
            <div class="w-full bg-slate-200 dark:bg-slate-700 h-1.5 rounded-full overflow-hidden">
              <div class="bg-blue-500 h-full transition-all duration-700" style="width: ${carrierPercent}%"></div>
            </div>
          </div>`;
      })
      .join("");

    const tableRows = deliveries
      .map((row, rowIndex) => {
        const status = String(row["STATUS"] || "PENDENTE").toUpperCase();
        const materialType = String(row["TYPE OF MATERIAL"] || "").toUpperCase();
        const isBattery = materialType.includes("BATTERY") || materialType.includes("BATERIA");

        let rowClass = "transition-colors hover:bg-slate-50 dark:hover:bg-slate-700/50 cursor-pointer";
        if (isBattery) rowClass += " is-battery";
        if (status === "ENTREGUE") rowClass += " bg-green-50/30 dark:bg-green-900/10 opacity-80";

        const batteryBadge = isBattery
          ? `<span class="ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 dark:bg-amber-900 dark:text-amber-200 uppercase"><i class="fas fa-bolt mr-1"></i>${t(
              "badgeBattery"
            )}</span>`
          : "";

        const options = ["PENDENTE", "A CAMINHO", "ADIADO", "ENTREGUE", "CANCELADO"];
        const selectHTML = `
          <select class="status-select bg-white dark:bg-slate-700 dark:text-slate-200 border border-slate-300 dark:border-slate-500 text-xs rounded-md p-1 w-full"
                  data-original-index="${row.originalIndex}">
            ${options
              .map((opt) => `<option value="${opt}" ${status === opt ? "selected" : ""}>${t(statusKeyMap[opt])}</option>`)
              .join("")}
          </select>`;

        return `<tr class="${rowClass}" data-original-index="${row.originalIndex}">
          <td class="px-4 py-3 text-xs text-center border-l-4 ${isBattery ? "border-amber-500" : "border-transparent"}">${rowIndex + 1}</td>
          <td class="px-4 py-3 text-xs font-semibold text-slate-800 dark:text-slate-100">${row["CONTAINER"] || "-"} ${batteryBadge}</td>
          <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-mono">${row["BL"] || "-"}</td>
          <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["TRANSPORTATION COMPANY"] || "-"}</td>
          <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["VESSEL"] || "-"}</td>
          <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["BONDED WAREHOUSE"] || "-"}</td>
          <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-medium">${row["LOT"] || "-"}</td>
          <td class="px-4 py-3 text-xs">${status === "ENTREGUE" || status === "CANCELADO" ? getStatusPill(status) : selectHTML}</td>
        </tr>`;
      })
      .join("");

    card.innerHTML = `
      <div class="p-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50 rounded-t-lg">
        <div class="flex justify-between items-center mb-2">
          <h3 class="font-bold text-lg text-slate-800 dark:text-slate-100">${formattedDate}</h3>
          <span class="text-sm font-medium text-slate-500 dark:text-slate-400">${t("containersDelivered", deliveredInCard, totalInCard)}</span>
        </div>
        <div class="progress-bar"><div class="progress-bar-inner" style="width: ${percentage}%"></div></div>
      </div>

      <div class="p-4 bg-slate-50/50 dark:bg-slate-900/30 border-b border-slate-200 dark:border-slate-700">
        <h4 class="text-xs font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest mb-4 flex items-center">
          <i class="fas fa-chart-line mr-2 text-blue-500"></i> ${t("performanceTitle")}
        </h4>
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
          ${carrierBreakdownHTML}
        </div>
      </div>

      <div class="table-responsive">
        <table class="min-w-full text-sm">
          <thead>
            <tr class="border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50">
              <th class="px-4 py-2 text-center text-slate-500 text-xs uppercase w-12">${t("tableHeaderRow")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderContainer")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderBL")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderCompany")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderVessel")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderWarehouse")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderLot")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase w-40">${t("tableHeaderStatus")}</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100 dark:divide-slate-700">${tableRows}</tbody>
        </table>
      </div>
    `;

    deliveryContent.appendChild(card);
  });
}

// --- DETAILS EXPAND (kept from your original structure) ---
function handleRowInteraction(row: HTMLTableRowElement) {
  if (!row || row.classList.contains("details-row")) return;

  const table = row.closest("table");
  if (!table) return;

  const currentlyExpandedRow = table.querySelector("tr.is-expanded") as HTMLTableRowElement | null;
  const isAlreadyExpanded = row.classList.contains("is-expanded");

  if (currentlyExpandedRow) {
    currentlyExpandedRow.classList.remove("is-expanded");
    const existingDetails = currentlyExpandedRow.nextElementSibling as HTMLTableRowElement | null;
    if (existingDetails && existingDetails.classList.contains("details-row")) {
      const wrapper = existingDetails.querySelector(".details-content-wrapper") as HTMLDivElement | null;
      if (wrapper) {
        wrapper.classList.remove("expanded");
        setTimeout(() => existingDetails.remove(), 350);
      } else existingDetails.remove();
    }
  }

  if (!isAlreadyExpanded) {
    row.classList.add("is-expanded");

    const originalIndex = parseInt(row.dataset.originalIndex || "", 10);
    const rowData = deliveryData[originalIndex];

    const newDetailsRow = document.createElement("tr");
    newDetailsRow.className = "details-row";

    const detailsCell = document.createElement("td");
    detailsCell.colSpan = 8;
    detailsCell.className = "details-cell";

    detailsCell.innerHTML = `
      <div class="details-content-wrapper bg-slate-50 dark:bg-slate-900/50">
        <div class="grid grid-cols-1 md:grid-cols-4 gap-x-6 gap-y-4">
          <div>
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsCompany")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${rowData["TRANSPORTATION COMPANY"] || "-"}</p>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsVessel")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${rowData["VESSEL"] || "-"}</p>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsWarehouse")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${rowData["BONDED WAREHOUSE"] || "-"}</p>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsLot")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${rowData["LOT"] || "-"}</p>
          </div>
          <div class="md:col-span-2">
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsMaterial")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${rowData["TYPE OF MATERIAL"] || "-"}</p>
          </div>
          <div class="md:col-span-2">
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsNotes")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100 italic">${rowData["NOTES"] || "-"}</p>
          </div>
        </div>
      </div>
    `;

    newDetailsRow.appendChild(detailsCell);
    row.after(newDetailsRow);

    setTimeout(() => {
      const wrapper = newDetailsRow.querySelector(".details-content-wrapper") as HTMLDivElement | null;
      if (wrapper) wrapper.classList.add("expanded");
    }, 10);
  }
}

deliveryContent?.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;
  const row = target.closest<HTMLTableRowElement>("tbody tr:not(.details-row)");
  if (row && !target.closest(".status-select")) handleRowInteraction(row);
});

// --- STATUS CHANGE (with confirmation + Firebase sync) ---
deliveryContent?.addEventListener("change", async (event) => {
  const target = event.target as HTMLElement;
  const select = target.closest<HTMLSelectElement>(".status-select");
  if (!select) return;

  const originalIndex = parseInt(select.dataset.originalIndex || "", 10);
  if (Number.isNaN(originalIndex) || !deliveryData[originalIndex]) return;

  const rowData = deliveryData[originalIndex];
  const containerId = rowData["CONTAINER"] || rowData["BL"] || "";
  const newStatus = String((select as HTMLSelectElement).value || "PENDENTE").toUpperCase();
  const oldStatus = String(rowData["STATUS"] || "PENDENTE").toUpperCase();

  if (newStatus === oldStatus) return;

  const ok = await showConfirmationDialog(
    t("confirmStatusChangeTitle"),
    t("confirmStatusChangeMessage", containerId, newStatus)
  );

  if (!ok) {
    (select as HTMLSelectElement).value = oldStatus;
    return;
  }

  rowData["STATUS"] = newStatus;
  showToast(t("statusUpdated", containerId, newStatus), "success");

  // Persist to Firebase (and all users update via onSnapshot)
  await saveStateToFirebase({ deliveryData });
});

// --- TABS CLICK ---
deliveryTabs?.addEventListener("click", (event) => {
  const btn = (event.target as HTMLElement).closest<HTMLButtonElement>(".tab-btn");
  if (!btn) return;

  const targetId = btn.dataset.target;
  if (!targetId) return;

  deliveryTabs.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  btn.classList.add("active");

  deliveryContent.querySelectorAll<HTMLElement>(".date-card").forEach((card) => {
    card.classList.toggle("hidden", card.id !== targetId);
  });
});

// --- SEARCH ---
searchInput?.addEventListener("input", () => {
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = window.setTimeout(() => applyFiltersAndRender(), 250);
});

// --- SUMMARY FILTER CLICK ---
summaryStats?.addEventListener("click", (event) => {
  const card = (event.target as HTMLElement).closest<HTMLDivElement>("[data-status]");
  if (!card || !card.dataset.status) return;

  const status = card.dataset.status;
  activeStatusFilter = status === "ALL" ? null : activeStatusFilter === status ? null : status;
  applyFiltersAndRender();
});

// --- FILE UPLOAD -> parse -> set state -> save to firebase ---
fileUpload?.addEventListener("change", (event) => {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  const uploadLabel = document.querySelector('label[for="file-upload"]');
  if (!file || !uploadLabel) return;

  const labelSpan = uploadLabel.querySelector("span");
  uploadLabel.classList.add("opacity-50", "cursor-not-allowed");
  if (labelSpan) labelSpan.textContent = t("processing");
  uploadLabel.querySelector("i")?.classList.add("fa-spin");

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      if (!e.target?.result) throw new Error(t("fileReadError"));

      const workbook = XLSX.read(new Uint8Array(e.target.result as ArrayBuffer), { type: "array" });
      const sheetName = findDeliverySheet(workbook);
      const sheet = workbook.Sheets[sheetName];

      const rawData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      // Find header row
      let headerRowIndex = -1;
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i].map((c) => String(c).toUpperCase().trim());
        if (row.includes("CONTAINER") || row.includes("DELIVERY AT BYD")) {
          headerRowIndex = i;
          break;
        }
      }
      if (headerRowIndex === -1) headerRowIndex = rawData.length > 7 ? 7 : 0;

      const headers = rawData[headerRowIndex].map((h) => String(h).toUpperCase().trim());
      const dataRows = rawData.slice(headerRowIndex + 1);

      const mapping: Record<string, number> = {
        "DELIVERY AT BYD": headers.indexOf("DELIVERY AT BYD"),
        CONTAINER: headers.indexOf("CONTAINER"),
        BL: headers.indexOf("BL"),
        VESSEL: headers.indexOf("VESSEL"),
        "BONDED WAREHOUSE": headers.indexOf("BONDED WAREHOUSE"),
        "TYPE OF MATERIAL": headers.indexOf("TYPE OF MATERIAL"),
        STATUS: headers.indexOf("STATUS"),
        "TRANSPORTATION COMPANY": headers.indexOf("TRANSPORTATION COMPANY"),
        LOT: headers.indexOf("LOT"),
        NOTES: headers.indexOf("NOTES"),
      };

      // Fallbacks
      if (mapping.CONTAINER === -1) mapping.CONTAINER = 10;
      if (mapping.BL === -1) mapping.BL = 11;
      if (mapping.VESSEL === -1) mapping.VESSEL = 12;
      if (mapping["BONDED WAREHOUSE"] === -1) mapping["BONDED WAREHOUSE"] = 13;
      if (mapping["DELIVERY AT BYD"] === -1) mapping["DELIVERY AT BYD"] = 0;
      if (mapping["TRANSPORTATION COMPANY"] === -1) mapping["TRANSPORTATION COMPANY"] = 3;
      if (mapping.LOT === -1) mapping.LOT = 18;

      deliveryData = dataRows
        .filter((row) => row[mapping.CONTAINER] || row[mapping.BL])
        .map((row, idx) => ({
          "DELIVERY AT BYD": row[mapping["DELIVERY AT BYD"]] || "",
          CONTAINER: row[mapping.CONTAINER] || "",
          BL: row[mapping.BL] || "",
          VESSEL: row[mapping.VESSEL] || "",
          "BONDED WAREHOUSE": row[mapping["BONDED WAREHOUSE"]] || "",
          "TYPE OF MATERIAL": row[mapping["TYPE OF MATERIAL"]] || "",
          "TRANSPORTATION COMPANY": row[mapping["TRANSPORTATION COMPANY"]] || "",
          LOT: row[mapping.LOT] || "",
          NOTES: row[mapping.NOTES] || "",
          STATUS: row[mapping.STATUS] || "PENDENTE",
          originalIndex: idx,
        }));

      if (deliveryData.length === 0) throw new Error(t("emptySheetError"));

      searchInput.value = "";
      activeStatusFilter = null;

      lastUpdate.dataset.sheetName = sheetName;
      lastUpdate.textContent = t("lastUpdateText", sheetName, new Date().toLocaleString(currentLanguage));
      showToast(t("sheetLoaded"), "success");

      applyFiltersAndRender();

      // Save to Firebase (and all users update)
      await saveStateToFirebase({
        deliveryData,
        lastUpdateSheetName: sheetName,
      });
    } catch (err: any) {
      console.error(err);
      showToast(err.message || t("fileProcessError"), "error");
      resetUI();
    } finally {
      uploadLabel.classList.remove("opacity-50", "cursor-not-allowed");
      if (labelSpan) labelSpan.textContent = t("uploadSheetButton");
      uploadLabel.querySelector("i")?.classList.remove("fa-spin");
      fileUpload.value = "";
    }
  };

  reader.readAsArrayBuffer(file);
});

// --- EXPORTS ---
async function exportToExcel() {
  if (!deliveryData.length) {
    showToast(t("noDataToExport"), "warning");
    return;
  }

  const ok = await showConfirmationDialog(t("exportExcelTitle"), t("exportExcelMessage"));
  if (!ok) return;

  const ws = XLSX.utils.json_to_sheet(deliveryData.map((d) => ({ ...d })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Deliveries");
  XLSX.writeFile(wb, "delivery_dashboard.xlsx");

  showToast(t("excelGenerated"), "success");
}

async function exportToPdf() {
  if (!deliveryData.length) {
    showToast(t("noDataToExport"), "warning");
    return;
  }

  const ok = await showConfirmationDialog(t("exportPdfTitle"), t("exportPdfMessage"));
  if (!ok) return;

  const { jsPDF } = jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

  // Try to use autoTable if available, otherwise print minimal info
  const hasAutoTable = typeof (doc as any).autoTable === "function";

  doc.setFontSize(14);
  doc.text(t("pdfTitle"), 40, 40);
  doc.setFontSize(9);
  doc.text(t("pdfGeneratedOn", new Date().toLocaleString(currentLanguage)), 40, 58);

  if (hasAutoTable) {
    const headers = [["#", "DELIVERY", "CONTAINER", "BL", "VESSEL", "CARRIER", "WAREHOUSE", "LOT", "STATUS"]];
    const body = deliveryData.map((d, idx) => [
      String(idx + 1),
      String(d["DELIVERY AT BYD"] || ""),
      String(d["CONTAINER"] || ""),
      String(d["BL"] || ""),
      String(d["VESSEL"] || ""),
      String(d["TRANSPORTATION COMPANY"] || ""),
      String(d["BONDED WAREHOUSE"] || ""),
      String(d["LOT"] || ""),
      String(d["STATUS"] || ""),
    ]);

    (doc as any).autoTable({
      head: headers,
      body,
      startY: 80,
      styles: { fontSize: 8 },
      headStyles: { fillColor: [30, 64, 175] },
      margin: { left: 40, right: 40 },
    });
  } else {
    doc.setFontSize(10);
    doc.text("autoTable plugin not found. Showing summary only.", 40, 90);
    doc.text(`Rows: ${deliveryData.length}`, 40, 110);
  }

  doc.save("delivery_dashboard.pdf");
  showToast(t("pdfGenerated"), "success");
}

exportExcelBtn?.addEventListener("click", exportToExcel);
exportPdfBtn?.addEventListener("click", exportToPdf);

// --- INIT ---
document.addEventListener("DOMContentLoaded", () => {
  // Theme
  const savedTheme = (localStorage.getItem("theme") as "light" | "dark") || (htmlEl.classList.contains("dark") ? "dark" : "light");
  setTheme(savedTheme);

  // Logo
  loadLogoFromStorage();

  // Language
  const savedLang = localStorage.getItem("language") as Language;
  const browserLang = navigator.language;
  let initialLang: Language = "pt-BR";
  if (savedLang && (translations as any)[savedLang]) initialLang = savedLang;
  else if (browserLang.startsWith("en") && (translations as any)["en-US"]) initialLang = "en-US";
  else if (browserLang.startsWith("zh") && (translations as any)["zh-CN"]) initialLang = "zh-CN";
  setLanguage(initialLang);

  // Start Firebase realtime listener
  listenForRealtimeUpdates();

  // Initial UI
  resetUI();
});
