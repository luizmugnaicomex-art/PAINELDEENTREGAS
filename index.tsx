/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * Container Delivery Dashboard — Studio AI (single file)
 * Improved to match the attached sheet (e.g. "DELIVERY SCHEDULE - 02.19.xlsx"):
 * - Robust header mapping (accent/case/spacing tolerant)
 * - Supports Excel Date objects / serials / "dd/mm/yyyy" strings
 * - Sanitizes Excel error strings (#REF!, #N/A, etc.)
 * - Keeps stable row id (_id) so status updates won’t break after filtering/sorting
 * - Expanded Details panel with the extra columns present in your sheet
 * - Adds Daily Goal (150 on weekdays; weekends show goal as “bonus”) per date card
 */

/* ------------------------- CDN typings (Studio AI) ------------------------- */
declare const firebase: any;
declare const XLSX: any;
declare const jspdf: any;

/* ----------------------------- FIREBASE SAFE ------------------------------ */
const env = (import.meta as any).env || (process as any).env || {};
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

let db: any = null;
try {
  if (firebaseConfig.apiKey && typeof firebase !== "undefined") {
    if (!firebase.apps || !firebase.apps.length) firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
  }
} catch (e) {
  console.warn("Firebase init skipped:", e);
  db = null;
}

/* ------------------------------- DOM Elements ------------------------------ */
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
const batteryFilterBtn = document.getElementById("battery-filter-btn") as HTMLButtonElement;
const kdFilterBtn = document.getElementById("kd-filter-btn") as HTMLButtonElement;
const projectFilterBtn = document.getElementById("project-filter-btn") as HTMLButtonElement;
const lotSearchInput = document.getElementById("lot-search-input") as HTMLInputElement;
const lotSearchContainer = document.getElementById("lot-search-container") as HTMLDivElement;
const htmlEl = document.documentElement;

/* ------------------------------- Logo Elements ----------------------------- */
const logoUpload = document.getElementById("logo-upload") as HTMLInputElement;
const logoContainer = document.getElementById("logo-container") as HTMLDivElement;
const companyLogo = document.getElementById("company-logo") as HTMLImageElement;

/* ------------------------------ Modal Elements ----------------------------- */
const modalContainer = document.getElementById("confirmation-modal-container") as HTMLDivElement;
const modalTitle = document.getElementById("modal-title") as HTMLHeadingElement;
const modalMessage = document.getElementById("modal-message") as HTMLParagraphElement;
const modalConfirmBtn = document.getElementById("modal-confirm-btn") as HTMLButtonElement;
const modalCancelBtn = document.getElementById("modal-cancel-btn") as HTMLButtonElement;

/* ------------------------------ Language Elements -------------------------- */
const languageSwitcher = document.getElementById("language-switcher") as HTMLDivElement;

/* ------------------------------- i18n ------------------------------------ */
const translations = {
  "pt-BR": {
    pageTitle: "Painel de Entregas de Contêineres",
    headerTitle: "Painel de Entregas",
    uploadPrompt: "Carregue sua planilha de agendamento para começar",
    searchInputPlaceholder: "Pesquisar container, BL, navio...",
    searchLotPlaceholder: "Pesquisar LOTE",
    uploadLogoTooltip: "Carregar logo da empresa",
    toggleThemeTooltip: "Alternar tema",
    uploadSheetButton: "Carregar",
    uploadSheetTooltip: "Carregar Planilha",
    filterBatteryTooltip: "Filtrar Baterias",
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
    awaitingUnload: "Aguardando Desova",
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
    tableHeaderModel: "Modelo",
    tableHeaderOperation: "Escopo da Operação",
    STATUS_PENDENTE: "Pendente",
    STATUS_A_CAMINHO: "A Caminho",
    STATUS_ADIADO: "Adiado",
    STATUS_ENTREGUE: "Entregue",
    STATUS_CANCELADO: "Cancelado",
    STATUS_AGUARDANDO_DESOVA: "Aguardando Desova",
    detailsTitle: "Detalhes",
    detailsVessel: "Navio (Vessel)",
    detailsWarehouse: "Armazém",
    detailsNotes: "Observações",
    detailsMaterial: "Tipo de Material",
    detailsLot: "Lote (LOT)",
    detailsCompany: "Transportadora",
    performanceTitle: "Desempenho por Transportadora",
    badgeBattery: "Bateria",
    deliveriesTab: "Entregas",
    arrivalsTab: "Arrivals per Lot",
    pdfTitle: "Programação de Entregas de Contêineres",
    pdfGeneratedOn: (date: string) => `Relatório gerado em: ${date}`,
    pdfPage: (page: number, total: number) => `Página ${page} de ${total}`,
    lastUpdateText: (sheet: string, date: string) => `Dados de "${sheet}" | Carregado em: ${date}`,
    clickToExpand: "Clique para expandir",
    changeStatusFor: (containerId: string) => `Alterar status do container ${containerId || ""}`,
    viewDetailsFor: (containerId: string) => `Ver detalhes do container ${containerId || "sem identificação"}`,
    goalLabel: "Meta",
    goalWeekday: "150/dia útil",
    goalWeekend: "Fim de semana (bônus)",
    reachedGoal: "Meta atingida",
    notReachedGoal: "Abaixo da meta",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
  },
  "en-US": {
    pageTitle: "Container Delivery Dashboard",
    headerTitle: "Delivery Dashboard",
    uploadPrompt: "Upload your schedule spreadsheet to begin",
    searchInputPlaceholder: "Search container, BL, vessel...",
    searchLotPlaceholder: "Search LOT",
    uploadLogoTooltip: "Upload company logo",
    toggleThemeTooltip: "Toggle theme",
    uploadSheetButton: "Upload",
    uploadSheetTooltip: "Upload Spreadsheet",
    filterBatteryTooltip: "Filter Batteries",
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
    awaitingUnload: "Awaiting Unload",
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
    tableHeaderModel: "Model",
    tableHeaderOperation: "Operation Scope",
    STATUS_PENDENTE: "Pending",
    STATUS_A_CAMINHO: "In Transit",
    STATUS_ADIADO: "Postponed",
    STATUS_ENTREGUE: "Delivered",
    STATUS_CANCELADO: "Canceled",
    STATUS_AGUARDANDO_DESOVA: "Awaiting Unload",
    detailsTitle: "Details",
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
    goalLabel: "Goal",
    goalWeekday: "150/weekday",
    goalWeekend: "Weekend (bonus)",
    reachedGoal: "Goal reached",
    notReachedGoal: "Below goal",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
  },
  "zh-CN": {
    pageTitle: "集装箱交付仪表板",
    headerTitle: "交付仪表板",
    uploadPrompt: "上传您的排程电子表格以开始",
    searchInputPlaceholder: "搜索集装箱、提单 (BL)、船名...",
    searchLotPlaceholder: "搜索批号 (LOT)",
    uploadLogoTooltip: "上传公司标志",
    toggleThemeTooltip: "切换主题",
    uploadSheetButton: "上传",
    uploadSheetTooltip: "上传电子表格",
    filterBatteryTooltip: "过滤电池",
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
    awaitingUnload: "等待卸货",
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
    tableHeaderModel: "型号",
    tableHeaderOperation: "操作范围",
    STATUS_PENDENTE: "待处理",
    STATUS_A_CAMINHO: "运输中",
    STATUS_ADIADO: "已推迟",
    STATUS_ENTREGUE: "已交付",
    STATUS_CANCELADO: "已取消",
    STATUS_AGUARDANDO_DESOVA: "等待卸货",
    detailsTitle: "详细信息",
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
    goalLabel: "目标",
    goalWeekday: "工作日150",
    goalWeekend: "周末（加分）",
    reachedGoal: "已达目标",
    notReachedGoal: "未达目标",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
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
  "AGUARDANDO DESOVA": "STATUS_AGUARDANDO_DESOVA",
};

/* ------------------------------ APP STATE --------------------------------- */
type DeliveryRow = Record<string, any> & {
  _id: string;
};

let deliveryData: DeliveryRow[] = [];
let searchDebounceTimer: number;
let activeStatusFilter: string | null = null;
let showOnlyBattery: boolean = false;
let showOnlyKd: boolean = false;
let showOnlyProject: boolean = false;

/* ------------------------------ STATIC TEXT -------------------------------- */
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

/* -------------------------------- THEME ---------------------------------- */
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

/* -------------------------------- TOAST ---------------------------------- */
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

/* ---------------------------- CONFIRM MODAL ------------------------------- */
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

/* --------------------------------- LOGO ---------------------------------- */
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

    if (db) await saveStateToFirebase({ companyLogo: dataUrl });
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

/* --------------------------- FIREBASE INTEGRATION -------------------------- */
let isUpdatingFromFirebase = false;

type FirebaseState = {
  deliveryData?: DeliveryRow[];
  lastUpdate?: any; // Firestore Timestamp
  lastUpdateSheetName?: string;
  companyLogo?: string;
};

const FIREBASE_COLLECTION = "delivery_dashboard";
const FIREBASE_DOC = "live_data";

async function saveStateToFirebase(patch: Partial<FirebaseState> = {}) {
  if (!db || isUpdatingFromFirebase) return;

  try {
    const stateToSave: FirebaseState = {
      deliveryData,
      lastUpdate: new Date(),
      lastUpdateSheetName: lastUpdate?.dataset?.sheetName || "",
      companyLogo: localStorage.getItem("companyLogo") || "",
      ...patch,
    };

    await db.collection(FIREBASE_COLLECTION).doc(FIREBASE_DOC).set(stateToSave, { merge: true });
  } catch (error) {
    console.error("Error saving state to Firebase:", error);
  }
}

function listenForRealtimeUpdates() {
  if (!db) return;
  db.collection(FIREBASE_COLLECTION)
    .doc(FIREBASE_DOC)
    .onSnapshot(
      (docSnap: any) => {
        isUpdatingFromFirebase = true;
        if (docSnap.exists) {
          const data: FirebaseState = docSnap.data() || {};
          deliveryData = Array.isArray(data.deliveryData) ? data.deliveryData : [];
          activeStatusFilter = null;
          if (searchInput) searchInput.value = "";

          if (data.companyLogo && typeof data.companyLogo === "string") {
            localStorage.setItem("companyLogo", data.companyLogo);
            companyLogo.src = data.companyLogo;
            logoContainer.classList.toggle("hidden", !data.companyLogo);
          }

          const lastUpdateDate = data.lastUpdate?.toDate ? data.lastUpdate.toDate() : null;
          const sheetName = data.lastUpdateSheetName || "Sheet";
          if (lastUpdateDate && lastUpdate) {
            lastUpdate.dataset.sheetName = sheetName;
            lastUpdate.textContent = t("lastUpdateText", sheetName, lastUpdateDate.toLocaleString(currentLanguage, { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" }));
          }

          if (deliveryData.length > 0) applyFiltersAndRender();
          else resetUI();
        }
        setTimeout(() => {
          isUpdatingFromFirebase = false;
        }, 250);
      },
      (error: any) => console.error("Firebase listener error:", error)
    );
}

/* ------------------------------ DATA HELPERS ------------------------------- */
function normalizeText(input: any): string {
  const s = String(input ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // remove accents
    .replace(/\s+/g, " "); // collapse spaces
  return s;
}

function isExcelErrorString(v: any): boolean {
  const s = String(v ?? "").trim().toUpperCase();
  return s === "#REF!" || s === "#N/A" || s === "#VALUE!" || s === "#DIV/0!" || s === "#NAME?" || s === "#NULL!";
}

function safeValue(v: any): any {
  if (v == null) return "";
  if (typeof v === "string" && isExcelErrorString(v)) return "";
  return v;
}

function toDateMaybe(v: any): Date | null {
  v = safeValue(v);
  if (!v) return null;

  // XLSX may return Date objects
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  // Excel time can come as Date with 1899/1900 base; still Date object -> ok
  // Serial number
  if (typeof v === "number" && v > 1) {
    // Excel serial dates are days since 1900-01-01.
    // 25569 is the offset between 1900-01-01 and 1970-01-01.
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  // dd/mm/yyyy or dd-mm-yyyy or yyyy-mm-dd
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;

    // if ISO-ish
    const iso = new Date(s);
    if (!isNaN(iso.getTime()) && /\d{4}/.test(s)) return iso;

    const parts = s.split(/[/\-]/).map((p) => p.trim());
    if (parts.length === 3) {
      let a = parseInt(parts[0], 10);
      let b = parseInt(parts[1], 10);
      let c = parseInt(parts[2], 10);
      if ([a, b, c].some((n) => isNaN(n))) return null;

      // detect yyyy-mm-dd
      if (c < 1000 && a > 1000) {
        const dt = new Date(a, b - 1, c);
        return isNaN(dt.getTime()) ? null : dt;
      }

      // dd/mm/yy
      if (c < 100) c += 2000;
      // if month > 12 swap
      if (b > 12 && a <= 12) {
        const tmp = a;
        a = b;
        b = tmp;
      }
      const dt = new Date(c, b - 1, a);
      return isNaN(dt.getTime()) ? null : dt;
    }
  }

  return null;
}

function formatDate(v: any): string {
  const d = toDateMaybe(v);
  if (!d) return String(safeValue(v) || t("dateNotAvailable"));
  return d.toLocaleDateString(currentLanguage, { day: "2-digit", month: "2-digit", year: "2-digit" });
}

function formatTime(v: any): string {
  v = safeValue(v);
  if (!v) return t("dateNotAvailable");
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.toLocaleTimeString(currentLanguage, { hour: "2-digit", minute: "2-digit" });
  }
  // XLSX may provide time as "08:00" or number (fraction of day)
  if (typeof v === "number") {
    const totalMinutes = Math.round(v * 24 * 60);
    const hh = Math.floor(totalMinutes / 60) % 24;
    const mm = totalMinutes % 60;
    return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
  }
  const s = String(v).trim();
  if (/^\d{1,2}:\d{2}/.test(s)) return s;
  return s || t("dateNotAvailable");
}

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
    "TERCA",
    "QUARTA",
    "QUINTA",
    "SEXTA",
    "SABADO",
    "DOMINGO",
  ];
  return (
    workbook.SheetNames.find((name: string) => {
      const upperName = normalizeText(name);
      return keywords.some((key) => upperName.includes(key));
    }) || workbook.SheetNames[0]
  );
}

function getStatusDetails(status: string) {
  const upperStatus = normalizeText(status || "PENDENTE");
  switch (upperStatus) {
    case "ENTREGUE":
      return { icon: "fa-check-circle", pillBg: "bg-green-100 dark:bg-green-900/50", pillText: "text-green-700 dark:text-green-300" };
    case "A CAMINHO":
      return { icon: "fa-truck", pillBg: "bg-yellow-100 dark:bg-yellow-900/50", pillText: "text-yellow-700 dark:text-yellow-300" };
    case "ADIADO":
      return { icon: "fa-calendar-alt", pillBg: "bg-blue-100 dark:bg-blue-900/50", pillText: "text-blue-700 dark:text-blue-300" };
    case "CANCELADO":
      return { icon: "fa-times-circle", pillBg: "bg-red-100 dark:bg-red-900/50", pillText: "text-red-700 dark:text-red-300" };
    case "AGUARDANDO DESOVA":
      return { icon: "fa-box", pillBg: "bg-purple-100 dark:bg-purple-900/50", pillText: "text-purple-700 dark:text-purple-300" };
    default:
      return { icon: "fa-hourglass-half", pillBg: "bg-slate-200 dark:bg-slate-700", pillText: "text-slate-700 dark:text-slate-200" };
  }
}

function getStatusPill(status: string): string {
  const upperStatus = normalizeText(status || "PENDENTE");
  const details = getStatusDetails(upperStatus);
  const labelKey = statusKeyMap[upperStatus] || "STATUS_PENDENTE";
  return `<span class="status-pill ${details.pillBg} ${details.pillText}">
    <i class="fas ${details.icon} fa-fw"></i>
    <span>${t(labelKey)}</span>
  </span>`;
}

function isWeekend(d: Date): boolean {
  const day = d.getDay(); // 0 sun ... 6 sat
  return day === 0 || day === 6;
}

const WEEKDAY_GOAL = 150;

/* ------------------------------- UI CORE ---------------------------------- */
function resetUI() {
  placeholder?.classList.remove("hidden");
  deliveryDashboard?.classList.add("hidden");
  summaryStats?.classList.add("hidden");
  lotSearchContainer?.classList.add("hidden");
  exportExcelBtn?.classList.add("hidden");
  exportPdfBtn?.classList.add("hidden");
  if (deliveryTabs) deliveryTabs.innerHTML = "";
  if (deliveryContent) deliveryContent.innerHTML = "";
  if (lastUpdate) lastUpdate.textContent = t("uploadPrompt");
}

function applyFiltersAndRender(activeTabId: string | null = null) {
  if (!activeTabId) {
    const activeTab = deliveryTabs?.querySelector(".tab-btn.active");
    activeTabId = (activeTab as HTMLElement)?.dataset.target || null;
  }
  const query = (searchInput?.value || "").trim().toLowerCase();
  const lotQuery = (lotSearchInput?.value || "").trim().toLowerCase();
  let filteredData = deliveryData;

  if (showOnlyBattery) {
    filteredData = filteredData.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return materialType.includes("BATTERY") || materialType.includes("BATERIA");
    });
  }

  if (showOnlyKd) {
    filteredData = filteredData.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return materialType.includes("KD");
    });
  }

  if (showOnlyProject) {
    filteredData = filteredData.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return !materialType.includes("BATTERY") && !materialType.includes("BATERIA") && !materialType.includes("KD");
    });
  }

  if (activeStatusFilter) {
    if (activeStatusFilter === "PENDENTE") {
      filteredData = filteredData.filter((row) => {
        const status = normalizeText(row["STATUS"] || "");
        return !["ENTREGUE", "A CAMINHO", "ADIADO", "CANCELADO", "AGUARDANDO DESOVA"].includes(status);
      });
    } else {
      filteredData = filteredData.filter((row) => normalizeText(row["STATUS"] || "") === activeStatusFilter);
    }
  }

  if (query) {
    const searchTerms = query.split(/[\s,\n\t]+/).filter(t => t.length > 0);
    filteredData = filteredData.filter((row) => {
      const rowValues = Object.values(row).map(v => String(v ?? "").toLowerCase());
      return searchTerms.some(term => rowValues.some(val => val.includes(term)));
    });
  }

  if (lotQuery) {
    const lotSearchTerms = lotQuery.split(/[\s,\n\t]+/).filter(t => t.length > 0);
    filteredData = filteredData.filter((row) => {
      const lotValue = String(row["LOT"] || "").toLowerCase();
      return lotSearchTerms.some(term => lotValue.includes(term));
    });
  }

  renderDeliveryDashboard(filteredData, activeTabId);
  updateStats();
}

function updateStats() {
  let dataForStats = deliveryData;
  if (showOnlyBattery) {
    dataForStats = dataForStats.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return materialType.includes("BATTERY") || materialType.includes("BATERIA");
    });
  }
  if (showOnlyKd) {
    dataForStats = dataForStats.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return materialType.includes("KD");
    });
  }
  if (showOnlyProject) {
    dataForStats = dataForStats.filter((row) => {
      const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
      return !materialType.includes("BATTERY") && !materialType.includes("BATERIA") && !materialType.includes("KD");
    });
  }

  const total = dataForStats.length;
  const delivered = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "ENTREGUE").length;
  const inTransit = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "A CAMINHO").length;
  const postponed = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "ADIADO").length;
  const canceled = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "CANCELADO").length;
  const awaitingUnload = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "AGUARDANDO DESOVA").length;
  const pending = Math.max(0, total - delivered - inTransit - postponed - canceled - awaitingUnload);

  const getPercentage = (count: number) => total === 0 ? "0%" : `${((count / total) * 100).toFixed(1)}%`;

  const getCardClasses = (cardStatus: string | null) => {
    const isAll = cardStatus === "ALL";
    const isActive = activeStatusFilter === cardStatus || (activeStatusFilter === null && isAll);
    let classes =
      "summary-card bg-white dark:bg-slate-800 p-5 rounded-lg shadow-sm border flex items-center cursor-pointer transition-all duration-200";
    if (isActive) classes += " border-blue-500 ring-2 ring-blue-500/50 scale-[1.02] z-10";
    else classes += " border-slate-200 dark:border-slate-700 hover:border-blue-300";
    return classes;
  };

  if (!summaryStats) return;
  summaryStats.innerHTML = `
    <div class="${getCardClasses("ALL")}" data-status="ALL">
      <div class="bg-blue-100 dark:bg-blue-900/50 text-blue-600 dark:text-blue-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-box-open text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("totalContainers")}">${t("totalContainers")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${total}</div>
      </div>
    </div>

    <div class="${getCardClasses("ENTREGUE")}" data-status="ENTREGUE">
      <div class="bg-green-100 dark:bg-green-900/50 text-green-600 dark:text-green-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-check-circle text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("delivered")}">${t("delivered")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${delivered}</div>
        <div class="text-[10px] font-bold text-green-600 dark:text-green-400">${getPercentage(delivered)}</div>
      </div>
    </div>

    <div class="${getCardClasses("AGUARDANDO DESOVA")}" data-status="AGUARDANDO DESOVA">
      <div class="bg-purple-100 dark:bg-purple-900/50 text-purple-600 dark:text-purple-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-box text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("awaitingUnload")}">${t("awaitingUnload")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${awaitingUnload}</div>
        <div class="text-[10px] font-bold text-purple-600 dark:text-purple-400">${getPercentage(awaitingUnload)}</div>
      </div>
    </div>

    <div class="${getCardClasses("A CAMINHO")}" data-status="A CAMINHO">
      <div class="bg-yellow-100 dark:bg-yellow-900/50 text-yellow-600 dark:text-yellow-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-truck text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("inTransit")}">${t("inTransit")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${inTransit}</div>
        <div class="text-[10px] font-bold text-yellow-600 dark:text-yellow-400">${getPercentage(inTransit)}</div>
      </div>
    </div>

    <div class="${getCardClasses("PENDENTE")}" data-status="PENDENTE">
      <div class="bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-hourglass-half text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("pending")}">${t("pending")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${pending}</div>
        <div class="text-[10px] font-bold text-slate-600 dark:text-slate-400">${getPercentage(pending)}</div>
      </div>
    </div>

    <div class="${getCardClasses("ADIADO")}" data-status="ADIADO">
      <div class="bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-calendar-alt text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[10px] font-semibold uppercase tracking-wider truncate" title="${t("postponed")}">${t("postponed")}</div>
        <div class="text-xl font-extrabold text-slate-800 dark:text-slate-100">${postponed}</div>
        <div class="text-[10px] font-bold text-indigo-600 dark:text-indigo-400">${getPercentage(postponed)}</div>
      </div>
    </div>

    <div class="${getCardClasses("CANCELADO")}" data-status="CANCELADO">
      <div class="bg-red-100 dark:bg-red-900/50 text-red-600 dark:text-red-400 rounded-full h-10 w-10 flex items-center justify-center mr-3 flex-shrink-0">
        <i class="fas fa-times-circle text-lg"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-xs font-semibold uppercase tracking-wider truncate">${t("canceled")}</div>
        <div class="text-2xl font-extrabold text-slate-800 dark:text-slate-100">${canceled}</div>
        <div class="text-xs font-bold text-red-600 dark:text-red-400">${getPercentage(canceled)}</div>
      </div>
    </div>
  `;
}

function renderDeliveryDashboard(data: DeliveryRow[], activeTabId: string | null = null) {
  placeholder?.classList.add("hidden");
  deliveryDashboard?.classList.remove("hidden");
  summaryStats?.classList.remove("hidden");
  lotSearchContainer?.classList.remove("hidden");
  exportExcelBtn?.classList.remove("hidden");
  exportPdfBtn?.classList.remove("hidden");
  if (deliveryTabs) deliveryTabs.innerHTML = "";
  if (deliveryContent) deliveryContent.innerHTML = "";

  if (!data || data.length === 0) {
    if (deliveryTabs) deliveryTabs.classList.add("hidden");
    if (deliveryContent) {
      deliveryContent.innerHTML = `<div class="text-center py-20 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
        <i class="fas fa-search text-6xl text-slate-300 dark:text-slate-600 mb-4"></i>
        <h2 class="text-2xl font-semibold text-slate-700 dark:text-slate-200">${t("noResultsTitle")}</h2>
        <p class="text-slate-500 dark:text-slate-400 mt-2">${t("noResultsMessage")}</p>
      </div>`;
    }
    return;
  }

  if (deliveryTabs) deliveryTabs.classList.remove("hidden");

  const groupedByDate = data.reduce((acc, row) => {
    const d = toDateMaybe(row["DELIVERY AT BYD"]);
    const key = d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : (String(row["DELIVERY AT BYD"] || "").trim() || t("undefinedDate"));
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {} as Record<string, DeliveryRow[]>);

  const sortedKeys = Object.keys(groupedByDate).sort((a, b) => {
    if (a === t("undefinedDate")) return 1;
    if (b === t("undefinedDate")) return -1;
    const da = new Date(a);
    const dbb = new Date(b);
    return da.getTime() - dbb.getTime();
  });

  sortedKeys.forEach((dateKey, index) => {
    const deliveries = groupedByDate[dateKey];

    let dateObj = new Date(dateKey);
    let hasRealDate = false;
    
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
      const [y, m, d] = dateKey.split("-").map(Number);
      dateObj = new Date(y, m - 1, d);
      hasRealDate = true;
    } else {
      hasRealDate = !isNaN(dateObj.getTime());
    }

    const formattedDate = hasRealDate
      ? dateObj.toLocaleDateString(currentLanguage, { day: "2-digit", month: "2-digit", year: "2-digit" })
      : dateKey;

    const contentId = `content-${index}`;
    const isActive = activeTabId ? contentId === activeTabId : index === 0;

    const tabBtn = document.createElement("button");
    tabBtn.className = `tab-btn flex-shrink-0 px-4 py-3 text-sm font-semibold transition-colors duration-200 flex items-center space-x-2 ${
      isActive ? "active" : ""
    }`;
    tabBtn.innerHTML = `<span class="font-bold">${formattedDate}</span>
      <span class="tab-count-badge bg-slate-200 dark:bg-slate-700 dark:text-slate-200 text-slate-600 font-bold">${deliveries.length}</span>`;
    tabBtn.dataset.target = contentId;
    deliveryTabs?.appendChild(tabBtn);

    const card = document.createElement("div");
    card.id = contentId;
    card.className = `date-card bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 ${
      !isActive ? "hidden" : ""
    }`;

    const deliveredInCard = deliveries.filter((d) => normalizeText(d["STATUS"] || "") === "ENTREGUE").length;
    const totalInCard = deliveries.length;
    const percentage = totalInCard > 0 ? (deliveredInCard / totalInCard) * 100 : 0;

    const cardGoal = hasRealDate ? (isWeekend(dateObj) ? 0 : WEEKDAY_GOAL) : 0;
    const goalMet = cardGoal > 0 ? deliveredInCard >= cardGoal : false;
    const goalLabel = hasRealDate ? (cardGoal > 0 ? t("goalWeekday") : t("goalWeekend")) : t("dateNotAvailable");

    const carrierStats: Record<string, { total: number; delivered: number }> = {};
    deliveries.forEach((d) => {
      const carrier = String(d["TRANSPORTATION COMPANY"] || "N/A").trim() || "N/A";
      if (!carrierStats[carrier]) carrierStats[carrier] = { total: 0, delivered: 0 };
      carrierStats[carrier].total++;
      if (normalizeText(d["STATUS"] || "") === "ENTREGUE") carrierStats[carrier].delivered++;
    });

    const carrierBreakdownHTML = Object.entries(carrierStats)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([carrier, stats]) => {
        const carrierPercent = stats.total > 0 ? (stats.delivered / stats.total) * 100 : 0;
        return `<button type="button" class="carrier-card-btn text-left bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col justify-between transition-all hover:border-blue-500 dark:hover:border-blue-500 hover:shadow-md cursor-pointer w-full" data-carrier="${carrier}">
          <div class="flex justify-between items-start mb-2">
            <span class="font-bold text-sm text-slate-700 dark:text-slate-200 truncate pr-2" title="${carrier}">${carrier}</span>
            <span class="text-[10px] font-bold text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/40 px-1.5 py-0.5 rounded">${carrierPercent.toFixed(0)}%</span>
          </div>
          <div class="flex items-center justify-between mb-2">
            <span class="text-xs text-slate-500 dark:text-slate-400">${t("totalContainers")}: <strong class="text-slate-700 dark:text-slate-200">${stats.total}</strong></span>
            <span class="text-xs text-slate-500 dark:text-slate-400">${t("delivered")}: <strong class="text-green-600 dark:text-green-400">${stats.delivered}</strong></span>
          </div>
          <div class="w-full bg-slate-200 dark:bg-slate-700 h-1.5 rounded-full overflow-hidden">
            <div class="bg-blue-500 h-full transition-all duration-700" style="width: ${carrierPercent}%"></div>
          </div>
        </button>`;
      })
      .join("");

    // Add click handlers for carrier cards
    setTimeout(() => {
      card.querySelectorAll(".carrier-card-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          const detailsContainers = btn.querySelectorAll(".lot-details");
          detailsContainers.forEach((d) => d.classList.toggle("hidden"));
        });
      });
      card.querySelectorAll(".carrier-name-filter").forEach((span) => {
        span.addEventListener("click", (e) => {
          e.stopPropagation();
          const carrier = (span as HTMLElement).dataset.carrier;
          const searchInput = document.getElementById("search-input") as HTMLInputElement;
          if (searchInput) {
            searchInput.value = carrier || "";
            applyFiltersAndRender();
          }
        });
      });
    }, 0);

    const goalBadge =
      hasRealDate && cardGoal > 0
        ? `<span class="ml-3 inline-flex items-center px-2 py-1 rounded text-[11px] font-bold ${
            goalMet ? "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-200" : "bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-200"
          }">
            <i class="fas ${goalMet ? "fa-bullseye" : "fa-flag"} mr-2"></i>
            ${t("goalLabel")}: ${t("kpiGoal", deliveredInCard, cardGoal)} — ${goalMet ? t("reachedGoal") : t("notReachedGoal")}
          </span>`
        : hasRealDate
          ? `<span class="ml-3 inline-flex items-center px-2 py-1 rounded text-[11px] font-bold bg-slate-100 text-slate-700 dark:bg-slate-700 dark:text-slate-200">
              <i class="fas fa-plus-circle mr-2"></i>${t("goalLabel")}: ${goalLabel}
            </span>`
          : "";

    card.innerHTML = `
      <div class="p-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50 rounded-t-lg">
        <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-2 mb-2">
          <div class="flex items-center flex-wrap">
            <h3 class="font-bold text-lg text-slate-800 dark:text-slate-100">${formattedDate}</h3>
            ${goalBadge}
          </div>
          <span class="text-sm font-medium text-slate-500 dark:text-slate-400">${t("containersDelivered", deliveredInCard, totalInCard)}</span>
        </div>
        <div class="progress-bar"><div class="progress-bar-inner" style="width: ${percentage}%"></div></div>
      </div>

      <div class="p-4 bg-slate-50/50 dark:bg-slate-900/30 border-b border-slate-200 dark:border-slate-700">
        <h4 class="text-xs font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest mb-4 flex items-center">
          <i class="fas fa-chart-line mr-2 text-blue-500"></i> ${t("performanceTitle")}
        </h4>
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
          ${Object.entries(
            deliveries.reduce((acc, d) => {
              const carrier = String(d["TRANSPORTATION COMPANY"] || "N/A").trim() || "N/A";
              const lot = String(d["LOT"] || "N/A");
              if (!acc[carrier]) acc[carrier] = {};
              if (!acc[carrier][lot]) acc[carrier][lot] = { total: 0, delivered: 0 };
              acc[carrier][lot].total++;
              if (normalizeText(d["STATUS"] || "") === "ENTREGUE") acc[carrier][lot].delivered++;
              return acc;
            }, {} as Record<string, Record<string, { total: number; delivered: number }>>)
          )
            .map(([carrier, lots]) => {
              const lotHTML = Object.entries(lots)
                .map(([lot, stats]) => {
                  return `
                    <div class="lot-details border-t border-slate-100 dark:border-slate-700 mt-2 pt-2 hidden">
                        <div class="text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Lote ${lot}</div>
                        <div class="flex items-center justify-between text-[10px] text-slate-500 dark:text-slate-400">
                           <span>Agendados: <strong class="text-slate-700 dark:text-slate-200">${stats.total}</strong></span>
                           <span>Entregues: <strong class="text-green-600 dark:text-green-400">${stats.delivered}</strong></span>
                        </div>
                    </div>`;
                })
                .join("");
              
              const totalItems = Object.values(lots).reduce((a, b) => a + b.total, 0);
              const totalDelivered = Object.values(lots).reduce((a, b) => a + b.delivered, 0);
              const carrierPercent = totalItems > 0 ? (totalDelivered / totalItems) * 100 : 0;
                
              return `<button type="button" class="carrier-card-btn text-left bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col justify-between transition-all hover:border-blue-500 dark:hover:border-blue-500 hover:shadow-md cursor-pointer w-full group" data-carrier="${carrier}">
                <div class="flex justify-between items-start mb-2 w-full">
                  <span class="carrier-name-filter font-bold text-sm text-slate-700 dark:text-slate-200 truncate pr-2 hover:text-blue-600 cursor-pointer" title="${carrier}" data-carrier="${carrier}">${carrier}</span>
                  <span class="text-[10px] font-bold text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/40 px-1.5 py-0.5 rounded">${carrierPercent.toFixed(0)}%</span>
                </div>
                <div class="flex flex-col gap-1 w-full">
                   <div class="flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 mt-1">
                      <span>Agendados: <strong class="text-slate-700 dark:text-slate-200">${totalItems}</strong></span>
                      <span>Entregues: <strong class="text-green-600 dark:text-green-400">${totalDelivered}</strong></span>
                   </div>
                  ${lotHTML}
                </div>
                <div class="mt-2 text-[10px] text-blue-600 dark:text-blue-400 font-semibold text-center italic">
                    ${t("clickToExpand")}
                </div>
              </button>`;
            })
            .join("")}
        </div>
      </div>

      <div class="table-responsive">
        <table class="min-w-full text-sm">
          <thead>
            <tr class="border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50">
              <th class="px-4 py-2 text-center text-slate-500 text-xs uppercase w-12">${t("tableHeaderRow")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderContainer")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderModel")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderOperation")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderBL")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderCompany")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderVessel")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderWarehouse")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase">${t("tableHeaderLot")}</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase w-40">${t("tableHeaderStatus")}</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100 dark:divide-slate-700">
            ${deliveries
              .map((row, rowIndex) => {
                const status = normalizeText(row["STATUS"] || "PENDENTE") || "PENDENTE";
                const materialType = normalizeText(row["TYPE OF MATERIAL"] || "");
                const isBattery = materialType.includes("BATTERY") || materialType.includes("BATERIA");
                const isKd = materialType.includes("KD");
                const rowClass = `transition-colors hover:bg-slate-50 dark:hover:bg-slate-700/50 cursor-pointer ${
                  isBattery ? "is-battery" : ""
                } ${isKd ? "is-kd" : ""} ${status === "ENTREGUE" ? "bg-green-100 dark:bg-green-900/30" : status === "CANCELADO" ? "bg-red-100 dark:bg-red-900/30" : ""}`;

                const plate = String(row["TRUCK LICENSE PLATE 1"] || row["PLATE"] || "").trim();

                return `<tr class="${rowClass}" data-row-id="${row._id}">
                  <td class="px-4 py-3 text-xs text-center border-l-8 ${isBattery ? "border-amber-600" : isKd ? "border-blue-700" : "border-transparent"}">${rowIndex + 1}</td>
                  <td class="px-4 py-3 text-xs font-semibold text-slate-800 dark:text-slate-100">
                    ${row["CONTAINER"] || "-"}
                    ${isBattery
                      ? `<span class="ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 dark:bg-amber-900 dark:text-amber-200 uppercase"><i class="fas fa-bolt mr-1"></i>${t(
                          "badgeBattery"
                        )}</span>`
                      : ""}
                    ${isKd
                      ? `<span class="ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200 uppercase">KD</span>`
                      : ""}
                  </td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["MODEL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["OPERATION SCOPE"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-mono">${row["BL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["TRANSPORTATION COMPANY"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["VESSEL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["BONDED WAREHOUSE"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-medium">${row["LOT"] || "-"}</td>
                  <td class="px-4 py-3 text-xs">
                    <select class="status-select bg-white dark:bg-slate-700 dark:text-slate-200 border border-slate-300 dark:border-slate-500 text-xs rounded-md p-1 w-full" data-row-id="${row._id}">
                      ${["PENDENTE", "AGUARDANDO DESOVA", "A CAMINHO", "ADIADO", "ENTREGUE", "CANCELADO"]
                        .map((opt) => `<option value="${opt}" ${status === opt ? "selected" : ""}>${t(statusKeyMap[opt])}</option>`)
                        .join("")}
                    </select>
                    ${plate ? `<div class="mt-1 text-[10px] text-slate-400 dark:text-slate-500"><i class="fas fa-id-card mr-1"></i>${plate}</div>` : ""}
                  </td>
                </tr>`;
              })
              .join("")}
          </tbody>
        </table>
      </div>
    `;

    deliveryContent?.appendChild(card);
  });
}

/* --------------------------- ROW DETAILS EXPAND ---------------------------- */
function kv(label: string, value: any) {
  const v = String(safeValue(value) ?? "").trim();
  return `<div>
    <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${label}</label>
    <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${v || "-"}</p>
  </div>`;
}

function handleRowInteraction(rowEl: HTMLTableRowElement) {
  if (!rowEl || rowEl.classList.contains("details-row")) return;

  const table = rowEl.closest("table")!;
  const expanded = table.querySelector("tr.is-expanded") as HTMLTableRowElement | null;

  if (expanded) {
    expanded.classList.remove("is-expanded");
    const existing = expanded.nextElementSibling as HTMLElement | null;
    if (existing && existing.classList.contains("details-row")) {
      const wrap = existing.querySelector(".details-content-wrapper") as HTMLElement;
      wrap?.classList.remove("expanded");
      setTimeout(() => existing.remove(), 350);
    }
  }

  if (expanded === rowEl) return;

  rowEl.classList.add("is-expanded");

  const rowId = rowEl.dataset.rowId || "";
  const rowData = deliveryData.find((d) => d._id === rowId);
  if (!rowData) return;

  const details = document.createElement("tr");
  details.className = "details-row";

  const plate1 = rowData["TRUCK LICENSE PLATE 1"] || "";
  const plate2 = rowData["TRUCK LICENSE PLATE 2"] || "";
  const plates = [plate1, plate2].filter(Boolean).join(" / ");

  details.innerHTML = `
    <td colspan="8" class="details-cell">
      <div class="details-content-wrapper bg-slate-50 dark:bg-slate-900/50">
        <div class="flex items-center justify-between mb-4">
          <h4 class="text-sm font-extrabold text-slate-700 dark:text-slate-200 flex items-center">
            <i class="fas fa-info-circle mr-2 text-blue-500"></i>${t("detailsTitle")}
          </h4>
          <div class="text-xs text-slate-500 dark:text-slate-400">
            <span class="font-bold">${rowData["CONTAINER"] || rowData["BL"] || "-"}</span>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-4 gap-x-6 gap-y-4">
          ${kv(t("detailsCompany"), rowData["TRANSPORTATION COMPANY"])}
          ${kv(t("detailsVessel"), rowData["VESSEL"])}
          ${kv(t("detailsWarehouse"), rowData["BONDED WAREHOUSE"])}
          ${kv(t("detailsLot"), rowData["LOT"])}

          ${kv("Delivery at BYD", formatDate(rowData["DELIVERY AT BYD"]))}
          ${kv("Unload Time (BYD)", formatTime(rowData["UNLOAD TIME BYD"]))}
          ${kv("Operation Scope", rowData["OPERATION SCOPE"])}
          ${kv("Return Depot Schedule", rowData["RETURN DEPOT SCHEDULE"])}

          ${kv("Driver", rowData["DRIVER NAME"])}
          ${kv("CPF", rowData["CPF"])}
          ${kv("Truck Type", rowData["TRUCK TYPE"])}
          ${kv("Plates", plates || "-")}

          ${kv("Model", rowData["MODEL"])}
          ${kv("ETA Salvador", formatDate(rowData["ETA SALVADOR"]))}
          ${kv("PO SAP", rowData["PO SAP"])}
          ${kv("NF", rowData["NF"])}

          ${kv("Port Arrival", formatDate(rowData["PORT ARRIVAL"]))}
          ${kv("Loading Window", rowData["DATA E HORARIO DE CARREGAMENTO (PREVISÃO / JANELA)"])}
          ${kv("Terminal Departure", rowData["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA."])}
          ${kv("ETA BYD (forecast)", rowData["PREVISÃO DATA E HORARIO DE CHEGADA NA BYD"])}

          ${kv("Unload at BYD", rowData["DATA E HORARIO DE DESCARGA NA BYD "])}
          ${kv("Empty Delivered", rowData["DATA E HORARIO DE ENTREGA CONTAINER VAZIO"])}
          ${kv("Depot", rowData["DEPOT"])}
          ${kv("Ref", rowData["REF"])}

          <div class="md:col-span-2">
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsMaterial")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100">${String(rowData["TYPE OF MATERIAL"] || "-")}</p>
          </div>

          <div class="md:col-span-2">
            <label class="block text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">${t("detailsNotes")}</label>
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100 italic">${String(rowData["NOTES"] || rowData["Conversamos amanhã sobre a referência"] || "-")}</p>
          </div>
        </div>
      </div>
    </td>
  `;

  rowEl.after(details);
  setTimeout(() => (details.querySelector(".details-content-wrapper") as HTMLElement)?.classList.add("expanded"), 10);
}

deliveryContent?.addEventListener("click", (e) => {
  const row = (e.target as HTMLElement).closest<HTMLTableRowElement>("tbody tr:not(.details-row)");
  if (row && !(e.target as HTMLElement).closest(".status-select")) handleRowInteraction(row);
});

/* ---------------------------- STATUS CHANGE ------------------------------- */
function sanitizeStatus(raw: any): string {
  const s = normalizeText(raw || "");
  if (!s) return "PENDENTE";
  if (s === "DELIVERED") return "ENTREGUE";
  if (s === "IN TRANSIT") return "A CAMINHO";
  if (s === "POSTPONED") return "ADIADO";
  if (s === "CANCELED" || s === "CANCELLED") return "CANCELADO";
  if (s === "AWAITING UNLOAD") return "AGUARDANDO DESOVA";
  // if user sheet has garbage like #REF!
  if (isExcelErrorString(raw)) return "PENDENTE";
  // keep only our options if unknown
  if (!["PENDENTE", "A CAMINHO", "ADIADO", "ENTREGUE", "CANCELADO", "AGUARDANDO DESOVA"].includes(s)) return "PENDENTE";
  return s;
}

deliveryContent?.addEventListener("change", async (e) => {
  const select = (e.target as HTMLElement).closest<HTMLSelectElement>(".status-select");
  if (!select) return;

  const rowId = select.dataset.rowId || "";
  const row = deliveryData.find((d) => d._id === rowId);
  if (!row) return;

  const next = sanitizeStatus(select.value);
  const prev = sanitizeStatus(row["STATUS"] || "PENDENTE");
  if (next === prev) return;

  const label = row["CONTAINER"] || row["BL"] || rowId;

  if (await showConfirmationDialog(t("confirmStatusChangeTitle"), t("confirmStatusChangeMessage", String(label), next))) {
    row["STATUS"] = next;
    showToast(t("statusUpdated", String(label), next), "success");
    saveStateToFirebase();
    applyFiltersAndRender();
  } else {
    select.value = prev;
  }
});

/* ------------------------------ TABS -------------------------------------- */
deliveryTabs?.addEventListener("click", (e) => {
  const btn = (e.target as HTMLElement).closest<HTMLButtonElement>(".tab-btn");
  if (btn) {
    deliveryTabs.querySelectorAll(".tab-btn").forEach((b) => {
      b.classList.remove("border-blue-500", "text-blue-600");
      b.classList.add("border-transparent", "text-slate-500");
    });
    btn.classList.add("border-blue-500", "text-blue-600");
    btn.classList.remove("border-transparent", "text-slate-500");

    const target = btn.dataset.tab;
    deliveryContent.classList.toggle("hidden", target !== "deliveries");
    const arrivalsContent = document.getElementById("arrivals-content");
    arrivalsContent?.classList.toggle("hidden", target !== "arrivals");
    
    if (target === "arrivals") {
      renderArrivalsTable();
    }
  }
});

/* ----------------------------- SEARCH & FILTER ---------------------------- */
searchInput?.addEventListener("input", () => {
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = window.setTimeout(applyFiltersAndRender, 250);
});

lotSearchInput?.addEventListener("input", () => {
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = window.setTimeout(applyFiltersAndRender, 250);
});

batteryFilterBtn?.addEventListener("click", () => {
  showOnlyBattery = !showOnlyBattery;
  batteryFilterBtn.classList.toggle("ring-2", showOnlyBattery);
  batteryFilterBtn.classList.toggle("ring-amber-500", showOnlyBattery);
  batteryFilterBtn.classList.toggle("bg-amber-50", showOnlyBattery);
  batteryFilterBtn.classList.toggle("dark:bg-amber-900/30", showOnlyBattery);
  applyFiltersAndRender();
});

kdFilterBtn?.addEventListener("click", () => {
  showOnlyKd = !showOnlyKd;
  kdFilterBtn.classList.toggle("ring-2", showOnlyKd);
  kdFilterBtn.classList.toggle("ring-blue-500", showOnlyKd);
  kdFilterBtn.classList.toggle("bg-blue-50", showOnlyKd);
  kdFilterBtn.classList.toggle("dark:bg-blue-900/30", showOnlyKd);
  applyFiltersAndRender();
});

projectFilterBtn?.addEventListener("click", () => {
  showOnlyProject = !showOnlyProject;
  projectFilterBtn.classList.toggle("ring-2", showOnlyProject);
  projectFilterBtn.classList.toggle("ring-purple-500", showOnlyProject);
  projectFilterBtn.classList.toggle("bg-purple-50", showOnlyProject);
  projectFilterBtn.classList.toggle("dark:bg-purple-900/30", showOnlyProject);
  applyFiltersAndRender();
});

/* -------------------------- STATUS FILTER CARDS ---------------------------- */
summaryStats?.addEventListener("click", (e) => {
  const card = (e.target as HTMLElement).closest<HTMLDivElement>("[data-status]");
  if (card) {
    const s = card.dataset.status!;
    activeStatusFilter = s === "ALL" ? null : activeStatusFilter === s ? null : s;
    applyFiltersAndRender();
  }
});

/* ----------------------- XLSX PARSER (IMPROVED) ---------------------------- */
function renderArrivalsTable() {
  const arrivalsContent = document.getElementById("arrivals-content");
  if (!arrivalsContent) return;

  const lotsFromData = Array.from(new Set(deliveryData.map((d) => String(d["LOT"] || "N/A")))).sort();
  const statuses = ["A CAMINHO", "ADIADO", "AGUARDANDO DESOVA", "ENTREGUE"];

  const tableHtml = `
    <div class="overflow-x-auto bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
      <table class="w-full text-xs text-left text-slate-600 dark:text-slate-300">
        <thead class="bg-slate-50 dark:bg-slate-700 border-b border-slate-200 dark:border-slate-600">
          <tr>
            <th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">Lote</th>
            ${statuses.map((s) => `<th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">${s}</th>`).join("")}
            <th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">Total</th>
          </tr>
        </thead>
        <tbody class="divide-y divide-slate-100 dark:divide-slate-700">
          ${lotsFromData
            .map((lot) => {
              const deliveriesInLot = deliveryData.filter((d) => String(d["LOT"] || "N/A") === lot);
              const statusCounts = statuses.reduce((acc, s) => {
                acc[s] = deliveriesInLot.filter((d) => normalizeText(d["STATUS"] || "") === normalizeText(s)).length;
                return acc;
              }, {} as Record<string, number>);
              const total = Object.values(statusCounts).reduce((a, b) => a + b, 0);

              if (total === 0) return "";

              return `<tr>
                <td class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">${lot}</td>
                ${statuses.map((s) => `<td class="px-4 py-3">${statusCounts[s] > 0 ? statusCounts[s] : ""}</td>`).join("")}
                <td class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">${total}</td>
              </tr>`;
            })
            .join("")}
        </tbody>
        <tfoot class="bg-slate-50 dark:bg-slate-700 border-t border-slate-200 dark:border-slate-600 font-bold">
          <tr>
            <td class="px-4 py-3 text-slate-800 dark:text-slate-100">Total Geral</td>
            ${statuses.map((s) => `<td class="px-4 py-3 text-slate-800 dark:text-slate-100">${deliveryData.filter((d) => normalizeText(d["STATUS"] || "") === normalizeText(s)).length}</td>`).join("")}
            <td class="px-4 py-3 text-slate-800 dark:text-slate-100">${deliveryData.length}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
  arrivalsContent.innerHTML = tableHtml;
}

function buildHeaderIndex(headers: any[]): Record<string, number> {
  const idx: Record<string, number> = {};
  headers.forEach((h, i) => {
    const n = normalizeText(h);
    if (n) idx[n] = i;
  });
  return idx;
}

function pickIndex(hIdx: Record<string, number>, aliases: string[]): number {
  for (const a of aliases) {
    const key = normalizeText(a);
    if (key in hIdx) return hIdx[key];
  }
  return -1;
}

function makeRowId(row: any): string {
  const c = String(row["CONTAINER"] || "").trim();
  const bl = String(row["BL"] || "").trim();
  const date = String(row["DELIVERY AT BYD"] || "").trim();
  const w = String(row["BONDED WAREHOUSE"] || "").trim();
  const tco = String(row["TRANSPORTATION COMPANY"] || "").trim();
  // stable enough
  return normalizeText(`${c}|${bl}|${date}|${w}|${tco}`) || String(Math.random());
}

fileUpload?.addEventListener("change", (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = async (ev) => {
    try {
      const workbook = XLSX.read(new Uint8Array(ev.target!.result as ArrayBuffer), { type: "array" });
      const sheetName = findDeliverySheet(workbook);
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) throw new Error("Sheet not found");

      // Read as array-of-arrays
      const rawData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      // Find header row: first row that has "CONTAINER" column label
      let hRow = rawData.findIndex((r) => r.some((c) => normalizeText(c).includes("CONTAINER") && normalizeText(c) === "CONTAINER"));
      if (hRow === -1) {
        // fallback: any row containing "CONTAINER" as a cell
        hRow = rawData.findIndex((r) => r.some((c) => normalizeText(c) === "CONTAINER"));
      }
      if (hRow === -1) hRow = 0;

      const headers = rawData[hRow] || [];
      const headerIndex = buildHeaderIndex(headers);

      // Aliases matching your sheet columns
      const col = {
        DELIVERY_AT_BYD: pickIndex(headerIndex, ["DELIVERY AT BYD", "DELIVERY", "DATA DE ENTREGA", "ENTREGA NA BYD"]),
        UNLOAD_TIME_BYD: pickIndex(headerIndex, ["UNLOAD TIME BYD", "UNLOAD TIME", "HORARIO DE DESCARGA"]),
        TRANSPORTATION_COMPANY: pickIndex(headerIndex, ["TRANSPORTATION COMPANY", "CARRIER", "TRANSPORTADORA"]),
        CPF: pickIndex(headerIndex, ["CPF"]),
        DRIVER_NAME: pickIndex(headerIndex, ["DRIVER NAME", "MOTORISTA"]),
        PLATE_1: pickIndex(headerIndex, ["TRUCK LICENSE PLATE 1", "PLATE 1", "PLACA 1", "TRUCK PLATE 1"]),
        PLATE_2: pickIndex(headerIndex, ["TRUCK LICENSE PLATE 2", "PLATE 2", "PLACA 2", "TRUCK PLATE 2"]),
        RETURN_DEPOT_SCHEDULE: pickIndex(headerIndex, ["RETURN DEPOT SCHEDULE"]),
        OPERATION_SCOPE: pickIndex(headerIndex, ["OPERATION SCOPE", "SCOPE"]),
        CONTAINER: pickIndex(headerIndex, ["CONTAINER"]),
        BL: pickIndex(headerIndex, ["BL", "B/L"]),
        VESSEL: pickIndex(headerIndex, ["VESSEL", "NAVIO"]),
        BONDED_WAREHOUSE: pickIndex(headerIndex, ["BONDED WAREHOUSE", "WAREHOUSE", "ARMAZEM"]),
        MODEL: pickIndex(headerIndex, ["MODEL", "MODELO"]),
        ETA_SALVADOR: pickIndex(headerIndex, ["ETA SALVADOR", "ETA"]),
        PO_SAP: pickIndex(headerIndex, ["PO SAP", "PO"]),
        NF: pickIndex(headerIndex, ["NF", "NOTA FISCAL"]),
        LOT: pickIndex(headerIndex, ["LOT", "LOTE"]),
        DEMURRAGE: pickIndex(headerIndex, ["DEMURRAGE"]),
        SHIP_OWNER: pickIndex(headerIndex, ["SHIP OWNER", "ARMADOR"]),
        TYPE_OF_MATERIAL: pickIndex(headerIndex, ["TYPE OF MATERIAL", "MATERIAL"]),
        CONTAINER_COST: pickIndex(headerIndex, ["CONTAINER COST", "COST"]),
        STATUS: pickIndex(headerIndex, ["STATUS", "SITUACAO"]),
        TRUCK_TYPE: pickIndex(headerIndex, ["TRUCK TYPE"]),
        LOADING_WINDOW: pickIndex(headerIndex, ["DATA E HORARIO DE CARREGAMENTO (PREVISÃO / JANELA)", "LOADING", "JANELA"]),
        PORT_ARRIVAL: pickIndex(headerIndex, ["PORT ARRIVAL"]),
        TERMINAL_DEPARTURE: pickIndex(headerIndex, ["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA.", "TERMINAL DEPARTURE"]),
        ETA_BYD_FORECAST: pickIndex(headerIndex, ["PREVISÃO DATA E HORARIO DE CHEGADA NA BYD", "ETA BYD"]),
        UNLOAD_AT_BYD: pickIndex(headerIndex, ["DATA E HORARIO DE DESCARGA NA BYD ", "DESCARGA NA BYD"]),
        EMPTY_DELIVERED: pickIndex(headerIndex, ["DATA E HORARIO DE ENTREGA CONTAINER VAZIO", "EMPTY DELIVERY"]),
        DEPOT: pickIndex(headerIndex, ["DEPOT"]),
        REF: pickIndex(headerIndex, ["REF", "REFERENCE", "REFERENCIA"]),
        NOTES: pickIndex(headerIndex, ["NOTES", "OBSERVACOES", "OBSERVAÇÕES", "Conversamos amanhã sobre a referência"]),
      };

      const dataRows = rawData.slice(hRow + 1);

      const parsed: DeliveryRow[] = dataRows
        .filter((r) => {
          const c = col.CONTAINER >= 0 ? safeValue(r[col.CONTAINER]) : "";
          const bl = col.BL >= 0 ? safeValue(r[col.BL]) : "";
          return String(c || "").trim() || String(bl || "").trim();
        })
        .map((r) => {
          const obj: any = {};
          // Keep the original header names as in the sheet, for export compatibility
          headers.forEach((h, i) => {
            const key = String(h || "").trim();
            if (!key) return;
            obj[key] = safeValue(r[i]);
          });

          // Ensure our main canonical keys exist (used by UI)
          obj["DELIVERY AT BYD"] = col.DELIVERY_AT_BYD >= 0 ? safeValue(r[col.DELIVERY_AT_BYD]) : "";
          obj["UNLOAD TIME BYD"] = col.UNLOAD_TIME_BYD >= 0 ? safeValue(r[col.UNLOAD_TIME_BYD]) : "";
          obj["TRANSPORTATION COMPANY"] = col.TRANSPORTATION_COMPANY >= 0 ? safeValue(r[col.TRANSPORTATION_COMPANY]) : "";
          obj["CPF"] = col.CPF >= 0 ? safeValue(r[col.CPF]) : "";
          obj["DRIVER NAME"] = col.DRIVER_NAME >= 0 ? safeValue(r[col.DRIVER_NAME]) : "";
          obj["TRUCK LICENSE PLATE 1"] = col.PLATE_1 >= 0 ? safeValue(r[col.PLATE_1]) : "";
          obj["TRUCK LICENSE PLATE 2"] = col.PLATE_2 >= 0 ? safeValue(r[col.PLATE_2]) : "";
          obj["RETURN DEPOT SCHEDULE"] = col.RETURN_DEPOT_SCHEDULE >= 0 ? safeValue(r[col.RETURN_DEPOT_SCHEDULE]) : "";
          obj["OPERATION SCOPE"] = col.OPERATION_SCOPE >= 0 ? safeValue(r[col.OPERATION_SCOPE]) : "";
          obj["CONTAINER"] = col.CONTAINER >= 0 ? safeValue(r[col.CONTAINER]) : "";
          obj["BL"] = col.BL >= 0 ? safeValue(r[col.BL]) : "";
          obj["VESSEL"] = col.VESSEL >= 0 ? safeValue(r[col.VESSEL]) : "";
          obj["BONDED WAREHOUSE"] = col.BONDED_WAREHOUSE >= 0 ? safeValue(r[col.BONDED_WAREHOUSE]) : "";
          obj["MODEL"] = col.MODEL >= 0 ? safeValue(r[col.MODEL]) : "";
          obj["ETA SALVADOR"] = col.ETA_SALVADOR >= 0 ? safeValue(r[col.ETA_SALVADOR]) : "";
          obj["PO SAP"] = col.PO_SAP >= 0 ? safeValue(r[col.PO_SAP]) : "";
          obj["NF"] = col.NF >= 0 ? safeValue(r[col.NF]) : "";
          obj["LOT"] = col.LOT >= 0 ? safeValue(r[col.LOT]) : "";
          obj["DEMURRAGE"] = col.DEMURRAGE >= 0 ? safeValue(r[col.DEMURRAGE]) : "";
          obj["SHIP OWNER"] = col.SHIP_OWNER >= 0 ? safeValue(r[col.SHIP_OWNER]) : "";
          obj["TYPE OF MATERIAL"] = col.TYPE_OF_MATERIAL >= 0 ? safeValue(r[col.TYPE_OF_MATERIAL]) : "";
          obj["CONTAINER COST"] = col.CONTAINER_COST >= 0 ? safeValue(r[col.CONTAINER_COST]) : "";
          obj["TRUCK TYPE"] = col.TRUCK_TYPE >= 0 ? safeValue(r[col.TRUCK_TYPE]) : "";
          obj["PORT ARRIVAL"] = col.PORT_ARRIVAL >= 0 ? safeValue(r[col.PORT_ARRIVAL]) : "";
          obj["DATA E HORARIO DE CARREGAMENTO (PREVISÃO / JANELA)"] = col.LOADING_WINDOW >= 0 ? safeValue(r[col.LOADING_WINDOW]) : "";
          obj["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA."] = col.TERMINAL_DEPARTURE >= 0 ? safeValue(r[col.TERMINAL_DEPARTURE]) : "";
          obj["PREVISÃO DATA E HORARIO DE CHEGADA NA BYD"] = col.ETA_BYD_FORECAST >= 0 ? safeValue(r[col.ETA_BYD_FORECAST]) : "";
          obj["DATA E HORARIO DE DESCARGA NA BYD "] = col.UNLOAD_AT_BYD >= 0 ? safeValue(r[col.UNLOAD_AT_BYD]) : "";
          obj["DATA E HORARIO DE ENTREGA CONTAINER VAZIO"] = col.EMPTY_DELIVERED >= 0 ? safeValue(r[col.EMPTY_DELIVERED]) : "";
          obj["DEPOT"] = col.DEPOT >= 0 ? safeValue(r[col.DEPOT]) : "";
          obj["REF"] = col.REF >= 0 ? safeValue(r[col.REF]) : "";

          const notesVal =
            col.NOTES >= 0 ? safeValue(r[col.NOTES]) : safeValue(obj["NOTES"] || obj["Conversamos amanhã sobre a referência"] || "");
          obj["NOTES"] = notesVal;

          // STATUS: if sheet formula is broken (#REF!) default to PENDENTE
          const statusRaw = col.STATUS >= 0 ? safeValue(r[col.STATUS]) : "";
          obj["STATUS"] = sanitizeStatus(statusRaw);

          // Stable id
          obj._id = makeRowId(obj);

          return obj as DeliveryRow;
        });

      deliveryData = parsed;

      if (lastUpdate) {
        lastUpdate.dataset.sheetName = sheetName;
        lastUpdate.textContent = t("lastUpdateText", sheetName, new Date().toLocaleString(currentLanguage, { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" }));
      }

      showToast(t("sheetLoaded"), "success");
      applyFiltersAndRender();
      saveStateToFirebase();
    } catch (err) {
      console.error(err);
      showToast(t("fileProcessError"), "error");
    }
  };

  reader.onerror = () => showToast(t("fileReadError"), "error");
  reader.readAsArrayBuffer(file);
});

/* ------------------------------- EXPORTS ---------------------------------- */
exportExcelBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) return showToast(t("noDataToExport"), "warning");

  // Keep everything except internal _id
  const out = deliveryData.map((d) => {
    const { _id, ...rest } = d;
    return rest;
  });

  const ws = XLSX.utils.json_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Deliveries");
  XLSX.writeFile(wb, "deliveries.xlsx");
  showToast(t("excelGenerated"), "success");
});

exportPdfBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) return showToast(t("noDataToExport"), "warning");
  try {
    const doc = new (jspdf as any).jsPDF({ orientation: "landscape" });
    doc.text(t("pdfTitle"), 40, 40);
    (doc as any).autoTable({
      head: [["#", "DELIVERY", "CONTAINER", "BL", "CARRIER", "VESSEL", "WAREHOUSE", "STATUS"]],
      body: deliveryData.map((d, i) => [
        i + 1,
        formatDate(d["DELIVERY AT BYD"]),
        d["CONTAINER"] || "",
        d["BL"] || "",
        d["TRANSPORTATION COMPANY"] || "",
        d["VESSEL"] || "",
        d["BONDED WAREHOUSE"] || "",
        sanitizeStatus(d["STATUS"]),
      ]),
      startY: 60,
      styles: { fontSize: 8 },
    });
    doc.save("deliveries.pdf");
    showToast(t("pdfGenerated"), "success");
  } catch (e) {
    console.error(e);
    showToast(t("fileProcessError"), "error");
  }
});

/* ------------------------------ STARTUP ----------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  setTheme(((localStorage.getItem("theme") as any) || "light") as any);
  loadLogoFromStorage();
  setLanguage((localStorage.getItem("language") as Language) || "pt-BR");
  listenForRealtimeUpdates();
  resetUI();
});
