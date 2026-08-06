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
 * - Adds Daily Goal (300 on weekdays; weekends show goal as “bonus”) per date card
 * - Fixed: Integrated Inventory and Operational Blog storage with online persistence (Firebase)
 * - Standardized BYD Corporate Minutes formatting layout for operational logs and reports
 */

declare const firebase: any;
declare const XLSX: any;
declare const jspdf: any;
declare const Chart: any;
declare const ChartDataLabels: any;

/* ----------------------------- FIREBASE SAFE ------------------------------ */
const getEnv = (key: string): string => {
  const env = (import.meta as any).env || (process as any).env || {};
  return env[key] || "";
};

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
const viewModeTabs = document.getElementById("view-mode-tabs") as HTMLDivElement;
const deliveriesWrapper = document.getElementById("deliveries-wrapper") as HTMLDivElement;
const deliveryTabs = document.getElementById("delivery-tabs") as HTMLDivElement;
const deliveryContent = document.getElementById("delivery-content") as HTMLDivElement;
const exportExcelBtn = document.getElementById("export-excel-btn") as HTMLButtonElement;
const exportPdfBtn = document.getElementById("export-pdf-btn") as HTMLButtonElement;
const themeToggleBtn = document.getElementById("theme-toggle") as HTMLButtonElement;
const batteryFilterBtn = document.getElementById("battery-filter-btn") as HTMLButtonElement;
const kdFilterBtn = document.getElementById("kd-filter-btn") as HTMLButtonElement;
const spFilterBtn = document.getElementById("sp-filter-btn") as HTMLButtonElement;
const pbpFilterBtn = document.getElementById("pbp-filter-btn") as HTMLButtonElement;
const projectFilterBtn = document.getElementById("project-filter-btn") as HTMLButtonElement;
const lotSearchInput = document.getElementById("lot-search-input") as HTMLInputElement;
const lotSearchContainer = document.getElementById("lot-search-container") as HTMLDivElement;
const monthFilterSelect = document.getElementById("month-filter-select") as HTMLSelectElement;
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
    pageTitle: "KD Monitor Dashboard",
    headerTitle: "KD Monitor Dashboard",
    uploadPrompt: "Carregue sua planilha de agendamento para começar",
    searchInputPlaceholder: "Pesquisar container, BL, navio, PO...",
    searchLotPlaceholder: "Pesquisar LOTE",
    uploadLogoTooltip: "Carregar logo da empresa",
    toggleThemeTooltip: "Alternar tema",
    uploadSheetButton: "Carregar",
    uploadSheetTooltip: "Carregar Planilha",
    filterBatteryTooltip: "Filtrar Baterias",
    filterKdTooltip: "Filtrar KD",
    filterSpTooltip: "Filtrar Peças de Reposição (SP)",
    filterPbpTooltip: "Filtrar Part by Part (PBP)",
    filterProjectTooltip: "Filtrar Project Cargo",
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
    statusBacklog: "Backlog",
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
    tableHeaderTotal: "Total",
    tableHeaderOverallTotal: "Total Geral",
    STATUS_PENDENTE: "Pendente",
    STATUS_A_CAMINHO: "A Caminho",
    STATUS_ADIADO: "Adiado",
    STATUS_ENTREGUE: "Entregue",
    STATUS_CANCELADO: "Cancelado",
    STATUS_AGUARDANDO_DESOVA: "Aguardando Desova",
    STATUS_BACKLOG: "Backlog",
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
    chartsTab: "Gráficos (Operação)",
    timeTab: "Tempo de Operação",
    modelsTitle: "Modelos",
    legendTitle: "Legenda",
    efic: "EFIC.",
    prog: "PROG.",
    pend: "PEND.",
    tableHeaderStart: "Início",
    tableHeaderEnd: "Fim",
    tableHeaderFullTime: "Tempo Total",
    tableHeaderTimeAvg: "Tempo Médio",
    avgPeriod1: "1º Período (06:30 - 15:00)",
    avgPeriod2: "2º Período (15:01 - 00:00)",
    chartsOverviewTitle: "Visão Geral da Operação",
    chartsLotProgressTitle: "Progresso por Lote",
    chartsCarrierTitle: "Desempenho por Transportadora",
    chartsWarehouseTitle: "Status por Armazém Afiançado",
    chartsJustificationTitle: "Justificativas por Lote",
    chartsJustificationPlaceholder: "Justificativa...",
    chartsOther: "Outros (Adiado/Cancelado)",
    pdfTitle: "Programação de Entregas de Contêineres",
    pdfGeneratedOn: (date: string) => `Relatório gerado em: ${date}`,
    pdfPage: (page: number, total: number) => `Página ${page} de ${total}`,
    lastUpdateText: (sheet: string, date: string) => `Dados de "${sheet}" | Carregado em: ${date}`,
    clickToExpand: "Clique para expandir",
    changeStatusFor: (containerId: string) => `Alterar status do container ${containerId || ""}`,
    viewDetailsFor: (containerId: string) => `Ver detalhes do container ${containerId || "sem identificação"}`,
    goalLabel: "Meta",
    goalWeekday: "300/dia útil",
    goalWeekend: "Fim de semana (bônus)",
    reachedGoal: "Meta atingida",
    notReachedGoal: "Abaixo da meta",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "Hoje",
    saveDayButton: "Fechar e Salvar Dia",
    historyTab: "Histórico Mensal",
    paretoTab: "Análise de Fenômeno - Paretos",
    saveDayConfirmTitle: "Arquivar o dia atual?",
    saveDayConfirmMsg: "Isso salvará a programação atual no histórico mensal e limpará o painel para a nova planilha do dia seguinte. Continuar?"
  },
  "en-US": {
    pageTitle: "KD Monitor Dashboard",
    headerTitle: "KD Monitor Dashboard",
    uploadPrompt: "Upload your schedule spreadsheet to begin",
    searchInputPlaceholder: "Search container, BL, vessel, PO...",
    searchLotPlaceholder: "Search LOT",
    uploadLogoTooltip: "Upload company logo",
    toggleThemeTooltip: "Toggle theme",
    uploadSheetButton: "Upload",
    uploadSheetTooltip: "Upload Spreadsheet",
    filterBatteryTooltip: "Filter Batteries",
    filterKdTooltip: "Filter KD",
    filterSpTooltip: "Filter Spare Parts (SP)",
    filterPbpTooltip: "Filter Part by Part (PBP)",
    filterProjectTooltip: "Filter Project Cargo",
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
    statusBacklog: "Backlog",
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
    tableHeaderTotal: "Total",
    tableHeaderOverallTotal: "Overall Total",
    STATUS_PENDENTE: "Pending",
    STATUS_A_CAMINHO: "In Transit",
    STATUS_ADIADO: "Postponed",
    STATUS_ENTREGUE: "Delivered",
    STATUS_CANCELADO: "Canceled",
    STATUS_AGUARDANDO_DESOVA: "Awaiting Unload",
    STATUS_BACKLOG: "Backlog",
    detailsTitle: "Details",
    detailsVessel: "Vessel",
    detailsWarehouse: "Warehouse",
    detailsNotes: "Notes",
    detailsMaterial: "Material Type",
    detailsLot: "LOT Number",
    detailsCompany: "Carrier",
    performanceTitle: "Carrier Performance",
    badgeBattery: "Battery",
    deliveriesTab: "Deliveries",
    chartsTab: "Charts (Operation)",
    timeTab: "Operation Time",
    modelsTitle: "Models",
    legendTitle: "Legend",
    efic: "EFFIC.",
    prog: "PROG.",
    pend: "PEND.",
    tableHeaderStart: "Start",
    tableHeaderEnd: "End",
    tableHeaderFullTime: "Total Time",
    tableHeaderTimeAvg: "Average Time",
    avgPeriod1: "1st Period (06:30 - 15:00)",
    avgPeriod2: "2nd Period (15:01 - 00:00)",
    chartsOverviewTitle: "Operation Overview",
    chartsLotProgressTitle: "Progress by Lot",
    chartsCarrierTitle: "Carrier Performance",
    chartsWarehouseTitle: "Bonded Warehouse Status",
    chartsJustificationTitle: "Lot Justifications",
    chartsJustificationPlaceholder: "Justification...",
    chartsOther: "Other (Postponed/Canceled)",
    pdfTitle: "Container Delivery Schedule",
    pdfGeneratedOn: (date: string) => `Report generated on: ${date}`,
    pdfPage: (page: number, total: number) => `Page ${page} of ${total}`,
    lastUpdateText: (sheet: string, date: string) => `Data from "${sheet}" | Loaded on: ${date}`,
    changeStatusFor: (containerId: string) => `Change status for container ${containerId || ""}`,
    viewDetailsFor: (containerId: string) => `View details for container ${containerId || "unidentified"}`,
    goalLabel: "Goal",
    goalWeekday: "300/weekday",
    goalWeekend: "Weekend (bonus)",
    reachedGoal: "Goal reached",
    notReachedGoal: "Below goal",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "Today",
    saveDayButton: "Save End of Day",
    historyTab: "Monthly History",
    paretoTab: "Phenomenon Analysis - Pareto",
    saveDayConfirmTitle: "Archive Current Day?",
    saveDayConfirmMsg: "This will save the current schedule to the monthly history and clear the dashboard for the new day's upload. Continue?"
  },
  "zh-CN": {
    pageTitle: "KD 监控仪表板",
    headerTitle: "KD 监控仪表板",
    uploadPrompt: "上传您的排程电子表格以开始",
    searchInputPlaceholder: "搜索集装箱、提单 (BL)、船名、PO...",
    searchLotPlaceholder: "搜索批号 (LOT)",
    uploadLogoTooltip: "上传公司标志",
    toggleThemeTooltip: "切换主题",
    uploadSheetButton: "上传",
    uploadSheetTooltip: "上传电子表格",
    filterBatteryTooltip: "过滤电池",
    filterKdTooltip: "过滤 KD",
    filterSpTooltip: "过滤备件 (SP)",
    filterPbpTooltip: "过滤按件 (PBP)",
    filterProjectTooltip: "过滤项目货物 (Project Cargo)",
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
    statusBacklog: "积压 (Backlog)",
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
    tableHeaderTotal: "总计",
    tableHeaderOverallTotal: "总计",
    STATUS_PENDENTE: "待处理",
    STATUS_A_CAMINHO: "运输中",
    STATUS_ADIADO: "已推迟",
    STATUS_ENTREGUE: "已交付",
    STATUS_CANCELADO: "已取消",
    STATUS_AGUARDANDO_DESOVA: "等待卸货",
    STATUS_BACKLOG: "积压 (Backlog)",
    detailsTitle: "详细信息",
    detailsVessel: "船名",
    detailsWarehouse: "仓库",
    detailsNotes: "备注",
    detailsMaterial: "物料类型",
    detailsLot: "批号",
    detailsCompany: "运输公司",
    performanceTitle: "承运人绩效",
    badgeBattery: "电池",
    deliveriesTab: "交货",
    chartsTab: "图表（运营）",
    timeTab: "运营时间",
    modelsTitle: "型号",
    legendTitle: "图例",
    efic: "效率",
    prog: "进度",
    pend: "待处理",
    tableHeaderStart: "起点",
    tableHeaderEnd: "终点",
    tableHeaderFullTime: "总时间",
    tableHeaderTimeAvg: "平均时间",
    avgPeriod1: "第一段 (06:30 - 15:00)",
    avgPeriod2: "第二段 (15:01 - 00:00)",
    chartsOverviewTitle: "运营概览",
    chartsLotProgressTitle: "按批次进度",
    chartsCarrierTitle: "承运人绩效",
    chartsWarehouseTitle: "保税仓库状态",
    chartsJustificationTitle: "批次说明",
    chartsJustificationPlaceholder: "说明...",
    chartsOther: "其他 (推迟/取消)",
    pdfTitle: "集装箱交付计划",
    pdfGeneratedOn: (date: string) => `报告生成于：${date}`,
    pdfPage: (page: number, total: number) => `第 ${page} 页，共 ${total} 页`,
    lastUpdateText: (sheet: string, date: string) => `数据来源 "${sheet}" | 加载于：${date}`,
    changeStatusFor: (containerId: string) => `更改集装箱 ${containerId || ""} 的状态`,
    viewDetailsFor: (containerId: string) => `查看集装箱 ${containerId || "未识别"} 的详细信息`,
    goalLabel: "目标",
    goalWeekday: "工作日300",
    goalWeekend: "周末（加分）",
    reachedGoal: "已达目标",
    notReachedGoal: "未达目标",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "今天",
    saveDayButton: "保存当天 (Save Day)",
    historyTab: "历史记录",
    paretoTab: "现象分析 - 帕累托",
    saveDayConfirmTitle: "归档当天数据？",
    saveDayConfirmMsg: "这会将当前计划保存到月度历史记录，并清空仪表板以便上传新一天的计划。是否继续？"
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
  BACKLOG: "STATUS_BACKLOG",
  "AGUARDANDO DESOVA": "STATUS_AGUARDANDO_DESOVA",
};

/* ------------------------------ APP STATE --------------------------------- */
type DeliveryRow = Record<string, any> & {
  _id: string;
};

let deliveryData: DeliveryRow[] = [];
let historicalData: DeliveryRow[] = [];
let dailyCarrierNotes: Record<string, Record<string, { motivo: string, impacto: string }>> = {};
let searchDebounceTimer: number;
let activeStatusFilter: string | null = null;
let showOnlyBattery: boolean = false;
let showOnlyKd: boolean = false;
let showOnlySp: boolean = false;
let showOnlyPbp: boolean = false;
let showOnlyProject: boolean = false;
let isMacroView: boolean = false;
let chartGroupBy: "lot" | "po" = "lot";
let overallChart: any = null;
let lotChart: any = null;
let maxLotChart: any = null;
let modelChart: any = null;
let historyDailyChart: any = null;
let historyWeeklyChart: any = null;
let operacionalDiariaChart: any = null;
let operacionalTransportadorasChart: any = null;
let selectedHistoryDate: string | null = null;
let selectedHistoryWeek: string | null = null;
let carrierCharts: any[] = [];
let warehouseCharts: any[] = [];

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
  historicalData?: DeliveryRow[];
  lastUpdate?: any; // Firestore Timestamp
  lastUpdateSheetName?: string;
  companyLogo?: string;
  dailyCarrierNotes?: Record<string, Record<string, { motivo: string, impacto: string }>>;
  paretoReasons?: string[];
  historyLastUpdate?: number;
};

const FIREBASE_COLLECTION = "delivery_dashboard";
const FIREBASE_DOC = "live_data";

let loadedHistoryLastUpdate: any = null;
let isFirstHistoryLoad = true;

async function saveStateToFirebase(patch: Partial<FirebaseState> = {}) {
  if (!db || isUpdatingFromFirebase) return;

  try {
    // 1. Separate logo data to save in its own document (prevents bloating active state)
    const logoStr = patch.companyLogo || localStorage.getItem("companyLogo") || "";
    if (logoStr) {
      await db.collection(FIREBASE_COLLECTION).doc("logo_data").set({ companyLogo: logoStr });
    }

    // 2. Separate historical data into small chunks of 250 rows each
    const finalHistoricalData = patch.hasOwnProperty("historicalData") ? patch.historicalData || [] : historicalData;
    const chunkCount = Math.ceil(finalHistoricalData.length / 250);
    
    const chunkPromises = [];
    for (let i = 0; i < chunkCount; i++) {
      const chunkRows = finalHistoricalData.slice(i * 250, (i + 1) * 250);
      chunkPromises.push(
        db.collection(FIREBASE_COLLECTION).doc(`history_chunk_${i}`).set({ rows: chunkRows })
      );
    }
    
    // Clean up any extra/dangling chunks that might have existed previously
    const deletePromises = [];
    for (let i = chunkCount; i < chunkCount + 10; i++) {
      deletePromises.push(
        db.collection(FIREBASE_COLLECTION).doc(`history_chunk_${i}`).delete().catch(() => {})
      );
    }

    await Promise.all(chunkPromises);
    await Promise.all(deletePromises);
    await db.collection(FIREBASE_COLLECTION).doc("history_metadata").set({ chunkCount });

    // 3. Save active live_data (clean of logo and raw historicalData)
    const liveDataToSave: any = {
      deliveryData,
      dailyCarrierNotes,
      paretoReasons: (window as any).__PARETO_REASONS__ || [
        "PRAZO CURTO PARA COLETA",
        "QUEBRA DE VEÍCULO",
        "INCIDENTE TERMINAL",
        "GREVE DOS CAMINHONEIROS",
        "GREVE SINDICAL",
        "ALTERAÇÃO DE PROGRAMAÇÃO",
        "ACIDENTE NA RODOVIA",
        "FILA NO TERMINAL",
        "PENDÊNCIA DOCUMENTAL"
      ],
      lastUpdate: new Date(),
      lastUpdateSheetName: lastUpdate?.dataset?.sheetName || "",
      historyLastUpdate: new Date().getTime(), // trigger other clients to reload chunked history
      ...patch,
    };

    // Strip huge keys to stay safe under 1MB limit
    delete liveDataToSave.historicalData;
    delete liveDataToSave.companyLogo;

    await db.collection(FIREBASE_COLLECTION).doc(FIREBASE_DOC).set(liveDataToSave, { merge: true });
  } catch (error) {
    console.error("Error saving state to Firebase:", error);
  }
}

function listenForRealtimeUpdates() {
  if (!db) return;

  // 1. Fetch custom logo once from logo_data on startup
  db.collection(FIREBASE_COLLECTION)
    .doc("logo_data")
    .get()
    .then((docSnap: any) => {
      if (docSnap.exists) {
        const logoData = docSnap.data();
        if (logoData && logoData.companyLogo) {
          localStorage.setItem("companyLogo", logoData.companyLogo);
          companyLogo.src = logoData.companyLogo;
          logoContainer.classList.toggle("hidden", !logoData.companyLogo);
        }
      }
    })
    .catch((err: any) => console.error("Error loading logo from Firebase:", err));

  // 2. Listen for active live data in real-time
  db.collection(FIREBASE_COLLECTION)
    .doc(FIREBASE_DOC)
    .onSnapshot(
      async (docSnap: any) => {
        isUpdatingFromFirebase = true;
        if (docSnap.exists) {
          const data: any = docSnap.data() || {};
          deliveryData = Array.isArray(data.deliveryData) ? data.deliveryData : [];
          dailyCarrierNotes = data.dailyCarrierNotes || {};
          
          if (Array.isArray(data.paretoReasons)) {
            (window as any).__PARETO_REASONS__ = data.paretoReasons;
          }

          const lastUpdateDate = data.lastUpdate?.toDate ? data.lastUpdate.toDate() : null;
          const sheetName = data.lastUpdateSheetName || "Sheet";
          if (lastUpdateDate && lastUpdate) {
            lastUpdate.dataset.sheetName = sheetName;
            lastUpdate.textContent = t("lastUpdateText", sheetName, lastUpdateDate.toLocaleString(currentLanguage, { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" }));
          }

          // 3. Dynamically fetch history chunks when the remote history state changes
          const remoteHistoryLastUpdate = data.historyLastUpdate || null;
          if (isFirstHistoryLoad || remoteHistoryLastUpdate !== loadedHistoryLastUpdate) {
            loadedHistoryLastUpdate = remoteHistoryLastUpdate;
            isFirstHistoryLoad = false;

            try {
              const metaSnap = await db.collection(FIREBASE_COLLECTION).doc("history_metadata").get();
              if (metaSnap.exists) {
                const meta = metaSnap.data();
                const chunkCount = meta.chunkCount || 0;

                const chunkPromises = [];
                for (let i = 0; i < chunkCount; i++) {
                  chunkPromises.push(db.collection(FIREBASE_COLLECTION).doc(`history_chunk_${i}`).get());
                }
                const chunkSnaps = await Promise.all(chunkPromises);

                let loadedRows: any[] = [];
                chunkSnaps.forEach((snap: any) => {
                  if (snap.exists) {
                    const chunkData = snap.data();
                    if (Array.isArray(chunkData.rows)) {
                      loadedRows = loadedRows.concat(chunkData.rows);
                    }
                  }
                });

                historicalData = loadedRows;
              } else if (Array.isArray(data.historicalData)) {
                // Fallback for backward compatibility
                historicalData = data.historicalData;
              }
            } catch (err) {
              console.error("Error loading historical chunks:", err);
              if (Array.isArray(data.historicalData)) {
                historicalData = data.historicalData;
              }
            }
          }

          if (deliveryData.length > 0 || historicalData.length > 0) applyFiltersAndRender();
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

function toDateTimeMaybe(v: any): Date | null {
  v = safeValue(v);
  if (v === null || v === undefined || v === "" || v === 0 || v === "0" || v === "-") return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 1) {
    // Excel base date is Dec 30, 1899 for SheetJS numbers
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      // Use UTC parts to avoid timezone shift on naive Excel dates
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());
    }
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    
    // Try native parsing first (works for ISO and some locales)
    const iso = new Date(s);
    if (!isNaN(iso.getTime()) && s.length > 8 && /\d{4}/.test(s)) return iso;

    // Custom parsing for common Brazilian/International formats: DD/MM/YYYY HH:MM:SS
    // Handles separators: / - . , and spaces
    const parts = s.split(/[\s,]+/);
    if (parts.length > 0) {
      const datePart = parts[0];
      const timePart = parts[1] || "";
      const dateParts = datePart.split(/[\/\-\.]/);
      
      if (dateParts.length === 3) {
        let day = parseInt(dateParts[0], 10);
        let month = parseInt(dateParts[1], 10);
        let year = parseInt(dateParts[2], 10);
        
        // Handle YYYY/MM/DD
        if (day > 1000) {
          const tmp = day; day = year; year = tmp;
        }
        if (year < 100) year += 2000;
        
        // Basic sanity check on month/day ordering (01..12)
        if (month > 12 && day <= 12) {
          const tmp = day; day = month; month = tmp;
        }

        let h = 0, m = 0, sec = 0;
        if (timePart) {
          const timeParts = timePart.split(':');
          h = parseInt(timeParts[0] || "0", 10);
          m = parseInt(timeParts[1] || "0", 10);
          sec = parseInt(timeParts[2] || "0", 10);
        }
        
        const dt = new Date(year, month - 1, day, h, m, sec);
        return isNaN(dt.getTime()) ? null : dt;
      }
    }
  }
  return null;
}

function toDateMaybe(v: any): Date | null {
  v = safeValue(v);
  if (!v) return null;

  // XLSX may return Date objects
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  // Serial number
  if (typeof v === "number" && v > 1) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  // dd/mm/yyyy or dd-mm-yyyy or yyyy-mm-dd
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;

    const iso = new Date(s);
    if (!isNaN(iso.getTime()) && /\d{4}/.test(s)) return iso;

    const parts = s.split(/[/\-]/).map((p) => p.trim());
    if (parts.length === 3) {
      let a = parseInt(parts[0], 10);
      let b = parseInt(parts[1], 10);
      let c = parseInt(parts[2], 10);
      if ([a, b, c].some((n) => isNaN(n))) return null;

      if (c < 1000 && a > 1000) {
        const dt = new Date(a, b - 1, c);
        return isNaN(dt.getTime()) ? null : dt;
      }

      if (c < 100) c += 2000;
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

function isRowCompletedInOperationalTime(row: any): boolean {
  const op = String(row["OPERATION SCOPE"] || "").trim().toUpperCase();
  const status = normalizeText(row["STATUS"] || "");
  const isBaixa = op.includes("SWAP") || op.includes("PUT DOWN") || op.includes("PUTDOWN") || op.includes("BAIXA") || op.includes("PISO");
  
  const startDt = toDateTimeMaybe(row["TERMINAL - INÍCIO DE ROTA"]);
  const endDt = toDateTimeMaybe(row["ENTREGA VAZIO"]) || toDateTimeMaybe(row["DATA E HORARIO DE DESCARGA"]);

  if (isBaixa) {
    return !!(startDt && status === "ENTREGUE");
  } else {
    return !!(startDt && endDt);
  }
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
    "DELIVERY", "SCHEDULE", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY",
    "SEGUNDA", "TERCA", "QUARTA", "QUINTA", "SEXTA", "SABADO", "DOMINGO",
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
    case "BACKLOG":
      return { icon: "fa-history", pillBg: "bg-orange-100 dark:bg-orange-900/50", pillText: "text-orange-700 dark:text-orange-300" };
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
  const day = d.getDay();
  return day === 0 || day === 6;
}

const WEEKDAY_GOAL = 300;

/* ------------------------------- UI CORE ---------------------------------- */
function resetUI() {
  if (deliveryData.length === 0 && historicalData.length === 0) {
    placeholder?.classList.remove("hidden");
    deliveryDashboard?.classList.add("hidden");
    summaryStats?.classList.add("hidden");
    lotSearchContainer?.classList.add("hidden");
    exportExcelBtn?.classList.add("hidden");
    exportPdfBtn?.classList.add("hidden");
    if (deliveryTabs) deliveryTabs.innerHTML = "";
    if (deliveryContent) deliveryContent.innerHTML = "";
    if (lastUpdate) lastUpdate.textContent = t("uploadPrompt");
  } else if (deliveryData.length === 0 && historicalData.length > 0) {
    // We have history but no active delivery data.
    placeholder?.classList.add("hidden");
    deliveryDashboard?.classList.remove("hidden");
    summaryStats?.classList.remove("hidden");
    
    // Switch to history tab visually if not already
    const histBtn = document.querySelector(".view-tab-btn[data-tab='history']") as HTMLElement;
    if (histBtn) histBtn.click();
    
    applyFiltersAndRender();
  }
}

function isBatteryRow(row: any): boolean {
  const mt = normalizeText(row["TYPE OF MATERIAL"] || "");
  const mod = normalizeText(row["MODEL"] || "");
  const rat = normalizeText(row["RATIONALIZATION"] || "");
  const searchStr = `${mt} ${mod} ${rat}`;
  return searchStr.includes("BATTERY") || searchStr.includes("BATERIA");
}

function isKdRow(row: any): boolean {
  if (isBatteryRow(row)) return false;
  const mt = normalizeText(row["TYPE OF MATERIAL"] || "");
  if (mt.includes("KD") || mt.includes("CKD") || mt.includes("SKD")) return true;
  
  const mod = normalizeText(row["MODEL"] || "");
  const rat = normalizeText(row["RATIONALIZATION"] || "");
  const isAssembly = mod.includes("BIG ASSEMBLY") || rat.includes("BIG ASSEMBLY");
  const hasKdKeyword = mod.includes("KD") || mod.includes("SKD") || mod.includes("CKD");
  return isAssembly || hasKdKeyword;
}

function isSpRow(row: any): boolean {
  if (isBatteryRow(row) || isKdRow(row)) return false;
  const mt = normalizeText(row["TYPE OF MATERIAL"] || "");
  const mod = normalizeText(row["MODEL"] || "");
  const rat = normalizeText(row["RATIONALIZATION"] || "");
  const op = normalizeText(row["OPERATION SCOPE"] || "");
  
  // Prioritize using the MODEL as a base
  if (mod.includes("SPARE") || mod.includes("PECAS") || mod.includes("REPOSIC") || mod.includes("SPARE PARTS")) {
    return true;
  }
  
  const combined = ` ${mt} ${mod} ${rat} ${op} `.replace(/[^A-Z0-9]/g, " ");
  const tokens = combined.split(/\s+/).filter(Boolean);
  return tokens.includes("SPARE") || tokens.includes("PARTS") || tokens.includes("PECAS") || tokens.includes("SP") || combined.includes("REPOSIC");
}

function isPbpRow(row: any): boolean {
  if (isBatteryRow(row) || isKdRow(row) || isSpRow(row)) return false;
  const mt = normalizeText(row["TYPE OF MATERIAL"] || "");
  const mod = normalizeText(row["MODEL"] || "");
  const rat = normalizeText(row["RATIONALIZATION"] || "");
  const op = normalizeText(row["OPERATION SCOPE"] || "");
  
  // Prioritize using the MODEL as a base
  if (mod.includes("PBP") || mod.includes("PART BY PART")) {
    return true;
  }
  
  const combined = ` ${mt} ${mod} ${rat} ${op} `.replace(/[^A-Z0-9]/g, " ");
  const tokens = combined.split(/\s+/).filter(Boolean);
  return tokens.includes("PBP") || combined.includes("PART BY PART");
}

function isProjectRow(row: any): boolean {
  return !isBatteryRow(row) && !isKdRow(row) && !isSpRow(row) && !isPbpRow(row);
}

function applyFiltersAndRender(activeTabId: string | null = null) {
  if (!activeTabId) {
    const activeTab = deliveryTabs?.querySelector(".tab-btn.active");
    activeTabId = (activeTab as HTMLElement)?.dataset.target || null;
  }
  const query = (searchInput?.value || "").trim().toLowerCase();
  const lotQuery = (lotSearchInput?.value || "").trim().toLowerCase();
  const selectedMonth = monthFilterSelect?.value;
  let filteredData = deliveryData;

  if (selectedMonth) {
    const monthIndex = parseInt(selectedMonth, 10);
    filteredData = filteredData.filter(row => {
      const d = toDateMaybe(row["DELIVERY AT BYD"]);
      if (!d) return false;
      return d.getMonth() === monthIndex;
    });
  }

  if (showOnlyBattery) {
    filteredData = filteredData.filter(row => isBatteryRow(row));
  }

  if (showOnlyKd) {
    filteredData = filteredData.filter(row => isKdRow(row));
  }

  if (showOnlySp) {
    filteredData = filteredData.filter(row => isSpRow(row));
  }

  if (showOnlyPbp) {
    filteredData = filteredData.filter(row => isPbpRow(row));
  }

  if (showOnlyProject) {
    filteredData = filteredData.filter(row => isProjectRow(row));
  }

  if (activeStatusFilter) {
    if (activeStatusFilter === "PENDENTE") {
      filteredData = filteredData.filter((row) => {
        const status = normalizeText(row["STATUS"] || "");
        return !["ENTREGUE", "A CAMINHO", "ADIADO", "CANCELADO", "AGUARDANDO DESOVA"].includes(status);
      });
    } else if (activeStatusFilter === "ENTREGUE") {
      filteredData = filteredData.filter((row) => {
        const status = normalizeText(row["STATUS"] || "");
        return status === "ENTREGUE" && isRowCompletedInOperationalTime(row);
      });
    } else if (activeStatusFilter === "AGUARDANDO DESOVA") {
      filteredData = filteredData.filter((row) => {
        const status = normalizeText(row["STATUS"] || "");
        return status === "AGUARDANDO DESOVA" || (status === "ENTREGUE" && !isRowCompletedInOperationalTime(row));
      });
    } else {
      filteredData = filteredData.filter((row) => normalizeText(row["STATUS"] || "") === activeStatusFilter);
    }
  }

  if (query) {
    const searchTerms = query.split(/[\s,\n\t]+/).filter(t => t.length > 0);
    filteredData = filteredData.filter((row) => {
      const poValues = [
        row["PO SAP"],
        row["PO"],
        row["PO NUMBER"],
        row["PO_SAP"],
        row["Pedido"],
        row["PO/SAP"]
      ].filter(v => v !== undefined && v !== null && v !== "").map(v => String(v).toLowerCase());

      const rowValues = Object.entries(row).map(([k, v]) => String(v ?? "").toLowerCase());
      return searchTerms.some(term => {
        if (rowValues.some(val => val.includes(term))) return true;
        if (poValues.some(val => val.includes(term))) return true;
        return false;
      });
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
  renderCharts(filteredData);
  renderHistoryTab();
  updateStats();
}

function updateStats() {
  const isHistoryTabActive = document.querySelector(".view-tab-btn[data-tab='history']")?.classList.contains("border-blue-500") ?? false;
  let dataForStats = isHistoryTabActive ? historicalData : deliveryData;

  const selectedMonth = monthFilterSelect?.value;
  if (selectedMonth) {
    const monthIndex = parseInt(selectedMonth, 10);
    dataForStats = dataForStats.filter(row => {
      const d = toDateMaybe(row["DELIVERY AT BYD"]);
      if (!d) return false;
      return d.getMonth() === monthIndex;
    });
  }

  if (showOnlyBattery) {
    dataForStats = dataForStats.filter(row => isBatteryRow(row));
  }
  if (showOnlyKd) {
    dataForStats = dataForStats.filter(row => isKdRow(row));
  }
  if (showOnlySp) {
    dataForStats = dataForStats.filter(row => isSpRow(row));
  }
  if (showOnlyPbp) {
    dataForStats = dataForStats.filter(row => isPbpRow(row));
  }
  if (showOnlyProject) {
    dataForStats = dataForStats.filter(row => isProjectRow(row));
  }

  const total = dataForStats.length;
  const delivered = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "ENTREGUE" && isRowCompletedInOperationalTime(d)).length;
  const inTransit = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "A CAMINHO").length;
  const postponed = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "ADIADO").length;
  const canceled = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "CANCELADO").length;
  const backlog = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "BACKLOG").length;
  const awaitingUnload = dataForStats.filter((d) => 
    normalizeText(d["STATUS"] || "") === "AGUARDANDO DESOVA" || 
    (normalizeText(d["STATUS"] || "") === "ENTREGUE" && !isRowCompletedInOperationalTime(d))
  ).length;
  const pending = Math.max(0, total - delivered - inTransit - postponed - canceled - awaitingUnload - backlog);

  const getPercentage = (count: number) => total === 0 ? "0%" : `${((count / total) * 100).toFixed(1)}%`;

  const getCardClasses = (cardStatus: string | null) => {
    const isAll = cardStatus === "ALL";
    const isActive = activeStatusFilter === cardStatus || (activeStatusFilter === null && isAll);
    let classes =
      "summary-card bg-white dark:bg-slate-800 p-3 rounded-lg shadow-sm border flex items-center cursor-pointer transition-all duration-200";
    if (isActive) classes += " border-blue-500 ring-2 ring-blue-500/50 scale-[1.02] z-10";
    else classes += " border-slate-200 dark:border-slate-700 hover:border-blue-300";
    return classes;
  };

  if (!summaryStats) return;
  summaryStats.innerHTML = `
    <div class="${getCardClasses("ALL")}" data-status="ALL">
      <div class="bg-blue-100 dark:bg-blue-900/50 text-blue-600 dark:text-blue-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-box-open text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("totalContainers")}">${t("totalContainers")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${total}</div>
      </div>
    </div>

    <div class="${getCardClasses("ENTREGUE")}" data-status="ENTREGUE">
      <div class="bg-green-100 dark:bg-green-900/50 text-green-600 dark:text-green-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-check-circle text-sm"></i>
      </div>
      <div class="min-w-0 flex-1">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("delivered")}">${t("delivered")}</div>
        <div class="flex items-baseline gap-1.5">
          <span class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${delivered}</span>
          <span class="text-[9px] font-bold text-green-600 dark:text-green-400">${getPercentage(delivered)}</span>
        </div>
        <div class="text-[8px] text-slate-400 dark:text-slate-500 font-medium leading-none mt-0.5 truncate" title="${currentLanguage === "pt" ? "Apenas containers já descarregados no Tempo de Operação" : "Only containers unloaded in Operation Time"}">${currentLanguage === "pt" ? "Descarregados" : "Unloaded"}</div>
      </div>
    </div>

    <div class="${getCardClasses("AGUARDANDO DESOVA")}" data-status="AGUARDANDO DESOVA">
      <div class="bg-purple-100 dark:bg-purple-900/50 text-purple-600 dark:text-purple-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-box text-sm"></i>
      </div>
      <div class="min-w-0 flex-1">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("awaitingUnload")}">${t("awaitingUnload")}</div>
        <div class="flex items-baseline gap-1.5">
          <span class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${awaitingUnload}</span>
          <span class="text-[9px] font-bold text-purple-600 dark:text-purple-400">${getPercentage(awaitingUnload)}</span>
        </div>
        <div class="text-[8px] text-slate-400 dark:text-slate-500 font-medium leading-none mt-0.5 truncate" title="${currentLanguage === "pt" ? "Manual + (Entregues não descarregados)" : "Manual + (Delivered but not unloaded)"}">${currentLanguage === "pt" ? "Não Descarregados" : "Not Unloaded"}</div>
      </div>
    </div>

    <div class="${getCardClasses("A CAMINHO")}" data-status="A CAMINHO">
      <div class="bg-yellow-100 dark:bg-yellow-900/50 text-yellow-600 dark:text-yellow-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-truck text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("inTransit")}">${t("inTransit")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${inTransit}</div>
        <div class="text-[9px] font-bold text-yellow-600 dark:text-yellow-400">${getPercentage(inTransit)}</div>
      </div>
    </div>

    <div class="${getCardClasses("PENDENTE")}" data-status="PENDENTE">
      <div class="bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-hourglass-half text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("pending")}">${t("pending")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${pending}</div>
        <div class="text-[9px] font-bold text-slate-600 dark:text-slate-400">${getPercentage(pending)}</div>
      </div>
    </div>

    <div class="${getCardClasses("ADIADO")}" data-status="ADIADO">
      <div class="bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-calendar-alt text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("postponed")}">${t("postponed")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${postponed}</div>
        <div class="text-[9px] font-bold text-indigo-600 dark:text-indigo-400">${getPercentage(postponed)}</div>
      </div>
    </div>

    <div class="${getCardClasses("BACKLOG")}" data-status="BACKLOG">
      <div class="bg-orange-100 dark:bg-orange-900/50 text-orange-600 dark:text-orange-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-history text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="Backlog">Backlog</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${backlog}</div>
        <div class="text-[9px] font-bold text-orange-600 dark:text-orange-400">${getPercentage(backlog)}</div>
      </div>
    </div>

    <div class="${getCardClasses("CANCELADO")}" data-status="CANCELADO">
      <div class="bg-red-100 dark:bg-red-900/50 text-red-600 dark:text-red-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-times-circle text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate">${t("canceled")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${canceled}</div>
        <div class="text-[9px] font-bold text-red-600 dark:text-red-400">${getPercentage(canceled)}</div>
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
              const bl = String(d["BL"] || "").trim();
              if (!acc[carrier]) acc[carrier] = {};
              if (!acc[carrier][lot]) acc[carrier][lot] = { total: 0, delivered: 0, bls: new Set<string>() };
              acc[carrier][lot].total++;
              if (normalizeText(d["STATUS"] || "") === "ENTREGUE") acc[carrier][lot].delivered++;
              if (bl) acc[carrier][lot].bls.add(bl);
              return acc;
            }, {} as Record<string, Record<string, { total: number; delivered: number; bls: Set<string> }>>)
          )
            .map(([carrier, lots]) => {
              const lotHTML = Object.entries(lots)
                .map(([lot, stats]) => {
                  return `
                    <div class="lot-details border-t border-slate-100 dark:border-slate-700 mt-2 pt-2 hidden">
                        <div class="text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Lote ${lot} ${stats.bls && stats.bls.size > 0 ? "- " + Array.from(stats.bls).join(", ") : ""}</div>
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
                <div class="flex flex-col gap-2 w-full">
                   <div class="w-full bg-slate-200 dark:bg-slate-700 h-1.5 rounded-full overflow-hidden mt-1">
                      <div class="bg-blue-500 h-full transition-all duration-700" style="width: ${carrierPercent}%"></div>
                   </div>
                   <div class="flex items-center justify-between text-xs text-slate-500 dark:text-slate-400">
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
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase w-40">Pareto</th>
              <th class="px-4 py-2 text-left text-slate-500 text-xs uppercase w-40">${t("tableHeaderStatus")}</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100 dark:divide-slate-700">
            ${deliveries
              .map((row, rowIndex) => {
                const status = normalizeText(row["STATUS"] || "PENDENTE") || "PENDENTE";
                const pareto = row["PARETO"] || "-";
                const isBattery = isBatteryRow(row);
                const isKd = isKdRow(row);
                const isSp = isSpRow(row);
                const isPbp = isPbpRow(row);
                const rowClass = `transition-colors hover:bg-slate-50 dark:hover:bg-slate-700/50 cursor-pointer ${
                  isBattery ? "is-battery" : ""
                } ${isKd ? "is-kd" : ""} ${isSp ? "is-sp" : ""} ${isPbp ? "is-pbp" : ""} ${status === "ENTREGUE" ? "bg-green-100 dark:bg-green-900/30" : status === "CANCELADO" ? "bg-red-100 dark:bg-red-900/30" : status === "BACKLOG" ? "bg-orange-100 dark:bg-orange-900/30" : ""}`;

                const plate = String(row["TRUCK LICENSE PLATE 1"] || row["PLATE"] || "").trim();

                return `<tr class="${rowClass}" data-row-id="${row._id}">
                  <td class="px-4 py-3 text-xs text-center border-l-8 ${isBattery ? "border-amber-600" : isKd ? "border-blue-700" : isSp ? "border-orange-600" : isPbp ? "border-emerald-600" : "border-transparent"}">${rowIndex + 1}</td>
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
                    ${isSp
                      ? `<span class="ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-orange-100 text-orange-800 dark:bg-orange-900 dark:text-orange-200 uppercase">SP</span>`
                      : ""}
                    ${isPbp
                      ? `<span class="ml-2 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-900 dark:text-emerald-200 uppercase">PBP</span>`
                      : ""}
                  </td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["MODEL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["OPERATION SCOPE"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-mono">${row["BL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["TRANSPORTATION COMPANY"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["VESSEL"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">${row["BONDED WAREHOUSE"] || "-"}</td>
                  <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-300 font-medium">${row["LOT"] || "-"}</td>
                  <td class="px-4 py-3 text-xs" onclick="event.stopPropagation()">
                    <select class="pareto-select bg-white dark:bg-slate-700 dark:text-slate-200 border border-slate-300 dark:border-slate-500 text-xs rounded-md p-1 w-full" data-row-id="${row._id}">
                      <option value="-" ${pareto === "-" ? "selected" : ""}>-</option>
                      ${(window as any).__PARETO_REASONS__ ? (window as any).__PARETO_REASONS__.map((opt: string) => `<option value="${opt}" ${pareto === opt ? "selected" : ""}>${opt}</option>`).join("") : [
                        "PRAZO CURTO PARA COLETA",
                        "QUEBRA DE VEÍCULO",
                        "INCIDENTE TERMINAL",
                        "GREVE DOS CAMINHONEIROS",
                        "GREVE SINDICAL",
                        "ALTERAÇÃO DE PROGRAMAÇÃO",
                        "ACIDENTE NA RODOVIA",
                        "FILA NO TERMINAL",
                        "PENDÊNCIA DOCUMENTAL"
                      ].map(opt => `<option value="${opt}" ${pareto === opt ? "selected" : ""}>${opt}</option>`).join("")}
                    </select>
                  </td>
                  <td class="px-4 py-3 text-xs" onclick="event.stopPropagation()">
                    <select class="status-select bg-white dark:bg-slate-700 dark:text-slate-200 border border-slate-300 dark:border-slate-500 text-xs rounded-md p-1 w-full" data-row-id="${row._id}">
                      ${["PENDENTE", "AGUARDANDO DESOVA", "A CAMINHO", "ADIADO", "BACKLOG", "ENTREGUE", "CANCELADO"]
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
    
    // Wire dynamic collapse / search events inside tab card mapping
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
        if (searchInput) {
          searchInput.value = carrier || "";
          applyFiltersAndRender();
        }
      });
    });
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
    <td colspan="10" class="details-cell">
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
            <p class="text-sm font-medium mt-1 text-slate-800 dark:text-slate-100 italic">${String(rowData["NOTES"] || "-")}</p>
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
  if (s === "BACKLOG") return "BACKLOG";
  if (s === "CANCELED" || s === "CANCELLED") return "CANCELADO";
  if (s === "AWAITING UNLOAD") return "AGUARDANDO DESOVA";
  if (isExcelErrorString(raw)) return "PENDENTE";
  if (!["PENDENTE", "A CAMINHO", "ADIADO", "ENTREGUE", "CANCELADO", "AGUARDANDO DESOVA", "BACKLOG"].includes(s)) return "PENDENTE";
  return s;
}

deliveryContent?.addEventListener("change", async (e) => {
  const target = e.target as HTMLElement;
  const statusSelect = target.closest<HTMLSelectElement>(".status-select");
  const paretoSelect = target.closest<HTMLSelectElement>(".pareto-select");
  
  if (paretoSelect) {
    const rowId = paretoSelect.dataset.rowId || "";
    const row = deliveryData.find((d) => d._id === rowId);
    if (!row) return;

    const next = paretoSelect.value;
    const prev = row["PARETO"] || "-";
    if (next === prev) return;

    row["PARETO"] = next;
    showToast(`Motivo Pareto atualizado para ${next}`, "success");
    await saveStateToFirebase();
    applyFiltersAndRender();
    return;
  }

  if (!statusSelect) return;

  const select = statusSelect;
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
    await saveStateToFirebase();
    applyFiltersAndRender();
  } else {
    select.value = prev;
  }
});

/* ------------------------------ TABS -------------------------------------- */
viewModeTabs?.addEventListener("click", (e) => {
  const btn = (e.target as HTMLElement).closest<HTMLButtonElement>(".view-tab-btn");
  if (btn) {
    viewModeTabs.querySelectorAll(".view-tab-btn").forEach((b) => {
      b.classList.remove("border-blue-500", "text-blue-600");
      b.classList.add("border-transparent", "text-slate-500");
    });
    btn.classList.add("border-blue-500", "text-blue-600");
    btn.classList.remove("border-transparent", "text-slate-500");

    const target = btn.dataset.tab;
    deliveriesWrapper?.classList.toggle("hidden", target !== "deliveries");
    
    const chartsContent = document.getElementById("charts-content");
    chartsContent?.classList.toggle("hidden", target !== "charts");
    
    const timeContent = document.getElementById("time-content");
    timeContent?.classList.toggle("hidden", target !== "time");
    
    const historyContent = document.getElementById("history-content");
    historyContent?.classList.toggle("hidden", target !== "history");

    const paretoContent = document.getElementById("pareto-content");
    paretoContent?.classList.toggle("hidden", target !== "pareto");

    if (target === "charts") {
      renderCharts(deliveryData);
    } else if (target === "time") {
      renderTimeTable(deliveryData);
    } else if (target === "history") {
      renderHistoryTab();
    } else if (target === "pareto") {
      renderParetoTab();
    }

    updateStats();
  }
});

deliveryTabs?.addEventListener("click", (e) => {
  const btn = (e.target as HTMLElement).closest<HTMLButtonElement>(".tab-btn");
  if (btn) {
    deliveryTabs.querySelectorAll(".tab-btn").forEach((b) => {
      b.classList.remove("active", "border-blue-500", "text-blue-600");
      b.classList.add("border-transparent", "text-slate-500");
    });
    btn.classList.add("active", "border-blue-500", "text-blue-600");
    btn.classList.remove("border-transparent", "text-slate-500");

    const target = btn.dataset.target;
    if (target) {
      document.querySelectorAll(".date-card").forEach((c) => c.classList.add("hidden"));
      document.getElementById(target)?.classList.remove("hidden");
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

monthFilterSelect?.addEventListener("change", () => {
  applyFiltersAndRender();
});

(window as any).pontoApoioQtd = 0;
(window as any).updatePontoApoio = (val: string) => { 
  (window as any).pontoApoioQtd = parseInt(val, 10) || 0; 
  applyFiltersAndRender(); 
};

function resetCategoryButtonsUI() {
  [batteryFilterBtn, kdFilterBtn, spFilterBtn, pbpFilterBtn, projectFilterBtn].forEach(btn => {
    if (!btn) return;
    btn.classList.remove("ring-2", "ring-amber-500", "ring-blue-500", "ring-orange-500", "ring-emerald-500", "ring-purple-500", "bg-amber-50", "bg-blue-50", "bg-orange-50", "bg-emerald-50", "bg-purple-50", "dark:bg-amber-900/30", "dark:bg-blue-900/30", "dark:bg-orange-900/30", "dark:bg-emerald-900/30", "dark:bg-purple-900/30");
  });
}

batteryFilterBtn?.addEventListener("click", () => {
  const targetState = !showOnlyBattery;
  showOnlyBattery = false; showOnlyKd = false; showOnlySp = false; showOnlyPbp = false; showOnlyProject = false;
  showOnlyBattery = targetState;
  
  resetCategoryButtonsUI();
  if (showOnlyBattery) {
    batteryFilterBtn.classList.add("ring-2", "ring-amber-500", "bg-amber-50", "dark:bg-amber-900/30");
  }
  applyFiltersAndRender();
});

kdFilterBtn?.addEventListener("click", () => {
  const targetState = !showOnlyKd;
  showOnlyBattery = false; showOnlyKd = false; showOnlySp = false; showOnlyPbp = false; showOnlyProject = false;
  showOnlyKd = targetState;
  
  resetCategoryButtonsUI();
  if (showOnlyKd) {
    kdFilterBtn.classList.add("ring-2", "ring-blue-500", "bg-blue-50", "dark:bg-blue-900/30");
  }
  applyFiltersAndRender();
});

spFilterBtn?.addEventListener("click", () => {
  const targetState = !showOnlySp;
  showOnlyBattery = false; showOnlyKd = false; showOnlySp = false; showOnlyPbp = false; showOnlyProject = false;
  showOnlySp = targetState;
  
  resetCategoryButtonsUI();
  if (showOnlySp) {
    spFilterBtn.classList.add("ring-2", "ring-orange-500", "bg-orange-50", "dark:bg-orange-900/30");
  }
  applyFiltersAndRender();
});

pbpFilterBtn?.addEventListener("click", () => {
  const targetState = !showOnlyPbp;
  showOnlyBattery = false; showOnlyKd = false; showOnlySp = false; showOnlyPbp = false; showOnlyProject = false;
  showOnlyPbp = targetState;
  
  resetCategoryButtonsUI();
  if (showOnlyPbp) {
    pbpFilterBtn.classList.add("ring-2", "ring-emerald-500", "bg-emerald-50", "dark:bg-emerald-900/30");
  }
  applyFiltersAndRender();
});

projectFilterBtn?.addEventListener("click", () => {
  const targetState = !showOnlyProject;
  showOnlyBattery = false; showOnlyKd = false; showOnlySp = false; showOnlyPbp = false; showOnlyProject = false;
  showOnlyProject = targetState;
  
  resetCategoryButtonsUI();
  if (showOnlyProject) {
    projectFilterBtn.classList.add("ring-2", "ring-purple-500", "bg-purple-50", "dark:bg-purple-900/30");
  }
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

/* --------------------------------- CHARTS ---------------------------------- */
function renderCharts(data: DeliveryRow[]) {
  const chartsContent = document.getElementById("charts-content");
  if (!chartsContent) return;

  if (overallChart) { overallChart.destroy(); overallChart = null; }
  if (lotChart) { lotChart.destroy(); lotChart = null; }
  if (modelChart) { modelChart.destroy(); modelChart = null; }
  carrierCharts.forEach(c => c.destroy()); carrierCharts = [];
  warehouseCharts.forEach(c => c.destroy()); warehouseCharts = [];

  if (typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
    Chart.defaults.set('plugins.datalabels', {
      color: '#ffffff',
      font: { weight: 'bold', size: 10 },
      formatter: (value: number, ctx: any) => {
        if (value === 0) return '';
        let sum = 0;
        let dataArr = ctx.chart.data.datasets[0].data;
        dataArr.map((data: number) => { sum += data; });
        return sum > 0 ? (value * 100 / sum).toFixed(1) + "%" : "0%";
      }
    });
  }

  if (typeof Chart === "undefined") {
     console.warn("Chart.js is not loaded.");
     return;
  }

    const statusColors: Record<string, string> = {
      "ENTREGUE": "#22c55e",
      "A CAMINHO": "#3b82f6",
      "AGUARDANDO DESOVA": "#a855f7",
      "BACKLOG": "#f97316",
      "PENDENTE": "#64748b",
      "ADIADO": "#6366f1",
      "CANCELADO": "#ef4444"
    };

    const statusLabels = [t("delivered"), t("inTransit"), t("awaitingUnload"), t("statusBacklog"), t("pending"), t("postponed"), t("canceled")];
    
    function getStatusIndex(s: string) {
      if (s === "ENTREGUE") return 0;
      if (s === "A CAMINHO") return 1;
      if (s === "AGUARDANDO DESOVA") return 2;
      if (s === "BACKLOG") return 3;
      if (s === "PENDENTE") return 4;
      if (s === "ADIADO") return 5;
      if (s === "CANCELADO") return 6;
      return 4;
    }

    const customLegendHTML = `
      <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4 sticky top-4">
        <h4 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-3 border-b border-slate-200 dark:border-slate-600 pb-2 uppercase tracking-wider" data-i18n="legendTitle">${t("legendTitle")}</h4>
        <div class="space-y-3 text-sm text-slate-600 dark:text-slate-300 font-medium">
          ${statusLabels.map((lbl, idx) => `
            <div class="flex items-center">
              <span class="w-4 h-4 rounded-md mr-3 shadow-sm border border-slate-200/20" style="background-color: ${Object.values(statusColors)[idx]}"></span>
              <span>${lbl}</span>
            </div>
          `).join('')}
        </div>
      </div>
    `;

    let overallCounts = [0, 0, 0, 0, 0, 0, 0];
  data.forEach((row) => {
    let s = normalizeText(row["STATUS"] || "PENDENTE");
    overallCounts[getStatusIndex(s)]++;
  });

  const total = data.length;
  const delivered = overallCounts[0];
  const inTransit = overallCounts[1];
  const waiting = overallCounts[2];
  const backlogStat = overallCounts[3];
  const pending = overallCounts[4];
  
  const efficiency = total > 0 ? ((delivered / total) * 100).toFixed(1) : "0.0";
  const progressPct = total > 0 ? (((delivered + inTransit + waiting + backlogStat) / total) * 100).toFixed(1) : "0.0";

  chartsContent.innerHTML = `
    <div class="space-y-6 pb-8">
      <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-6 gap-6 p-4">
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4 flex flex-col md:col-span-1 xl:col-span-1">
          <div class="flex gap-2 border-b border-slate-100 pb-2 mb-2 justify-around">
            <div class="bg-slate-50 dark:bg-slate-900 rounded p-1 text-center shadow-sm flex-1">
              <div class="text-[#0f172a] dark:text-slate-100 text-base font-black">${efficiency}%</div>
              <div class="text-[9px] font-bold text-slate-400" data-i18n="efic">${t("efic")}</div>
            </div>
            <div class="bg-slate-50 dark:bg-slate-900 rounded p-1 text-center shadow-sm flex-1">
              <div class="text-[#0f172a] dark:text-slate-100 text-base font-black">${progressPct}%</div>
              <div class="text-[9px] font-bold text-slate-400" data-i18n="prog">${t("prog")}</div>
            </div>
            <div class="bg-slate-50 dark:bg-slate-900 rounded p-1 text-center shadow-sm flex-1 flex flex-col justify-center">
              <div class="text-blue-600 dark:text-blue-400 text-base font-black">${pending}</div>
              <div class="text-[9px] font-bold text-slate-400" data-i18n="pend">${t("pend")}</div>
            </div>
          </div>
          <div class="flex-grow min-w-0">
            <h3 class="text-xs font-bold text-slate-700 dark:text-slate-200 mb-2 text-center" data-i18n="chartsOverviewTitle">${t("chartsOverviewTitle")}</h3>
            <div class="relative h-40">
               <canvas id="overallChartCanvas"></canvas>
            </div>
          </div>
        </div>

        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4 md:col-span-1 xl:col-span-4 overflow-hidden relative">
          <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-4 border-b border-slate-100 dark:border-slate-700 pb-3">
             <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200" id="charts-group-progress-title">
               ${chartGroupBy === 'lot' ? t("chartsLotProgressTitle") : "Progresso por PO SAP"}
             </h3>
             <div class="flex items-center gap-2">
                <!-- Group By Selector -->
                <div class="inline-flex rounded-md shadow-sm border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 p-0.5">
                  <button id="group-by-lot-btn" type="button" class="text-[10px] font-bold px-2 py-1 rounded transition-all cursor-pointer ${chartGroupBy === 'lot' ? 'bg-white dark:bg-slate-700 text-blue-600 shadow-sm border border-slate-200/50 dark:border-slate-600' : 'text-slate-500 hover:text-slate-700 dark:hover:text-slate-200'}">
                    Lote
                  </button>
                  <button id="group-by-po-btn" type="button" class="text-[10px] font-bold px-2 py-1 rounded transition-all cursor-pointer ${chartGroupBy === 'po' ? 'bg-white dark:bg-slate-700 text-blue-600 shadow-sm border border-slate-200/50 dark:border-slate-600' : 'text-slate-500 hover:text-slate-700 dark:hover:text-slate-200'}">
                    PO SAP
                  </button>
                </div>
                
                <button id="toggle-macro-view-btn" type="button" class="text-xs font-bold bg-slate-50 text-slate-500 hover:text-blue-600 px-2.5 py-1.5 flex items-center justify-center rounded hover:bg-slate-100 dark:bg-slate-700 dark:text-slate-200 dark:hover:bg-slate-600 transition border border-slate-200 dark:border-slate-600 shadow-sm" title="Toggle Macro View">
                  <i class="fas fa-layer-group mr-1"></i> Macro View
                </button>
                
                <button id="maximize-chart-btn" type="button" class="text-xs font-bold bg-slate-50 text-slate-500 hover:text-blue-600 px-2.5 py-1.5 flex items-center justify-center rounded hover:bg-slate-100 dark:bg-slate-700 dark:text-slate-200 dark:hover:bg-slate-600 transition border border-slate-200 dark:border-slate-600 shadow-sm cursor-pointer" title="Maximizar Visualização">
                  <i class="fas fa-expand mr-1"></i> Maximizar
                </button>
             </div>
          </div>
          <div class="relative h-64 w-full cursor-grab active:cursor-grabbing overflow-x-auto pb-2 custom-scrollbar">
             <div style="min-width: 800px; height: 100%;">
                <canvas id="lotChartCanvas"></canvas>
             </div>
          </div>
          ${isMacroView ? `
            <div class="mt-4 flex flex-wrap justify-center gap-4 border-t border-slate-100 dark:border-slate-700 pt-4">
              ${statusLabels.map((lbl, idx) => `
                <div class="flex items-center text-[10px] font-bold text-slate-500">
                  <span class="w-3 h-3 rounded-sm mr-2" style="background-color: ${Object.values(statusColors)[idx]}"></span>
                  ${lbl}
                </div>
              `).join('')}
            </div>
          ` : ''}
        </div>

        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4 md:col-span-2 xl:col-span-1">
          <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-4 text-center" data-i18n="modelsTitle">${t("modelsTitle")}</h3>
          <div class="relative h-64">
             <canvas id="modelChartCanvas"></canvas>
          </div>
        </div>
      </div>

      <div class="flex flex-col lg:flex-row gap-6 p-4">
        <div class="flex-grow space-y-8 min-w-0">
          <div>
            <h3 class="text-lg font-bold text-slate-800 dark:text-slate-100 mb-4 border-b border-slate-200 dark:border-slate-700 pb-2" data-i18n="chartsCarrierTitle">${t("chartsCarrierTitle")}</h3>
            <div class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-4" id="carrier-charts-grid"></div>
          </div>
          <div>
            <h3 class="text-lg font-bold text-slate-800 dark:text-slate-100 mb-4 border-b border-slate-200 dark:border-slate-700 pb-2" data-i18n="chartsWarehouseTitle">${t("chartsWarehouseTitle")}</h3>
            <div class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-4" id="warehouse-charts-grid"></div>
          </div>
        </div>
        <div class="w-full lg:w-64 shrink-0">
          ${customLegendHTML}
        </div>
      </div>
    </div>
  `;

  const macroBtn = document.getElementById("toggle-macro-view-btn");
  if (macroBtn) {
    if (isMacroView) macroBtn.classList.add("bg-blue-50", "text-blue-600");
    macroBtn.addEventListener("click", () => {
      isMacroView = !isMacroView;
      renderCharts(data);
    });
  }

  const lotBtn = document.getElementById("group-by-lot-btn");
  const poBtn = document.getElementById("group-by-po-btn");
  if (lotBtn) {
    lotBtn.addEventListener("click", () => {
      chartGroupBy = "lot";
      renderCharts(data);
    });
  }
  if (poBtn) {
    poBtn.addEventListener("click", () => {
      chartGroupBy = "po";
      renderCharts(data);
    });
  }

  const maximizeBtn = document.getElementById("maximize-chart-btn");
  if (maximizeBtn) {
    maximizeBtn.addEventListener("click", () => {
      const maxContainer = document.getElementById("chart-max-modal-container");
      const maxModal = document.getElementById("chart-max-modal");
      if (maxContainer && maxModal) {
        maxContainer.classList.remove("hidden");
        setTimeout(() => {
          maxModal.classList.remove("scale-95", "opacity-0");
          maxModal.classList.add("scale-100", "opacity-100");
          renderCharts(data);
        }, 10);
      }
    });
  }

  // Handle modal close buttons statically using onclick
  const maxCloseBtn = document.getElementById("chart-max-close-btn");
  const maxContainer = document.getElementById("chart-max-modal-container");
  const maxModal = document.getElementById("chart-max-modal");
  if (maxCloseBtn && maxContainer && maxModal) {
    const closeModal = () => {
      maxModal.classList.add("scale-95", "opacity-0");
      maxModal.classList.remove("scale-100", "opacity-100");
      setTimeout(() => {
        maxContainer.classList.add("hidden");
        if (maxLotChart) {
          maxLotChart.destroy();
          maxLotChart = null;
        }
      }, 200);
    };
    maxCloseBtn.onclick = closeModal;
    maxContainer.onclick = (e) => {
      if (e.target === maxContainer) {
        closeModal();
      }
    };
  }

  const ctxOverall = document.getElementById("overallChartCanvas") as HTMLCanvasElement;
  if (ctxOverall) {
    overallChart = new Chart(ctxOverall, {
      type: "doughnut",
      data: {
        labels: statusLabels,
        datasets: [{
          data: overallCounts,
          backgroundColor: Object.values(statusColors),
          borderWidth: 1,
          borderColor: "#ffffff"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context: any) {
                let value = context.parsed || 0;
                let totalSum = context.chart._metasets[context.datasetIndex].total;
                let pct = totalSum > 0 ? Math.round(value / totalSum * 100) : 0;
                return `${context.label}: ${value} (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  const groupStats: Record<string, { total: number; done: number; statusCounts: number[]; carriers: Set<string>; operations: Set<string> }> = {};
  data.forEach((row) => {
    const rawVal = chartGroupBy === "lot" ? row["LOT"] : row["PO SAP"];
    const key = String(rawVal || "N/A").trim() || "N/A";
    if (!groupStats[key]) groupStats[key] = { total: 0, done: 0, statusCounts: [0,0,0,0,0,0,0], carriers: new Set(), operations: new Set() };
    groupStats[key].total++;
    const status = normalizeText(row["STATUS"] || "PENDENTE");
    if (status === "ENTREGUE") groupStats[key].done++;
    groupStats[key].statusCounts[getStatusIndex(status)]++;
    const carrier = String(row["TRANSPORTATION COMPANY"] || "").trim().toUpperCase();
    if (carrier) groupStats[key].carriers.add(carrier);
    let operation = String(row["OPERATION SCOPE"] || "").trim().toUpperCase();
    if (operation) {
      if (operation.includes("UNLOAD") || operation.includes("DESOVA")) operation = "UNLOAD";
      else if (operation.includes("SWAP")) operation = "SWAP";
      groupStats[key].operations.add(operation);
    }
  });

  const sortedGroups = Object.keys(groupStats).sort((a, b) => {
    if (a === "N/A") return 1;
    if (b === "N/A") return -1;
    return a.localeCompare(b, undefined, { numeric: true });
  });
  const groupLabels = sortedGroups;

  const ctxLot = document.getElementById("lotChartCanvas") as HTMLCanvasElement;
  if (ctxLot) {
    const minW = Math.max(800, groupLabels.length * 45);
    ctxLot.parentElement!.style.minWidth = `${minW}px`;
    
    let chartData, chartOptions;

    if (isMacroView) {
      chartData = {
        labels: groupLabels,
        datasets: statusLabels.map((lbl, idx) => ({
          label: lbl,
          data: sortedGroups.map(grp => groupStats[grp].statusCounts[idx]),
          backgroundColor: Object.values(statusColors)[idx]
        }))
      };
      chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, ticks: { color: "#64748b" }, grid: { display: false } },
          y: { stacked: true, beginAtZero: true, ticks: { color: "#64748b" }, grid: { color: "rgba(100, 116, 139, 0.1)" } }
        },
        plugins: {
          legend: { display: false },
          datalabels: {
            color: '#fff',
            font: { weight: 'bold', size: 10 },
            formatter: (value: number) => value > 0 ? value : ''
          },
          tooltip: {
            mode: 'index',
            intersect: false,
            backgroundColor: 'rgba(15, 23, 42, 0.9)',
            titleFont: { size: 14, weight: 'bold' },
            padding: 12,
            cornerRadius: 8,
            callbacks: {
              title: (items: any) => {
                const label = items[0].label;
                return (chartGroupBy === "lot" ? "Lote: " : "PO SAP: ") + label;
              },
              afterBody: (items: any) => {
                const grp = items[0].label;
                const stats = groupStats[grp];
                const carriers = Array.from(stats.carriers || []).join(", ") || "N/A";
                const ops = Array.from(stats.operations || []).join(", ") || "N/A";
                return `\nTransportadora: ${carriers}\nEscopo da Operação: ${ops}`;
              },
              label: (item: any) => {
                return ` ${item.dataset.label}: ${item.parsed.y}`;
              }
            }
          }
        }
      };
    } else {
      const groupData = sortedGroups.map((grp) => groupStats[grp].total > 0 ? (groupStats[grp].done / groupStats[grp].total) * 100 : 0);
      chartData = {
        labels: groupLabels,
        datasets: [{
          label: "% " + t("delivered"),
          data: groupData,
          backgroundColor: groupData.map(v => v === 100 ? "#22c55e" : "#3b82f6"),
          borderRadius: 4
        }]
      };
      chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, max: 100, ticks: { color: "#64748b" } },
          x: { ticks: { color: "#64748b" }, grid: { display: false } }
        },
        plugins: {
          legend: { display: false },
          datalabels: {
            formatter: (value: number) => value > 0 ? value.toFixed(0) + '%' : ''
          },
          tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.9)',
            titleFont: { size: 14, weight: 'bold' },
            padding: 12,
            cornerRadius: 8,
            callbacks: {
              title: (items: any) => {
                const label = items[0].label;
                return (chartGroupBy === "lot" ? "Lote: " : "PO SAP: ") + label;
              },
              afterBody: (items: any) => {
                const grp = items[0].label;
                const stats = groupStats[grp];
                const carriers = Array.from(stats.carriers || []).join(", ") || "N/A";
                const ops = Array.from(stats.operations || []).join(", ") || "N/A";
                return `\nQuantidade total: ${stats.total}\nEntregue: ${stats.done}\nTransportadora: ${carriers}\nEscopo da Operação: ${ops}`;
              }
            }
          }
        }
      };
    }

    lotChart = new Chart(ctxLot, { type: "bar", data: chartData as any, options: chartOptions as any });
    
    // --- Maximized Chart Render & Sync ---
    const maxContainer = document.getElementById("chart-max-modal-container");
    const ctxMaxLot = document.getElementById("maxLotChartCanvas") as HTMLCanvasElement;
    if (maxContainer && !maxContainer.classList.contains("hidden") && ctxMaxLot) {
      const maxMinW = Math.max(1000, groupLabels.length * 60);
      ctxMaxLot.parentElement!.style.minWidth = `${maxMinW}px`;

      const maxTitle = document.getElementById("max-chart-title");
      if (maxTitle) {
        maxTitle.textContent = chartGroupBy === "lot" ? t("chartsLotProgressTitle") : "Progresso por PO SAP";
      }

      // Sync active button styles in modal
      const modalLotBtn = document.getElementById("modal-group-by-lot-btn");
      const modalPoBtn = document.getElementById("modal-group-by-po-btn");
      const modalToggleBtn = document.getElementById("modal-toggle-macro-view-btn");

      if (modalLotBtn && modalPoBtn) {
        const activeClass = "text-[10px] font-bold px-2 py-1 rounded transition-all cursor-pointer bg-white dark:bg-slate-700 text-blue-600 shadow-sm border border-slate-200/50 dark:border-slate-600";
        const inactiveClass = "text-[10px] font-bold px-2 py-1 rounded transition-all cursor-pointer text-slate-500 hover:text-slate-700 dark:hover:text-slate-200";
        if (chartGroupBy === "lot") {
          modalLotBtn.className = activeClass;
          modalPoBtn.className = inactiveClass;
        } else {
          modalLotBtn.className = inactiveClass;
          modalPoBtn.className = activeClass;
        }
      }

      if (modalToggleBtn) {
        if (isMacroView) {
          modalToggleBtn.classList.add("bg-blue-50", "text-blue-600");
        } else {
          modalToggleBtn.classList.remove("bg-blue-50", "text-blue-600");
        }
      }

      const modalLegendContainer = document.getElementById("modal-macro-legend-container");
      if (modalLegendContainer) {
        if (isMacroView) {
          modalLegendContainer.classList.remove("hidden");
          modalLegendContainer.innerHTML = statusLabels.map((lbl, idx) => `
            <div class="flex items-center text-[10px] font-bold text-slate-500">
              <span class="w-3 h-3 rounded-sm mr-2" style="background-color: ${Object.values(statusColors)[idx]}"></span>
              ${lbl}
            </div>
          `).join('');
        } else {
          modalLegendContainer.classList.add("hidden");
        }
      }

      if (maxLotChart) {
        maxLotChart.destroy();
      }
      maxLotChart = new Chart(ctxMaxLot, { type: "bar", data: chartData as any, options: chartOptions as any });

      // Listeners for buttons inside modal (using onclick to avoid duplication)
      const maxLotBtn = document.getElementById("modal-group-by-lot-btn");
      const maxPoBtn = document.getElementById("modal-group-by-po-btn");
      const maxToggleBtn = document.getElementById("modal-toggle-macro-view-btn");

      if (maxLotBtn) {
        maxLotBtn.onclick = () => {
          chartGroupBy = "lot";
          renderCharts(data);
        };
      }
      if (maxPoBtn) {
        maxPoBtn.onclick = () => {
          chartGroupBy = "po";
          renderCharts(data);
        };
      }
      if (maxToggleBtn) {
        maxToggleBtn.onclick = () => {
          isMacroView = !isMacroView;
          renderCharts(data);
        };
      }
    }
  }

  const modelStats: Record<string, number> = {};
  data.forEach((row) => {
    const model = String(row["MODEL"] || "").trim().toUpperCase() || "OUTROS";
    modelStats[model] = (modelStats[model] || 0) + 1;
  });

  const sortedModels = Object.keys(modelStats).sort((a,b) => modelStats[b] - modelStats[a]);
  const ctxModel = document.getElementById("modelChartCanvas") as HTMLCanvasElement;
  if (ctxModel) {
    modelChart = new Chart(ctxModel, {
      type: "bar",
      data: {
        labels: sortedModels,
        datasets: [{ data: sortedModels.map(m => modelStats[m]), backgroundColor: "#8b5cf6", borderRadius: 4 }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } }
      }
    });
  }

  const carrierStats: Record<string, number[]> = {};
  data.forEach((row) => {
    const carrier = String(row["TRANSPORTATION COMPANY"] || "N/A").trim().toUpperCase() || "N/A";
    if (!carrierStats[carrier]) carrierStats[carrier] = [0, 0, 0, 0, 0];
    carrierStats[carrier][getStatusIndex(normalizeText(row["STATUS"] || "PENDENTE"))]++;
  });

  const carrierGrid = document.getElementById("carrier-charts-grid");
  if (carrierGrid) {
    Object.keys(carrierStats).sort().forEach((carrier, idx) => {
      const containerId = `carrier-chart-${idx}`;
      const carrierTotal = carrierStats[carrier].reduce((a, b) => a + b, 0);
      const cEfficiency = carrierTotal > 0 ? ((carrierStats[carrier][0] / carrierTotal) * 100).toFixed(1) : "0.0";

      carrierGrid.insertAdjacentHTML("beforeend", `
        <div class="flex flex-col items-center">
          <h4 class="text-xs font-bold text-slate-700 dark:text-slate-200 mb-2 w-full text-center truncate">${carrier} (${carrierTotal})</h4>
          <div class="relative h-48 w-full"><canvas id="${containerId}"></canvas></div>
          <div class="mt-2 text-center"><span class="text-lg font-black text-[#0f172a] dark:text-slate-100">${cEfficiency}%</span></div>
        </div>
      `);
      
      const ctx = document.getElementById(containerId) as HTMLCanvasElement;
      if (ctx) {
        const cChart = new Chart(ctx, {
          type: "doughnut",
          data: { labels: statusLabels, datasets: [{ data: carrierStats[carrier], backgroundColor: Object.values(statusColors) }] },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });
        carrierCharts.push(cChart);
      }
    });
  }

  const warehouseStats: Record<string, number[]> = {};
  data.forEach((row) => {
    const wh = String(row["BONDED WAREHOUSE"] || "N/A").trim().toUpperCase() || "N/A";
    if (!warehouseStats[wh]) warehouseStats[wh] = [0, 0, 0, 0, 0];
    warehouseStats[wh][getStatusIndex(normalizeText(row["STATUS"] || "PENDENTE"))]++;
  });

  const warehouseGrid = document.getElementById("warehouse-charts-grid");
  if (warehouseGrid) {
    Object.keys(warehouseStats).sort().forEach((wh, idx) => {
      const containerId = `warehouse-chart-${idx}`;
      const whTotal = warehouseStats[wh].reduce((a, b) => a + b, 0);

      warehouseGrid.insertAdjacentHTML("beforeend", `
        <div class="flex flex-col items-center">
          <h4 class="text-xs font-bold text-slate-700 dark:text-slate-200 mb-2 w-full text-center truncate">${wh} (${whTotal})</h4>
          <div class="relative h-48 w-full"><canvas id="${containerId}"></canvas></div>
        </div>
      `);
      
      const ctx = document.getElementById(containerId) as HTMLCanvasElement;
      if (ctx) {
        const wChart = new Chart(ctx, {
          type: "doughnut",
          data: { labels: statusLabels, datasets: [{ data: warehouseStats[wh], backgroundColor: Object.values(statusColors) }] },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });
        warehouseCharts.push(wChart);
      }
    });
  }
}


/* ----------------------- ARRIVALS & METRICS ---------------------------- */
function renderHistoryTab() {
  const historyContent = document.getElementById("history-content");
  if (!historyContent) return;

  const activeElem = document.activeElement as HTMLElement;
  const isTypingInHistory = activeElem?.classList.contains("history-note-input");
  let activeData = { date: '', carrier: '', field: '', value: '' };
  if (isTypingInHistory) {
    activeData = {
      date: activeElem.dataset.date || '',
      carrier: activeElem.dataset.carrier || '',
      field: activeElem.dataset.field || '',
      value: (activeElem as HTMLInputElement).value
    };
  }

  const selectedMonth = monthFilterSelect?.value;
  let filteredHistory = historicalData;
  if (selectedMonth) {
    const monthIndex = parseInt(selectedMonth, 10);
    filteredHistory = filteredHistory.filter(row => {
      const d = toDateMaybe(row["DELIVERY AT BYD"]);
      if (!d) return false;
      return d.getMonth() === monthIndex;
    });
  }

  if (filteredHistory.length === 0) {
    historyContent.innerHTML = `<div class="text-center py-20 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
      <i class="fas fa-archive text-6xl text-slate-300 dark:text-slate-600 mb-4"></i>
      <h2 class="text-2xl font-semibold text-slate-700 dark:text-slate-200">${t("noResultsTitle")}</h2>
      <p class="text-slate-500 dark:text-slate-400 mt-2">Nenhum dado arquivado encontrado.</p>
    </div>`;
    return;
  }

  // Group historical data by Date
  const groupedByDate = filteredHistory.reduce((acc, row) => {
    const d = toDateMaybe(row["DELIVERY AT BYD"]);
    const key = d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : (String(row["DELIVERY AT BYD"] || "").trim() || t("undefinedDate"));
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {} as Record<string, DeliveryRow[]>);

  const sortedDatesAsc = Object.keys(groupedByDate).sort((a, b) => a.localeCompare(b));

  const formatDateLabel = (dKey: string) => dKey === t("undefinedDate") ? dKey : dKey.split("-").reverse().join("/");

  const getWeekLabel = (dateStr: string) => {
    if (dateStr === t("undefinedDate")) return dateStr;
    const d = new Date(dateStr + "T12:00:00");
    if (isNaN(d.getTime())) return dateStr;
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    const start = new Date(d);
    start.setDate(diff);
    const end = new Date(start);
    end.setDate(end.getDate() + 6);
    return `${start.toLocaleDateString("pt-BR", {day: '2-digit', month: '2-digit'})} a ${end.toLocaleDateString("pt-BR", {day: '2-digit', month: '2-digit'})}`;
  };

  const weekLabels = sortedDatesAsc.map(getWeekLabel);
  const uniqueWeeks = Array.from(new Set(weekLabels));
  
  if (!selectedHistoryWeek || !uniqueWeeks.includes(selectedHistoryWeek)) {
    selectedHistoryWeek = uniqueWeeks[uniqueWeeks.length - 1]; // newest week
  }

  const datesInSelectedWeek = sortedDatesAsc.filter(d => getWeekLabel(d) === selectedHistoryWeek);

  if (!selectedHistoryDate || !datesInSelectedWeek.includes(selectedHistoryDate)) {
    selectedHistoryDate = datesInSelectedWeek[datesInSelectedWeek.length - 1] || sortedDatesAsc[sortedDatesAsc.length - 1];
  }

  const dailyItems = groupedByDate[selectedHistoryDate] || [];
  
  const dailyByCarrier = dailyItems.reduce((acc, row) => {
    const c = String(row["TRANSPORTATION COMPANY"] || "N/A").trim().toUpperCase();
    if (!acc[c]) acc[c] = { programados: 0, entregues: 0, backlog: 0, motivos: new Set<string>() };
    acc[c].programados++;
    const status = normalizeText(row["STATUS"] || "");
    if (status === "ENTREGUE") {
      acc[c].entregues++;
    } else {
      acc[c].backlog++;
      if (row["NOTES"] && String(row["NOTES"]).trim().length > 0) acc[c].motivos.add(String(row["NOTES"]).trim());
    }
    return acc;
  }, {} as Record<string, { programados: number, entregues: number, backlog: number, motivos: Set<string> }>);

  // For weekly, we aggregate over all items in the selected week
  const weeklyItems = datesInSelectedWeek.flatMap(d => groupedByDate[d]);
  
  const weeklyByCarrier = weeklyItems.reduce((acc, row) => {
    const c = String(row["TRANSPORTATION COMPANY"] || "N/A").trim().toUpperCase();
    if (!acc[c]) acc[c] = { programados: 0, entregues: 0, backlog: 0 };
    acc[c].programados++;
    const status = normalizeText(row["STATUS"] || "");
    if (status === "ENTREGUE") {
      acc[c].entregues++;
    } else {
      acc[c].backlog++;
    }
    return acc;
  }, {} as Record<string, { programados: number, entregues: number, backlog: number }>);

  const weeklyTrendData = datesInSelectedWeek.map(dateKey => {
    const items = groupedByDate[dateKey] || [];
    const total = items.length;
    const delivered = items.filter(r => normalizeText(r["STATUS"] || "") === "ENTREGUE").length;
    const perf = total > 0 ? (delivered / total) * 100 : 0;
    return { date: dateKey, perf, total, delivered };
  });

  const dailyCarriers = Object.keys(dailyByCarrier).sort((a, b) => a.localeCompare(b));
  const weeklyCarriers = Object.keys(weeklyByCarrier).sort((a, b) => a.localeCompare(b));

  let dailyTotal = { programados: 0, entregues: 0, backlog: 0 };
  dailyCarriers.forEach(c => {
    dailyTotal.programados += dailyByCarrier[c].programados;
    dailyTotal.entregues += dailyByCarrier[c].entregues;
    dailyTotal.backlog += dailyByCarrier[c].backlog;
  });

  let weeklyTotal = { programados: 0, entregues: 0, backlog: 0 };
  weeklyCarriers.forEach(c => {
    weeklyTotal.programados += weeklyByCarrier[c].programados;
    weeklyTotal.entregues += weeklyByCarrier[c].entregues;
    weeklyTotal.backlog += weeklyByCarrier[c].backlog;
  });

  historyContent.innerHTML = `
    <div class="flex flex-col gap-6 w-full" id="history-report-container">
      <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white dark:bg-slate-800 p-4 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
        <div class="flex items-center gap-4 flex-wrap flex-1">
          <span class="text-sm font-bold text-slate-700 dark:text-slate-200 whitespace-nowrap"><i class="fas fa-calendar-alt mr-2 text-blue-500"></i>FILTROS:</span>
          
          <select id="history-week-select" class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-700 dark:text-slate-200 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block p-2 cursor-pointer font-medium">
            ${uniqueWeeks.map(w => `<option value="${w}" ${w === selectedHistoryWeek ? 'selected' : ''}>Semana: ${w}</option>`).join("")}
          </select>
          
          <span class="text-slate-300 dark:text-slate-600">|</span>
          
          <div class="flex flex-wrap gap-2" id="history-date-tabs">
            ${datesInSelectedWeek.map(dateKey => {
              const isSelected = dateKey === selectedHistoryDate;
              const btnClass = isSelected 
                ? "bg-blue-600 text-white border-blue-600 shadow-sm" 
                : "bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 border-slate-200 dark:border-slate-600 hover:bg-slate-200 dark:hover:bg-slate-600";
              return `<button type="button" class="px-3 py-1.5 text-xs font-bold rounded border transition-colors cursor-pointer history-date-btn ${btnClass}" data-date="${dateKey}">${formatDateLabel(dateKey)}</button>`;
            }).join('')}
          </div>
        </div>
        
        <div class="flex gap-2">
          <button id="delete-history-btn" class="flex-none bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-lg font-bold flex items-center transition-colors shadow-sm text-sm">
            <i class="fas fa-trash-alt mr-2"></i> Excluir Data
          </button>
          <button id="export-history-pdf" class="flex-none bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-bold flex items-center transition-colors shadow-sm text-sm">
            <i class="fas fa-file-pdf mr-2"></i> Exportar PDF
          </button>
        </div>
      </div>

      <!-- NOVOS GRÁFICOS DE PERFORMANCE OPERACIONAL -->
      <div class="grid grid-cols-1 xl:grid-cols-2 gap-6">
        <!-- Gráfico 1: Análise de Performance Operacional Diária -->
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[350px]">
          <div class="bg-slate-700 text-white font-bold text-center py-2 text-sm uppercase">Análise de Performance Operacional Diária</div>
          <div class="p-4 flex-1 relative min-h-[260px]">
             <canvas id="operacionalDiariaChartCanvas"></canvas>
          </div>
        </div>
        
        <!-- Gráfico 2: Análise de Performance Operacional das Transportadoras -->
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[350px]">
          <div class="bg-slate-700 text-white font-bold text-center py-2 text-sm uppercase">Análise de Performance Operacional das Transportadoras</div>
          <div class="p-4 flex-1 relative min-h-[260px]">
             <canvas id="operacionalTransportadorasChartCanvas"></canvas>
          </div>
        </div>
      </div>
      
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[400px]">
          <div class="bg-slate-700 text-white font-bold text-center py-2 text-sm uppercase">DESEMPENHO DIÁRIO POR TRANSPORTADORA</div>
          <div class="overflow-y-auto flex-1 custom-scrollbar">
            <table class="w-full text-xs text-center border-collapse">
              <thead class="bg-slate-100 dark:bg-slate-900 text-slate-800 dark:text-slate-200 sticky top-0 z-10">
                 <tr>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Transportadora</th>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Programados</th>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Entregues</th>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Backlog</th>
                   <th class="py-2 border-b dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Performance</th>
                 </tr>
              </thead>
              <tbody class="text-slate-700 dark:text-slate-300">
                ${dailyCarriers.map(c => {
                  const data = dailyByCarrier[c];
                  const perf = data.programados > 0 ? ((data.entregues / data.programados) * 100).toFixed(1) + "%" : "0.0%";
                  return `
                    <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                      <td class="py-2 border-b border-r dark:border-slate-700 font-bold">${c}</td>
                      <td class="py-2 border-b border-r dark:border-slate-700">${data.programados}</td>
                      <td class="py-2 border-b border-r dark:border-slate-700">${data.entregues}</td>
                      <td class="py-2 border-b border-r dark:border-slate-700">${data.backlog}</td>
                      <td class="py-2 border-b dark:border-slate-700">${perf}</td>
                    </tr>
                  `;
                }).join("")}
                <tr class="bg-slate-100 dark:bg-slate-900 font-bold text-slate-800 dark:text-slate-200">
                  <td class="py-2 border-t dark:border-slate-700 border-r">TOTAL</td>
                  <td class="py-2 border-t dark:border-slate-700 border-r">${dailyTotal.programados}</td>
                  <td class="py-2 border-t dark:border-slate-700 border-r">${dailyTotal.entregues}</td>
                  <td class="py-2 border-t dark:border-slate-700 border-r">${dailyTotal.backlog}</td>
                  <td class="py-2 border-t dark:border-slate-700">${dailyTotal.programados > 0 ? ((dailyTotal.entregues / dailyTotal.programados) * 100).toFixed(1) + "%" : "0.0%"}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[400px]">
          <div class="bg-blue-600 text-white font-bold text-center py-2 text-sm uppercase">PROGRAMADOS X BACKLOG (DIÁRIO)</div>
          <div class="p-4 flex-1 relative min-h-[300px]">
             <canvas id="historyDailyChartCanvas"></canvas>
          </div>
        </div>
      </div>
      
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[400px]">
          <div class="bg-slate-700 text-white font-bold text-center py-2 text-sm uppercase">BACKLOG DO DIA - MOTIVOS</div>
          <div class="overflow-y-auto flex-1 custom-scrollbar">
            <table class="w-full text-xs text-center border-collapse">
              <thead class="bg-slate-100 dark:bg-slate-900 text-slate-800 dark:text-slate-200 sticky top-0 z-10">
                 <tr>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Transportadora</th>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Backlog</th>
                   <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Principal Motivo</th>
                   <th class="py-2 border-b dark:border-slate-700 bg-slate-100 dark:bg-slate-900 w-24">Impacto</th>
                 </tr>
              </thead>
              <tbody class="text-slate-700 dark:text-slate-300">
                ${dailyCarriers.map(c => {
                  const data = dailyByCarrier[c];
                  if (data.backlog === 0) return '';
                  let autoMotivos = Array.from(data.motivos).join("; ");
                  if (!autoMotivos) autoMotivos = "-";
                  const savedNotes = (dailyCarrierNotes[selectedHistoryDate!] || {})[c] || { motivo: '', impacto: '' };
                  return `
                    <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                      <td class="py-2 border-b border-r dark:border-slate-700 font-bold">${c}</td>
                      <td class="py-2 border-b border-r dark:border-slate-700">${data.backlog}</td>
                      <td class="py-1 px-2 border-b border-r dark:border-slate-700 text-left">
                        <input type="text" class="w-full bg-transparent border-0 border-b border-transparent hover:border-slate-300 dark:hover:border-slate-600 focus:border-blue-500 focus:outline-none transition-colors history-note-input px-1 py-0.5 text-xs text-slate-700 dark:text-slate-300 placeholder:text-slate-400 dark:placeholder:text-slate-500" data-date="${selectedHistoryDate}" data-carrier="${c}" data-field="motivo" placeholder="${autoMotivos}" value="${savedNotes.motivo}">
                      </td>
                      <td class="py-1 px-2 border-b dark:border-slate-700 text-center">
                        <input type="text" class="w-full bg-transparent border-0 border-b border-transparent hover:border-slate-300 dark:hover:border-slate-600 focus:border-blue-500 focus:outline-none transition-colors history-note-input px-1 py-0.5 text-xs text-slate-700 dark:text-slate-300 placeholder:text-slate-400 dark:placeholder:text-slate-500 text-center" data-date="${selectedHistoryDate}" data-carrier="${c}" data-field="impacto" placeholder="-" value="${savedNotes.impacto}">
                      </td>
                    </tr>
                  `;
                }).join("")}
                <tr class="bg-slate-100 dark:bg-slate-900 font-bold text-slate-800 dark:text-slate-200">
                  <td class="py-2 border-t dark:border-slate-700 border-r">TOTAL</td>
                  <td class="py-2 border-t dark:border-slate-700 border-r">${dailyTotal.backlog}</td>
                  <td class="py-2 border-t dark:border-slate-700 border-r"></td>
                  <td class="py-2 border-t dark:border-slate-700"></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        
        <div class="flex flex-col gap-6 h-full">
          <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col min-h-[220px]">
            <div class="bg-green-700 text-white font-bold text-center py-2 text-sm uppercase">PERFORMANCE SEMANAL</div>
            <div class="overflow-y-auto flex-1 custom-scrollbar">
               <table class="w-full text-xs text-center border-collapse">
                 <thead class="bg-slate-100 dark:bg-slate-900 text-slate-800 dark:text-slate-200 sticky top-0 z-10">
                   <tr>
                     <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Transportadora</th>
                     <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Programados</th>
                     <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Entregues</th>
                     <th class="py-2 border-b border-r dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Backlog</th>
                     <th class="py-2 border-b dark:border-slate-700 bg-slate-100 dark:bg-slate-900">Performance</th>
                   </tr>
                 </thead>
                 <tbody class="text-slate-700 dark:text-slate-300">
                    ${weeklyCarriers.map(c => {
                      const data = weeklyByCarrier[c];
                      const perf = data.programados > 0 ? ((data.entregues / data.programados) * 100).toFixed(1) + "%" : "0.0%";
                      return `
                        <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                          <td class="py-1 border-b border-r dark:border-slate-700 font-bold">${c}</td>
                          <td class="py-1 border-b border-r dark:border-slate-700">${data.programados}</td>
                          <td class="py-1 border-b border-r dark:border-slate-700">${data.entregues}</td>
                          <td class="py-1 border-b border-r dark:border-slate-700">${data.backlog}</td>
                          <td class="py-1 border-b dark:border-slate-700">${perf}</td>
                        </tr>
                      `;
                    }).join("")}
                    <tr class="bg-slate-100 dark:bg-slate-900 font-bold text-slate-800 dark:text-slate-200">
                      <td class="py-2 border-t dark:border-slate-700 border-r">TOTAL</td>
                      <td class="py-2 border-t dark:border-slate-700 border-r">${weeklyTotal.programados}</td>
                      <td class="py-2 border-t dark:border-slate-700 border-r">${weeklyTotal.entregues}</td>
                      <td class="py-2 border-t dark:border-slate-700 border-r">${weeklyTotal.backlog}</td>
                      <td class="py-2 border-t dark:border-slate-700">${weeklyTotal.programados > 0 ? ((weeklyTotal.entregues / weeklyTotal.programados) * 100).toFixed(1) + "%" : "0.0%"}</td>
                    </tr>
                 </tbody>
               </table>
            </div>
          </div>
          
          <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col flex-1 min-h-[160px]">
            <div class="bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-200 font-bold text-center py-2 text-sm uppercase border-b dark:border-slate-700">PERFORMANCE SEMANA GERAL</div>
            <div class="p-2 relative flex-1 min-h-[120px]">
               <canvas id="historyWeeklyChartCanvas"></canvas>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;

  document.getElementById("history-week-select")?.addEventListener("change", (e) => {
    selectedHistoryWeek = (e.target as HTMLSelectElement).value;
    // reset selected date when week changes
    selectedHistoryDate = null;
    renderHistoryTab();
  });

  document.getElementById("delete-history-btn")?.addEventListener("click", () => {
    if (!selectedHistoryDate) return;
    
    const container = document.getElementById("history-report-container");
    if (!container) return;
    
    container.innerHTML = `
      <div class="bg-white dark:bg-slate-800 p-6 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 max-w-md mx-auto text-center mt-10">
        <h3 class="text-lg font-bold text-slate-800 dark:text-slate-200 mb-4"><i class="fas fa-exclamation-triangle text-amber-500 mr-2"></i> Excluir Dados Históricos</h3>
        <p class="text-sm text-slate-600 dark:text-slate-400 mb-4">Por favor, insira a senha de exclusão para confirmar a remoção da data <strong>${formatDateLabel(selectedHistoryDate)}</strong>.</p>
        <input type="password" id="delete-pwd-input" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-700 dark:text-slate-200 rounded-lg p-2 mb-4 focus:ring-blue-500 focus:border-blue-500" placeholder="Senha" />
        <div class="flex justify-center gap-2">
          <button id="cancel-delete-btn" class="bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-lg font-bold transition-colors">Cancelar</button>
          <button id="confirm-delete-btn" class="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-bold transition-colors">Confirmar Exclusão</button>
        </div>
      </div>
    `;
    
    document.getElementById("cancel-delete-btn")?.addEventListener("click", () => renderHistoryTab());
    
    document.getElementById("confirm-delete-btn")?.addEventListener("click", () => {
      const pwd = (document.getElementById("delete-pwd-input") as HTMLInputElement).value;
      if (pwd === "Byd@N1") {
        historicalData = historicalData.filter(row => {
          const d = toDateMaybe(row["DELIVERY AT BYD"]);
          const key = d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : (String(row["DELIVERY AT BYD"] || "").trim() || t("undefinedDate"));
          return key !== selectedHistoryDate;
        });
        if (dailyCarrierNotes[selectedHistoryDate!]) {
          delete dailyCarrierNotes[selectedHistoryDate!];
        }
        saveStateToFirebase();
        selectedHistoryDate = null;
        renderHistoryTab();
        showToast("Data excluída com sucesso.", "success");
      } else {
        showToast("Senha incorreta.", "error");
      }
    });
  });

  document.getElementById("export-history-pdf")?.addEventListener("click", () => {
    try {
      const doc = new (jspdf as any).jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
      
      doc.setFontSize(16);
      doc.text(`Relatório Histórico - Semana: ${selectedHistoryWeek}`, 14, 15);
      
      let currentY = 25;
      
      const ctxDaily = document.getElementById("historyDailyChartCanvas") as HTMLCanvasElement;
      if (ctxDaily) {
        doc.setFontSize(12);
        doc.text("Desempenho Diário por Transportadora", 14, currentY);
        const imgData = ctxDaily.toDataURL("image/png");
        doc.addImage(imgData, "PNG", 14, currentY + 5, 130, 65);
      }
      
      const ctxWeekly = document.getElementById("historyWeeklyChartCanvas") as HTMLCanvasElement;
      if (ctxWeekly) {
        doc.setFontSize(12);
        doc.text("Performance Semanal Geral", 150, currentY);
        const imgDataWeekly = ctxWeekly.toDataURL("image/png");
        doc.addImage(imgDataWeekly, "PNG", 150, currentY + 5, 130, 65);
      }
      
      currentY += 80;
      
      // AutoTable for daily performance
      doc.setFontSize(12);
      doc.text(`Transportadoras Diário (${formatDateLabel(selectedHistoryDate!)})`, 14, currentY);
      currentY += 5;
      
      const dailyTableData = dailyCarriers.map(c => [
        c,
        dailyByCarrier[c].programados,
        dailyByCarrier[c].entregues,
        dailyByCarrier[c].backlog,
        dailyByCarrier[c].programados > 0 ? ((dailyByCarrier[c].entregues / dailyByCarrier[c].programados) * 100).toFixed(1) + "%" : "0.0%"
      ]);
      dailyTableData.push(["TOTAL", dailyTotal.programados, dailyTotal.entregues, dailyTotal.backlog, dailyTotal.programados > 0 ? ((dailyTotal.entregues / dailyTotal.programados) * 100).toFixed(1) + "%" : "0.0%"]);
      
      (doc as any).autoTable({
        startY: currentY,
        head: [["Transportadora", "Programados", "Entregues", "Backlog", "Performance"]],
        body: dailyTableData,
        theme: "grid",
        styles: { fontSize: 8 },
        headStyles: { fillColor: [51, 65, 85] }
      });
      
      currentY = (doc as any).lastAutoTable.finalY + 15;
      
      // AutoTable for weekly performance
      doc.text("Transportadoras Semanal", 14, currentY);
      currentY += 5;
      
      const weeklyTableData = weeklyCarriers.map(c => [
        c,
        weeklyByCarrier[c].programados,
        weeklyByCarrier[c].entregues,
        weeklyByCarrier[c].backlog,
        weeklyByCarrier[c].programados > 0 ? ((weeklyByCarrier[c].entregues / weeklyByCarrier[c].programados) * 100).toFixed(1) + "%" : "0.0%"
      ]);
      weeklyTableData.push(["TOTAL", weeklyTotal.programados, weeklyTotal.entregues, weeklyTotal.backlog, weeklyTotal.programados > 0 ? ((weeklyTotal.entregues / weeklyTotal.programados) * 100).toFixed(1) + "%" : "0.0%"]);
      
      (doc as any).autoTable({
        startY: currentY,
        head: [["Transportadora", "Programados", "Entregues", "Backlog", "Performance"]],
        body: weeklyTableData,
        theme: "grid",
        styles: { fontSize: 8 },
        headStyles: { fillColor: [21, 128, 61] } // green-700
      });
      
      doc.save(`Historico_${selectedHistoryWeek?.replace(/ /g, "_")}.pdf`);
      showToast(t("pdfGenerated"), "success");
    } catch (err) {
      console.error("PDF generation error:", err);
      showToast("Erro ao gerar PDF", "error");
    }
  });

  document.querySelectorAll(".history-date-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      selectedHistoryDate = (e.currentTarget as HTMLElement).dataset.date || null;
      renderHistoryTab();
    });
  });

  document.querySelectorAll(".history-note-input").forEach(input => {
    input.addEventListener("change", (e) => {
      const target = e.target as HTMLInputElement;
      const date = target.dataset.date;
      const carrier = target.dataset.carrier;
      const field = target.dataset.field; // "motivo" or "impacto"
      const value = target.value;
      
      if (date && carrier && field) {
        if (!dailyCarrierNotes[date]) dailyCarrierNotes[date] = {};
        if (!dailyCarrierNotes[date][carrier]) dailyCarrierNotes[date][carrier] = { motivo: "", impacto: "" };
        
        if (field === "motivo") {
          dailyCarrierNotes[date][carrier].motivo = value;
        } else if (field === "impacto") {
          dailyCarrierNotes[date][carrier].impacto = value;
        }
        
        saveStateToFirebase();
      }
    });
  });

  if (isTypingInHistory) {
    const inputToFocus = document.querySelector(`.history-note-input[data-date="${activeData.date}"][data-carrier="${activeData.carrier}"][data-field="${activeData.field}"]`) as HTMLInputElement;
    if (inputToFocus) {
      inputToFocus.focus();
      inputToFocus.value = activeData.value;
      // move cursor to end
      const len = inputToFocus.value.length;
      inputToFocus.setSelectionRange(len, len);
    }
  }

  // Render Charts
  const ctxDaily = document.getElementById("historyDailyChartCanvas") as HTMLCanvasElement;
  if (ctxDaily) {
    if (historyDailyChart) historyDailyChart.destroy();
    historyDailyChart = new Chart(ctxDaily, {
      type: "bar",
      data: {
        labels: dailyCarriers,
        datasets: [
          {
            label: "Entregues",
            data: dailyCarriers.map(c => dailyByCarrier[c].entregues),
            backgroundColor: "#2563eb",
            barPercentage: 0.8,
            categoryPercentage: 0.8,
          },
          {
            label: "Adiado (Backlog)",
            data: dailyCarriers.map(c => dailyByCarrier[c].backlog),
            backgroundColor: "#dc2626",
            barPercentage: 0.8,
            categoryPercentage: 0.8,
          }
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "right", labels: { boxWidth: 12, font: { size: 10 } } },
          tooltip: { mode: "index", intersect: false },
          datalabels: {
            color: '#fff',
            anchor: 'end',
            align: 'start',
            font: { size: 10, weight: 'bold' },
            formatter: (val) => val > 0 ? val : ""
          }
        },
        scales: {
          x: { stacked: false, beginAtZero: true, grid: { color: "rgba(0,0,0,0.05)" } },
          y: { stacked: false, grid: { display: false }, ticks: { font: { size: 10, weight: 'bold' } } }
        }
      },
      plugins: [ChartDataLabels]
    });
  }

  const ctxWeekly = document.getElementById("historyWeeklyChartCanvas") as HTMLCanvasElement;
  if (ctxWeekly) {
    if (historyWeeklyChart) historyWeeklyChart.destroy();
    historyWeeklyChart = new Chart(ctxWeekly, {
      type: "line",
      data: {
        labels: weeklyTrendData.map(d => formatDateLabel(d.date)),
        datasets: [
          {
            label: "Performance %",
            data: weeklyTrendData.map(d => d.perf),
            borderColor: "#dc2626",
            backgroundColor: "#dc2626",
            borderWidth: 2,
            tension: 0,
            pointBackgroundColor: "#2563eb",
            pointBorderColor: "#fff",
            pointRadius: 4,
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.parsed.y.toFixed(1)}%`
            }
          },
          datalabels: {
            align: 'top',
            color: '#334155',
            font: { size: 10, weight: 'bold' },
            formatter: (val) => val.toFixed(1) + "%"
          }
        },
        scales: {
          y: { beginAtZero: true, max: 100, grid: { color: "rgba(0,0,0,0.05)" }, ticks: { stepSize: 20, callback: (v) => v + "%" } },
          x: { grid: { display: false }, ticks: { font: { size: 10 } } }
        }
      },
      plugins: [ChartDataLabels]
    });
  }

  // Render Operational Charts (Análise de Performance Operacional Diária & Transportadoras)
  const formatToDayMonth = (dateStr: string) => {
    if (dateStr === t("undefinedDate")) return dateStr;
    const d = new Date(dateStr + "T12:00:00");
    if (isNaN(d.getTime())) return dateStr;
    const day = String(d.getDate()).padStart(2, "0");
    const months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
    return `${day}/${months[d.getMonth()]}`;
  };

  const isDark = document.documentElement.classList.contains("dark");
  const textColor = isDark ? "#cbd5e1" : "#334155";
  const gridColor = isDark ? "rgba(255,255,255,0.08)" : "rgba(0,0,0,0.05)";

  // 1. Análise de Performance Operacional Diária
  const ctxOperacionalDiaria = document.getElementById("operacionalDiariaChartCanvas") as HTMLCanvasElement;
  if (ctxOperacionalDiaria) {
    if (operacionalDiariaChart) operacionalDiariaChart.destroy();

    const dailyDataPoints = datesInSelectedWeek.map(dateKey => {
      const items = groupedByDate[dateKey] || [];
      const total = items.length;
      const delivered = items.filter(r => normalizeText(r["STATUS"] || "") === "ENTREGUE").length;
      const backlog = total - delivered;
      const perf = total > 0 ? (delivered / total) * 100 : 0;
      return {
        label: formatToDayMonth(dateKey),
        entregues: delivered,
        backlog: backlog,
        performance: perf
      };
    });

    let sumEntregues = 0;
    let sumBacklog = 0;
    let sumPerformance = 0;
    const daysCount = dailyDataPoints.length;
    dailyDataPoints.forEach(pt => {
      sumEntregues += pt.entregues;
      sumBacklog += pt.backlog;
      sumPerformance += pt.performance;
    });

    const avgEntregues = daysCount > 0 ? Math.round((sumEntregues / daysCount) * 10) / 10 : 0;
    const avgBacklog = daysCount > 0 ? Math.round((sumBacklog / daysCount) * 10) / 10 : 0;
    const avgPerformance = daysCount > 0 ? Math.round((sumPerformance / daysCount) * 10) / 10 : 0;

    const labelsDiaria = dailyDataPoints.map(pt => pt.label).concat(["Average"]);
    const entreguesDiaria = dailyDataPoints.map(pt => pt.entregues).concat([avgEntregues]);
    const backlogDiaria = dailyDataPoints.map(pt => pt.backlog).concat([avgBacklog]);
    const performanceDiaria = dailyDataPoints.map(pt => pt.performance).concat([avgPerformance]);

    let maxTotalVal = 300;
    dailyDataPoints.forEach(pt => {
      const total = pt.entregues + pt.backlog;
      if (total > maxTotalVal) maxTotalVal = total;
    });
    const computedYMax = Math.ceil((maxTotalVal + 50) / 50) * 50;

    operacionalDiariaChart = new Chart(ctxOperacionalDiaria, {
      type: "bar",
      data: {
        labels: labelsDiaria,
        datasets: [
          {
            type: "bar",
            label: "Entregues",
            data: entreguesDiaria,
            backgroundColor: "#1e3a8a", // Dark blue
            stack: "stack1",
            barPercentage: 0.55,
            categoryPercentage: 0.8,
            order: 2,
            datalabels: {
              color: "#ffffff",
              anchor: "center",
              align: "center",
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val : ""
            }
          },
          {
            type: "bar",
            label: "Backlog",
            data: backlogDiaria,
            backgroundColor: "#fdba74", // Soft orange
            stack: "stack1",
            barPercentage: 0.55,
            categoryPercentage: 0.8,
            order: 3,
            datalabels: {
              color: "#7c2d12", // Dark orange text
              anchor: "center",
              align: "center",
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val : ""
            }
          },
          {
            type: "line",
            label: "Performance Dia",
            data: performanceDiaria,
            borderColor: "#94a3b8", // Slate-400
            backgroundColor: "#94a3b8",
            borderWidth: 2,
            tension: 0.1,
            yAxisID: "y1",
            order: 1,
            pointBackgroundColor: "#475569",
            pointBorderColor: "#ffffff",
            pointRadius: 5,
            pointHoverRadius: 7,
            datalabels: {
              color: textColor,
              anchor: "end",
              align: "top",
              offset: 8,
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val.toFixed(1) + "%" : "0.0%"
            }
          },
          {
            type: "line",
            label: "Meta 300 Ctnr",
            data: labelsDiaria.map(() => 300),
            borderColor: "#ef4444", // Red
            borderWidth: 2,
            borderDash: [6, 6],
            fill: false,
            pointRadius: 0,
            pointHoverRadius: 0,
            order: 4,
            datalabels: {
              display: false
            }
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              boxWidth: 12,
              font: { size: 10, weight: "bold" },
              color: textColor
            }
          },
          tooltip: {
            mode: "index",
            intersect: false
          }
        },
        scales: {
          x: {
            stacked: true,
            grid: { display: false },
            ticks: { font: { size: 10, weight: "bold" }, color: textColor }
          },
          y: {
            stacked: true,
            beginAtZero: true,
            max: computedYMax,
            grid: { color: gridColor },
            ticks: { font: { size: 10 }, color: textColor }
          },
          y1: {
            position: "right",
            beginAtZero: true,
            max: 120,
            grid: { drawOnChartArea: false },
            ticks: {
              stepSize: 20,
              font: { size: 10 },
              color: textColor,
              callback: (val: any) => val + "%"
            }
          }
        }
      },
      plugins: [ChartDataLabels]
    });
  }

  // 2. Análise de Performance Operacional das Transportadoras
  const ctxOperacionalTransportadoras = document.getElementById("operacionalTransportadorasChartCanvas") as HTMLCanvasElement;
  if (ctxOperacionalTransportadoras) {
    if (operacionalTransportadorasChart) operacionalTransportadorasChart.destroy();

    const labelsCarrier = weeklyCarriers;
    const entreguesCarrier = weeklyCarriers.map(c => weeklyByCarrier[c].entregues);
    const backlogCarrier = weeklyCarriers.map(c => weeklyByCarrier[c].backlog);
    const performanceCarrier = weeklyCarriers.map(c => {
      const data = weeklyByCarrier[c];
      const total = data.entregues + data.backlog;
      return total > 0 ? (data.entregues / total) * 100 : 0;
    });

    operacionalTransportadorasChart = new Chart(ctxOperacionalTransportadoras, {
      type: "bar",
      data: {
        labels: labelsCarrier,
        datasets: [
          {
            type: "bar",
            label: "Entregues",
            data: entreguesCarrier,
            backgroundColor: "#15803d", // Green
            barPercentage: 0.6,
            categoryPercentage: 0.6,
            order: 2,
            datalabels: {
              color: "#ffffff",
              anchor: "end",
              align: "top",
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val : ""
            }
          },
          {
            type: "bar",
            label: "Backlog",
            data: backlogCarrier,
            backgroundColor: "#6b7280", // Gray
            barPercentage: 0.6,
            categoryPercentage: 0.6,
            order: 3,
            datalabels: {
              color: "#ffffff",
              anchor: "end",
              align: "top",
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val : ""
            }
          },
          {
            type: "line",
            label: "Performance",
            data: performanceCarrier,
            borderColor: "#94a3b8", // Slate-400
            backgroundColor: "#94a3b8",
            borderWidth: 2,
            tension: 0.1,
            yAxisID: "y1",
            order: 1,
            pointBackgroundColor: "#475569",
            pointBorderColor: "#ffffff",
            pointRadius: 5,
            pointHoverRadius: 7,
            datalabels: {
              color: textColor,
              anchor: "end",
              align: "top",
              offset: 8,
              font: { size: 10, weight: "bold" },
              formatter: (val: any) => val > 0 ? val.toFixed(1) + "%" : "0.0%"
            }
          },
          {
            type: "line",
            label: "Meta",
            data: labelsCarrier.map(() => 100),
            borderColor: "#1d4ed8", // Blue
            borderWidth: 2,
            borderDash: [6, 6],
            fill: false,
            pointRadius: 0,
            pointHoverRadius: 0,
            order: 4,
            datalabels: {
              display: false
            }
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              boxWidth: 12,
              font: { size: 10, weight: "bold" },
              color: textColor
            }
          },
          tooltip: {
            mode: "index",
            intersect: false
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: 10, weight: "bold" }, color: textColor }
          },
          y: {
            beginAtZero: true,
            grid: { color: gridColor },
            ticks: { font: { size: 10 }, color: textColor }
          },
          y1: {
            position: "right",
            beginAtZero: true,
            max: 120,
            grid: { drawOnChartArea: false },
            ticks: {
              stepSize: 20,
              font: { size: 10 },
              color: textColor,
              callback: (val: any) => val + "%"
            }
          }
        }
      },
      plugins: [ChartDataLabels]
    });
  }
}


function renderTimeTable(data: DeliveryRow[]) {
  const timeContent = document.getElementById("time-content");
  if (!timeContent) return;

  let totalTimeSum = 0, validRecords = 0;
  let totalTimeSumP1 = 0, validRecordsP1 = 0;
  let totalTimeSumP2 = 0, validRecordsP2 = 0;
  
  let desovaTotal = 0, desova1 = 0, desova2 = 0, desovaCross = 0;
  let baixaTotal = 0, baixa1 = 0, baixa2 = 0;
  let desovaTotalScheduled = 0, baixaTotalScheduled = 0;

  const shiftLists: Record<string, any[]> = {
    desova1: [],
    desova2: [],
    desovaCross: [],
    baixa1: [],
    baixa2: []
  };

  const rowsHtml = data.map((row) => {
    const op = String(row["OPERATION SCOPE"] || "").trim().toUpperCase();
    const status = normalizeText(row["STATUS"] || "");
    const isBaixa = op.includes("SWAP") || op.includes("PUT DOWN") || op.includes("PUTDOWN") || op.includes("BAIXA") || op.includes("PISO");
    const isDesova = !isBaixa; // All non-Baixa containers default to DESOVAS

    if (status !== "CANCELADO") {
      if (isBaixa) baixaTotalScheduled++;
      if (isDesova) desovaTotalScheduled++;
    }

    const startDt = toDateTimeMaybe(row["TERMINAL - INÍCIO DE ROTA"]);
    let endDt = toDateTimeMaybe(row["ENTREGA VAZIO"]) || toDateTimeMaybe(row["DATA E HORARIO DE DESCARGA"]);
    let fullTimeString = "-";
    let durationHours = 0;

    const isBaixaCompleted = isBaixa && startDt && status === "ENTREGUE";
    const isDesovaCompleted = isDesova && startDt && endDt;

    if (isBaixaCompleted) {
      baixaTotal++;
      const sTime = startDt.getHours() * 100 + startDt.getMinutes();
      const s1 = (sTime < 1500);
      if (s1) { baixa1++; shiftLists.baixa1.push(row); }
      else { baixa2++; shiftLists.baixa2.push(row); }
    }

    if (isDesovaCompleted) {
      desovaTotal++;
      const sTime = startDt.getHours() * 100 + startDt.getMinutes();
      const eTime = endDt.getHours() * 100 + endDt.getMinutes();
      const s1 = (sTime < 1500);
      const e1 = (eTime < 1500);
      
      if (s1 && e1) { desova1++; shiftLists.desova1.push(row); }
      else if (!s1 && !e1) { desova2++; shiftLists.desova2.push(row); }
      else if (s1 && !e1) { desovaCross++; shiftLists.desovaCross.push(row); }
      else { desova2++; shiftLists.desova2.push(row); } // Fallback for 2nd->1st
    }

    if (startDt && endDt) {
      const diffMs = endDt.getTime() - startDt.getTime();
      if (diffMs > -3600000) { // Small threshold for slight negative values due to clock drift
        durationHours = Math.max(0, diffMs / (1000 * 60 * 60));
        const dDays = Math.floor(durationHours / 24);
        const dHours = Math.floor(durationHours % 24);
        const dMins = Math.round((durationHours - Math.floor(durationHours)) * 60);
        fullTimeString = dDays > 0 ? `${dDays}v ${dHours}h ${dMins}m` : `${dHours}h ${dMins}m`;

        totalTimeSum += durationHours; validRecords++;
        const timeVal = startDt.getHours() * 100 + startDt.getMinutes();
        if (timeVal < 1500) { 
          totalTimeSumP1 += durationHours; 
          validRecordsP1++; 
        } else { 
          totalTimeSumP2 += durationHours; 
          validRecordsP2++; 
        }
      }
    }

    return `<tr class="hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors">
        <td class="px-4 py-3 font-medium text-slate-800 dark:text-slate-100">${row["CONTAINER"] || "-"}</td>
        <td class="px-4 py-3 text-slate-600 dark:text-slate-300 font-mono">${row["BL"] || "-"}</td>
        <td class="px-4 py-3 text-slate-600 dark:text-slate-300">${row["TRANSPORTATION COMPANY"] || "-"}</td>
        <td class="px-4 py-3 text-slate-600 dark:text-slate-300">${row["LOT"] || "-"}</td>
        <td class="px-4 py-3 text-slate-500 dark:text-slate-400">${startDt ? startDt.toLocaleString() : "-"}</td>
        <td class="px-4 py-3 text-slate-500 dark:text-slate-400">${endDt ? endDt.toLocaleString() : "-"}</td>
        <td class="px-4 py-3 font-bold text-blue-600 dark:text-blue-400">${fullTimeString}</td>
      </tr>`;
  }).join("");

  const totalCompletedOps = desovaTotal + baixaTotal;

  const scheduled = deliveryData.filter(d => normalizeText(d["STATUS"] || "") !== "CANCELADO").length;
  const delivered = deliveryData.filter(d => normalizeText(d["STATUS"] || "") === "ENTREGUE").length;
  
  const kpiEntregues = delivered;
  const deltaApoio = kpiEntregues - totalCompletedOps;
  const taxaAbsorcao = kpiEntregues > 0 ? (totalCompletedOps / kpiEntregues * 100).toFixed(1) : 0;
  const remaining = deliveryData.filter(d => {
    const s = normalizeText(d["STATUS"] || "");
    return s === "PENDENTE" || s === "AGUARDANDO DESOVA" || s === "A CAMINHO";
  }).length;

  const avgHours = validRecords > 0 ? totalTimeSum / validRecords : 0;
  const hoursNeeded = remaining * avgHours;

  const formatDuration = (h: number) => {
    const hr = Math.floor(h), m = Math.round((h - Math.floor(h)) * 60);
    return `${hr}h ${m}m`;
  };

  const now = new Date();
  // Assume a standard operational shift ending at 22:00
  const shiftEnd = new Date(now);
  shiftEnd.setHours(22, 0, 0, 0);

  // If it's already past shift end, assume it's for 22:00 tomorrow or just for the current window
  let remainingMs = shiftEnd.getTime() - now.getTime();
  if (remainingMs < 0) remainingMs = 0; // Or handle next day shift

  const hoursRemainingInShift = remainingMs / (1000 * 60 * 60);

  // Standard starting time for shift is 06:30
  const shiftStart = new Date(now);
  shiftStart.setHours(6, 30, 0, 0);

  let elapsedHrs = (now.getTime() - shiftStart.getTime()) / (1000 * 60 * 60);
  if (elapsedHrs <= 0) {
    elapsedHrs = 0.1; // Safety fallback
  }
  if (elapsedHrs > 15.5) {
    elapsedHrs = 15.5; // Max shift is 15.5h (06:30 to 22:00)
  }

  // Throughput Calculations (containers/hour)
  const currentThroughput = delivered / elapsedHrs;
  const requiredThroughput = hoursRemainingInShift > 0 ? remaining / hoursRemainingInShift : 0;
  const pctIncrease = currentThroughput > 0 ? ((requiredThroughput / currentThroughput) - 1) * 100 : 0;

  const projectedAdditional = currentThroughput * hoursRemainingInShift;
  const estimatedRemainingBacklog = Math.max(0, remaining - projectedAdditional);

  const pctIncreaseFormatted = pctIncrease > 100 
    ? `mais do que duplicar (+${pctIncrease.toFixed(0)}%)` 
    : pctIncrease > 0 
      ? `aumentar em +${pctIncrease.toFixed(0)}%`
      : `manter (ritmo atual com folga de ${Math.abs(pctIncrease).toFixed(0)}%)`;

  const alertOrSuccess = pctIncrease > 0
    ? `⚠️ ALERTA DE CAPACIDADE: O time precisaria ${pctIncreaseFormatted} a velocidade de escoamento atual para cumprir o plano de hoje. Mantendo o ritmo de ${currentThroughput.toFixed(1)} cont./h, a projeção é entregar apenas ~${Math.round(projectedAdditional)} unidades, gerando um Backlog estimado de ${Math.ceil(estimatedRemainingBacklog)} containers para amanhã.`
    : `✅ DESEMPENHO SEGURO: A operação segue dentro do ritmo planejado com folga de capacidade. Mantendo o ritmo de ${currentThroughput.toFixed(1)} cont./h, a projeção é entregar com tranquilidade os ${remaining} restantes, sem gerar backlog para amanhã.`;

  // Takt Time Logic
  const activeFronts = 15; // Calculation base for active teams/fronts
  const targetTaktTimeHrs = (hoursRemainingInShift > 0 && remaining > 0) ? (hoursRemainingInShift * activeFronts) / remaining : 0;
  const targetTaktMins = Math.round(targetTaktTimeHrs * 60);
  const currentAvgMins = Math.round(avgHours * 60);
  const deviationMins = currentAvgMins - targetTaktMins;

  (window as any).__SHIFT_LISTS__ = shiftLists;

  const pontoApoioQtd = (window as any).pontoApoioQtd || 0;
  const transitTime = 0.5; // 30 mins
  const consumoTransito = currentThroughput * transitTime;
  const taxaAbsorcaoNum = parseFloat(String(taxaAbsorcao));
  
  const isPontoApoioHold = taxaAbsorcaoNum > 0 && taxaAbsorcaoNum <= 60;
  const loteSugerido = isPontoApoioHold ? 0 : Math.min(pontoApoioQtd, Math.ceil(currentThroughput * 1.0));
  
  const saldoProjetado = pontoApoioQtd - loteSugerido;
  const isPontoApoioWarning = !isPontoApoioHold && pontoApoioQtd < consumoTransito;
  const isPontoApoioHealthy = !isPontoApoioHold && !isPontoApoioWarning && pontoApoioQtd >= loteSugerido;

  let stateColor = 'slate';
  let stateIcon = 'fa-info-circle';
  let stateTitle = 'Sugestão de Liberação Agora:';
  let stateMessage = `Liberar ${loteSugerido} contêineres`;
  let stateExtra = '';

  if (isPontoApoioHold) {
    stateColor = 'red';
    stateIcon = 'fa-hand-paper';
    stateTitle = 'AÇÃO NECESSÁRIA:';
    stateMessage = 'SEGURAR NO PONTO DE APOIO';
    stateExtra = `<div class="mt-2 text-[10px] font-bold text-red-400 bg-red-400/10 p-1.5 rounded border border-red-400/20">
      Baixa velocidade de drenagem no terminal (${taxaAbsorcao}%). Aguarde a absorção do fluxo atual antes de liberar novas carretas.
    </div>`;
  } else if (isPontoApoioWarning) {
    stateColor = 'yellow';
    stateIcon = 'fa-exclamation-triangle';
    stateExtra = `<div class="mt-2 text-[10px] font-bold text-yellow-400 bg-yellow-400/10 p-1.5 rounded border border-yellow-400/20">
      Risco de parada na BYD em 30 min por falta de carretas
    </div>`;
  } else if (isPontoApoioHealthy) {
    stateColor = 'emerald';
    stateIcon = 'fa-check-circle';
  }

  timeContent.innerHTML = `
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
      <div class="lg:col-span-2 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-5">
        <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-4 uppercase tracking-wider flex items-center">
          <div class="relative group cursor-help flex items-center mr-2">
            <span class="relative flex h-2.5 w-2.5">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-blue-500"></span>
            </span>
            <!-- Tooltip Popup -->
            <div class="pointer-events-none absolute bottom-full left-0 mb-2 w-80 origin-bottom-left scale-0 transition-all group-hover:scale-100 z-50 bg-slate-900 dark:bg-slate-950 border border-slate-700 text-slate-200 rounded-lg p-3 shadow-xl text-[10px] normal-case tracking-normal">
              <div class="font-bold border-b border-slate-700 pb-1 mb-2 uppercase text-blue-400 text-[10px] tracking-wider flex items-center gap-1.5">
                <svg class="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                  <path stroke-linecap="round" stroke-linejoin="round" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 11h.01M12 14h.01M12 17h.01M15 11h.01M15 14h.01M15 17h.01M18 11h.01M18 14h.01M18 17h.01" />
                </svg>
                Fórmulas e Métricas Operacionais
              </div>
              <div class="space-y-2 text-slate-300">
                <div>
                  <strong class="text-blue-300">📊 Vazão Média Atual:</strong>
                  <div class="font-mono bg-slate-800/80 px-1 py-0.5 rounded text-[9px] mt-0.5 select-all">Vazão = Entregues / Horas Decorridas desde 06:30</div>
                </div>
                <div>
                  <strong class="text-blue-300">🎯 Vazão Meta Necessária:</strong>
                  <div class="font-mono bg-slate-800/80 px-1 py-0.5 rounded text-[9px] mt-0.5 select-all">Meta = Restantes / Relógio Restante até 22:00</div>
                </div>
                <div>
                  <strong class="text-blue-300">⏳ Meta Takt Time (15 frentes):</strong>
                  <div class="font-mono bg-slate-800/80 px-1 py-0.5 rounded text-[9px] mt-0.5 select-all">Takt = (Horas Shift Rest. * 15 Frentes) / Restantes</div>
                </div>
                <div>
                  <strong class="text-blue-300">🔮 Previsão de Backlog:</strong>
                  <div class="font-mono bg-slate-800/80 px-1 py-0.5 rounded text-[9px] mt-0.5 select-all">Backlog = Restantes - (Vazão Atual * Horas Shift Rest.)</div>
                </div>
              </div>
            </div>
          </div>
          Capacidade Operacional (Análise de Vazão e Tempo)
        </h3>
        <div class="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-6 gap-3">
          <div class="p-3 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-100 dark:border-slate-600">
            <span class="text-[10px] font-bold text-slate-500 block uppercase mb-1">Restantes</span>
            <span class="text-xl font-black text-slate-800 dark:text-slate-100">${remaining}</span>
          </div>
          <div class="p-3 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-100 dark:border-blue-800/30">
            <span class="text-[10px] font-bold text-blue-600 block uppercase mb-1">Relógio Rest.</span>
            <span class="text-xl font-black text-blue-700 dark:text-blue-300">${hoursRemainingInShift.toFixed(1)}h</span>
          </div>
          <div class="p-3 bg-emerald-50 dark:bg-emerald-900/20 rounded-lg border border-emerald-100 dark:border-emerald-800/30 font-mono">
            <span class="text-[10px] font-bold text-emerald-600 block uppercase mb-1">Vazão Atual</span>
            <span class="text-lg font-black text-emerald-700 dark:text-emerald-300">${currentThroughput.toFixed(1)} <span class="text-[10px]">c/h</span></span>
          </div>
          <div class="p-3 bg-orange-50 dark:bg-orange-900/20 rounded-lg border border-orange-100 dark:border-orange-800/30 font-mono">
            <span class="text-[10px] font-bold text-orange-600 block uppercase mb-1">Vazão Meta</span>
            <span class="text-lg font-black text-orange-700 dark:text-orange-300">${requiredThroughput.toFixed(1)} <span class="text-[10px]">c/h</span></span>
          </div>
          <div class="p-3 bg-teal-50 dark:bg-teal-900/20 rounded-lg border border-teal-100 dark:border-teal-800/30 font-mono">
            <span class="text-[10px] font-bold text-teal-600 block uppercase mb-1">Meta Takt (15F)</span>
            <span class="text-lg font-black text-teal-700 dark:text-teal-300">${targetTaktMins} min</span>
          </div>
          <div class="p-3 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-100 dark:border-slate-600 font-mono">
            <span class="text-[10px] font-bold text-slate-500 block uppercase mb-1">H Trabalho Tot.</span>
            <span class="text-lg font-black text-slate-700 dark:text-slate-300">${remaining > 0 ? hoursNeeded.toFixed(1) + 'h' : '0h'}</span>
          </div>
        </div>
        
        <div class="mt-5 pt-4 border-t border-slate-100 dark:border-slate-700 space-y-3">
          <p class="text-xs text-slate-700 dark:text-slate-300 leading-relaxed">
            <strong class="text-slate-900 dark:text-white uppercase text-[10px] bg-slate-100 dark:bg-slate-700 px-1.5 py-0.5 rounded mr-1">Veredito Operacional:</strong> 
            Desde o início do turno às <span class="font-bold text-slate-800 dark:text-slate-200 text-xs">06:30</span>, a equipe mantém uma média de <span class="font-bold text-slate-800 dark:text-slate-200 text-xs underline">${currentThroughput.toFixed(1)}</span> containers concluídos por hora. Nas próximas <span class="font-bold underline text-xs">${hoursRemainingInShift.toFixed(1)}h</span> restantes, para liquidar os <span class="font-bold text-blue-600 dark:text-blue-400 text-xs">${remaining}</span> containers pendentes, o ritmo da operação precisaria subir para <span class="font-bold text-orange-600 dark:text-orange-400 font-mono text-xs">${requiredThroughput.toFixed(1)}</span> containers por hora.
          </p>
          <div class="p-3 rounded-md text-[11px] font-medium leading-relaxed bg-slate-50 dark:bg-slate-700/30 border ${pctIncrease > 0 ? 'border-red-200/30 text-red-600 dark:text-red-400 bg-red-50/10' : 'border-emerald-200/30 text-emerald-600 dark:text-emerald-400 bg-emerald-50/10'}">
            ${alertOrSuccess}
          </div>
        </div>
      </div>

      <div class="space-y-4">
        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4">
          <span class="text-[10px] font-bold text-blue-500 block uppercase mb-2">Resumo de Médias</span>
          <div class="space-y-3">
             <div class="flex justify-between items-end border-b border-slate-50 dark:border-slate-700 pb-2">
               <span class="text-xs font-medium text-slate-500">Média Geral:</span>
               <span class="text-lg font-black text-blue-600 dark:text-blue-400">${formatDuration(avgHours)}</span>
             </div>
             <div class="flex justify-between items-end border-b border-slate-50 dark:border-slate-700 pb-2">
               <span class="text-xs font-medium text-slate-500">1º Período:</span>
               <span class="text-base font-bold text-emerald-600 dark:text-emerald-400">${validRecordsP1 > 0 ? formatDuration(totalTimeSumP1/validRecordsP1) : "-"}</span>
             </div>
             <div class="flex justify-between items-end">
               <span class="text-xs font-medium text-slate-500">2º Período:</span>
               <span class="text-base font-bold text-amber-600 dark:text-amber-400">${validRecordsP2 > 0 ? formatDuration(totalTimeSumP2/validRecordsP2) : "-"}</span>
             </div>
          </div>
        </div>
        <p class="text-[10px] text-slate-400 px-2 italic">Ref. Horário Relógio: ${now.toLocaleTimeString()}</p>
      </div>
    </div>

    <div class="mb-6">
      <h3 class="text-sm font-black text-slate-800 dark:text-slate-100 uppercase tracking-wide mb-4">Recebimentos por Turno e Tipo de Recebimento</h3>
      <div class="flex flex-col xl:flex-row gap-4">
        
        <!-- Card 1: DESOVAS -->
        <div class="flex-1 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-5">
          <div class="flex items-center gap-3 mb-6 border-b border-slate-100 dark:border-slate-700 pb-3">
            <i class="fas fa-box-open text-blue-600 text-2xl"></i>
            <h4 class="text-base font-bold text-blue-800 dark:text-blue-400">DESOVAS (${desovaTotal})</h4>
          </div>
          <div class="space-y-5">
            
            <div class="flex items-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-700/50 p-1.5 -mx-1.5 rounded transition-colors" onclick="window.showShiftDetails('desova1', 'DESOVAS - Realizadas inteiramente no 1º turno')">
              <div class="w-32 text-xs font-semibold text-slate-600 dark:text-slate-400 pr-2 leading-tight">Realizadas inteiramente no 1º turno</div>
              <div class="flex-1 flex items-center gap-3">
                <div class="flex-1 h-6 bg-slate-100 dark:bg-slate-700 rounded-r-md overflow-hidden relative">
                  <div class="absolute top-0 left-0 h-full bg-green-500 rounded-r-md transition-all duration-500" style="width: ${desovaTotalScheduled > 0 ? (desova1 / desovaTotalScheduled * 100) : 0}%"></div>
                </div>
                <div class="w-16 text-right flex flex-col items-end leading-tight">
                  <span class="text-sm font-black text-green-600">${desova1}</span>
                  <span class="text-[10px] font-bold text-green-600/70">(${desovaTotalScheduled > 0 ? (desova1 / desovaTotalScheduled * 100).toFixed(1).replace('.', ',') : "0,0"}%)</span>
                </div>
              </div>
            </div>
            
            <div class="flex items-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-700/50 p-1.5 -mx-1.5 rounded transition-colors" onclick="window.showShiftDetails('desova2', 'DESOVAS - Realizadas inteiramente no 2º turno')">
              <div class="w-32 text-xs font-semibold text-slate-600 dark:text-slate-400 pr-2 leading-tight">Realizadas inteiramente no 2º turno</div>
              <div class="flex-1 flex items-center gap-3">
                <div class="flex-1 h-6 bg-slate-100 dark:bg-slate-700 rounded-r-md overflow-hidden relative">
                  <div class="absolute top-0 left-0 h-full bg-blue-500 rounded-r-md transition-all duration-500" style="width: ${desovaTotalScheduled > 0 ? (desova2 / desovaTotalScheduled * 100) : 0}%"></div>
                </div>
                <div class="w-16 text-right flex flex-col items-end leading-tight">
                  <span class="text-sm font-black text-blue-600">${desova2}</span>
                  <span class="text-[10px] font-bold text-blue-600/70">(${desovaTotalScheduled > 0 ? (desova2 / desovaTotalScheduled * 100).toFixed(1).replace('.', ',') : "0,0"}%)</span>
                </div>
              </div>
            </div>

            <div class="flex items-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-700/50 p-1.5 -mx-1.5 rounded transition-colors" onclick="window.showShiftDetails('desovaCross', 'DESOVAS - Iniciadas no 1º turno e finalizadas no 2º turno')">
              <div class="w-32 text-xs font-semibold text-slate-600 dark:text-slate-400 pr-2 leading-tight">Iniciadas no 1º turno e finalizadas no 2º turno</div>
              <div class="flex-1 flex items-center gap-3">
                <div class="flex-1 h-6 bg-slate-100 dark:bg-slate-700 rounded-r-md overflow-hidden relative">
                  <div class="absolute top-0 left-0 h-full bg-orange-500 rounded-r-md transition-all duration-500" style="width: ${desovaTotalScheduled > 0 ? (desovaCross / desovaTotalScheduled * 100) : 0}%"></div>
                </div>
                <div class="w-16 text-right flex flex-col items-end leading-tight">
                  <span class="text-sm font-black text-orange-600">${desovaCross}</span>
                  <span class="text-[10px] font-bold text-orange-600/70">(${desovaTotalScheduled > 0 ? (desovaCross / desovaTotalScheduled * 100).toFixed(1).replace('.', ',') : "0,0"}%)</span>
                </div>
              </div>
            </div>

          </div>
        </div>

        <!-- Card 2: BAIXA DE PISO -->
        <div class="flex-1 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-5">
          <div class="flex items-center gap-3 mb-6 border-b border-slate-100 dark:border-slate-700 pb-3">
            <i class="fas fa-exchange-alt text-purple-600 text-2xl"></i>
            <h4 class="text-base font-bold text-purple-800 dark:text-purple-400">BAIXA DE PISO (SWAP / PUT DOWN) (${baixaTotal})</h4>
          </div>
          <div class="space-y-5">
            
            <div class="flex items-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-700/50 p-1.5 -mx-1.5 rounded transition-colors" onclick="window.showShiftDetails('baixa1', 'BAIXA DE PISO - Realizadas inteiramente no 1º turno')">
              <div class="w-32 text-xs font-semibold text-slate-600 dark:text-slate-400 pr-2 leading-tight">Realizadas inteiramente no 1º turno</div>
              <div class="flex-1 flex items-center gap-3">
                <div class="flex-1 h-6 bg-slate-100 dark:bg-slate-700 rounded-r-md overflow-hidden relative">
                  <div class="absolute top-0 left-0 h-full bg-green-500 rounded-r-md transition-all duration-500" style="width: ${baixaTotalScheduled > 0 ? (baixa1 / baixaTotalScheduled * 100) : 0}%"></div>
                </div>
                <div class="w-16 text-right flex flex-col items-end leading-tight">
                  <span class="text-sm font-black text-green-600">${baixa1}</span>
                  <span class="text-[10px] font-bold text-green-600/70">(${baixaTotalScheduled > 0 ? (baixa1 / baixaTotalScheduled * 100).toFixed(1).replace('.', ',') : "0,0"}%)</span>
                </div>
              </div>
            </div>
            
            <div class="flex items-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-700/50 p-1.5 -mx-1.5 rounded transition-colors" onclick="window.showShiftDetails('baixa2', 'BAIXA DE PISO - Realizadas inteiramente no 2º turno')">
              <div class="w-32 text-xs font-semibold text-slate-600 dark:text-slate-400 pr-2 leading-tight">Realizadas inteiramente no 2º turno</div>
              <div class="flex-1 flex items-center gap-3">
                <div class="flex-1 h-6 bg-slate-100 dark:bg-slate-700 rounded-r-md overflow-hidden relative">
                  <div class="absolute top-0 left-0 h-full bg-blue-500 rounded-r-md transition-all duration-500" style="width: ${baixaTotalScheduled > 0 ? (baixa2 / baixaTotalScheduled * 100) : 0}%"></div>
                </div>
                <div class="w-16 text-right flex flex-col items-end leading-tight">
                  <span class="text-sm font-black text-blue-600">${baixa2}</span>
                  <span class="text-[10px] font-bold text-blue-600/70">(${baixaTotalScheduled > 0 ? (baixa2 / baixaTotalScheduled * 100).toFixed(1).replace('.', ',') : "0,0"}%)</span>
                </div>
              </div>
            </div>

          </div>
        </div>

        <!-- Card 3: TOTAL & APOIO -->
        <div class="flex-none xl:w-80 bg-slate-900 dark:bg-slate-950 rounded-lg shadow-sm border border-slate-700 dark:border-slate-800 p-5 text-white flex flex-col justify-center relative overflow-hidden">
          <div class="absolute top-0 right-0 p-4 opacity-10">
            <i class="fas fa-truck-loading text-6xl"></i>
          </div>
          
          <div class="relative z-10">
            <h4 class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Total de Recebimentos</h4>
            <div class="text-4xl font-black text-white mb-6">${totalCompletedOps}</div>
            
            <div class="pt-4 border-t border-slate-700/50 space-y-3">
              <h4 class="text-[10px] font-bold text-blue-400 uppercase tracking-wider mb-2">Controle do Ponto de Apoio (Staging)</h4>
              
              <div class="flex items-center justify-between bg-slate-800/50 p-2 rounded-md border border-slate-700">
                <label for="ponto-apoio-input" class="text-xs font-semibold text-slate-300">No Ponto de Apoio (Qtd. Manual):</label>
                <input id="ponto-apoio-input" type="number" min="0" value="${pontoApoioQtd}" 
                       oninput="window.updatePontoApoio(this.value)"
                       class="w-16 bg-slate-900 border border-slate-600 rounded px-2 py-1 text-sm font-bold text-white text-center focus:outline-none focus:border-blue-500" />
              </div>

              <div class="bg-${stateColor}-500/10 border border-${stateColor}-500/30 rounded-md p-3 mt-3">
                <div class="flex flex-col gap-1.5">
                  <div class="flex items-start gap-2">
                    <i class="fas ${stateIcon} text-${stateColor}-400 mt-0.5 text-xs"></i>
                    <div>
                      <span class="block text-xs font-bold text-${stateColor}-300">${stateTitle}</span>
                      <span class="block text-lg font-black text-white mt-1">
                        ${stateMessage}
                      </span>
                    </div>
                  </div>
                  <div class="mt-2 text-[10px] text-slate-300 space-y-1">
                    <div class="flex justify-between">
                      <span class="text-slate-400">Tempo Estimado (ETA):</span>
                      <span class="font-semibold text-white">20–30 min (Just-In-Time)</span>
                    </div>
                    <div class="flex justify-between">
                      <span class="text-slate-400">Saldo Restante no Apoio:</span>
                      <span class="font-semibold text-white">${saldoProjetado}</span>
                    </div>
                  </div>
                  ${stateExtra}
                </div>
              </div>
            </div>
          </div>
        </div>

      </div>
    </div>

    <div class="overflow-x-auto bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 max-h-[450px]">
      <table class="w-full text-xs text-left">
        <thead class="bg-slate-50 dark:bg-slate-700 sticky top-0">
          <tr>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Container</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">BL</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Carrier</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Lot</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Start (Terminal)</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Finish (BYD/Empty)</th>
            <th class="px-4 py-2 font-bold uppercase text-slate-500">Duration</th>
          </tr>
        </thead>
        <tbody class="divide-y">${rowsHtml}</tbody>
      </table>
    </div>
  `;
}

/* ----------------------- XLSX PARSER & SAVE ---------------------------- */
function buildHeaderIndex(headers: any[]): Record<string, number> {
  const idx: Record<string, number> = {};
  headers.forEach((h, i) => { const n = normalizeText(h); if (n) idx[n] = i; });
  return idx;
}

function pickIndex(hIdx: Record<string, number>, aliases: string[]): number {
  for (const a of aliases) { const key = normalizeText(a); if (key in hIdx) return hIdx[key]; }
  return -1;
}

function makeRowId(row: any): string {
  return normalizeText(`${row["CONTAINER"]}|${row["BL"]}|${row["DELIVERY AT BYD"]}|${row["BONDED WAREHOUSE"]}`) || String(Math.random());
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

      const rawData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      let hRow = rawData.findIndex((r) => r.some((c) => normalizeText(c) === "CONTAINER"));
      if (hRow === -1) hRow = 0;

      const headers = rawData[hRow] || [];
      const headerIndex = buildHeaderIndex(headers);

      const col = {
        DELIVERY_AT_BYD: pickIndex(headerIndex, ["DELIVERY AT BYD", "DELIVERY AT", "DELIVERY", "DATA DE ENTREGA", "ENTREGA"]),
        UNLOAD_TIME_BYD: pickIndex(headerIndex, ["UNLOAD TIME BYD", "UNLOAD TIME", "TEMPO DESCARGA"]),
        TRANSPORTATION_COMPANY: pickIndex(headerIndex, ["TRANSPORTATION COMPANY", "CARRIER", "TRANSPORTADORA", "TRANSPORTADOR"]),
        CONTAINER: pickIndex(headerIndex, ["CONTAINER", "CNTR"]),
        BL: pickIndex(headerIndex, ["BL", "B/L", "BL NUMBER"]),
        VESSEL: pickIndex(headerIndex, ["VESSEL", "NAVIO", "NAVIO/VIAGEM"]),
        BONDED_WAREHOUSE: pickIndex(headerIndex, ["BONDED WAREHOUSE", "ARMAZEM", "RECINTO", "WAREHOUSE"]),
        MODEL: pickIndex(headerIndex, ["MODEL", "MODELO"]),
        RATIONALIZATION: pickIndex(headerIndex, ["RATIONALIZATION", "RATIONALIZACAO", "RACIONALIZACAO"]),
        LOT: pickIndex(headerIndex, ["LOT", "LOTE"]),
        DRIVER_NAME: pickIndex(headerIndex, ["DRIVER NAME", "MOTORISTA"]),
        CPF: pickIndex(headerIndex, ["CPF"]),
        PLATE1: pickIndex(headerIndex, ["LICENSE P (Plate 1)", "PLACA 1", "PLATE 1"]),
        PLATE2: pickIndex(headerIndex, ["LICENSE P (Plate 2)", "PLACA 2", "PLATE 2"]),
        TRUCK_TYPE: pickIndex(headerIndex, ["TRUCK TYPE", "TIPO CAMINHAO", "TIPO"]),
        LOAD_TIME: pickIndex(headerIndex, ["LOAD TIME", "TEMPO CARGA"]),
        DEPOT_SCA: pickIndex(headerIndex, ["DEPOT SCA"]),
        SALVADOR: pickIndex(headerIndex, ["SALVADOR"]),
        PO_SAP: pickIndex(headerIndex, ["PO SAP", "PO"]),
        NF: pickIndex(headerIndex, ["NF", "NOTA FISCAL", "NFE"]),
        EMISSION_NF: pickIndex(headerIndex, ["EMISSÃO NF", "EMISSAO NF", "DATA NF"]),
        TYPE_OF_MATERIAL: pickIndex(headerIndex, ["TYPE OF MATERIAL", "MATERIAL", "TIPO DE MATERIAL", "TIPO", "MATERIAL TYPE"]),
        STATUS: pickIndex(headerIndex, ["STATUS", "SITUACAO"]),
        TERMINAL_DEPARTURE: pickIndex(headerIndex, ["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA.", "TERMINAL - INICIO", "TERMINAL - INÍCIO", "SAÍDA TERMINAL", "TERMINAL - INÍCIO DE ROTA", "TERMINAL - INICIO DE ROTA"]),
        EMPTY_DELIVERED: pickIndex(headerIndex, ["DATA E HORARIO DE ENTREGA CONTAINER VAZIO", "ENTREGA VAZIO", "ENTREGA CONTAINER VAZIO", "DATA DEVOLUCAO VAZIO"]),
        UNLOAD_AT_BYD: pickIndex(headerIndex, ["DATA E HORARIO DE DESCARGA NA BYD ", "DESCARGA BYD", "DESCARGA", "DATA E HORARIO DE DESCARGA"]),
        NOTES: pickIndex(headerIndex, ["NOTES", "OBSERVACOES", "OBSERVAÇÕES", "OBS"]),
        PARETO: pickIndex(headerIndex, ["PARETO", "MOTIVO PARETO", "MOTIVO", "REASON"]),
      };

      deliveryData = rawData.slice(hRow + 1).filter(r => safeValue(r[col.CONTAINER]) || safeValue(r[col.BL])).map((r) => {
        const obj: any = {};
        headers.forEach((h, i) => { if (h) obj[String(h).trim()] = safeValue(r[i]); });
        
        obj["DELIVERY AT BYD"] = col.DELIVERY_AT_BYD >= 0 ? safeValue(r[col.DELIVERY_AT_BYD]) : "";
        obj["UNLOAD TIME BYD"] = col.UNLOAD_TIME_BYD >= 0 ? safeValue(r[col.UNLOAD_TIME_BYD]) : "";
        obj["TRANSPORTATION COMPANY"] = col.TRANSPORTATION_COMPANY >= 0 ? safeValue(r[col.TRANSPORTATION_COMPANY]) : "";
        obj["CONTAINER"] = col.CONTAINER >= 0 ? safeValue(r[col.CONTAINER]) : "";
        obj["BL"] = col.BL >= 0 ? safeValue(r[col.BL]) : "";
        obj["VESSEL"] = col.VESSEL >= 0 ? safeValue(r[col.VESSEL]) : "";
        obj["BONDED WAREHOUSE"] = col.BONDED_WAREHOUSE >= 0 ? safeValue(r[col.BONDED_WAREHOUSE]) : "";
        obj["MODEL"] = col.MODEL >= 0 ? safeValue(r[col.MODEL]) : "";
        obj["RATIONALIZATION"] = col.RATIONALIZATION >= 0 ? safeValue(r[col.RATIONALIZATION]) : "";
        obj["LOT"] = col.LOT >= 0 ? safeValue(r[col.LOT]) : "";
        obj["DRIVER NAME"] = col.DRIVER_NAME >= 0 ? safeValue(r[col.DRIVER_NAME]) : "";
        obj["CPF"] = col.CPF >= 0 ? safeValue(r[col.CPF]) : "";
        obj["LICENSE P (Plate 1)"] = col.PLATE1 >= 0 ? safeValue(r[col.PLATE1]) : "";
        obj["LICENSE P (Plate 2)"] = col.PLATE2 >= 0 ? safeValue(r[col.PLATE2]) : "";
        obj["TRUCK TYPE"] = col.TRUCK_TYPE >= 0 ? safeValue(r[col.TRUCK_TYPE]) : "";
        obj["LOAD TIME"] = col.LOAD_TIME >= 0 ? safeValue(r[col.LOAD_TIME]) : "";
        obj["DEPOT SCA"] = col.DEPOT_SCA >= 0 ? safeValue(r[col.DEPOT_SCA]) : "";
        obj["SALVADOR"] = col.SALVADOR >= 0 ? safeValue(r[col.SALVADOR]) : "";
        obj["PO SAP"] = col.PO_SAP >= 0 ? safeValue(r[col.PO_SAP]) : "";
        obj["NF"] = col.NF >= 0 ? safeValue(r[col.NF]) : "";
        obj["EMISSÃO NF"] = col.EMISSION_NF >= 0 ? safeValue(r[col.EMISSION_NF]) : "";
        obj["TYPE OF MATERIAL"] = col.TYPE_OF_MATERIAL >= 0 ? safeValue(r[col.TYPE_OF_MATERIAL]) : "";
        obj["TERMINAL - INÍCIO DE ROTA"] = col.TERMINAL_DEPARTURE >= 0 ? safeValue(r[col.TERMINAL_DEPARTURE]) : "";
        obj["ENTREGA VAZIO"] = col.EMPTY_DELIVERED >= 0 ? safeValue(r[col.EMPTY_DELIVERED]) : "";
        obj["DATA E HORARIO DE DESCARGA"] = col.UNLOAD_AT_BYD >= 0 ? safeValue(r[col.UNLOAD_AT_BYD]) : "";
        obj["NOTES"] = col.NOTES >= 0 ? safeValue(r[col.NOTES]) : "";
        obj["PARETO"] = col.PARETO >= 0 ? safeValue(r[col.PARETO]) : "";
        obj["STATUS"] = sanitizeStatus(col.STATUS >= 0 ? safeValue(r[col.STATUS]) : "");
        obj._id = makeRowId(obj);
        return obj;
      });

      if (lastUpdate) {
        lastUpdate.dataset.sheetName = sheetName;
        lastUpdate.textContent = t("lastUpdateText", sheetName, new Date().toLocaleString());
      }

      showToast(t("sheetLoaded"), "success");
      await saveStateToFirebase();
      applyFiltersAndRender();
    } catch (err) {
      console.error(err);
      showToast(t("fileProcessError"), "error");
    }
  };
  reader.readAsArrayBuffer(file);
});

(window as any).showShiftDetails = (listKey: string, title: string) => {
  const lists = (window as any).__SHIFT_LISTS__;
  if (!lists || !lists[listKey] || lists[listKey].length === 0) {
    showToast("Nenhum container encontrado para este filtro.", "info");
    return;
  }
  const data = lists[listKey];

  let modal = document.getElementById("shift-details-modal");
  if (!modal) {
    modal = document.createElement("div");
    modal.id = "shift-details-modal";
    modal.className = "fixed inset-0 bg-black/60 dark:bg-black/80 z-50 flex items-center justify-center p-4 transition-all duration-300 hidden";
    modal.innerHTML = `
      <div class="bg-white dark:bg-slate-800 rounded-lg shadow-2xl w-full max-w-5xl max-h-[85vh] flex flex-col relative transform transition-all duration-300">
        <div class="flex justify-between items-center p-4 border-b border-slate-200 dark:border-slate-700">
          <h2 id="shift-modal-title" class="text-lg font-bold text-slate-800 dark:text-slate-100">Containers</h2>
          <button onclick="document.getElementById('shift-details-modal').classList.add('hidden')" class="text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 text-2xl font-bold cursor-pointer leading-none">&times;</button>
        </div>
        <div class="overflow-y-auto p-4 custom-scrollbar">
          <table class="w-full text-xs text-left">
            <thead class="bg-slate-50 dark:bg-slate-700 sticky -top-4 shadow-sm z-10">
              <tr>
                <th class="px-4 py-3 font-bold uppercase text-slate-500">Container</th>
                <th class="px-4 py-3 font-bold uppercase text-slate-500">Status</th>
                <th class="px-4 py-3 font-bold uppercase text-slate-500">Escopo</th>
                <th class="px-4 py-3 font-bold uppercase text-slate-500">Início</th>
                <th class="px-4 py-3 font-bold uppercase text-slate-500">Fim</th>
              </tr>
            </thead>
            <tbody id="shift-modal-tbody" class="divide-y divide-slate-100 dark:divide-slate-700">
            </tbody>
          </table>
        </div>
      </div>
    `;
    document.body.appendChild(modal);
  }
  
  modal.classList.remove("hidden");
  const titleEl = document.getElementById("shift-modal-title");
  if (titleEl) titleEl.textContent = title + ` (${data.length})`;

  const tbody = document.getElementById("shift-modal-tbody");
  if (tbody) {
    tbody.innerHTML = data.map((row: any) => {
      const startDt = toDateTimeMaybe(row["TERMINAL - INÍCIO DE ROTA"]);
      let endDt = toDateTimeMaybe(row["ENTREGA VAZIO"]) || toDateTimeMaybe(row["DATA E HORARIO DE DESCARGA"]);
      
      return `
        <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
          <td class="px-4 py-3 font-mono font-bold text-slate-800 dark:text-slate-200">${row["CONTAINER"] || "-"}</td>
          <td class="px-4 py-3 font-bold text-slate-700 dark:text-slate-300">${row["STATUS"] || "-"}</td>
          <td class="px-4 py-3 text-[10px] text-slate-600 dark:text-slate-400">${row["OPERATION SCOPE"] || "-"}</td>
          <td class="px-4 py-3 font-mono text-slate-600 dark:text-slate-400">${startDt ? startDt.toLocaleString() : "-"}</td>
          <td class="px-4 py-3 font-mono text-slate-600 dark:text-slate-400">${endDt ? endDt.toLocaleString() : "-"}</td>
        </tr>
      `;
    }).join("");
  }
};

/* ------------------------------- EXPORTS ---------------------------------- */
exportExcelBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) return showToast(t("noDataToExport"), "warning");

  const exportColumns = [
    "STATUS", "DELIVERY AT BYD", "CONTAINER", "BL", "LOT", "MODEL", "OPERATION SCOPE", "RATIONALIZATION", "TRANSPORTATION COMPANY", 
    "VESSEL", "BONDED WAREHOUSE", "DRIVER NAME", "CPF", "LICENSE P (Plate 1)", "LICENSE P (Plate 2)", "TRUCK TYPE",
    "TERMINAL - INÍCIO DE ROTA", 
    "DATA E HORARIO DE DESCARGA", 
    "ENTREGA VAZIO", 
    "TIME OF OPERATION",
    "LOAD TIME", "DEPOT SCA", "SALVADOR", "PO SAP", "NF", "EMISSÃO NF", "NOTES", "PARETO"
  ];

  // Create Title Row
  const aoa = [
    ["KD Monitor Dashboard - Supervisors View"],
    [], // Empty row
    exportColumns // Header row at index 2
  ];

  deliveryData.forEach(d => {
    // Calculate Time of Operation column
    const startDt = toDateTimeMaybe(d["TERMINAL - INÍCIO DE ROTA"]);
    const endDt = toDateTimeMaybe(d["ENTREGA VAZIO"]) || toDateTimeMaybe(d["DATA E HORARIO DE DESCARGA"]);
    let timeOfOp = "-";
    if (startDt && endDt) {
      const diffMs = endDt.getTime() - startDt.getTime();
      if (diffMs > -3600000) {
        const durationHours = Math.max(0, diffMs / (1000 * 60 * 60));
        const dDays = Math.floor(durationHours / 24);
        const dHours = Math.floor(durationHours % 24);
        const dMins = Math.round((durationHours - Math.floor(durationHours)) * 60);
        timeOfOp = dDays > 0 ? `${dDays}v ${dHours}h ${dMins}m` : `${dHours}h ${dMins}m`;
      }
    }

    const rowData = exportColumns.map(col => {
      if (col === "TIME OF OPERATION") return timeOfOp;
      return d[col] ?? "";
    });
    aoa.push(rowData);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Styling
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  
  // Style Header Row (index 2)
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: 2, c });
    if (!ws[cellRef]) continue;
    ws[cellRef].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "4472C4" } },
      alignment: { horizontal: "center", vertical: "center" }
    };
  }

  // Style Title (Row 0)
  const titleCell = XLSX.utils.encode_cell({ r: 0, c: 0 });
  if (ws[titleCell]) {
    ws[titleCell].s = {
      font: { bold: true, size: 16, color: { rgb: "4472C4" } },
      alignment: { horizontal: "left" }
    };
  }

  // Find column indices for Start and Finish
  const startColIdx = exportColumns.indexOf("TERMINAL - INÍCIO DE ROTA");
  const finishColIdx = exportColumns.indexOf("ENTREGA VAZIO");
  const durationColIdx = exportColumns.indexOf("TIME OF OPERATION");

  // Apply colors and borders to data rows (starting from row 3)
  for (let r = 3; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      if (!ws[cellRef]) ws[cellRef] = { v: "", t: "s" };
      if (!ws[cellRef].s) ws[cellRef].s = {};
      
      // Default common style: thin border for all data cells
      ws[cellRef].s.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };

      // Start Time (Yellow)
      if (c === startColIdx) {
        ws[cellRef].s.fill = { fgColor: { rgb: "FFFF00" } };
        ws[cellRef].s.alignment = { horizontal: "center" };
      }
      // Finish Time (Green)
      else if (c === finishColIdx) {
        ws[cellRef].s.fill = { fgColor: { rgb: "92D050" } };
        ws[cellRef].s.font = { color: { rgb: "000000" } };
        ws[cellRef].s.alignment = { horizontal: "center" };
      }
      // Duration (Bold + formatting)
      else if (c === durationColIdx) {
        ws[cellRef].s.font = { bold: true };
        ws[cellRef].s.alignment = { horizontal: "center" };
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, t("deliveriesTab"));

  // [rest of tab additions...]


  XLSX.writeFile(wb, `KD_Monitor_Report_${new Date().toISOString().split("T")[0]}.xlsx`);
  showToast(t("excelGenerated"), "success");
});

exportPdfBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) return showToast(t("noDataToExport"), "warning");
  try {
    const doc = new (jspdf as any).jsPDF({ orientation: "landscape" });
    doc.text(t("pdfTitle"), 14, 15);
    (doc as any).autoTable({
      head: [["#", "DELIVERY", "CONTAINER", "BL", "CARRIER", "LOT", "STATUS"]],
      body: deliveryData.map((d, i) => [i + 1, formatDate(d["DELIVERY AT BYD"]), d["CONTAINER"] || "", d["BL"] || "", d["TRANSPORTATION COMPANY"] || "", d["LOT"] || "", sanitizeStatus(d["STATUS"])]),
      startY: 25,
      styles: { fontSize: 8 },
    });
    doc.save("KD_Deliveries_Report.pdf");
    showToast(t("pdfGenerated"), "success");
  } catch (e) { console.error(e); showToast(t("fileProcessError"), "error"); }
});

const saveDayBtn = document.getElementById("save-day-btn") as HTMLButtonElement;
saveDayBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) {
    showToast(t("noDataToExport"), "warning");
    return;
  }
  const confirmed = await showConfirmationDialog(t("saveDayConfirmTitle"), t("saveDayConfirmMsg"));
  if (confirmed) {
    const existingIds = new Set(historicalData.map(d => d._id));
    deliveryData.forEach(d => {
      if (existingIds.has(d._id)) {
        const idx = historicalData.findIndex(h => h._id === d._id);
        if (idx !== -1) historicalData[idx] = d;
      } else {
        historicalData.push(d);
        existingIds.add(d._id);
      }
    });
    
    deliveryData = [];
    
    await saveStateToFirebase();
    showToast("Dia salvo e arquivado com sucesso!", "success");
    applyFiltersAndRender();
    if (deliveryData.length === 0) resetUI();
  }
});

let paretoSelectedWeek: string | null = null;
let paretoSelectedCarrier: string = "TODOS";
let paretoChartInstance: any = null;
let paretoChartMode: "reason" | "carrier" = "reason";
let paretoTableView: "daily" | "matrix" = "daily";

function renderParetoTab() {
  const paretoContent = document.getElementById("pareto-content");
  if (!paretoContent) return;

  let allData = [...historicalData, ...deliveryData];
  const selectedMonth = monthFilterSelect?.value;
  if (selectedMonth) {
    const monthIndex = parseInt(selectedMonth, 10);
    allData = allData.filter(row => {
      const d = toDateMaybe(row["DELIVERY AT BYD"]);
      if (!d) return false;
      return d.getMonth() === monthIndex;
    });
  }
  
  if (allData.length === 0) {
    paretoContent.innerHTML = `<div class="text-center py-20 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
      <i class="fas fa-chart-bar text-6xl text-slate-300 dark:text-slate-600 mb-4"></i>
      <h2 class="text-2xl font-semibold text-slate-700 dark:text-slate-200">Sem dados</h2>
      <p class="text-slate-500 dark:text-slate-400 mt-2">Nenhum dado encontrado para gerar o Pareto.</p>
    </div>`;
    return;
  }

  // 1. Group all available dates
  const groupedByDate = allData.reduce((acc, row) => {
    const d = toDateMaybe(row["DELIVERY AT BYD"]);
    const key = d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : t("undefinedDate");
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {} as Record<string, DeliveryRow[]>);

  const sortedDatesAsc = Object.keys(groupedByDate).sort((a, b) => a.localeCompare(b));
  
  const getWeekLabel = (dateStr: string) => {
    if (dateStr === t("undefinedDate")) return dateStr;
    const d = new Date(dateStr + "T12:00:00");
    const day = d.getDay();
    const diffToMonday = d.getDate() - day + (day === 0 ? -6 : 1);
    const startOfWeek = new Date(d.setDate(diffToMonday));
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    return `Semana: ${startOfWeek.getDate().toString().padStart(2, '0')}/${(startOfWeek.getMonth() + 1).toString().padStart(2, '0')} a ${endOfWeek.getDate().toString().padStart(2, '0')}/${(endOfWeek.getMonth() + 1).toString().padStart(2, '0')}`;
  };

  const getMonthLabel = (dateStr: string) => {
    if (dateStr === t("undefinedDate")) return dateStr;
    const d = new Date(dateStr + "T12:00:00");
    const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    return `Mês: ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
  };

  const weeksMap: Record<string, string[]> = {};
  sortedDatesAsc.forEach(d => {
    const w = getWeekLabel(d);
    if (!weeksMap[w]) weeksMap[w] = [];
    if (!weeksMap[w].includes(d)) weeksMap[w].push(d);
    
    const m = getMonthLabel(d);
    if (!weeksMap[m]) weeksMap[m] = [];
    if (!weeksMap[m].includes(d)) weeksMap[m].push(d);
  });

  const availablePeriods = Object.keys(weeksMap);
  const availableWeeks = availablePeriods.filter(p => p.startsWith("Semana"));
  const availableMonths = availablePeriods.filter(p => p.startsWith("Mês"));
  const sortedPeriods = [...availableWeeks, ...availableMonths];

  if (!paretoSelectedWeek && sortedPeriods.length > 0) {
    paretoSelectedWeek = availableWeeks.length > 0 ? availableWeeks[availableWeeks.length - 1] : sortedPeriods[sortedPeriods.length - 1]; // default to latest week
  }

  const datesInSelectedWeek = paretoSelectedWeek ? (weeksMap[paretoSelectedWeek] || []) : [];
  
  // 2. Identify carriers
  const carriers = new Set<string>();
  allData.forEach(r => {
    const c = String(r["TRANSPORTATION COMPANY"] || "").trim().toUpperCase();
    if (c) carriers.add(c);
  });
  const carrierOptions = Array.from(carriers).sort();

  // 3. Filter data for the selected week and carrier
  const weekData = datesInSelectedWeek.flatMap(d => groupedByDate[d] || []);
  const filteredData = paretoSelectedCarrier === "TODOS" 
    ? weekData 
    : weekData.filter(r => String(r["TRANSPORTATION COMPANY"] || "").trim().toUpperCase() === paretoSelectedCarrier);
  
  const occurrencesByDateAndReason: Record<string, Record<string, number>> = {};
  const occurrencesByCarrierAndReason: Record<string, Record<string, number>> = {};
  const defaultReasons = [
    "PRAZO CURTO PARA COLETA",
    "QUEBRA DE VEÍCULO",
    "INCIDENTE TERMINAL",
    "GREVE DOS CAMINHONEIROS",
    "GREVE SINDICAL",
    "ALTERAÇÃO DE PROGRAMAÇÃO",
    "ACIDENTE NA RODOVIA",
    "FILA NO TERMINAL",
    "PENDÊNCIA DOCUMENTAL"
  ];
  const allReasons = new Set<string>(defaultReasons);
  if ((window as any).__PARETO_REASONS__) {
    (window as any).__PARETO_REASONS__.forEach((r: string) => allReasons.add(r));
  }
  
  const matrixCarriers = new Set<string>();

  filteredData.forEach(row => {
    const pareto = String(row["PARETO"] || "").trim().toUpperCase();
    if (pareto && pareto !== "-") {
      const d = toDateMaybe(row["DELIVERY AT BYD"]);
      const dateKey = d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : t("undefinedDate");
      const carrier = String(row["TRANSPORTATION COMPANY"] || "").trim().toUpperCase() || "DESCONHECIDO";
      
      if (datesInSelectedWeek.includes(dateKey)) {
        if (!occurrencesByDateAndReason[dateKey]) occurrencesByDateAndReason[dateKey] = {};
        if (!occurrencesByDateAndReason[dateKey][pareto]) occurrencesByDateAndReason[dateKey][pareto] = 0;
        occurrencesByDateAndReason[dateKey][pareto]++;
        
        if (!occurrencesByCarrierAndReason[carrier]) occurrencesByCarrierAndReason[carrier] = {};
        if (!occurrencesByCarrierAndReason[carrier][pareto]) occurrencesByCarrierAndReason[carrier][pareto] = 0;
        occurrencesByCarrierAndReason[carrier][pareto]++;
        
        matrixCarriers.add(carrier);
        allReasons.add(pareto);
      }
    }
  });

  const matrixCarrierList = Array.from(matrixCarriers).sort();

  let reasonTotals: { reason: string, total: number }[] = Array.from(allReasons).map(reason => {
    let sum = 0;
    datesInSelectedWeek.forEach(d => {
      sum += (occurrencesByDateAndReason[d]?.[reason] || 0);
    });
    return { reason, total: sum };
  }).filter(r => r.total > 0);

  if (reasonTotals.length === 0) {
    // leave it empty
  }

  reasonTotals.sort((a, b) => {
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return a.reason.localeCompare(b.reason);
  });
  const grandTotal = reasonTotals.reduce((acc, r) => acc + r.total, 0);

  let cumulative = 0;
  const tableRows = reasonTotals.map(r => {
    const pct = grandTotal > 0 ? (r.total / grandTotal) * 100 : 0;
    cumulative += pct;
    return {
      reason: r.reason,
      total: r.total,
      pct,
      cumulative: Math.min(cumulative, 100)
    };
  });

  // Calculate carrier totals for the carrier chart mode
  let carrierTotals: { carrier: string, total: number }[] = matrixCarrierList.map(carrier => {
    let sum = 0;
    reasonTotals.forEach(r => {
      sum += (occurrencesByCarrierAndReason[carrier]?.[r.reason] || 0);
    });
    return { carrier, total: sum };
  }).filter(c => c.total > 0);

  carrierTotals.sort((a, b) => {
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return a.carrier.localeCompare(b.carrier);
  });

  let carrierCumulative = 0;
  const carrierChartData = carrierTotals.map(c => {
    const pct = grandTotal > 0 ? (c.total / grandTotal) * 100 : 0;
    carrierCumulative += pct;
    return {
      carrier: c.carrier,
      total: c.total,
      pct,
      cumulative: Math.min(carrierCumulative, 100)
    };
  });

  const formatDateLabel = (dKey: string) => dKey === t("undefinedDate") ? dKey : dKey.split("-").reverse().join("/");

  let html = `
    <div class="bg-red-600 text-white p-4 rounded-t-lg shadow-sm">
      <h2 class="text-center text-2xl font-bold">Análise de Fenômeno - Paretos</h2>
    </div>
    
    <div class="bg-white dark:bg-slate-800 p-4 border-x border-b border-slate-200 dark:border-slate-700 flex flex-wrap gap-4 items-center justify-between mb-6 rounded-b-lg shadow-sm">
      <div class="flex flex-wrap gap-4 items-center">
        <div class="flex items-center gap-2">
          <i class="fas fa-calendar-alt text-slate-400"></i>
          <select id="pareto-week-select" class="border border-slate-300 dark:border-slate-600 rounded p-2 text-sm bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-200 font-semibold focus:ring-2 focus:ring-blue-500 outline-none">
            <optgroup label="Semanas">
              ${availableWeeks.map(w => `<option value="${w}" ${w === paretoSelectedWeek ? 'selected' : ''}>${w}</option>`).join('')}
            </optgroup>
            <optgroup label="Meses">
              ${availableMonths.map(m => `<option value="${m}" ${m === paretoSelectedWeek ? 'selected' : ''}>${m}</option>`).join('')}
            </optgroup>
          </select>
        </div>
        
        <div class="flex items-center gap-2">
          <i class="fas fa-truck text-slate-400"></i>
          <select id="pareto-carrier-select" class="border border-slate-300 dark:border-slate-600 rounded p-2 text-sm bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-200 font-semibold focus:ring-2 focus:ring-blue-500 outline-none">
            <option value="TODOS" ${paretoSelectedCarrier === "TODOS" ? 'selected' : ''}>TODAS TRANSPORTADORAS</option>
            ${carrierOptions.map(c => `<option value="${c}" ${c === paretoSelectedCarrier ? 'selected' : ''}>${c}</option>`).join('')}
          </select>
        </div>
      </div>
      <div>
        <button id="pareto-config-btn" class="bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-md font-semibold text-sm transition-colors flex items-center gap-2">
          <i class="fas fa-cog"></i> Configurar Motivos
        </button>
      </div>
    </div>

    <div class="flex flex-col xl:flex-row gap-6 mb-6">
      <div class="flex-1 overflow-x-auto bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
        <div class="bg-slate-200 dark:bg-slate-700 p-2 border-b border-slate-300 dark:border-slate-600 flex justify-between items-center">
          <h3 class="font-bold text-slate-700 dark:text-slate-200 ml-2">Estratificação 1</h3>
          <div class="flex">
            <button id="pareto-table-daily" class="px-3 py-1 text-xs font-bold rounded-l-md border ${paretoTableView === 'daily' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300'}">Visão Diária</button>
            <button id="pareto-table-matrix" class="px-3 py-1 text-xs font-bold rounded-r-md border ${paretoTableView === 'matrix' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300'}">Matriz Transportadoras</button>
          </div>
        </div>
        ${paretoTableView === 'daily' ? `
        <table class="w-full text-sm text-left whitespace-nowrap">
          <thead class="bg-slate-400 dark:bg-slate-600 text-white font-bold text-xs">
            <tr>
              <th class="px-2 py-3 border border-slate-300 dark:border-slate-500 text-center w-10">#</th>
              <th class="px-4 py-3 border border-slate-300 dark:border-slate-500">${paretoSelectedCarrier !== "TODOS" ? paretoSelectedCarrier : "MOTIVOS"}</th>
              ${datesInSelectedWeek.map(d => `<th class="px-2 py-3 border border-slate-300 dark:border-slate-500 text-center text-[10px] sm:text-xs bg-[#c2d69b] dark:bg-lime-800 text-slate-800 dark:text-lime-100">${formatDateLabel(d)}</th>`).join('')}
              <th class="px-3 py-3 border border-slate-300 dark:border-slate-500 text-center bg-slate-500 dark:bg-slate-700">Qtd de Cntr</th>
              <th class="px-3 py-3 border border-slate-300 dark:border-slate-500 text-center bg-slate-500 dark:bg-slate-700">%</th>
              <th class="px-3 py-3 border border-slate-300 dark:border-slate-500 text-center bg-slate-500 dark:bg-slate-700">% Acum.</th>
            </tr>
          </thead>
          <tbody class="text-slate-700 dark:text-slate-300">
            ${tableRows.length > 0 ? tableRows.map((r, idx) => `
              <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                <td class="px-2 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold">${idx + 1}</td>
                <td class="px-4 py-2 border border-slate-200 dark:border-slate-700 font-semibold">${r.reason}</td>
                ${datesInSelectedWeek.map(d => {
                  const val = occurrencesByDateAndReason[d]?.[r.reason] || 0;
                  return `<td class="px-2 py-2 border border-slate-200 dark:border-slate-700 text-center ${val > 0 ? '' : 'text-transparent'}">${val > 0 ? val : 0}</td>`;
                }).join('')}
                <td class="px-3 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold bg-yellow-50 dark:bg-yellow-900/20 text-slate-800">${r.total}</td>
                <td class="px-3 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold">${r.pct.toFixed(2).replace('.', ',')}%</td>
                <td class="px-3 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold">${r.cumulative.toFixed(0)}%</td>
              </tr>
            `).join('') : `<tr><td colspan="${4 + datesInSelectedWeek.length}" class="text-center py-4">Nenhum dado de ocorrência para os filtros selecionados.</td></tr>`}
          </tbody>
          ${tableRows.length > 0 ? `
          <tfoot class="bg-slate-600 dark:bg-slate-800 text-white font-bold">
            <tr>
              <td colspan="2" class="px-4 py-3 border border-slate-500 text-right">TOTAL</td>
              ${datesInSelectedWeek.map(d => {
                const totalForDay = tableRows.reduce((sum, r) => sum + (occurrencesByDateAndReason[d]?.[r.reason] || 0), 0);
                return `<td class="px-2 py-3 border border-slate-500 text-center">${totalForDay > 0 ? totalForDay : ''}</td>`;
              }).join('')}
              <td class="px-3 py-3 border border-slate-500 text-center">${grandTotal}</td>
              <td class="px-3 py-3 border border-slate-500 text-center">100%</td>
              <td class="px-3 py-3 border border-slate-500 text-center"></td>
            </tr>
          </tfoot>
          ` : ''}
        </table>
        ` : `
        <table class="w-full text-sm text-left whitespace-nowrap">
          <thead class="bg-slate-400 dark:bg-slate-600 text-white font-bold text-xs">
            <tr>
              <th class="px-4 py-3 border border-slate-300 dark:border-slate-500">MOTIVO</th>
              ${matrixCarrierList.map(c => `<th class="px-2 py-3 border border-slate-300 dark:border-slate-500 text-center bg-[#c2d69b] dark:bg-lime-800 text-slate-800 dark:text-lime-100 max-w-[100px] truncate" title="${c}">${c}</th>`).join('')}
              <th class="px-3 py-3 border border-slate-300 dark:border-slate-500 text-center bg-slate-500 dark:bg-slate-700">Qtd</th>
              <th class="px-3 py-3 border border-slate-300 dark:border-slate-500 text-center bg-slate-500 dark:bg-slate-700">%</th>
            </tr>
          </thead>
          <tbody class="text-slate-700 dark:text-slate-300">
            ${tableRows.length > 0 ? tableRows.map((r) => `
              <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                <td class="px-4 py-2 border border-slate-200 dark:border-slate-700 font-semibold">${r.reason}</td>
                ${matrixCarrierList.map(c => {
                  const val = occurrencesByCarrierAndReason[c]?.[r.reason] || 0;
                  return `<td class="px-2 py-2 border border-slate-200 dark:border-slate-700 text-center ${val > 0 ? '' : 'text-transparent'}">${val > 0 ? val : 0}</td>`;
                }).join('')}
                <td class="px-3 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold bg-yellow-50 dark:bg-yellow-900/20 text-slate-800">${r.total}</td>
                <td class="px-3 py-2 border border-slate-200 dark:border-slate-700 text-center font-bold">${r.pct.toFixed(2).replace('.', ',')}%</td>
              </tr>
            `).join('') : `<tr><td colspan="${3 + matrixCarrierList.length}" class="text-center py-4">Nenhum dado de ocorrência para os filtros selecionados.</td></tr>`}
          </tbody>
          ${tableRows.length > 0 ? `
          <tfoot class="bg-slate-600 dark:bg-slate-800 text-white font-bold">
            <tr>
              <td class="px-4 py-3 border border-slate-500 text-right">TOTAL</td>
              ${matrixCarrierList.map(c => {
                const totalForCarrier = tableRows.reduce((sum, r) => sum + (occurrencesByCarrierAndReason[c]?.[r.reason] || 0), 0);
                return `<td class="px-2 py-3 border border-slate-500 text-center">${totalForCarrier > 0 ? totalForCarrier : ''}</td>`;
              }).join('')}
              <td class="px-3 py-3 border border-slate-500 text-center">${grandTotal}</td>
              <td class="px-3 py-3 border border-slate-500 text-center">100%</td>
            </tr>
          </tfoot>
          ` : ''}
        </table>
        `}
      </div>
      
      <div class="flex-1 min-w-[300px] bg-slate-100 dark:bg-slate-800/80 rounded-lg shadow-sm border border-slate-300 dark:border-slate-700 p-4 relative flex flex-col">
        <div class="flex justify-between items-center mb-4">
          <div class="flex rounded-md shadow-sm">
            <button id="pareto-chart-mode-reason" class="px-4 py-1.5 text-xs font-bold rounded-l-md border ${paretoChartMode === 'reason' ? 'bg-blue-600 text-white border-blue-600 z-10' : 'bg-white text-slate-600 border-slate-300'}">Por Motivo</button>
            <button id="pareto-chart-mode-carrier" class="px-4 py-1.5 text-xs font-bold rounded-r-md border -ml-px ${paretoChartMode === 'carrier' ? 'bg-blue-600 text-white border-blue-600 z-10' : 'bg-white text-slate-600 border-slate-300'}">Por Transportadora</button>
          </div>
          ${paretoChartMode === 'reason' ? '<span class="text-xs text-slate-500">💡 Clique nas barras para ver o detalhamento</span>' : ''}
        </div>
        <div class="flex-1 min-h-[360px]">
          <canvas id="paretoChartCanvas"></canvas>
        </div>
      </div>
  `;

  paretoContent.innerHTML = html;

  const ctx = document.getElementById("paretoChartCanvas") as HTMLCanvasElement;
  if (ctx && tableRows.length > 0) {
    if (paretoChartInstance) {
      paretoChartInstance.destroy();
    }
    
    const isReasonMode = paretoChartMode === 'reason';
    const chartLabels = isReasonMode ? tableRows.map(r => r.reason) : carrierChartData.map(c => c.carrier);
    const chartCumulative = isReasonMode ? tableRows.map(r => r.cumulative) : carrierChartData.map(c => c.cumulative);
    const chartTotals = isReasonMode ? tableRows.map(r => r.total) : carrierChartData.map(c => c.total);
    
    paretoChartInstance = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: [
          {
            type: 'line',
            label: '% Acumulado',
            data: chartCumulative,
            borderColor: '#ef4444',
            backgroundColor: '#ef4444',
            borderWidth: 2,
            pointBackgroundColor: '#ef4444',
            pointBorderColor: '#fff',
            pointRadius: 5,
            yAxisID: 'y1',
            datalabels: {
              align: 'top',
              anchor: 'center',
              formatter: (val: number) => val.toFixed(0) + '%',
              color: '#333',
              font: { weight: 'bold', size: 10 },
              offset: 4
            }
          },
          {
            type: 'bar',
            label: 'Qtd de Cntr',
            data: chartTotals,
            backgroundColor: isReasonMode ? '#c2410c' : '#3b82f6', 
            borderColor: isReasonMode ? '#9a3412' : '#2563eb',
            borderWidth: 1,
            yAxisID: 'y',
            datalabels: {
              align: 'start',
              anchor: 'start',
              formatter: (val: number) => val,
              color: '#fff',
              font: { weight: 'bold', size: 12 },
              offset: -20 // Push label inside bar
            }
          }
        ]
      },
      plugins: [ChartDataLabels],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        onClick: (evt: any, elements: any[]) => {
          if (isReasonMode && elements.length > 0) {
            const index = elements[0].index;
            const clickedReason = chartLabels[index];
            showParetoDrillDownModal(clickedReason, matrixCarrierList, occurrencesByCarrierAndReason);
          }
        },
        plugins: {
          legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: (ctx: any) => ctx.datasetIndex === 0 ? ` ${ctx.raw.toFixed(0)}%` : ` ${ctx.raw} cntr`
            }
          },
          datalabels: {
            display: true
          }
        },
        scales: {
          x: {
            ticks: {
              autoSkip: false,
              maxRotation: 45,
              minRotation: 45,
              font: { size: 10 }
            },
            grid: { display: false }
          },
          y: {
            type: 'linear',
            display: true,
            position: 'left',
            title: { display: false },
            grid: { drawOnChartArea: true, color: '#e2e8f0' }
          },
          y1: {
            type: 'linear',
            display: true,
            position: 'right',
            min: 0,
            max: 100,
            grid: { drawOnChartArea: false },
            ticks: {
              stepSize: 10,
              callback: (val: number) => val + '%'
            }
          }
        }
      }
    });
  }

  document.getElementById("pareto-week-select")?.addEventListener("change", (e) => {
    paretoSelectedWeek = (e.target as HTMLSelectElement).value;
    renderParetoTab();
  });
  
  document.getElementById("pareto-carrier-select")?.addEventListener("change", (e) => {
    paretoSelectedCarrier = (e.target as HTMLSelectElement).value;
    renderParetoTab();
  });

  document.getElementById("pareto-config-btn")?.addEventListener("click", () => {
    renderParetoConfigModal();
  });

  document.getElementById("pareto-table-daily")?.addEventListener("click", () => {
    paretoTableView = "daily";
    renderParetoTab();
  });

  document.getElementById("pareto-table-matrix")?.addEventListener("click", () => {
    paretoTableView = "matrix";
    renderParetoTab();
  });

  document.getElementById("pareto-chart-mode-reason")?.addEventListener("click", () => {
    paretoChartMode = "reason";
    renderParetoTab();
  });

  document.getElementById("pareto-chart-mode-carrier")?.addEventListener("click", () => {
    paretoChartMode = "carrier";
    renderParetoTab();
  });
}

function showParetoDrillDownModal(reason: string, matrixCarrierList: string[], occurrencesByCarrierAndReason: Record<string, Record<string, number>>) {
  const modal = document.createElement("div");
  modal.className = "fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4 animate-in fade-in";
  
  let carrierData: { carrier: string, count: number }[] = [];
  let totalCount = 0;
  
  matrixCarrierList.forEach(c => {
    const count = occurrencesByCarrierAndReason[c]?.[reason] || 0;
    if (count > 0) {
      carrierData.push({ carrier: c, count });
      totalCount += count;
    }
  });
  
  carrierData.sort((a, b) => b.count - a.count);

  modal.innerHTML = `
    <div class="bg-white dark:bg-slate-800 rounded-lg shadow-xl max-w-2xl w-full flex flex-col max-h-[90vh]">
      <div class="flex items-center justify-between p-4 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 rounded-t-lg">
        <div>
          <h3 class="text-lg font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
            <i class="fas fa-search-plus text-blue-500"></i> Detalhamento: ${reason}
          </h3>
          <p class="text-sm text-slate-500 mt-1">Ocorrências por transportadora nesta semana</p>
        </div>
        <button id="close-drilldown-modal" class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-300">
          <i class="fas fa-times text-xl"></i>
        </button>
      </div>
      <div class="p-6 overflow-y-auto">
        ${carrierData.length > 0 ? `
          <div class="mb-4">
            <div class="flex justify-between items-end mb-2">
              <span class="text-sm font-semibold text-slate-600 dark:text-slate-400">Total de contêineres:</span>
              <span class="text-2xl font-bold text-slate-800 dark:text-slate-200">${totalCount}</span>
            </div>
            <div class="h-12 w-full bg-slate-100 dark:bg-slate-900 rounded overflow-hidden flex border border-slate-200 dark:border-slate-700">
              ${carrierData.map((cd, i) => {
                const colors = ['bg-blue-500', 'bg-orange-500', 'bg-emerald-500', 'bg-purple-500', 'bg-pink-500', 'bg-yellow-500', 'bg-cyan-500'];
                const bg = colors[i % colors.length];
                const pct = (cd.count / totalCount) * 100;
                return `<div class="${bg}" style="width: ${pct}%" title="${cd.carrier}: ${cd.count} (${pct.toFixed(1)}%)"></div>`;
              }).join('')}
            </div>
          </div>
          <table class="w-full text-sm text-left mt-6">
            <thead class="bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-bold text-xs uppercase">
              <tr>
                <th class="px-4 py-2 border-y border-slate-200 dark:border-slate-600">Transportadora</th>
                <th class="px-4 py-2 border-y border-slate-200 dark:border-slate-600 text-center w-24">Qtd</th>
                <th class="px-4 py-2 border-y border-slate-200 dark:border-slate-600 text-center w-24">%</th>
              </tr>
            </thead>
            <tbody>
              ${carrierData.map((cd, i) => {
                const colors = ['text-blue-500', 'text-orange-500', 'text-emerald-500', 'text-purple-500', 'text-pink-500', 'text-yellow-500', 'text-cyan-500'];
                const color = colors[i % colors.length];
                return `
                <tr class="border-b border-slate-100 dark:border-slate-700/50 hover:bg-slate-50 dark:hover:bg-slate-700/30">
                  <td class="px-4 py-2 font-semibold text-slate-700 dark:text-slate-300 flex items-center gap-2">
                    <div class="w-3 h-3 rounded-full ${color.replace('text-', 'bg-')}"></div>
                    ${cd.carrier}
                  </td>
                  <td class="px-4 py-2 text-center font-bold text-slate-800 dark:text-slate-200">${cd.count}</td>
                  <td class="px-4 py-2 text-center text-slate-500">${((cd.count / totalCount) * 100).toFixed(1)}%</td>
                </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        ` : `
          <div class="text-center py-8">
            <i class="fas fa-check-circle text-4xl text-emerald-400 mb-3"></i>
            <p class="text-slate-600 dark:text-slate-300 font-medium">Nenhuma ocorrência registrada.</p>
          </div>
        `}
      </div>
      <div class="p-4 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 rounded-b-lg flex justify-end">
        <button id="close-drilldown-btn" class="px-6 py-2 bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 rounded font-semibold transition-colors">
          Fechar
        </button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const closeModal = () => modal.remove();
  document.getElementById("close-drilldown-modal")?.addEventListener("click", closeModal);
  document.getElementById("close-drilldown-btn")?.addEventListener("click", closeModal);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModal();
  });
}

function renderParetoConfigModal(currentReasons?: string[]) {
  const container = document.createElement("div");
  container.className = "fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4 transition-opacity";
  container.id = "pareto-config-modal-container";
  
  const reasons = currentReasons ? [...currentReasons] : [...((window as any).__PARETO_REASONS__ || [
    "PRAZO CURTO PARA COLETA",
    "QUEBRA DE VEÍCULO",
    "INCIDENTE TERMINAL",
    "GREVE DOS CAMINHONEIROS",
    "GREVE SINDICAL",
    "ALTERAÇÃO DE PROGRAMAÇÃO",
    "ACIDENTE NA RODOVIA",
    "FILA NO TERMINAL",
    "PENDÊNCIA DOCUMENTAL"
  ])];

  container.innerHTML = `
    <div class="bg-white dark:bg-slate-800 rounded-lg shadow-xl w-full max-w-lg overflow-hidden flex flex-col max-h-[80vh]">
      <div class="bg-slate-100 dark:bg-slate-700 p-4 border-b border-slate-200 dark:border-slate-600 flex justify-between items-center">
        <h3 class="font-bold text-lg text-slate-800 dark:text-slate-100">Configurar Motivos (Pareto)</h3>
        <button id="close-pareto-config" class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"><i class="fas fa-times text-xl"></i></button>
      </div>
      <div class="p-6 overflow-y-auto flex-1 custom-scrollbar">
        <div class="mb-4">
          <label class="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-2">Adicionar Novo Motivo</label>
          <div class="flex gap-2">
            <input type="text" id="new-pareto-reason" class="flex-1 border border-slate-300 dark:border-slate-600 rounded p-2 text-sm bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-200 outline-none focus:ring-2 focus:ring-blue-500 uppercase" placeholder="DIGITE O MOTIVO...">
            <button id="add-pareto-reason-btn" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded font-semibold text-sm transition-colors"><i class="fas fa-plus"></i> Adicionar</button>
          </div>
        </div>
        
        <label class="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-2 mt-6">Motivos Cadastrados</label>
        <ul id="pareto-reasons-list" class="divide-y divide-slate-100 dark:divide-slate-700 border border-slate-200 dark:border-slate-700 rounded-md">
          ${reasons.map((r: string, i: number) => `
            <li class="p-3 flex justify-between items-center bg-slate-50 dark:bg-slate-800/50 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors">
              <span class="text-sm font-medium text-slate-700 dark:text-slate-300">${r}</span>
              <button class="text-red-500 hover:text-red-700 p-1 delete-reason-btn" data-index="${i}"><i class="fas fa-trash-alt"></i></button>
            </li>
          `).join('')}
        </ul>
      </div>
      <div class="p-4 border-t border-slate-200 dark:border-slate-600 bg-slate-50 dark:bg-slate-700/50 flex justify-end gap-3">
        <button id="save-pareto-config-btn" class="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-2 rounded font-bold transition-colors">Salvar Alterações</button>
      </div>
    </div>
  `;

  document.body.appendChild(container);

  const closeBtn = document.getElementById("close-pareto-config");
  const saveBtn = document.getElementById("save-pareto-config-btn");
  const addBtn = document.getElementById("add-pareto-reason-btn");
  const input = document.getElementById("new-pareto-reason") as HTMLInputElement;

  const close = () => {
    container.remove();
  };

  closeBtn?.addEventListener("click", close);

  addBtn?.addEventListener("click", () => {
    const val = input.value.trim().toUpperCase();
    if (val && !reasons.includes(val)) {
      reasons.unshift(val); // add to top
      close();
      renderParetoConfigModal(reasons); // re-render with updated reasons
    }
  });

  input?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") addBtn?.click();
  });

  container.querySelectorAll(".delete-reason-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const idx = parseInt((e.currentTarget as HTMLElement).dataset.index || "-1", 10);
      if (idx >= 0) {
        reasons.splice(idx, 1);
        close();
        renderParetoConfigModal(reasons); // re-render with updated reasons
      }
    });
  });

  saveBtn?.addEventListener("click", async () => {
    (window as any).__PARETO_REASONS__ = reasons;
    await saveStateToFirebase();
    showToast("Motivos atualizados com sucesso", "success");
    close();
    applyFiltersAndRender();
    renderParetoTab();
  });
}

/* ------------------------------ STARTUP ----------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  setTheme(((localStorage.getItem("theme") as any) || "light") as any);
  loadLogoFromStorage();
  setLanguage((localStorage.getItem("language") as Language) || "pt-BR");
  listenForRealtimeUpdates();
  resetUI();

  // Escape key handler to close maximized chart modal
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      const maxContainer = document.getElementById("chart-max-modal-container");
      const maxModal = document.getElementById("chart-max-modal");
      if (maxContainer && !maxContainer.classList.contains("hidden") && maxModal) {
        maxModal.classList.add("scale-95", "opacity-0");
        maxModal.classList.remove("scale-100", "opacity-100");
        setTimeout(() => {
          maxContainer.classList.add("hidden");
          if (maxLotChart) {
            maxLotChart.destroy();
            maxLotChart = null;
          }
        }, 200);
      }
    }
  });
});
