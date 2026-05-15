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
 * - Fixed: Integrated Inventory and Operational Blog storage with online persistence (Firebase)
 * - Standardized BYD Corporate Minutes formatting layout for operational logs and reports
 */

import { mountStorageInventory } from './StorageInventory';

function renderInventory(data: any[]) {
    const container = document.getElementById("inventory-content");
    if (container) {
        mountStorageInventory(container, data);
    }
}
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
    pageTitle: "KD Monitor Dashboard",
    headerTitle: "KD Monitor Dashboard",
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
    arrivalsTab: "Chegadas por Lote",
    chartsTab: "Gráficos (Operação)",
    timeTab: "Tempo de Operação",
    inventoryTab: "Estoque",
    newsTab: "Diário de Operação",
    newsBreaking: "ÚLTIMAS",
    newsNewUpdate: "Nova Atualização",
    newsNoPosts: "Nenhuma notícia postada ainda.",
    newsWaiting: "Aguardando atualizações da operação...",
    newsOlder: "Atualizações Anteriores",
    newsStayConnected: "Fique Conectado",
    newsEffTitle: "Eficiência Operacional",
    newsTransitTitle: "Trânsito Ativo",
    newsBacklogTitle: "Carga em Backlog",
    newsLotJust: "Justificativas por Lote",
    newsAddInfo: "Informações Adicionais",
    newsSaveNotes: "Salvar Notas",
    newsSaveAlerts: "Salvar Avisos",
    newsPostNew: "Publicar Nova Atualização",
    newsPublish: "Publicar História",
    newsModalTitle: "Nova Atualização Editorial",
    newsPostPhoto: "Capa do Post (Foto)",
    newsClickToPhoto: "CLIQUE PARA ADICIONAR FOTO",
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
    goalWeekday: "150/dia útil",
    goalWeekend: "Fim de semana (bônus)",
    reachedGoal: "Meta atingida",
    notReachedGoal: "Abaixo da meta",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "Hoje"
  },
  "en-US": {
    pageTitle: "KD Monitor Dashboard",
    headerTitle: "KD Monitor Dashboard",
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
    arrivalsTab: "Arrivals per Lot",
    chartsTab: "Charts (Operation)",
    timeTab: "Operation Time",
    inventoryTab: "Inventory",
    newsTab: "Operation Blog",
    newsBreaking: "BREAKING",
    newsNewUpdate: "New Update",
    newsNoPosts: "No news posted yet.",
    newsWaiting: "Waiting for operation updates...",
    newsOlder: "Older Updates",
    newsStayConnected: "Stay Connected",
    newsEffTitle: "Operation Efficiency",
    newsTransitTitle: "Active Transit",
    newsBacklogTitle: "Backlog Load",
    newsLotJust: "Lot Justifications",
    newsAddInfo: "Additional Information",
    newsSaveNotes: "Save Notes",
    newsSaveAlerts: "Save Notices",
    newsPostNew: "Post New Update",
    newsPublish: "Publish Story",
    newsModalTitle: "Post New Editorial Update",
    newsPostPhoto: "Post Cover (Photo)",
    newsClickToPhoto: "CLICK TO ADD PHOTO",
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
    goalWeekday: "150/weekday",
    goalWeekend: "Weekend (bonus)",
    reachedGoal: "Goal reached",
    notReachedGoal: "Below goal",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "Today"
  },
  "zh-CN": {
    pageTitle: "KD 监控仪表板",
    headerTitle: "KD 监控仪表板",
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
    arrivalsTab: "每批到达",
    chartsTab: "图表（运营）",
    timeTab: "运营时间",
    inventoryTab: "库存",
    newsTab: "运营日志",
    newsBreaking: "快讯",
    newsNewUpdate: "新增更新",
    newsNoPosts: "尚无发布消息。",
    newsWaiting: "正在等待运营更新...",
    newsOlder: "历史更新",
    newsStayConnected: "保持关注",
    newsEffTitle: "运营效率",
    newsTransitTitle: "在途运输",
    newsBacklogTitle: "积压负荷",
    newsLotJust: "各批次说明",
    newsAddInfo: "附加信息",
    newsSaveNotes: "保存笔记",
    newsSaveAlerts: "保存通知",
    newsPostNew: "发布新动态",
    newsPublish: "发布故事",
    newsModalTitle: "发布新编辑更新",
    newsPostPhoto: "文章封面 (照片)",
    newsClickToPhoto: "点击添加照片",
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
    goalWeekday: "工作日150",
    goalWeekend: "周末（加分）",
    reachedGoal: "已达目标",
    notReachedGoal: "未达目标",
    kpiGoal: (del: number, goal: number) => `${del}/${goal}`,
    today: "今天"
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
let blogPosts: BlogPost[] = [];
let lotJustifications: Record<string, string> = {};
let generalNotes: string = "";
let searchDebounceTimer: number;
let activeStatusFilter: string | null = null;
let showOnlyBattery: boolean = false;
let showOnlyKd: boolean = false;
let showOnlyProject: boolean = false;
let isMacroView: boolean = false;
let overallChart: any = null;
let lotChart: any = null;
let modelChart: any = null;
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

export type BlogPost = {
  id: string;
  title: string;
  text: string;
  image?: string;
  category: string;
  author: string;
  createdAt: any;
};

type FirebaseState = {
  deliveryData?: DeliveryRow[];
  lastUpdate?: any; // Firestore Timestamp
  lastUpdateSheetName?: string;
  companyLogo?: string;
  blogPosts?: BlogPost[];
  lotJustifications?: Record<string, string>;
  generalNotes?: string;
};

const FIREBASE_COLLECTION = "delivery_dashboard";
const FIREBASE_DOC = "live_data";

async function saveStateToFirebase(patch: Partial<FirebaseState> = {}) {
  if (!db || isUpdatingFromFirebase) return;

  try {
    const stateToSave: FirebaseState = {
      deliveryData,
      blogPosts,
      lotJustifications,
      generalNotes,
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
          blogPosts = Array.isArray(data.blogPosts) ? data.blogPosts : [];
          lotJustifications = data.lotJustifications || {};
          generalNotes = data.generalNotes || "";
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

function toDateTimeMaybe(v: any): Date | null {
  v = safeValue(v);
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 1) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes());
    }
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    const iso = new Date(s);
    if (!isNaN(iso.getTime()) && /\d{4}/.test(s)) return iso;
    // dd/mm/yyyy hh:mm or dd/mm/yyyy
    const parts = s.split(/[\s-]/);
    if (parts.length > 0) {
      const dateParts = parts[0].split('/');
      if (dateParts.length === 3) {
        let h = 0, m = 0;
        if (parts.length > 1) {
          const timeParts = parts[1].split(':');
          if (timeParts.length >= 2) {
            h = parseInt(timeParts[0], 10);
            m = parseInt(timeParts[1], 10);
          }
        }
        let a = parseInt(dateParts[0], 10);
        let b = parseInt(dateParts[1], 10);
        let c = parseInt(dateParts[2], 10);
        if (c < 100) c+=2000;
        if (b > 12 && a <= 12) { const tmp = a; a = b; b = tmp; }
        const dt = new Date(c, b - 1, a, h, m);
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
  renderArrivalsTable();
  renderCharts(filteredData);
  renderInventory(filteredData);
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
  const backlog = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "BACKLOG").length;
  const awaitingUnload = dataForStats.filter((d) => normalizeText(d["STATUS"] || "") === "AGUARDANDO DESOVA").length;
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
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("delivered")}">${t("delivered")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${delivered}</div>
        <div class="text-[9px] font-bold text-green-600 dark:text-green-400">${getPercentage(delivered)}</div>
      </div>
    </div>

    <div class="${getCardClasses("AGUARDANDO DESOVA")}" data-status="AGUARDANDO DESOVA">
      <div class="bg-purple-100 dark:bg-purple-900/50 text-purple-600 dark:text-purple-400 rounded-full h-8 w-8 flex items-center justify-center mr-2 flex-shrink-0">
        <i class="fas fa-box text-sm"></i>
      </div>
      <div class="min-w-0">
        <div class="text-slate-500 dark:text-slate-400 text-[9px] font-semibold uppercase tracking-wider truncate" title="${t("awaitingUnload")}">${t("awaitingUnload")}</div>
        <div class="text-lg font-extrabold text-slate-800 dark:text-slate-100">${awaitingUnload}</div>
        <div class="text-[9px] font-bold text-purple-600 dark:text-purple-400">${getPercentage(awaitingUnload)}</div>
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
    
    const arrivalsContent = document.getElementById("arrivals-content");
    arrivalsContent?.classList.toggle("hidden", target !== "arrivals");
    
    const chartsContent = document.getElementById("charts-content");
    chartsContent?.classList.toggle("hidden", target !== "charts");
    
    const timeContent = document.getElementById("time-content");
    timeContent?.classList.toggle("hidden", target !== "time");

    const inventoryContent = document.getElementById("inventory-content");
    inventoryContent?.classList.toggle("hidden", target !== "inventory");

    const newsContent = document.getElementById("news-content");
    newsContent?.classList.toggle("hidden", target !== "news");
    
    if (target === "arrivals") {
      renderArrivalsTable();
    } else if (target === "charts") {
      renderCharts(deliveryData);
    } else if (target === "time") {
      renderTimeTable(deliveryData);
    } else if (target === "inventory") {
      renderInventory(deliveryData);
    } else if (target === "news") {
      renderNewsTab(deliveryData);
    }
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
    "AGUARDANDO DESOVA": "#eab308",
    "BACKLOG": "#f97316",
    "PENDENTE": "#64748b",
    "OUTROS": "#ef4444"
  };

  const statusLabels = [t("delivered"), t("inTransit"), t("awaitingUnload"), "Backlog", t("pending"), t("chartsOther")];
  
  function getStatusIndex(s: string) {
    if (s === "ENTREGUE") return 0;
    if (s === "A CAMINHO") return 1;
    if (s === "AGUARDANDO DESOVA") return 2;
    if (s === "BACKLOG") return 3;
    if (s === "PENDENTE") return 4;
    return 5;
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

  let overallCounts = [0, 0, 0, 0, 0, 0];
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
          <button id="toggle-macro-view-btn" class="absolute top-4 right-4 text-xs font-bold bg-slate-50 text-slate-500 hover:text-blue-600 px-2 py-1 flex items-center justify-center rounded hover:bg-slate-100 transition border border-slate-200 shadow-sm z-10" title="Toggle Macro View">
            <i class="fas fa-layer-group mr-1"></i> Macro View
          </button>
          <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-4 text-center" data-i18n="chartsLotProgressTitle">${t("chartsLotProgressTitle")}</h3>
          <div class="relative h-64 w-full cursor-grab active:cursor-grabbing overflow-x-auto pb-2">
             <div style="min-width: 800px; height: 100%;">
                <canvas id="lotChartCanvas"></canvas>
             </div>
          </div>
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

  const lotStats: Record<string, { total: number; done: number; statusCounts: number[]; carriers: Set<string>; operations: Set<string> }> = {};
  data.forEach((row) => {
    const lot = String(row["LOT"] || "N/A");
    if (!lotStats[lot]) lotStats[lot] = { total: 0, done: 0, statusCounts: [0,0,0,0,0,0], carriers: new Set(), operations: new Set() };
    lotStats[lot].total++;
    const status = normalizeText(row["STATUS"] || "PENDENTE");
    if (status === "ENTREGUE") lotStats[lot].done++;
    lotStats[lot].statusCounts[getStatusIndex(status)]++;
    const carrier = String(row["TRANSPORTATION COMPANY"] || "").trim().toUpperCase();
    if (carrier) lotStats[lot].carriers.add(carrier);
    let operation = String(row["OPERATION SCOPE"] || "").trim().toUpperCase();
    if (operation) {
      if (operation.includes("UNLOAD") || operation.includes("DESOVA")) operation = "UNLOAD";
      else if (operation.includes("SWAP")) operation = "SWAP";
      lotStats[lot].operations.add(operation);
    }
  });

  const sortedLots = Object.keys(lotStats).sort();
  const lotLabels = sortedLots;

  const ctxLot = document.getElementById("lotChartCanvas") as HTMLCanvasElement;
  if (ctxLot) {
    const minW = Math.max(800, lotLabels.length * 40);
    ctxLot.parentElement!.style.minWidth = `${minW}px`;
    
    let chartData, chartOptions;

    if (isMacroView) {
      chartData = {
        labels: lotLabels,
        datasets: statusLabels.map((lbl, idx) => ({
          label: lbl,
          data: sortedLots.map(lot => lotStats[lot].statusCounts[idx]),
          backgroundColor: Object.values(statusColors)[idx],
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
            formatter: (value: number) => value > 0 ? value : ''
          }
        }
      };
    } else {
      const lotData = sortedLots.map((lot) => lotStats[lot].total > 0 ? (lotStats[lot].done / lotStats[lot].total) * 100 : 0);
      chartData = {
        labels: lotLabels,
        datasets: [{
          label: "% " + t("delivered"),
          data: lotData,
          backgroundColor: lotData.map(v => v === 100 ? "#22c55e" : "#3b82f6"),
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
          }
        }
      };
    }

    lotChart = new Chart(ctxLot, { type: "bar", data: chartData as any, options: chartOptions as any });
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
    if (!carrierStats[carrier]) carrierStats[carrier] = [0, 0, 0, 0, 0, 0];
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
    if (!warehouseStats[wh]) warehouseStats[wh] = [0, 0, 0, 0, 0, 0];
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

/* --------------------------------- BLOG ----------------------------------- */
function renderNewsTab(data: DeliveryRow[]) {
  const newsContent = document.getElementById("news-content");
  if (!newsContent) return;

  const sortedPosts = [...blogPosts].sort((a, b) => {
    const dateA = a.createdAt?.toDate ? a.createdAt.toDate().getTime() : (a.createdAt ? new Date(a.createdAt).getTime() : 0);
    const dateB = b.createdAt?.toDate ? b.createdAt.toDate().getTime() : (b.createdAt ? new Date(b.createdAt).getTime() : 0);
    return dateB - dateA;
  });

  const featuredPost = sortedPosts[0];
  const secondaryPosts = sortedPosts.slice(1, 4);
  const remainingPosts = sortedPosts.slice(4);

  const total = data.length;
  const delivered = data.filter(d => normalizeText(d["STATUS"] || "") === "ENTREGUE").length;
  const transit = data.filter(d => normalizeText(d["STATUS"] || "") === "A CAMINHO").length;
  const efficiency = total > 0 ? ((delivered / total) * 100).toFixed(1) : "0.0";
  const uniqueLots = Array.from(new Set(data.map(d => String(d["LOT"] || "N/A")))).sort();

  newsContent.innerHTML = `
    <div class="bg-slate-50 dark:bg-slate-900/50 min-h-screen">
      <div class="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 py-1 px-4 mb-6">
        <div class="max-w-7xl mx-auto flex items-center overflow-hidden">
          <div class="bg-slate-900 text-white px-2 py-0.5 text-[10px] font-black uppercase tracking-widest mr-4 flex-shrink-0">${t("newsBreaking")}</div>
          <div class="text-xs font-bold text-slate-600 dark:text-slate-300 truncate">
            ${sortedPosts.length > 0 ? sortedPosts[0].title : t("newsWaiting")}
          </div>
          <div class="ml-auto flex gap-2">
             <button id="add-news-btn" class="text-blue-600 hover:text-blue-700 font-black text-[10px] uppercase tracking-wider flex items-center">
               <i class="fas fa-plus-circle mr-1"></i> ${t("newsNewUpdate")}
             </button>
          </div>
        </div>
      </div>

      <div class="max-w-7xl mx-auto px-4 pb-12">
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-8">
          <div class="lg:col-span-8 flex flex-col gap-8">
            ${featuredPost ? `
              <div class="relative group cursor-pointer overflow-hidden rounded-2xl shadow-2xl bg-slate-200 aspect-[16/9] lg:aspect-[21/9]">
                ${featuredPost.image ? `<img src="${featuredPost.image}" class="absolute inset-0 w-full h-full object-cover transition-transform duration-700 group-hover:scale-110">` : `<div class="absolute inset-0 flex items-center justify-center bg-slate-800"><i class="fas fa-newspaper text-6xl text-slate-700"></i></div>`}
                <div class="absolute inset-0 bg-gradient-to-t from-black via-black/40 to-transparent"></div>
                <div class="absolute bottom-0 left-0 p-8 w-full">
                  <span class="bg-blue-600 text-white px-2 py-1 text-[10px] font-black uppercase tracking-widest rounded mb-3 inline-block">${featuredPost.category || 'GENERAL'}</span>
                  <h1 class="text-3xl lg:text-5xl font-black text-white leading-tight mb-4">${featuredPost.title}</h1>
                  <p class="text-slate-200 text-sm mb-4 line-clamp-2">${featuredPost.text}</p>
                  <div class="flex items-center text-slate-300 text-xs font-bold gap-4">
                    <span><i class="fas fa-user-edit mr-1"></i> ${featuredPost.author || 'Supervisão'}</span>
                    <span><i class="fas fa-clock mr-1"></i> ${featuredPost.createdAt?.toDate ? featuredPost.createdAt.toDate().toLocaleDateString() : (featuredPost.createdAt ? new Date(featuredPost.createdAt).toLocaleDateString() : t("today"))}</span>
                  </div>
                </div>
                <button class="delete-post-btn absolute top-4 right-4 bg-white/10 hover:bg-red-600 text-white p-2 rounded-full backdrop-blur-md transition-all opacity-0 group-hover:opacity-100" data-id="${featuredPost.id}">
                  <i class="fas fa-trash-alt"></i>
                </button>
              </div>
            ` : `
              <div class="bg-white dark:bg-slate-800 p-20 rounded-2xl border-2 border-dashed border-slate-200 dark:border-slate-700 text-center">
                 <i class="fas fa-newspaper text-5xl text-slate-300 mb-4 block"></i>
                 <p class="text-slate-500 font-bold">${t("newsNoPosts")}</p>
              </div>
            `}

            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
              ${secondaryPosts.map(post => `
                <div class="bg-white dark:bg-slate-800 rounded-xl overflow-hidden shadow-md border border-slate-100 dark:border-slate-700 group hover:shadow-xl transition-all">
                  <div class="relative h-48 bg-slate-200 overflow-hidden">
                    ${post.image ? `<img src="${post.image}" class="absolute inset-0 w-full h-full object-cover group-hover:scale-105 transition-all">` : `<div class="absolute inset-0 flex items-center justify-center bg-slate-100 dark:bg-slate-900"><i class="fas fa-image text-2xl text-slate-300"></i></div>`}
                    <div class="absolute top-3 left-3 bg-slate-900 text-white px-2 py-0.5 text-[9px] font-black uppercase tracking-widest rounded">${post.category || 'GERAL'}</div>
                  </div>
                  <div class="p-5">
                    <h3 class="text-lg font-black text-slate-800 dark:text-slate-100 mb-2 leading-tight">${post.title}</h3>
                    <p class="text-slate-500 dark:text-slate-400 text-xs leading-relaxed line-clamp-3 mb-4">${post.text}</p>
                    <div class="flex justify-between items-center">
                      <span class="text-[10px] text-slate-400 font-bold uppercase tracking-tighter">${post.createdAt?.toDate ? post.createdAt.toDate().toLocaleDateString() : 'Recent'}</span>
                      <button class="delete-post-btn text-red-500 hover:text-red-600 text-xs" data-id="${post.id}"><i class="fas fa-trash-alt"></i></button>
                    </div>
                  </div>
                </div>
              `).join("")}
            </div>

            <div class="flex flex-col gap-4 mt-4">
              <h4 class="text-xs font-black uppercase tracking-widest text-slate-400 border-b border-slate-200 dark:border-slate-700 pb-2">${t("newsOlder")}</h4>
              ${remainingPosts.map(post => `
                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl flex gap-4 items-center shadow-sm border border-slate-100 dark:border-slate-700 group hover:shadow-md transition-shadow">
                  <div class="w-16 h-16 rounded-lg overflow-hidden flex-shrink-0 bg-slate-100 dark:bg-slate-700">
                    ${post.image ? `<img src="${post.image}" class="w-full h-full object-cover">` : `<div class="w-full h-full flex items-center justify-center"><i class="fas fa-image text-slate-300"></i></div>`}
                  </div>
                  <div class="min-w-0 flex-grow">
                    <h5 class="font-bold text-slate-800 dark:text-slate-100 text-sm line-clamp-1">${post.title}</h5>
                    <p class="text-xs text-slate-500 dark:text-slate-400 truncate">${post.text}</p>
                    <div class="text-[9px] text-slate-400 font-bold uppercase mt-1">${post.author || 'Supervisão'}</div>
                  </div>
                  <div class="text-[10px] text-slate-300 font-bold ml-auto">${post.createdAt?.toDate ? post.createdAt.toDate().toLocaleDateString() : (post.createdAt instanceof Date ? post.createdAt.toLocaleDateString() : '')}</div>
                  <button class="delete-post-btn text-slate-300 hover:text-red-500 transition-colors ml-2" data-id="${post.id}"><i class="fas fa-times"></i></button>
                </div>
              `).join("")}
            </div>
          </div>

          <div class="lg:col-span-4 flex flex-col gap-8">
            <div class="bg-slate-900 rounded-2xl p-6 text-white shadow-xl">
              <h3 class="text-xs font-black uppercase tracking-widest text-slate-400 mb-6">${t("newsStayConnected")}</h3>
              <div class="space-y-4">
                <div class="flex items-center justify-between group p-2 rounded">
                  <div class="flex items-center gap-3">
                    <div class="w-10 h-10 bg-blue-600 rounded flex items-center justify-center font-black">OP</div>
                    <div>
                      <div class="text-xs font-bold">${t("newsEffTitle")}</div>
                      <div class="text-[10px] text-blue-400 font-medium">${efficiency}% Successful</div>
                    </div>
                  </div>
                </div>
                <div class="flex items-center justify-between group p-2 rounded">
                  <div class="flex items-center gap-3">
                    <div class="w-10 h-10 bg-emerald-600 rounded flex items-center justify-center font-black">TR</div>
                    <div>
                      <div class="text-xs font-bold">${t("newsTransitTitle")}</div>
                      <div class="text-[10px] text-emerald-400 font-medium">${transit} Units Moving</div>
                    </div>
                  </div>
                </div>
              </div>
              <div class="mt-8 pt-8 border-t border-white/10">
                 <div class="bg-blue-600 hover:bg-blue-700 p-4 rounded-xl text-center font-black text-xs uppercase tracking-widest cursor-pointer transition-all active:scale-95" id="sidebar-add-btn">
                   <i class="fas fa-plus mr-2"></i> ${t("newsPostNew")}
                 </div>
              </div>
            </div>

            <div class="bg-white dark:bg-slate-800 rounded-2xl border-2 border-slate-200 dark:border-slate-700 overflow-hidden">
               <div class="p-6">
                 <h4 class="text-xs font-black tracking-widest text-slate-400 uppercase mb-4">${t("newsLotJust")}</h4>
                 <div class="max-h-[300px] overflow-y-auto pr-2 space-y-4 mb-4 custom-scrollbar">
                    ${uniqueLots.map(lot => `
                      <div>
                        <label class="text-[10px] font-black uppercase text-slate-500 mb-1 block">Lote ${lot}</label>
                        <textarea class="lot-side-note w-full p-3 rounded-lg border border-slate-100 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-xs text-slate-700 dark:text-slate-300 resize-none focus:ring-2 focus:ring-blue-500 outline-none transition-all" rows="2" data-lot="${lot}" placeholder="Nota do lote...">${lotJustifications[lot] || ""}</textarea>
                      </div>
                    `).join("")}
                 </div>
                 <button id="save-lot-notes-btn" class="w-full bg-slate-900 dark:bg-slate-700 text-white py-3 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-black transition-all active:scale-95">
                    ${t("newsSaveNotes")} <i class="fas fa-save ml-2"></i>
                 </button>
               </div>
            </div>

            <div class="bg-white dark:bg-slate-800 rounded-2xl border-2 border-slate-200 dark:border-slate-700 overflow-hidden">
               <div class="p-6">
                 <h4 class="text-xs font-black tracking-widest text-slate-400 uppercase mb-4">${t("newsAddInfo")}</h4>
                 <textarea id="general-operation-notes" class="w-full p-4 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-sm text-slate-700 dark:text-slate-300 min-h-[150px] resize-none focus:ring-2 focus:ring-blue-500 outline-none transition-all" placeholder="Qualquer informação necessária...">${generalNotes}</textarea>
                 <button id="save-general-notes-btn" class="w-full mt-4 bg-blue-600 text-white py-3 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg shadow-blue-500/20 active:scale-95">
                    ${t("newsSaveAlerts")} <i class="fas fa-check ml-2"></i>
                 </button>
               </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div id="blog-modal" class="hidden fixed inset-0 z-[1000] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div class="bg-white dark:bg-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        <div class="bg-slate-900 p-6 flex justify-between items-center">
           <h3 class="text-white font-black uppercase tracking-widest text-sm">${t("newsModalTitle")}</h3>
           <button class="text-white/50 hover:text-white" id="close-blog-modal"><i class="fas fa-times text-xl"></i></button>
        </div>
        <div class="p-8 flex-grow overflow-y-auto">
          <div class="space-y-6">
            <div>
              <label class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">Título do Post</label>
              <input id="post-title" type="text" class="w-full p-4 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 font-bold outline-none">
            </div>
            <div class="grid grid-cols-2 gap-4">
               <div>
                  <label class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">Categoria</label>
                  <select id="post-category" class="w-full p-4 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 font-bold outline-none">
                    <option value="GERAL">GERAL</option>
                    <option value="OPERAÇÃO">OPERAÇÃO</option>
                    <option value="ALERTA">ALERTA</option>
                    <option value="SUCESSO">SUCESSO</option>
                  </select>
               </div>
               <div>
                  <label class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">Autor</label>
                  <input id="post-author" type="text" class="w-full p-4 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 font-bold outline-none" value="Supervisão">
               </div>
            </div>
            <div>
              <label class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">${t("newsPostPhoto")}</label>
              <label class="w-full h-40 border-2 border-dashed border-slate-200 dark:border-slate-700 rounded-2xl flex flex-col items-center justify-center cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-900 transition-all overflow-hidden relative">
                 <div id="image-preview-area" class="text-center">
                    <i class="fas fa-camera text-3xl text-slate-300 mb-2"></i>
                    <p class="text-[10px] font-bold text-slate-400">${t("newsClickToPhoto")}</p>
                 </div>
                 <img id="post-image-preview" class="hidden absolute inset-0 w-full h-full object-cover">
                 <input type="file" id="post-image-input" class="hidden" accept="image/*">
              </label>
            </div>
            <div>
              <label class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">Conteúdo</label>
              <textarea id="post-text" class="w-full p-4 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 outline-none resize-none min-h-[200px]"></textarea>
            </div>
          </div>
        </div>
        <div class="p-6 bg-slate-50 dark:bg-slate-900 border-t border-slate-200 dark:border-slate-700 flex gap-4">
           <button id="cancel-post-btn" class="flex-1 py-4 font-black uppercase tracking-widest text-sm text-slate-500 hover:text-slate-800">Descartar</button>
           <button id="save-post-btn" class="flex-[2] bg-blue-600 hover:bg-blue-700 text-white py-4 rounded-2xl font-black uppercase tracking-widest text-sm">${t("newsPublish")} <i class="fas fa-paper-plane ml-2"></i></button>
        </div>
      </div>
    </div>
  `;

  // Attach dynamic event triggers inside structural template literals
  const modal = document.getElementById("blog-modal");
  const imgInput = document.getElementById("post-image-input") as HTMLInputElement;
  const imgPreview = document.getElementById("post-image-preview") as HTMLImageElement;
  const previewArea = document.getElementById("image-preview-area");
  let currentPostImage: string | undefined = undefined;

  const hideModal = () => {
    modal?.classList.add("hidden");
    currentPostImage = undefined;
    if(imgPreview) imgPreview.classList.add("hidden");
    if(previewArea) previewArea.classList.remove("hidden");
    if(imgInput) imgInput.value = "";
  };

  document.getElementById("add-news-btn")?.addEventListener("click", () => modal?.classList.remove("hidden"));
  document.getElementById("sidebar-add-btn")?.addEventListener("click", () => modal?.classList.remove("hidden"));
  document.getElementById("close-blog-modal")?.addEventListener("click", hideModal);
  document.getElementById("cancel-post-btn")?.addEventListener("click", hideModal);

  imgInput?.addEventListener("change", (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (re) => {
        const img = new Image();
        img.src = re.target?.result as string;
        img.onload = () => {
          const canvas = document.createElement("canvas");
          const MAX_WIDTH = 1000;
          let width = img.width;
          let height = img.height;
          if (width > MAX_WIDTH) { height *= MAX_WIDTH / width; width = MAX_WIDTH; }
          canvas.width = width; canvas.height = height;
          const ctx = canvas.getContext("2d");
          ctx?.drawImage(img, 0, 0, width, height);
          currentPostImage = canvas.toDataURL("image/jpeg", 0.7);
          if(imgPreview) { imgPreview.src = currentPostImage; imgPreview.classList.remove("hidden"); }
          if(previewArea) previewArea.classList.add("hidden");
        }
      };
      reader.readAsDataURL(file);
    }
  });

  document.getElementById("save-post-btn")?.addEventListener("click", async () => {
    const title = (document.getElementById("post-title") as HTMLInputElement).value;
    const text = (document.getElementById("post-text") as HTMLTextAreaElement).value;
    const category = (document.getElementById("post-category") as HTMLSelectElement).value;
    const author = (document.getElementById("post-author") as HTMLInputElement).value;

    if (!title || !text) return showToast("Por favor, preencha o título e o conteúdo.", "warning");

    blogPosts.unshift({ 
      id: Date.now().toString(), 
      title, 
      text, 
      category, 
      author, 
      image: currentPostImage,
      createdAt: new Date() 
    });
    showToast("Notícia publicada com sucesso!", "success");
    hideModal();
    await saveStateToFirebase();
    renderNewsTab(deliveryData);
  });

  document.getElementById("save-lot-notes-btn")?.addEventListener("click", async () => {
    const lotNotes: Record<string, string> = {};
    newsContent.querySelectorAll(".lot-side-note").forEach(ta => {
      const el = ta as HTMLTextAreaElement;
      if (el.dataset.lot) lotNotes[el.dataset.lot] = el.value;
    });
    lotJustifications = { ...lotJustifications, ...lotNotes };
    showToast("Notas dos lotes salvas com sucesso!", "success");
    await saveStateToFirebase();
  });

  document.getElementById("save-general-notes-btn")?.addEventListener("click", async () => {
    generalNotes = (document.getElementById("general-operation-notes") as HTMLTextAreaElement).value;
    showToast("Avisos gerais salvos!", "success");
    await saveStateToFirebase();
  });

  newsContent.querySelectorAll(".delete-post-btn").forEach(btn => {
    btn.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!confirm("Tem certeza que deseja excluir esta notícia?")) return;
      const id = (e.currentTarget as HTMLElement).dataset.id;
      blogPosts = blogPosts.filter(p => p.id !== id);
      await saveStateToFirebase();
      renderNewsTab(deliveryData);
    });
  });
}

/* ----------------------- ARRIVALS & METRICS ---------------------------- */
function renderArrivalsTable() {
  const arrivalsContent = document.getElementById("arrivals-content");
  if (!arrivalsContent) return;

  const lotsFromData = Array.from(new Set(deliveryData.map((d) => String(d["LOT"] || "N/A")))).sort();
  const statuses = ["A CAMINHO", "ADIADO", "AGUARDANDO DESOVA", "ENTREGUE"];
  const statusKeys: Record<string, string> = {
    "A CAMINHO": "STATUS_A_CAMINHO", "ADIADO": "STATUS_ADIADO", "AGUARDANDO DESOVA": "STATUS_AGUARDANDO_DESOVA", "ENTREGUE": "STATUS_ENTREGUE"
  };

  arrivalsContent.innerHTML = `
    <div class="overflow-x-auto bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700">
      <table class="w-full text-xs text-left text-slate-600 dark:text-slate-300">
        <thead class="bg-slate-50 dark:bg-slate-700 border-b border-slate-200 dark:border-slate-600">
          <tr>
            <th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100 uppercase tracking-wider">${t("tableHeaderLot")}</th>
            ${statuses.map((s) => `<th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100 uppercase tracking-wider">${t(statusKeys[s] as TranslationKey)}</th>`).join("")}
            <th class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100 uppercase tracking-wider">${t("tableHeaderTotal")}</th>
          </tr>
        </thead>
        <tbody class="divide-y divide-slate-100 dark:divide-slate-700">
          ${lotsFromData.map((lot) => {
              const deliveriesInLot = deliveryData.filter((d) => String(d["LOT"] || "N/A") === lot);
              const statusCounts = statuses.reduce((acc, s) => {
                acc[s] = deliveriesInLot.filter((d) => normalizeText(d["STATUS"] || "") === normalizeText(s)).length;
                return acc;
              }, {} as Record<string, number>);
              const total = Object.values(statusCounts).reduce((a, b) => a + b, 0);
              if (total === 0) return "";
              const models = Array.from(new Set(deliveriesInLot.map(d => String(d["MODEL"] || "")).filter(m => m !== ""))).join(", ");
              return `<tr>
                <td class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">${models ? `${lot} - ${models}` : lot}</td>
                ${statuses.map((s) => `<td class="px-4 py-3">${statusCounts[s] > 0 ? statusCounts[s] : ""}</td>`).join("")}
                <td class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100">${total}</td>
              </tr>`;
            }).join("")}
        </tbody>
        <tfoot class="bg-slate-50 dark:bg-slate-700 border-t border-slate-200 dark:border-slate-600 font-bold">
          <tr>
            <td class="px-4 py-3 text-slate-800 dark:text-slate-100">${t("tableHeaderOverallTotal")}</td>
            ${statuses.map((s) => `<td class="px-4 py-3 text-slate-800 dark:text-slate-100">${deliveryData.filter((d) => normalizeText(d["STATUS"] || "") === normalizeText(s)).length}</td>`).join("")}
            <td class="px-4 py-3 text-slate-800 dark:text-slate-100">${deliveryData.length}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
}

function renderTimeTable(data: DeliveryRow[]) {
  const timeContent = document.getElementById("time-content");
  if (!timeContent) return;

  let totalTimeSum = 0, validRecords = 0;
  let totalTimeSumP1 = 0, validRecordsP1 = 0;
  let totalTimeSumP2 = 0, validRecordsP2 = 0;

  const rowsHtml = data.map((row) => {
    const startDt = toDateTimeMaybe(row["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA."]);
    let endDt = toDateTimeMaybe(row["DATA E HORARIO DE ENTREGA CONTAINER VAZIO"]) || toDateTimeMaybe(row["DATA E HORARIO DE DESCARGA NA BYD "]);
    let fullTimeString = "-", durationHours = 0;

    if (startDt && endDt) {
      const diffMs = endDt.getTime() - startDt.getTime();
      if (diffMs > 0) {
        durationHours = diffMs / (1000 * 60 * 60);
        const dDays = Math.floor(durationHours / 24);
        const dHours = Math.floor(durationHours % 24);
        const dMins = Math.round((durationHours - Math.floor(durationHours)) * 60);
        fullTimeString = dDays > 0 ? `${dDays}d ${dHours}h ${dMins}m` : `${dHours}h ${dMins}m`;

        totalTimeSum += durationHours; validRecords++;
        const timeVal = startDt.getHours() * 100 + startDt.getMinutes();
        if (timeVal >= 630 && timeVal <= 1500) { totalTimeSumP1 += durationHours; validRecordsP1++; }
        else if (timeVal >= 1501 || (startDt.getHours() === 0 && startDt.getMinutes() === 0)) { totalTimeSumP2 += durationHours; validRecordsP2++; }
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

  const formatAvg = (sum: number, count: number) => {
    if (count === 0) return "-";
    const avgH = sum / count;
    const aDays = Math.floor(avgH / 24), aHours = Math.floor(avgH % 24), aMins = Math.round((avgH - Math.floor(avgH)) * 60);
    return aDays > 0 ? `${aDays}d ${aHours}h ${aMins}m` : `${aHours}h ${aMins}m`;
  };

  timeContent.innerHTML = `
    <div class="mb-6 bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 p-4 flex flex-col md:flex-row justify-between items-center gap-4">
       <div>
         <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-1" data-i18n="timeTab">${t("timeTab")}</h3>
         <p class="text-xs text-slate-500 dark:text-slate-400">Registros válidos: <strong class="text-slate-700 dark:text-slate-200">${validRecords}</strong></p>
       </div>
       <div class="flex flex-wrap gap-4 justify-end">
         <div class="bg-blue-50 dark:bg-blue-900/30 px-4 py-2 rounded-lg border border-blue-100">
           <span class="text-[9px] font-bold text-blue-500 block uppercase">Média Geral</span>
           <span class="text-base font-black text-blue-700 dark:text-blue-300">${formatAvg(totalTimeSum, validRecords)}</span>
         </div>
         <div class="bg-emerald-50 dark:bg-emerald-900/30 px-4 py-2 rounded-lg border border-emerald-100">
           <span class="text-[9px] font-bold text-emerald-600 block uppercase">${t("avgPeriod1")}</span>
           <span class="text-base font-black text-emerald-700 dark:text-emerald-300">${formatAvg(totalTimeSumP1, validRecordsP1)}</span>
         </div>
         <div class="bg-amber-50 dark:bg-amber-900/30 px-4 py-2 rounded-lg border border-amber-100">
           <span class="text-[9px] font-bold text-amber-600 block uppercase">${t("avgPeriod2")}</span>
           <span class="text-base font-black text-amber-700 dark:text-amber-300">${formatAvg(totalTimeSumP2, validRecordsP2)}</span>
         </div>
       </div>
    </div>
    <div class="overflow-x-auto bg-white dark:bg-slate-800 rounded-lg shadow-sm border border-slate-200 dark:border-slate-700 max-h-[400px]">
      <table class="w-full text-xs text-left"><tbody class="divide-y">${rowsHtml}</tbody></table>
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
        DELIVERY_AT_BYD: pickIndex(headerIndex, ["DELIVERY AT BYD", "DELIVERY", "DATA DE ENTREGA"]),
        UNLOAD_TIME_BYD: pickIndex(headerIndex, ["UNLOAD TIME BYD", "UNLOAD TIME"]),
        TRANSPORTATION_COMPANY: pickIndex(headerIndex, ["TRANSPORTATION COMPANY", "CARRIER", "TRANSPORTADORA"]),
        CONTAINER: pickIndex(headerIndex, ["CONTAINER"]),
        BL: pickIndex(headerIndex, ["BL", "B/L"]),
        VESSEL: pickIndex(headerIndex, ["VESSEL", "NAVIO"]),
        BONDED_WAREHOUSE: pickIndex(headerIndex, ["BONDED WAREHOUSE", "ARMAZEM"]),
        MODEL: pickIndex(headerIndex, ["MODEL", "MODELO"]),
        LOT: pickIndex(headerIndex, ["LOT", "LOTE"]),
        TYPE_OF_MATERIAL: pickIndex(headerIndex, ["TYPE OF MATERIAL", "MATERIAL"]),
        STATUS: pickIndex(headerIndex, ["STATUS", "SITUACAO"]),
        TERMINAL_DEPARTURE: pickIndex(headerIndex, ["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA."]),
        EMPTY_DELIVERED: pickIndex(headerIndex, ["DATA E HORARIO DE ENTREGA CONTAINER VAZIO"]),
        UNLOAD_AT_BYD: pickIndex(headerIndex, ["DATA E HORARIO DE DESCARGA NA BYD "]),
        NOTES: pickIndex(headerIndex, ["NOTES", "OBSERVACOES", "OBSERVAÇÕES"]),
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
        obj["LOT"] = col.LOT >= 0 ? safeValue(r[col.LOT]) : "";
        obj["TYPE OF MATERIAL"] = col.TYPE_OF_MATERIAL >= 0 ? safeValue(r[col.TYPE_OF_MATERIAL]) : "";
        obj["DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA."] = col.TERMINAL_DEPARTURE >= 0 ? safeValue(r[col.TERMINAL_DEPARTURE]) : "";
        obj["DATA E HORARIO DE ENTREGA CONTAINER VAZIO"] = col.EMPTY_DELIVERED >= 0 ? safeValue(r[col.EMPTY_DELIVERED]) : "";
        obj["DATA E HORARIO DE DESCARGA NA BYD "] = col.UNLOAD_AT_BYD >= 0 ? safeValue(r[col.UNLOAD_AT_BYD]) : "";
        obj["NOTES"] = col.NOTES >= 0 ? safeValue(r[col.NOTES]) : "";
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

/* ------------------------------- EXPORTS ---------------------------------- */
exportExcelBtn?.addEventListener("click", async () => {
  if (!deliveryData || deliveryData.length === 0) return showToast(t("noDataToExport"), "warning");

  const exportColumns = [
    "STATUS", "DELIVERY AT BYD", "CONTAINER", "BL", "LOT", "MODEL", "OPERATION SCOPE", "TRANSPORTATION COMPANY", 
    "VESSEL", "BONDED WAREHOUSE", "DATA E HORRÁRIO DA SAÍDA DO TERMINAL - INICIO DA ROTA NA PISTA EXPRESSA.", 
    "DATA E HORARIO DE DESCARGA NA BYD ", "DATA E HORARIO DE ENTREGA CONTAINER VAZIO", "UNLOAD TIME BYD", "NOTES"
  ];

  const ws = XLSX.utils.aoa_to_sheet([exportColumns, ...deliveryData.map(d => exportColumns.map(col => d[col] ?? ""))]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, t("deliveriesTab"));

  // Append Standardized Minutes layout style logic inside spreadsheet generation
  try {
    const lotsFromData = Array.from(new Set(deliveryData.map((d) => String(d["LOT"] || "N/A")))).sort();
    const statuses = ["A CAMINHO", "ADIADO", "AGUARDANDO DESOVA", "ENTREGUE"];
    
    const arrivalsOut = lotsFromData.map((lot) => {
      const deliveriesInLot = deliveryData.filter((d) => String(d["LOT"] || "N/A") === lot);
      const row: any = { "Lote / Modelo": lot };
      statuses.forEach(s => row[s] = deliveriesInLot.filter(d => normalizeText(d["STATUS"] || "") === normalizeText(s)).length || "");
      return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(arrivalsOut), t("arrivalsTab"));
  } catch(e) { console.error(e); }

  // Append Inventory tab data to master excel file seamlessly
  try {
    const inventoryData = (window as any).inventoryExportData;
    if (inventoryData && inventoryData.sections) {
      const inventoryRows = [["CATEGORIA", "LOCALIZAÇÃO", "VAZIO", "CHEIO", "TOTAL", "CAPACIDADE"]];
      inventoryData.sections.forEach((sec: any) => {
        sec.locations.forEach((loc: any) => {
          inventoryRows.push([sec.title, loc.name, loc.empty, loc.full, (loc.empty + loc.full), loc.capacity]);
        });
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(inventoryRows), t("inventoryTab"));
    }
  } catch (e) { console.error(e); }

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

/* ------------------------------ STARTUP ----------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  setTheme(((localStorage.getItem("theme") as any) || "light") as any);
  loadLogoFromStorage();
  setLanguage((localStorage.getItem("language") as Language) || "pt-BR");
  listenForRealtimeUpdates();
  resetUI();
});
