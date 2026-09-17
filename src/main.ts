import {
  FieldType,
  ToastType,
  bitable,
  type IAttachmentField,
  type IDateTimeField,
  type IGridView,
  type INumberField,
  type IOpenAttachment,
  type ITextField,
} from "@lark-base-open/js-sdk";
import { PDFDocument } from "pdf-lib";
import "./style.css";

const FIELD = {
  reason: "报销事由",
  category: "报销类目",
  header: "发票抬头",
  approval: "审批",
  earliestDate: "最早开票日期",
  latestDate: "最晚开票日期",
  amount: "报销金额",
  source: "发票+订单截图+支付截图",
  batch: "PDF合并批次",
  result: "合并后材料PDF",
} as const;

const MAX_FILES = 30;
const MAX_PAGES = 30;
const MAX_TOTAL_BYTES = 80 * 1024 * 1024;
const SELECTION_POLL_MS = 800;
const VIEW_ORDER_CACHE_MS = 30_000;

const app = document.querySelector<HTMLElement>("#app");
if (!app) throw new Error("页面初始化失败");

app.innerHTML = `
  <section class="shell">
    <header>
      <span class="eyebrow">X-LAB 财务工具</span>
      <h1>合并所选 PDF</h1>
      <p>勾选记录后自动统计每条材料页数；抬头、类目、日期、金额和审批状态校验通过后即可合并。</p>
    </header>
    <div class="card summary">
      <div><span class="label">当前选择</span><strong id="selected-count">正在读取……</strong></div>
      <button id="refresh" class="secondary" type="button">重新检查</button>
    </div>
    <div id="selection-detail" class="card detail muted">请在当前表格左侧勾选记录。</div>
    <div class="rules">
      <span>✓ 实时统计合并页数</span><span>✓ 审批均为“同意”</span>
      <span>✓ 发票抬头一致</span><span>✓ 报销类目一致</span>
    </div>
    <button id="merge" class="primary" type="button" disabled>合并并写回第一条记录</button>
    <div id="status" class="status" aria-live="polite"></div>
    <p class="footnote">文件仅在当前浏览器内存中处理，不上传到第三方合并服务。</p>
  </section>`;

const selectedCount = document.querySelector<HTMLElement>("#selected-count")!;
const detail = document.querySelector<HTMLElement>("#selection-detail")!;
const status = document.querySelector<HTMLElement>("#status")!;
const refreshButton = document.querySelector<HTMLButtonElement>("#refresh")!;
const mergeButton = document.querySelector<HTMLButtonElement>("#merge")!;

type RecordDetail = {
  recordId: string;
  rowNumber: string;
  reason: string;
  pdfCount: number;
  pageCount: number;
};

type SelectionState = {
  recordIds: string[];
  category: string;
  header: string;
  rowNumbers: string[];
  records: RecordDetail[];
  fileCount: number;
  totalPages: number;
  earliestDate: number;
  latestDate: number;
  totalAmount: number;
};

type CachedPdf = { bytes: Uint8Array; pageCount: number };

let state: SelectionState | null = null;
let refreshVersion = 0;
let lastSelectionSignature = "";
let pollingSelection = false;
let merging = false;
const pdfCache = new Map<string, CachedPdf>();
const pdfPromiseCache = new Map<string, Promise<CachedPdf>>();
const viewOrderCache = new Map<string, { recordIds: string[]; expiresAt: number }>();
let fieldsCache: Awaited<ReturnType<typeof resolveFields>> | null = null;
let refreshTimer: number | undefined;

function setStatus(message: string, kind: "info" | "success" | "error" = "info") {
  status.className = `status ${kind}`;
  status.textContent = message;
}

function safeText(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;",
  })[char]!);
}

function formatDate(timestamp: number): string {
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date(timestamp));
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}.${values.month}.${values.day}`;
}

function formatFilenameDate(timestamp: number): string {
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date(timestamp));
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.month}.${values.day}`;
}

function cleanFilenamePart(value: string): string {
  return value.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "").slice(0, 40);
}

function buildFilename(selection: SelectionState): string {
  return `${cleanFilenamePart(selection.header)}+${cleanFilenamePart(selection.category)}+【${formatFilenameDate(selection.earliestDate)}-${formatFilenameDate(selection.latestDate)}】${selection.totalAmount.toFixed(2)}元.pdf`;
}

function pdfCacheKey(recordId: string, attachment: IOpenAttachment): string {
  return `${recordId}:${attachment.token}:${attachment.size}:${attachment.timeStamp}`;
}

async function resolveFields(table: Awaited<ReturnType<typeof bitable.base.getActiveTable>>) {
  const [reason, category, header, approval, earliestDate, latestDate, amount, source, batch, result] = await Promise.all([
    table.getFieldByName(FIELD.reason),
    table.getFieldByName(FIELD.category),
    table.getFieldByName(FIELD.header),
    table.getFieldByName(FIELD.approval),
    table.getFieldByName<IDateTimeField>(FIELD.earliestDate),
    table.getFieldByName<IDateTimeField>(FIELD.latestDate),
    table.getFieldByName<INumberField>(FIELD.amount),
    table.getFieldByName<IAttachmentField>(FIELD.source),
    table.getFieldByName<ITextField>(FIELD.batch),
    table.getFieldByName<IAttachmentField>(FIELD.result),
  ]);
  if (await source.getType() !== FieldType.Attachment) throw new Error(`“${FIELD.source}”必须是附件字段`);
  if (await result.getType() !== FieldType.Attachment) throw new Error(`“${FIELD.result}”必须是附件字段`);
  return { table, reason, category, header, approval, earliestDate, latestDate, amount, source, batch, result };
}

async function getFields() {
  const table = await bitable.base.getActiveTable();
  if (fieldsCache?.table.id === table.id) return fieldsCache;
  fieldsCache = await resolveFields(table);
  return fieldsCache;
}

async function getOrderedSelection(
  table: Awaited<ReturnType<typeof bitable.base.getActiveTable>>,
  view: IGridView,
  selectedRecordIds: string[],
) {
  const cacheKey = `${table.id}:${view.id}`;
  const cachedOrder = viewOrderCache.get(cacheKey);
  let orderedRecordIds: string[];
  if (cachedOrder && cachedOrder.expiresAt > Date.now()) {
    orderedRecordIds = cachedOrder.recordIds;
  } else {
    orderedRecordIds = [];
    let pageToken: number | undefined;
    do {
      const page = await table.getRecordIdListByPage({ pageSize: 200, pageToken, viewId: view.id });
      orderedRecordIds.push(...page.recordIds);
      pageToken = page.hasMore ? page.pageToken : undefined;
    } while (pageToken !== undefined);
    viewOrderCache.set(cacheKey, { recordIds: orderedRecordIds, expiresAt: Date.now() + VIEW_ORDER_CACHE_MS });
  }

  const selectedSet = new Set(selectedRecordIds);
  const recordIds = orderedRecordIds.filter((recordId) => selectedSet.has(recordId));
  recordIds.push(...selectedRecordIds.filter((recordId) => !recordIds.includes(recordId)));
  return {
    recordIds,
    rowNumbers: recordIds.map((recordId) => {
      const index = orderedRecordIds.indexOf(recordId);
      return index >= 0 ? String(index + 1) : recordId;
    }),
  };
}

async function loadPdf(recordId: string, attachment: IOpenAttachment, url: string): Promise<CachedPdf> {
  const key = pdfCacheKey(recordId, attachment);
  const cached = pdfCache.get(key);
  if (cached) return cached;
  const pending = pdfPromiseCache.get(key);
  if (pending) return pending;

  const request = (async () => {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`附件下载失败：${attachment.name}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const pdf = await PDFDocument.load(bytes, { ignoreEncryption: false });
    const loaded = { bytes, pageCount: pdf.getPageCount() };
    pdfCache.set(key, loaded);
    return loaded;
  })().finally(() => pdfPromiseCache.delete(key));
  pdfPromiseCache.set(key, request);
  return request;
}

async function getRecordPdfInfo(recordId: string, source: IAttachmentField, attachments: IOpenAttachment[]) {
  const urls = await source.getAttachmentUrls(recordId);
  if (attachments.length !== urls.length) throw new Error("附件读取数量不一致，请重新检查");
  const pdfs = await Promise.all(attachments.map((attachment, index) => loadPdf(recordId, attachment, urls[index])));
  return { pdfCount: pdfs.length, pageCount: pdfs.reduce((sum, pdf) => sum + pdf.pageCount, 0) };
}

async function readSelection(): Promise<SelectionState> {
  const { table, reason, category, header, approval, earliestDate, latestDate, amount, source } = await getFields();
  const view = await table.getActiveView() as IGridView;
  if (typeof view.getSelectedRecordIdList !== "function") throw new Error("请在表格视图中使用本插件");

  const selectedRecordIds = await view.getSelectedRecordIdList();
  if (selectedRecordIds.length === 0) throw new Error("请至少勾选一条记录");
  const { recordIds, rowNumbers } = await getOrderedSelection(table, view, selectedRecordIds);
  const rows = await Promise.all(recordIds.map(async (recordId, index) => {
    const [reasonText, categoryText, headerText, approvalText, earliestValue, latestValue, amountValue, files] = await Promise.all([
      reason.getCellString(recordId), category.getCellString(recordId), header.getCellString(recordId), approval.getCellString(recordId),
      earliestDate.getValue(recordId) as Promise<number | null>,
      latestDate.getValue(recordId) as Promise<number | null>,
      amount.getValue(recordId) as Promise<number | null>,
      source.getValue(recordId),
    ]);
    return {
      recordId, rowNumber: rowNumbers[index], reason: reasonText.trim() || "未填写报销事由",
      category: categoryText.trim(), header: headerText.trim(), approval: approvalText.trim(),
      earliestDate: earliestValue, latestDate: latestValue, amount: amountValue, files,
    };
  }));

  if (rows.some((row) => row.approval !== "同意")) throw new Error("所选记录中存在尚未审批同意的记录");
  const categories = new Set(rows.map((row) => row.category));
  if (categories.size !== 1 || categories.has("")) throw new Error("所选记录的报销类目必须一致且不能为空");
  const headers = new Set(rows.map((row) => row.header));
  if (headers.size !== 1 || headers.has("")) throw new Error("所选记录的发票抬头必须一致且不能为空");
  if (rows.some((row) => row.files.length === 0)) throw new Error("所选记录中存在未上传报销材料的记录");

  const files = rows.flatMap((row) => row.files);
  if (files.length > MAX_FILES) throw new Error(`单次最多处理 ${MAX_FILES} 个 PDF`);
  if (files.some((file) => !file.name.toLowerCase().endsWith(".pdf"))) throw new Error("所选记录中包含非 PDF 附件");
  if (files.reduce((sum, file) => sum + file.size, 0) > MAX_TOTAL_BYTES) throw new Error("附件总量超过80MB，请拆分合并");
  if (rows.some((row) => !Number.isFinite(row.earliestDate) || !Number.isFinite(row.latestDate))) {
    throw new Error("请先补全所选记录的最早、最晚开票日期");
  }
  if (rows.some((row) => !Number.isFinite(row.amount))) throw new Error("请先补全所选记录的报销金额");
  if (rows.some((row) => (row.earliestDate as number) > (row.latestDate as number))) {
    throw new Error("存在最早开票日期晚于最晚开票日期的记录");
  }

  const pdfInfo = await Promise.all(rows.map((row) => getRecordPdfInfo(row.recordId, source, row.files)));
  const records = rows.map((row, index) => ({
    recordId: row.recordId, rowNumber: row.rowNumber, reason: row.reason,
    pdfCount: pdfInfo[index].pdfCount, pageCount: pdfInfo[index].pageCount,
  }));
  return {
    recordIds, category: rows[0].category, header: rows[0].header, rowNumbers, records,
    fileCount: files.length,
    totalPages: records.reduce((sum, record) => sum + record.pageCount, 0),
    earliestDate: Math.min(...rows.map((row) => row.earliestDate as number)),
    latestDate: Math.max(...rows.map((row) => row.latestDate as number)),
    totalAmount: Math.round(rows.reduce((sum, row) => sum + (row.amount as number), 0) * 100) / 100,
  };
}

async function refreshSelection() {
  if (merging) return;
  const version = ++refreshVersion;
  mergeButton.disabled = true;
  setStatus("正在读取材料页数……");
  try {
    const nextState = await readSelection();
    if (version !== refreshVersion) return;
    state = nextState;
    selectedCount.textContent = `${state.recordIds.length} 条记录 · ${state.fileCount} 个 PDF · 共 ${state.totalPages} 页`;
    detail.classList.remove("muted");
    detail.innerHTML = `
      <dl>
        <div><dt>发票抬头</dt><dd>${safeText(state.header)}</dd></div>
        <div><dt>报销类目</dt><dd>${safeText(state.category)}</dd></div>
        <div><dt>日期范围</dt><dd>${formatDate(state.earliestDate)} - ${formatDate(state.latestDate)}</dd></div>
        <div><dt>合计金额</dt><dd>¥${state.totalAmount.toFixed(2)}</dd></div>
        <div><dt>文件名</dt><dd class="filename">${safeText(buildFilename(state))}</dd></div>
      </dl>
      <ol class="record-list">${state.records.map((record) => `
        <li><strong>第${safeText(record.rowNumber)}条</strong><span>${record.pdfCount}个 PDF · ${record.pageCount}页</span><small>${safeText(record.reason)}</small></li>
      `).join("")}</ol>`;

    if (state.recordIds.length < 2) {
      setStatus(`当前材料共${state.totalPages}页，请继续勾选需要合并的记录。`);
    } else if (state.totalPages > MAX_PAGES) {
      mergeButton.disabled = false;
      setStatus(`当前共${state.totalPages}页，超过建议的30页，但仍可继续合并。`);
    } else {
      mergeButton.disabled = false;
      setStatus(`页数校验通过：共${state.totalPages}页，可以合并。`, "success");
    }
  } catch (error) {
    if (version !== refreshVersion) return;
    state = null;
    selectedCount.textContent = "未满足合并条件";
    detail.classList.add("muted");
    detail.textContent = error instanceof Error ? error.message : "读取选择失败";
    setStatus("请调整勾选记录或补全信息。", "error");
  }
}

async function mergeSelected() {
  merging = true;
  mergeButton.disabled = true;
  refreshButton.disabled = true;
  setStatus("正在复核页数并合并 PDF……");
  try {
    const latest = await readSelection();
    if (latest.recordIds.length < 2) throw new Error("请至少勾选两条记录");
    const { source, batch, result } = await getFields();
    const output = await PDFDocument.create();
    let mergedFiles = 0;
    for (const recordId of latest.recordIds) {
      const [attachments, urls] = await Promise.all([source.getValue(recordId), source.getAttachmentUrls(recordId)]);
      if (attachments.length !== urls.length) throw new Error("附件读取数量不一致，请重新检查");
      for (let index = 0; index < urls.length; index += 1) {
        setStatus(`正在处理第 ${mergedFiles + 1}/${latest.fileCount} 个 PDF……`);
        const cached = await loadPdf(recordId, attachments[index], urls[index]);
        const sourcePdf = await PDFDocument.load(cached.bytes, { ignoreEncryption: false });
        const pages = await output.copyPages(sourcePdf, sourcePdf.getPageIndices());
        pages.forEach((page) => output.addPage(page));
        mergedFiles += 1;
      }
    }
    setStatus("正在按新文件名写回飞书……");
    const mergedBytes = await output.save({ useObjectStreams: true });
    const buffer = new ArrayBuffer(mergedBytes.byteLength);
    new Uint8Array(buffer).set(mergedBytes);
    const filename = buildFilename(latest);
    await result.setValue(latest.recordIds[0], new File([buffer], filename, { type: "application/pdf" }));
    await batch.setValue(latest.recordIds[0], `${latest.rowNumbers.join(" ")}合并`);
    await bitable.ui.showToast({ toastType: ToastType.success, message: "PDF 已合并并写回" });
    state = latest;
    setStatus(`完成：${latest.totalPages}页，文件名为 ${filename}`, "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : "合并失败";
    await bitable.ui.showToast({ toastType: ToastType.error, message });
    setStatus(message, "error");
  } finally {
    merging = false;
    refreshButton.disabled = false;
    mergeButton.disabled = !state || state.recordIds.length < 2;
  }
}

async function pollSelection() {
  if (pollingSelection || merging) return;
  pollingSelection = true;
  try {
    const table = await bitable.base.getActiveTable();
    const view = await table.getActiveView() as IGridView;
    if (typeof view.getSelectedRecordIdList !== "function") return;
    const selectedIds = await view.getSelectedRecordIdList();
    const signature = `${table.id}:${view.id}:${[...selectedIds].sort().join(",")}`;
    if (signature !== lastSelectionSignature) {
      lastSelectionSignature = signature;
      if (refreshTimer !== undefined) window.clearTimeout(refreshTimer);
      refreshTimer = window.setTimeout(() => void refreshSelection(), 180);
    }
  } catch {
    // 切换表格或视图时可能短暂无法读取，下一轮自动重试。
  } finally {
    pollingSelection = false;
  }
}

refreshButton.addEventListener("click", () => void refreshSelection());
mergeButton.addEventListener("click", () => void mergeSelected());
bitable.base.onSelectionChange(() => void pollSelection());
window.setInterval(() => void pollSelection(), SELECTION_POLL_MS);
void pollSelection();
