import {
  FieldType,
  ToastType,
  bitable,
  type IAttachmentField,
  type IGridView,
} from "@lark-base-open/js-sdk";
import { PDFDocument } from "pdf-lib";
import "./style.css";

const FIELD = {
  reason: "报销事由",
  category: "报销类目",
  header: "发票抬头",
  approval: "审批",
  source: "发票+订单截图+支付截图",
  result: "合并后材料PDF",
} as const;

const MAX_FILES = 30;
const MAX_TOTAL_BYTES = 80 * 1024 * 1024;

const app = document.querySelector<HTMLElement>("#app");
if (!app) throw new Error("页面初始化失败");

app.innerHTML = `
  <section class="shell">
    <header>
      <span class="eyebrow">X-LAB 财务工具</span>
      <h1>合并所选 PDF</h1>
      <p>在当前表格勾选记录，插件会校验抬头、类目与审批状态，再把材料按视图顺序合并。</p>
    </header>

    <div class="card summary">
      <div>
        <span class="label">当前选择</span>
        <strong id="selected-count">尚未读取</strong>
      </div>
      <button id="refresh" class="secondary" type="button">刷新选择</button>
    </div>

    <div id="selection-detail" class="card detail muted">
      请先在当前表格左侧勾选至少两条记录。
    </div>

    <div class="rules">
      <span>✓ 审批均为“同意”</span>
      <span>✓ 发票抬头一致</span>
      <span>✓ 报销类目一致</span>
      <span>✓ 附件均为 PDF</span>
    </div>

    <button id="merge" class="primary" type="button" disabled>合并并写回第一条记录</button>
    <div id="status" class="status" aria-live="polite"></div>
    <p class="footnote">文件仅在当前浏览器内存中处理，不上传到第三方合并服务。</p>
  </section>
`;

const selectedCount = document.querySelector<HTMLElement>("#selected-count")!;
const detail = document.querySelector<HTMLElement>("#selection-detail")!;
const status = document.querySelector<HTMLElement>("#status")!;
const refreshButton = document.querySelector<HTMLButtonElement>("#refresh")!;
const mergeButton = document.querySelector<HTMLButtonElement>("#merge")!;

type SelectionState = {
  recordIds: string[];
  category: string;
  header: string;
  reasons: string[];
  fileCount: number;
};

let state: SelectionState | null = null;

function setStatus(message: string, kind: "info" | "success" | "error" = "info") {
  status.className = `status ${kind}`;
  status.textContent = message;
}

function safeText(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]!);
}

function buildFilename(header: string, category: string): string {
  const stamp = new Intl.DateTimeFormat("zh-CN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(new Date()).replace(/\D/g, "");
  const clean = (value: string) => value.replace(/[\\/:*?"<>|]+/g, "_").slice(0, 30);
  return `${clean(header)}_${clean(category)}_${stamp}.pdf`;
}

async function getFields() {
  const table = await bitable.base.getActiveTable();
  const [reason, category, header, approval, source, result] = await Promise.all([
    table.getFieldByName(FIELD.reason),
    table.getFieldByName(FIELD.category),
    table.getFieldByName(FIELD.header),
    table.getFieldByName(FIELD.approval),
    table.getFieldByName<IAttachmentField>(FIELD.source),
    table.getFieldByName<IAttachmentField>(FIELD.result),
  ]);

  if (await source.getType() !== FieldType.Attachment) {
    throw new Error(`“${FIELD.source}”必须是附件字段`);
  }
  if (await result.getType() !== FieldType.Attachment) {
    throw new Error(`“${FIELD.result}”必须是附件字段`);
  }
  return { table, reason, category, header, approval, source, result };
}

async function readSelection(): Promise<SelectionState> {
  const { table, reason, category, header, approval, source } = await getFields();
  const view = await table.getActiveView() as IGridView;
  if (typeof view.getSelectedRecordIdList !== "function") {
    throw new Error("请在表格视图中使用本插件");
  }

  const recordIds = await view.getSelectedRecordIdList();
  if (recordIds.length < 2) throw new Error("请至少勾选两条记录");

  const rows = await Promise.all(recordIds.map(async (recordId) => {
    const [reasonText, categoryText, headerText, approvalText, files] = await Promise.all([
      reason.getCellString(recordId),
      category.getCellString(recordId),
      header.getCellString(recordId),
      approval.getCellString(recordId),
      source.getValue(recordId),
    ]);
    return {
      recordId,
      reason: reasonText.trim() || "未填写报销事由",
      category: categoryText.trim(),
      header: headerText.trim(),
      approval: approvalText.trim(),
      files,
    };
  }));

  if (rows.some((row) => row.approval !== "同意")) {
    throw new Error("所选记录中存在尚未审批同意的记录");
  }
  const categories = new Set(rows.map((row) => row.category));
  if (categories.size !== 1 || categories.has("")) {
    throw new Error("所选记录的报销类目必须一致且不能为空");
  }
  const headers = new Set(rows.map((row) => row.header));
  if (headers.size !== 1 || headers.has("")) {
    throw new Error("所选记录的发票抬头必须一致且不能为空");
  }
  const files = rows.flatMap((row) => row.files);
  if (files.length === 0) throw new Error("所选记录中没有可合并的附件");
  if (files.length > MAX_FILES) throw new Error(`单次最多合并 ${MAX_FILES} 个 PDF`);
  if (rows.some((row) => row.files.length === 0)) {
    throw new Error("所选记录中存在未上传报销材料的记录");
  }
  if (files.some((file) => !file.name.toLowerCase().endsWith(".pdf"))) {
    throw new Error("所选记录中包含非 PDF 附件");
  }

  return {
    recordIds,
    category: rows[0].category,
    header: rows[0].header,
    reasons: rows.map((row) => row.reason),
    fileCount: files.length,
  };
}

async function refreshSelection() {
  mergeButton.disabled = true;
  setStatus("正在读取当前选择……");
  try {
    state = await readSelection();
    selectedCount.textContent = `${state.recordIds.length} 条记录 · ${state.fileCount} 个 PDF`;
    detail.classList.remove("muted");
    detail.innerHTML = `
      <dl>
        <div><dt>发票抬头</dt><dd>${safeText(state.header)}</dd></div>
        <div><dt>报销类目</dt><dd>${safeText(state.category)}</dd></div>
        <div><dt>写回位置</dt><dd>第一条选中记录</dd></div>
      </dl>
      <ol>${state.reasons.map((item) => `<li>${safeText(item)}</li>`).join("")}</ol>
    `;
    mergeButton.disabled = false;
    setStatus("校验通过，可以合并。", "success");
  } catch (error) {
    state = null;
    selectedCount.textContent = "未满足合并条件";
    detail.classList.add("muted");
    detail.textContent = error instanceof Error ? error.message : "读取选择失败";
    setStatus("请调整勾选记录后重试。", "error");
  }
}

async function mergeSelected() {
  mergeButton.disabled = true;
  refreshButton.disabled = true;
  setStatus("正在校验并下载 PDF……");
  try {
    const latest = await readSelection();
    const { source, result } = await getFields();
    const output = await PDFDocument.create();
    let totalBytes = 0;
    let mergedFiles = 0;

    for (const recordId of latest.recordIds) {
      const [attachments, urls] = await Promise.all([
        source.getValue(recordId),
        source.getAttachmentUrls(recordId),
      ]);
      if (attachments.length !== urls.length) {
        throw new Error("附件读取数量不一致，请刷新后重试");
      }
      for (let index = 0; index < urls.length; index += 1) {
        setStatus(`正在处理第 ${mergedFiles + 1}/${latest.fileCount} 个 PDF……`);
        const response = await fetch(urls[index]);
        if (!response.ok) throw new Error(`附件下载失败：${attachments[index].name}`);
        const bytes = new Uint8Array(await response.arrayBuffer());
        totalBytes += bytes.byteLength;
        if (totalBytes > MAX_TOTAL_BYTES) throw new Error("附件总量超过 80MB，请拆分合并");
        const sourcePdf = await PDFDocument.load(bytes, { ignoreEncryption: false });
        const pages = await output.copyPages(sourcePdf, sourcePdf.getPageIndices());
        pages.forEach((page) => output.addPage(page));
        mergedFiles += 1;
      }
    }

    setStatus("正在写回飞书附件字段……");
    const mergedBytes = await output.save({ useObjectStreams: true });
    const buffer = new ArrayBuffer(mergedBytes.byteLength);
    new Uint8Array(buffer).set(mergedBytes);
    const filename = buildFilename(latest.header, latest.category);
    const file = new File([buffer], filename, { type: "application/pdf" });
    await result.setValue(latest.recordIds[0], file);
    await bitable.ui.showToast({ toastType: ToastType.success, message: "PDF 已合并并写回" });
    setStatus(`完成：${mergedFiles} 个 PDF 已合并为 ${filename}`, "success");
    state = latest;
  } catch (error) {
    const message = error instanceof Error ? error.message : "合并失败";
    await bitable.ui.showToast({ toastType: ToastType.error, message });
    setStatus(message, "error");
  } finally {
    mergeButton.disabled = state === null;
    refreshButton.disabled = false;
  }
}

refreshButton.addEventListener("click", refreshSelection);
mergeButton.addEventListener("click", mergeSelected);
void refreshSelection();
