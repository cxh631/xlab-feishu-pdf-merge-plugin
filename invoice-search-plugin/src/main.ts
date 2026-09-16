import { bitable, type IAttachmentField, type IDateTimeField, type IOpenAttachment } from "@lark-base-open/js-sdk";
import * as pdfjs from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { createWorker, type Worker } from "tesseract.js";
import "./style.css";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorker;

const FIELD = { reason:"报销事由", amount:"报销金额", source:"发票+订单截图+支付截图", treasury:"小金库", school:"报销学校", earliestDate:"最早开票日期", latestDate:"最晚开票日期" } as const;
const DB_NAME = "xlab-invoice-search-v2";
const STORE = "pages";
const app = document.querySelector<HTMLElement>("#app")!;

app.innerHTML = `
<main class="shell">
  <header><span class="eyebrow">X-LAB 财务工具</span><h1>发票内容定位</h1><p>输入商品名称、金额或开票日期，定位到报销记录、附件和具体页码。</p></header>
  <section class="card index-card"><div><strong id="index-summary">尚未建立索引</strong><small id="index-note">仅检索：小金库＝否、报销学校＝已勾选</small></div><button id="build" class="secondary">建立/更新索引</button></section>
  <section class="card form">
    <div class="fields"><label>商品/内容关键词<input id="keyword" placeholder="例如：开发板、打印纸"></label><label>发票金额<input id="amount" inputmode="decimal" placeholder="例如：299.00"></label><label>开票日期<input id="invoice-date" type="date"></label></div>
    <div class="actions"><button id="search" class="primary">搜索附件</button><button id="clear" class="secondary">清空</button></div>
    <div id="status" class="status" aria-live="polite"></div><div class="progress"><span id="bar"></span></div>
  </section>
  <section id="results" class="results"></section>
  <p class="footnote">优先读取 PDF 自带文字；仅对无文字页面进行 OCR。识别结果保存在当前浏览器，不上传到第三方业务服务器。</p>
</main>`;

const el = <T extends HTMLElement>(selector:string) => document.querySelector<T>(selector)!;
const buildBtn=el<HTMLButtonElement>("#build"), searchBtn=el<HTMLButtonElement>("#search"), clearBtn=el<HTMLButtonElement>("#clear");
const keywordInput=el<HTMLInputElement>("#keyword"), amountInput=el<HTMLInputElement>("#amount"), dateInput=el<HTMLInputElement>("#invoice-date");
const statusEl=el<HTMLElement>("#status"), bar=el<HTMLElement>("#bar"), resultsEl=el<HTMLElement>("#results");
const summaryEl=el<HTMLElement>("#index-summary"), noteEl=el<HTMLElement>("#index-note");

type PageEntry={ key:string; scope:string; recordId:string; rowNumber:number; reason:string; recordAmount:string; earliestDate:number|null; latestDate:number|null; attachmentName:string; pageNumber:number; text:string; ocr:boolean };
let currentEntries:PageEntry[]=[];
let currentScope="";
let ocrWorker:Promise<Worker>|null=null;

function status(message:string,kind:"info"|"error"|"success"="info"){statusEl.textContent=message;statusEl.className=`status ${kind}`;}
function progress(done:number,total:number){bar.style.width=total?`${Math.min(100,done/total*100)}%`:"0%";}
function escapeHtml(value:string){return value.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]!));}
function normalize(value:string){return value.toLowerCase().replace(/[，,￥¥\s]/g,"");}
function amountVariants(raw:string){const n=Number(raw.replace(/[,，￥¥\s]/g,""));if(!Number.isFinite(n))return[];return [...new Set([String(n),n.toFixed(2),n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})].map(normalize))];}
function cellText(value:unknown):string{if(value==null)return"";if(typeof value==="string"||typeof value==="number"||typeof value==="boolean")return String(value);if(Array.isArray(value))return value.map(cellText).join(" ");if(typeof value==="object"){const item=value as Record<string,unknown>;return cellText(item.text??item.name??item.value??"");}return"";}
function isChecked(value:unknown){return value===true||["true","是","已勾选","1"].includes(cellText(value).trim());}
function dateLabel(value:number|null){return value?new Intl.DateTimeFormat("zh-CN",{timeZone:"Asia/Shanghai",year:"numeric",month:"2-digit",day:"2-digit"}).format(value):"未填写";}

function openDb():Promise<IDBDatabase>{return new Promise((resolve,reject)=>{const req=indexedDB.open(DB_NAME,1);req.onupgradeneeded=()=>{const db=req.result;const store=db.createObjectStore(STORE,{keyPath:"key"});store.createIndex("scope","scope")};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
async function getCached(scope:string){const db=await openDb();return new Promise<PageEntry[]>((resolve,reject)=>{const tx=db.transaction(STORE,"readonly");const req=tx.objectStore(STORE).index("scope").getAll(scope);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
async function putEntries(entries:PageEntry[]){if(!entries.length)return;const db=await openDb();await new Promise<void>((resolve,reject)=>{const tx=db.transaction(STORE,"readwrite");for(const entry of entries)tx.objectStore(STORE).put(entry);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);});}

async function getOcrWorker(){if(!ocrWorker)ocrWorker=createWorker("chi_sim+eng",1,{logger:m=>{if(m.status==="recognizing text")status(`正在 OCR：${Math.round((m.progress||0)*100)}%`);}});return ocrWorker;}
async function pageText(page:Awaited<ReturnType<Awaited<ReturnType<typeof pdfjs.getDocument>["promise"]>["getPage"]>>){
  const content=await page.getTextContent();const text=content.items.map(item=>("str" in item?item.str:"")).join(" ").trim();
  if(text.replace(/\s/g,"").length>=10)return {text,ocr:false};
  const viewport=page.getViewport({scale:1.45});const canvas=document.createElement("canvas");canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
  const context=canvas.getContext("2d");if(!context)return {text,ocr:false};await page.render({canvasContext:context,viewport}).promise;
  const worker=await getOcrWorker();const result=await worker.recognize(canvas);return {text:`${text} ${result.data.text}`.trim(),ocr:true};
}

async function context(){const table=await bitable.base.getActiveTable();const view=await table.getActiveView();const [reason,amount,source,treasury,school,earliestDate,latestDate]=await Promise.all([table.getFieldByName(FIELD.reason),table.getFieldByName(FIELD.amount),table.getFieldByName<IAttachmentField>(FIELD.source),table.getFieldByName(FIELD.treasury),table.getFieldByName(FIELD.school),table.getFieldByName<IDateTimeField>(FIELD.earliestDate),table.getFieldByName<IDateTimeField>(FIELD.latestDate)]);return {table,view,reason,amount,source,treasury,school,earliestDate,latestDate,scope:`${table.id}:${view.id}:school-only-v2`};}
async function loadRows(){
  const ctx=await context();const rows:{recordId:string;rowNumber:number;reason:string;amount:string;earliestDate:number|null;latestDate:number|null;treasury:string;school:boolean;files:IOpenAttachment[]}[]=[];let pageToken:number|undefined;let row=0;
  do{const page=await ctx.table.getRecordsByPage({pageSize:200,pageToken,viewId:ctx.view.id});for(const record of page.records){row++;rows.push({recordId:record.recordId,rowNumber:row,reason:cellText(record.fields[ctx.reason.id]),amount:cellText(record.fields[ctx.amount.id]),earliestDate:typeof record.fields[ctx.earliestDate.id]==="number"?record.fields[ctx.earliestDate.id] as number:null,latestDate:typeof record.fields[ctx.latestDate.id]==="number"?record.fields[ctx.latestDate.id] as number:null,treasury:cellText(record.fields[ctx.treasury.id]).trim(),school:isChecked(record.fields[ctx.school.id]),files:(record.fields[ctx.source.id]??[]) as IOpenAttachment[]});}pageToken=page.hasMore?page.pageToken:undefined;}while(pageToken!==undefined);
  const eligible=rows.filter(item=>item.treasury==="否"&&item.school).reverse();
  return {ctx,rows:eligible,totalRows:rows.length};
}

async function buildIndex(){
  buildBtn.disabled=true;searchBtn.disabled=true;resultsEl.innerHTML="";try{
    status("正在从最后一行筛选可报学校记录……");const {ctx,rows,totalRows}=await loadRows();currentScope=ctx.scope;const cached=await getCached(currentScope);const cachedByKey=new Map<string,PageEntry[]>();for(const entry of cached){const prefix=entry.key.split(":p")[0];const list=cachedByKey.get(prefix)||[];list.push(entry);cachedByKey.set(prefix,list);}
    const targets=rows.flatMap(r=>r.files.filter(f=>/pdf/i.test(f.type)||f.name.toLowerCase().endsWith(".pdf")).map(f=>({row:r,file:f,prefix:`${currentScope}:${r.recordId}:${f.token}:${f.size}:${f.timeStamp}`})));
    currentEntries=[];let done=0;for(const target of targets){const old=cachedByKey.get(target.prefix);if(old?.length){currentEntries.push(...old);done++;progress(done,targets.length);continue;}
      status(`正在读取 ${done+1}/${targets.length}：${target.file.name}`);const urls=await ctx.source.getAttachmentUrls(target.row.recordId);const index=target.row.files.findIndex(f=>f.token===target.file.token);const response=await fetch(urls[index]);if(!response.ok)throw new Error(`附件下载失败：${target.file.name}`);
      const pdf=await pdfjs.getDocument({data:await response.arrayBuffer()}).promise;const fresh:PageEntry[]=[];for(let p=1;p<=pdf.numPages;p++){status(`正在识别 ${done+1}/${targets.length}：${target.file.name}（第${p}/${pdf.numPages}页）`);const extracted=await pageText(await pdf.getPage(p));fresh.push({key:`${target.prefix}:p${p}`,scope:currentScope,recordId:target.row.recordId,rowNumber:target.row.rowNumber,reason:target.row.reason,recordAmount:target.row.amount,earliestDate:target.row.earliestDate,latestDate:target.row.latestDate,attachmentName:target.file.name,pageNumber:p,text:extracted.text,ocr:extracted.ocr});}
      await putEntries(fresh);currentEntries.push(...fresh);done++;progress(done,targets.length);
    }
    summaryEl.textContent=`已索引 ${targets.length} 个 PDF、${currentEntries.length} 页`;noteEl.textContent=`${totalRows} 条中筛出 ${rows.length} 条；按最新记录优先读取`;status("索引更新完成，现在可以直接搜索。","success");
  }catch(error){status(error instanceof Error?error.message:String(error),"error");}finally{buildBtn.disabled=false;searchBtn.disabled=false;}
}

function excerpt(text:string,needle:string){const plain=text.replace(/\s+/g," ").trim();const at=needle?normalize(plain).indexOf(normalize(needle)):-1;return at<0?plain.slice(0,150):plain.slice(Math.max(0,at-45),at+105);}
async function search(){
  const keyword=keywordInput.value.trim(),amount=amountInput.value.trim(),invoiceDate=dateInput.value;if(!keyword&&!amount&&!invoiceDate){status("请输入商品名称、金额或开票日期。","error");return;}
  if(!currentEntries.length){const ctx=await context();currentScope=ctx.scope;currentEntries=await getCached(currentScope);}
  if(!currentEntries.length){status("请先点击“建立/更新索引”。","error");return;}
  const key=normalize(keyword),amounts=amountVariants(amount),dateValue=invoiceDate?new Date(`${invoiceDate}T00:00:00+08:00`).getTime():0;const hits=currentEntries.filter(e=>{const text=normalize(e.text);const dateOk=!dateValue||(!!e.earliestDate&&!!e.latestDate&&dateValue>=e.earliestDate&&dateValue<=e.latestDate);return(!key||text.includes(key))&&(!amount||amounts.some(v=>text.includes(v)))&&dateOk;});
  status(`找到 ${hits.length} 个匹配页面（最新记录优先）。`,hits.length?"success":"info");resultsEl.innerHTML=`<h2>搜索结果（${hits.length}）</h2>`+(hits.length?hits.map((h,i)=>`<article class="result"><div class="result-head"><strong>第 ${h.rowNumber} 条｜${escapeHtml(h.reason||"未填写报销事由")}</strong><span class="tag">第 ${h.pageNumber} 页${h.ocr?" · OCR":""}</span></div><div class="meta">${escapeHtml(h.attachmentName)}｜记录金额 ${escapeHtml(h.recordAmount||"未填写")}｜开票 ${dateLabel(h.earliestDate)}—${dateLabel(h.latestDate)}</div><div class="excerpt">${escapeHtml(excerpt(h.text,keyword||amount))}</div><button class="open" data-hit="${i}">打开报销记录</button></article>`).join(""):`<div class="card empty">没有找到匹配内容，可尝试减少检索条件。</div>`);
  resultsEl.querySelectorAll<HTMLButtonElement>(".open").forEach(button=>button.onclick=async()=>{const hit=hits[Number(button.dataset.hit)];const {table}=await context();await bitable.ui.showRecordDetailDialog({tableId:table.id,recordId:hit.recordId});});
}

buildBtn.onclick=buildIndex;searchBtn.onclick=search;clearBtn.onclick=()=>{keywordInput.value="";amountInput.value="";dateInput.value="";resultsEl.innerHTML="";status("");};
void (async()=>{try{const ctx=await context();currentScope=ctx.scope;currentEntries=await getCached(currentScope);if(currentEntries.length){summaryEl.textContent=`本机已有 ${currentEntries.length} 页索引`;noteEl.textContent="可直接搜索；新增附件后点击更新索引";}}catch(error){status(error instanceof Error?error.message:String(error),"error");}})();
