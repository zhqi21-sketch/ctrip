// ============================================================
// 携程签证订单抓取工具 - Node.js 后端 v2
// 基于真实页面结构：order-basic-row + order-extra-row
// ============================================================
const express = require('express');
const puppeteer = require('puppeteer-core');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const http = require('http');

const app = express();
const server = http.createServer(app);

// ── 极简 WebSocket（无需 ws 包）──
const clients = new Set();
server.on('upgrade', (req, socket) => {
  const key = req.headers['sec-websocket-key'];
  const crypto = require('crypto');
  const accept = crypto.createHash('sha1')
    .update(key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
  socket.write(
    'HTTP/1.1 101 Switching Protocols\r\n' +
    'Upgrade: websocket\r\nConnection: Upgrade\r\n' +
    `Sec-WebSocket-Accept: ${accept}\r\n\r\n`
  );
  socket.on('data', () => {});
  socket.on('close', () => clients.delete(socket));
  socket.on('error', () => clients.delete(socket));
  clients.add(socket);
});

function wsSend(obj) {
  const buf = Buffer.from(JSON.stringify(obj));
  const frame = Buffer.alloc(buf.length + 10);
  frame[0] = 0x81;
  let offset = 2;
  if (buf.length < 126) {
    frame[1] = buf.length;
  } else {
    frame[1] = 126;
    frame.writeUInt16BE(buf.length, 2);
    offset = 4;
  }
  buf.copy(frame, offset);
  const f = frame.slice(0, offset + buf.length);
  clients.forEach(s => { try { s.write(f); } catch(e) {} });
}

app.use(express.json());
app.use(express.static(__dirname));

// ── 飞书配置 ──
const FEISHU_APP_ID     = 'cli_a948bc98d3a31bc0';
const FEISHU_APP_SECRET = 'UfeFovKSrJhSbCXYBxlb8du2nAUy6TWN';
const FEISHU_APP_TOKEN  = 'BLJhbKvFFaSzUFs69sTcRZeGnob';
const FEISHU_TABLE_ID   = 'tblrTkq7l4Sao3od'; // 测试表格：加急进度看板 自动化

// ── 本地数据文件 ──
const DATA_FILE = path.join(__dirname, 'data.json');

// ── 全局状态 ──
let isRunning = false;
let shouldStop = false;
let allRows = [];
let browser = null;

// ── 启动时加载本地数据 ──
function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf-8');
      allRows = JSON.parse(raw);
      console.log(`[INFO] Loaded ${allRows.length} rows from data.json`);
    }
  } catch(e) {
    console.log('[WARN] Failed to load data.json:', e.message);
    allRows = [];
  }
}

// ── 保存数据到本地文件 ──
function saveData() {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(allRows, null, 2), 'utf-8');
  } catch(e) {
    console.log('[WARN] Failed to save data.json:', e.message);
  }
}

// ── 合并新数据（同订单号+出行人覆盖，其余保留）──
function mergeRows(newRows) {
  for (const newRow of newRows) {
    const key = String(newRow.orderNo) + '_' + String(newRow.travelerName);
    const idx = allRows.findIndex(r => String(r.orderNo) + '_' + String(r.travelerName) === key);
    if (idx >= 0) {
      allRows[idx] = newRow; // 覆盖
    } else {
      allRows.push(newRow); // 新增
    }
  }
  saveData();
}

function log(msg, type = 'default') {
  console.log(`[${type.toUpperCase()}] ${msg}`);
  wsSend({ type: 'log', level: type, msg });
}
function progress(text, pct, phase = '') {
  wsSend({ type: 'progress', text, pct, phase });
}
function pushRows() { wsSend({ type: 'rows', rows: allRows }); }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── API ──
app.post('/api/start', async (req, res) => {
  if (isRunning) return res.json({ ok: false, msg: 'Already running' });
  const { startDate, endDate } = req.body || {};
  res.json({ ok: true });
  runScraper({ startDate, endDate });
});
app.post('/api/stop', (req, res) => { shouldStop = true; res.json({ ok: true }); });
app.post('/api/clear', (req, res) => { allRows = []; pushRows(); res.json({ ok: true }); });

// ---- API: 重新抓取指定订单号 ----
app.post('/api/retry', async (req, res) => {
  if (isRunning) return res.json({ ok: false, msg: 'Already running' });
  const { orderNos } = req.body || {};
  if (!orderNos || orderNos.length === 0) return res.json({ ok: false, msg: 'No orders selected' });
  res.json({ ok: true });
  retryOrders(orderNos);
});

app.post('/api/export', async (req, res) => {
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('visa_orders');
    ws.columns = [
      { header: 'No.',         key: 'idx',         width: 5  },
      { header: 'Order No.',   key: 'orderNo',     width: 22 },
      { header: 'Pay Date',    key: 'payDate',     width: 22 },
      { header: 'Pay Status',  key: 'payStatus',   width: 10 },
      { header: 'Confirm Status', key: 'confirmStatus', width: 12 },
      { header: 'Dest Country',key: 'destCountry', width: 14 },
      { header: 'Sign Country',key: 'signCountry', width: 14 },
      { header: 'Service',     key: 'service',     width: 24 },
      { header: 'Traveler',    key: 'traveler',    width: 14 },
      { header: 'Amount',      key: 'amount',      width: 12 },
      { header: 'Sub Orders',  key: 'subCount',    width: 10 },
      { header: 'Sub Status',  key: 'subStatus',   width: 40 },
    ];
    ws.getRow(1).eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1677FF' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 11 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    ws.getRow(1).height = 22;
    allRows.forEach((row, i) => {
      const r = ws.addRow({
        idx: i + 1,
        orderNo: row.orderNo || '',
        payDate: row.payDate || '',
        payStatus: row.payStatus || '',
        confirmStatus: row.confirmStatus || '',
        destCountry: row.destCountry || '',
        signCountry: row.signCountry || '',
        service: row.service || '',
        traveler: row.travelerName || '',
        amount: row.amount ? parseFloat(row.amount) : '',
        subCount: row.subOrders ? row.subOrders.length : 0,
        subStatus: row.subOrders ? row.subOrders.map(s => s.status || '').join(' | ') : '',
      });
      if (i % 2 === 1) {
        r.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F5FF' } };
        });
      }
    });
    const filePath = path.join(__dirname, `orders_${Date.now()}.xlsx`);
    await wb.xlsx.writeFile(filePath);
    res.download(filePath, 'ctrip_visa_orders.xlsx', () => { fs.unlink(filePath, () => {}); });
  } catch (e) {
    res.status(500).json({ ok: false, msg: e.message });
  }
});

// ============================================================
// 飞书多维表格同步
// ============================================================

// ── 飞书字段定义（只写入7个有用字段）──
const FEISHU_FIELDS = [
  { field_name: '订单号',     type: 1 }, // 文本（表格已有）
  { field_name: '付款时间',   type: 1 },
  { field_name: '客人姓名',   type: 1 },
  { field_name: '所选服务',   type: 1 },
  { field_name: '国家或地区', type: 1 },
  { field_name: '收款',       type: 2 }, // 数字
  { field_name: '城市',       type: 1 }, // 送签国/地
];

// 自动创建缺失的字段
async function ensureFeishuFields(token) {
  // 获取现有字段
  const res = await fetch(
    `https://open.feishu.cn/open-apis/bitable/v1/apps/${FEISHU_APP_TOKEN}/tables/${FEISHU_TABLE_ID}/fields`,
    { headers: { 'Authorization': 'Bearer ' + token } }
  );
  const data = await res.json();
  if (data.code !== 0) { log('Get fields failed: ' + data.msg, 'warn'); return; }

  const existingNames = new Set((data.data.items || []).map(f => f.field_name));
  log(`Existing fields: ${[...existingNames].join(', ')}`, 'default');

  // 创建缺失的字段
  for (const field of FEISHU_FIELDS) {
    if (existingNames.has(field.field_name)) continue;
    const r = await fetch(
      `https://open.feishu.cn/open-apis/bitable/v1/apps/${FEISHU_APP_TOKEN}/tables/${FEISHU_TABLE_ID}/fields`,
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify(field)
      }
    );
    const d = await r.json();
    log(`Create field "${field.field_name}": ${d.code === 0 ? 'OK' : d.msg}`,
      d.code === 0 ? 'success' : 'warn');
    await new Promise(r => setTimeout(r, 200)); // 避免频率限制
  }
}

// 获取飞书 tenant_access_token
async function getFeishuToken() {
  const res = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ app_id: FEISHU_APP_ID, app_secret: FEISHU_APP_SECRET })
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error('Feishu auth failed: ' + data.msg);
  return data.tenant_access_token;
}

// 获取多维表格所有记录（返回 Map<订单号_出行人, record_id>）
async function getExistingRecords(token) {
  const map = new Map();
  let pageToken = '';
  while (true) {
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${FEISHU_APP_TOKEN}/tables/${FEISHU_TABLE_ID}/records?page_size=500${pageToken ? '&page_token=' + pageToken : ''}`;
    const res = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    const data = await res.json();
    if (data.code !== 0) break;
    for (const rec of (data.data.items || [])) {
      const orderNo = rec.fields['订单号'] || '';
      // 客人姓名字段可能是数组（飞书文本字段返回格式）
      let traveler = rec.fields['客人姓名'] || '';
      if (Array.isArray(traveler)) traveler = traveler.map(t => t.text || '').join('');
      const key = String(orderNo) + '_' + String(traveler);
      map.set(key, rec.record_id);
    }
    if (!data.data.has_more) break;
    pageToken = data.data.page_token;
  }
  return map;
}

// 把一行数据转成飞书字段格式（只写入7个有用字段）
function rowToFeishuFields(row) {
  return {
    '订单号':     String(row.orderNo || ''),
    '付款时间':   String(row.payDate || ''),
    '客人姓名':   String(row.travelerName || ''),
    '所选服务':   String(row.service || ''),
    '国家或地区': String(row.destCountry || ''),
    '收款':       row.amount ? parseFloat(row.amount) : 0,
    '城市':       String(row.signCountry || ''), // 送签国/地
  };
}

// 同步到飞书（批量更新+新增）
async function syncToFeishu(rows) {
  log('Syncing to Feishu...', 'info');
  try {
    const token = await getFeishuToken();
    log('Feishu token OK', 'default');

    // 自动创建缺失字段
    await ensureFeishuFields(token);

    const existing = await getExistingRecords(token);
    log(`Feishu existing records: ${existing.size}`, 'default');

    const toUpdate = [];
    const toCreate = [];

    for (const row of rows) {
      const key = String(row.orderNo || '') + '_' + String(row.travelerName || '');
      const fields = rowToFeishuFields(row);
      if (existing.has(key)) {
        toUpdate.push({ record_id: existing.get(key), fields });
      } else {
        toCreate.push({ fields });
      }
    }

    // 批量新增（每批 500 条）
    for (let i = 0; i < toCreate.length; i += 500) {
      const batch = toCreate.slice(i, i + 500);
      const res = await fetch(
        `https://open.feishu.cn/open-apis/bitable/v1/apps/${FEISHU_APP_TOKEN}/tables/${FEISHU_TABLE_ID}/records/batch_create`,
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ records: batch })
        }
      );
      const data = await res.json();
      log(`Feishu create batch ${i/500+1}: ${data.code === 0 ? 'OK '+batch.length+'条' : 'Error '+data.msg}`,
        data.code === 0 ? 'success' : 'error');
    }

    // 批量更新（每批 500 条）
    for (let i = 0; i < toUpdate.length; i += 500) {
      const batch = toUpdate.slice(i, i + 500);
      const res = await fetch(
        `https://open.feishu.cn/open-apis/bitable/v1/apps/${FEISHU_APP_TOKEN}/tables/${FEISHU_TABLE_ID}/records/batch_update`,
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ records: batch })
        }
      );
      const data = await res.json();
      log(`Feishu update batch ${i/500+1}: ${data.code === 0 ? 'OK '+batch.length+'条' : 'Error '+data.msg}`,
        data.code === 0 ? 'success' : 'error');
    }

    log(`Feishu sync done: +${toCreate.length} new, ~${toUpdate.length} updated`, 'success');
    return { ok: true, created: toCreate.length, updated: toUpdate.length };
  } catch(e) {
    log('Feishu sync error: ' + e.message, 'error');
    return { ok: false, msg: e.message };
  }
}

// ── API: 手动触发飞书同步 ──
app.post('/api/sync-feishu', async (req, res) => {
  const result = await syncToFeishu(allRows);
  res.json(result);
});

// ── API: 从本地文件重新加载数据 ──
app.get('/api/load', (req, res) => {
  loadData();
  pushRows();
  res.json({ ok: true, count: allRows.length });
});

// ============================================================
// 重新抓取指定订单（更新 allRows 里已有的数据）
// ============================================================
async function retryOrders(orderNos) {
  isRunning = true;
  shouldStop = false;

  try {
    log(`Retrying ${orderNos.length} orders...`, 'info');
    progress('Connecting...', 2);

    browser = await puppeteer.connect({
      browserURL: 'http://127.0.0.1:9222',
      defaultViewport: null,
    });

    const pages = await browser.pages();
    let orderPage = pages.find(p => p.url().includes('visaOrderList'))
      || pages.find(p => p.url().includes('vbooking.ctrip.com'));

    if (!orderPage) {
      orderPage = await browser.newPage();
      await orderPage.goto('https://vbooking.ctrip.com/order/visaOrderList?obsolete=1', {
        waitUntil: 'networkidle2', timeout: 30000
      });
    }

    for (let i = 0; i < orderNos.length && !shouldStop; i++) {
      const orderNo = orderNos[i];
      log(`[${i+1}/${orderNos.length}] Retrying order ${orderNo}`, 'info');
      progress(`重新抓取 ${orderNo}...`, Math.round((i / orderNos.length) * 100), `${i+1}/${orderNos.length}`);

      // 从现有 allRows 里找这个订单的基本信息
      const existingRow = allRows.find(r => r.orderNo === orderNo);
      if (!existingRow) {
        log(`  Order ${orderNo} not in current data, skipping`, 'warn');
        continue;
      }

      // 用订单基本信息重新走详情页流程
      const order = {
        orderNo: existingRow.orderNo,
        productName: existingRow.productName,
        amount: existingRow.amount,
        payDate: existingRow.payDate,
        payStatus: existingRow.payStatus,
        confirmStatus: existingRow.confirmStatus,
        destCountry: existingRow.destCountry,
        signCountry: existingRow.signCountry,
        service: existingRow.service,
        detailHref: `https://vbooking.ctrip.com/order/visaOrderList?obsolete=1`,
      };

      // 在列表页找到该订单的"查看订单"按钮
      // 先确保列表页上能看到这个订单（可能需要搜索）
      const found = await orderPage.evaluate((no) => {
        const rows = document.querySelectorAll('tr.order-basic-row');
        for (const row of rows) {
          const noEl = row.querySelector('.desc-item .value.link');
          if (noEl && noEl.innerText.trim() === no) return true;
        }
        return false;
      }, orderNo);

      if (!found) {
        // 在订单号搜索框里搜索
        log(`  Searching for order ${orderNo}...`, 'default');
        await orderPage.evaluate((no) => {
          const cols = document.querySelectorAll('.top-search-col');
          for (const col of cols) {
            const title = col.querySelector('.search-list-title');
            if (title && title.innerText.trim() === '订单号') {
              const input = col.querySelector('input.ant-input');
              if (input) {
                input.value = no;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
              }
              break;
            }
          }
        }, orderNo);
        await sleep(500);

        // 点查询
        await orderPage.evaluate(() => {
          const btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.trim() === '查询');
          if (btn) btn.click();
        });
        await sleep(2500);
      }

      // 删除 allRows 里该订单的所有行
      const before = allRows.length;
      allRows = allRows.filter(r => r.orderNo !== orderNo);
      log(`  Removed ${before - allRows.length} old rows for ${orderNo}`, 'default');

      // 重新处理订单详情
      await processOrder(orderPage, order);
      pushRows();

      // 如果搜索了，重置搜索框回订单列表
      if (!found) {
        await orderPage.evaluate(() => {
          const cols = document.querySelectorAll('.top-search-col');
          for (const col of cols) {
            const title = col.querySelector('.search-list-title');
            if (title && title.innerText.trim() === '订单号') {
              const input = col.querySelector('input.ant-input');
              if (input) {
                input.value = '';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
              }
              break;
            }
          }
          // 点查询恢复列表
          const btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.trim() === '查询');
          if (btn) btn.click();
        });
        await sleep(2000);
      }
    }

  } catch(e) {
    log('Retry error: ' + e.message, 'error');
    console.error(e);
  }

  isRunning = false;
  shouldStop = false;
  log(`Retry complete! Total rows: ${allRows.length}`, 'success');
  progress('重新抓取完成', 100, '完成');

  saveData();
  log('Data saved to data.json', 'default');
  await syncToFeishu(allRows);

  wsSend({ type: 'done', total: allRows.length });
  pushRows();
}

// ============================================================
// 主流程
// ============================================================
async function runScraper({ startDate, endDate } = {}) {
  isRunning = true;
  shouldStop = false;
  allRows = [];

  try {
    log('Connecting to Chrome (port 9222)...', 'info');
    progress('Connecting...', 2);

    browser = await puppeteer.connect({
      browserURL: 'http://127.0.0.1:9222',
      defaultViewport: null,
    });
    log('Chrome connected', 'success');

    const pages = await browser.pages();
    // 精确找订单列表页（URL包含 visaOrderList 或 order）
    let orderPage = pages.find(p => p.url().includes('visaOrderList'))
      || pages.find(p => p.url().includes('vbooking.ctrip.com/order'))
      || pages.find(p => p.url().includes('vbooking.ctrip.com'));

    // 打印所有标签页，方便调试
    log('All tabs: ' + pages.map(p => p.url()).join(' | '), 'warn');

    if (!orderPage) {
      log('No Ctrip page found, opening...', 'warn');
      orderPage = await browser.newPage();
      await orderPage.goto('https://vbooking.ctrip.com/order/visaOrderList?obsolete=1', {
        waitUntil: 'networkidle2', timeout: 30000
      });
    }

    log('Order list page found', 'success');

    // 如果有日期筛选，先设置付款日期
    if (startDate && endDate) {
      log(`Setting date filter: ${startDate} ~ ${endDate}`, 'info');
      await setDateFilter(orderPage, startDate, endDate);
    }

    await scrapeAllPages(orderPage);

  } catch (e) {
    log('Error: ' + e.message, 'error');
    console.error(e);
  }

  isRunning = false;
  shouldStop = false;

  // 清理：关闭所有除订单列表页以外的标签页
  try {
    const allPages = await browser.pages();
    for (const p of allPages) {
      const url = p.url();
      if (!url.includes('visaOrderList') && url.includes('vbooking.ctrip.com')) {
        try { await p.close(); } catch(e) {}
      }
    }
    log(`Cleanup: closed extra tabs`, 'default');
  } catch(e) {}

  log(`Done! Total ${allRows.length} rows`, allRows.length > 0 ? 'success' : 'warn');
  progress('Complete', 100, 'Done');

  // 保存到本地文件
  saveData();
  log('Data saved to data.json', 'default');

  // 自动同步到飞书
  if (allRows.length > 0) {
    await syncToFeishu(allRows);
  }

  wsSend({ type: 'done', total: allRows.length });
  pushRows();
}

// ============================================================
// 设置付款日期筛选
// 策略：点击 .ant-picker-range → 等日历弹出 →
//       找 td[title="YYYY-MM-DD"] 点击开始日期 →
//       再找结束日期点击 → 点查询按钮
// ============================================================
async function setDateFilter(page, startDate, endDate) {
  try {
    // 辅助：点击日历里指定日期的格子
    // 如果目标月份不在当前视图，需要翻月
    async function clickDateCell(targetDate) {
      // 最多翻12个月
      for (let attempt = 0; attempt < 12; attempt++) {
        const clicked = await page.evaluate((date) => {
          const cell = document.querySelector(`.ant-picker-dropdown:not(.ant-picker-dropdown-hidden) td[title="${date}"]`);
          if (cell) { cell.click(); return true; }
          return false;
        }, targetDate);
        if (clicked) return true;

        // 没找到，点下一月箭头
        const turned = await page.evaluate(() => {
          const btns = document.querySelectorAll(
            '.ant-picker-dropdown:not(.ant-picker-dropdown-hidden) .ant-picker-header-next-btn, ' +
            '.ant-picker-dropdown:not(.ant-picker-dropdown-hidden) button.ant-picker-next-btn'
          );
          if (btns[0]) { btns[0].click(); return true; }
          return false;
        });
        if (!turned) break;
        await sleep(400);
      }
      return false;
    }

    // 1. 点击付款日期的 picker 容器，弹出日历
    const opened = await page.evaluate(() => {
      const cols = document.querySelectorAll('.top-search-col');
      for (const col of cols) {
        const title = col.querySelector('.search-list-title');
        if (title && title.innerText.trim() === '付款日期') {
          const picker = col.querySelector('.ant-picker-range');
          if (picker) { picker.click(); return true; }
        }
      }
      return false;
    });

    if (!opened) { log('Date picker not found', 'warn'); return; }
    await sleep(700);

    // 2. 点击开始日期
    const startOk = await clickDateCell(startDate);
    log(`Start date ${startDate}: ${startOk ? 'clicked' : 'not found'}`, startOk ? 'default' : 'warn');
    await sleep(500);

    // 3. 点击结束日期（日历可能已切换到第二个面板）
    const endOk = await clickDateCell(endDate);
    log(`End date ${endDate}: ${endOk ? 'clicked' : 'not found'}`, endOk ? 'default' : 'warn');
    await sleep(500);

    // 4. 验证填入结果
    const filled = await page.evaluate(() => {
      const cols = document.querySelectorAll('.top-search-col');
      for (const col of cols) {
        const title = col.querySelector('.search-list-title');
        if (title && title.innerText.trim() === '付款日期') {
          const inputs = col.querySelectorAll('input');
          return { start: inputs[0]?.value || '', end: inputs[1]?.value || '' };
        }
      }
      return { start: '', end: '' };
    });
    log(`Date filled: ${filled.start} ~ ${filled.end}`, filled.start ? 'success' : 'warn');

    // 5. 先按 Escape 关闭日历，再点查询按钮
    await page.keyboard.press('Escape');
    await sleep(500);

    // 点击查询按钮（注意按钮文字是"查 询"中间有空格）
    const queryClicked = await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      // 匹配"查询"或"查 询"（去掉空格比较）
      const btn = btns.find(b => b.innerText.replace(/\s/g, '') === '查询');
      if (btn) { btn.click(); return true; }
      return false;
    });
    log(`Query button clicked: ${queryClicked}`, queryClicked ? 'success' : 'warn');

    // 6. 等待页面刷新：先等 loading 出现，再等 loading 消失
    await sleep(1000); // 等查询请求发出
    // 等待 ant-spin（加载中）消失，最多等15秒
    let loadWait = 0;
    while (loadWait < 15000) {
      await sleep(500);
      loadWait += 500;
      const loading = await page.evaluate(() => {
        // 检查是否还在加载
        const spinning = document.querySelector('.ant-spin-spinning');
        const skeleton = document.querySelector('.ant-skeleton-active');
        return !!(spinning || skeleton);
      });
      if (!loading) break;
    }
    await sleep(1000); // 额外等1秒确保渲染完成

    // 验证筛选结果
    const resultCount = await page.evaluate(() => {
      const rows = document.querySelectorAll('tr.order-basic-row');
      const empty = document.querySelector('.ant-table-placeholder');
      return { rows: rows.length, empty: !!empty };
    });
    log(`Date filter done: ${startDate} ~ ${endDate} | result rows: ${resultCount.rows}, empty: ${resultCount.empty}`, 'success');

  } catch(e) {
    log('Date filter error: ' + e.message, 'warn');
  }
}

// ============================================================
// 翻页循环
// ============================================================
async function scrapeAllPages(page) {
  let pageNum = 1;
  while (!shouldStop) {
    log(`===== Page ${pageNum} =====`, 'info');
    progress(`Scraping page ${pageNum}...`, Math.min(pageNum * 5, 80), `Page ${pageNum}`);

    const orders = await scrapeListPage(page);
    log(`Found ${orders.length} orders`, orders.length > 0 ? 'success' : 'warn');
    if (orders.length === 0) break;

    for (let i = 0; i < orders.length && !shouldStop; i++) {
      log(`[${i+1}/${orders.length}] Order ${orders[i].orderNo}`, 'default');
      await processOrder(page, orders[i]);
      pushRows();
    }

    wsSend({ type: 'pageCount', count: pageNum });
    const turned = await nextPage(page);
    if (!turned) { log('Last page reached', 'info'); break; }
    pageNum++;
    await sleep(2500);
  }
}

// ============================================================
// 抓取列表页 — 精确匹配真实 HTML 结构
//
// 每个订单由两行组成：
//   tr.order-basic-row  → 订单号、产品名、"查看订单"按钮
//   tr.order-extra-row  → [0]产品名 [1]子订单ID [2]确认状态 [3]数量
//                         [4]城市 [5]金额 [6]创建时间 [7]更新时间
//                         [8]支付状态 [9]支付时间
// ============================================================
async function scrapeListPage(page) {
  // 等待 Ant Design 表格渲染
  try {
    await page.waitForSelector('.ant-table-tbody tr', { timeout: 10000 });
  } catch(e) {
    log('Table wait timed out, proceeding...', 'warn');
  }
  await sleep(2000);

  // 调试：看Puppeteer实际看到的DOM结构
  const dbg = await page.evaluate(() => ({
    url: location.href,
    basicRows: document.querySelectorAll('tr.order-basic-row').length,
    antTr: document.querySelectorAll('.ant-table-tbody tr').length,
    allTr: document.querySelectorAll('tr').length,
    valueLink: document.querySelectorAll('.value.link').length,
    descItem: document.querySelectorAll('.desc-item').length,
    orderItem: document.querySelectorAll('.order-item').length,
    bodyLen: document.body.innerHTML.length,
    firstTrClass: document.querySelector('.ant-table-tbody tr')
      ? document.querySelector('.ant-table-tbody tr').className : 'NONE',
    sample: document.body.innerHTML.slice(0, 500),
  }));
  log('DEBUG basicRows=' + dbg.basicRows + ' antTr=' + dbg.antTr + ' allTr=' + dbg.allTr + ' valueLink=' + dbg.valueLink + ' descItem=' + dbg.descItem + ' bodyLen=' + dbg.bodyLen, 'warn');
  log('DEBUG firstTrClass: ' + dbg.firstTrClass, 'warn');
  log('DEBUG url: ' + dbg.url, 'warn');

  return await page.evaluate(() => {
    const orders = [];
    const basicRows = document.querySelectorAll('tr.order-basic-row');

    basicRows.forEach(row => {
      try {
        // 订单号
        const noEl = row.querySelector('.desc-item .value.link');
        const orderNo = noEl ? noEl.innerText.trim() : '';
        if (!orderNo || !/^\d{10,20}$/.test(orderNo)) return;

        // 产品名称
        const prodEl = row.querySelector('.desc-item .value[data-ignorechecktext]');
        const productName = prodEl ? prodEl.innerText.trim() : '';

        // 找紧随的 extra-row
        let extraRow = row.nextElementSibling;
        while (extraRow && !extraRow.classList.contains('order-extra-row')) {
          extraRow = extraRow.nextElementSibling;
        }

        let amount = '', payDate = '', payStatus = '', confirmStatus = '';
        if (extraRow) {
          const cells = extraRow.querySelectorAll('td.ant-table-cell');
          if (cells[5]) amount = cells[5].innerText.trim();
          if (cells[2]) {
            // 确认状态：去掉圆点只保留文字
            confirmStatus = cells[2].innerText.trim().replace(/\s+/g, '');
          }
          if (cells[8]) {
            const el = cells[8].querySelector('.status') || cells[8];
            payStatus = el.innerText.trim();
          }
          if (cells[9]) {
            const lines = cells[9].innerText.trim().split('\n').map(s => s.trim()).filter(Boolean);
            payDate = lines[0] || '';
          }
        }
        orders.push({ orderNo, productName, amount, payDate, payStatus, confirmStatus });
      } catch(e) {}
    });
    return orders;
  });
}

// ============================================================
// 解析产品名称 → 目的国、办签地、服务类型
// 典型格式：意大利个人旅游签证境外办签美国送签·单抢约服务
// ============================================================
function parseProduct(name = '') {
  const regions = [
    '意大利','西班牙','法国','德国','荷兰','比利时','葡萄牙','奥地利','希腊',
    '瑞士','捷克','波兰','丹麦','瑞典','挪威','芬兰','匈牙利','爱尔兰',
    '日本','韩国','美国','英国','澳大利亚','加拿大','新西兰',
    '新加坡','泰国','马来西亚','印度尼西亚','菲律宾','越南','印度',
    '土耳其','俄罗斯','阿联酋','埃及','南非','巴西','墨西哥',
    '申根','香港','澳门','台湾'
  ];

  let destCountry = '';
  let signCountry = '';

  // 目的国 = 产品名开头的国家
  for (const r of regions) {
    if (name.startsWith(r)) { destCountry = r; break; }
  }
  if (!destCountry) {
    for (const r of regions) {
      if (name.includes(r)) { destCountry = r; break; }
    }
  }

  // 办签地 = "境外办签XXX送签"里的XXX
  const m = name.match(/境外办签(.+?)送签/);
  if (m) {
    const mid = m[1];
    for (const r of regions) {
      if (mid.includes(r)) { signCountry = r; break; }
    }
    if (!signCountry) signCountry = mid.trim();
  }

  // 服务类型 = ·后面的内容
  let service = '';
  const dotIdx = name.indexOf('·');
  if (dotIdx >= 0) {
    service = name.slice(dotIdx + 1).replace(/【.*?】/g, '').trim().slice(0, 20);
  } else {
    if (name.includes('旅游')) service = '旅游签';
    else if (name.includes('商务')) service = '商务签';
    else service = '签证服务';
  }

  return { destCountry, signCountry, service };
}

// ============================================================
// 处理单个订单详情页
// ============================================================
async function processOrder(listPage, order) {
  const { destCountry, signCountry, service } = parseProduct(order.productName);
  order.destCountry = destCountry;
  order.signCountry = signCountry;
  order.service = service;

  let detailPage = null;
  try {
    const prevCount = (await browser.pages()).length;

    // 点击"查看订单"按钮
    await listPage.evaluate((orderNo) => {
      for (const row of document.querySelectorAll('tr.order-basic-row')) {
        const noEl = row.querySelector('.desc-item .value.link');
        if (noEl && noEl.innerText.trim() === orderNo) {
          const btn = Array.from(row.querySelectorAll('button'))
            .find(b => b.innerText.includes('查看订单'));
          if (btn) btn.click();
          return;
        }
      }
    }, order.orderNo);

    // 等待新标签页
    let waited = 0;
    while (waited < 15000) {
      await sleep(500);
      waited += 500;
      const nowPages = await browser.pages();
      if (nowPages.length > prevCount) {
        detailPage = nowPages[nowPages.length - 1];
        break;
      }
    }
    // 再等一次
    if (!detailPage) {
      await sleep(2000);
      const nowPages2 = await browser.pages();
      if (nowPages2.length > prevCount) {
        detailPage = nowPages2[nowPages2.length - 1];
      }
    }

    if (!detailPage) {
      log(`  Order ${order.orderNo}: no detail page opened`, 'warn');
      allRows.push({ ...order, travelerName: '详情页未打开', subOrders: [] });
      return;
    }

    await detailPage.waitForFunction(() => document.readyState === 'complete', { timeout: 15000 });
    await sleep(1500);

    // 提取出行人
    const travelers = await extractTravelers(detailPage);
    log(`  Travelers: ${travelers.map(t => t.name).join(', ')}`, 'success');

    for (const traveler of travelers) {
      if (shouldStop) break;
      const subOrders = await extractSubOrders(detailPage, traveler.name);
      // 合并到全局数据（同订单+出行人则覆盖）
      const newRow = {
        ...order,
        travelerName: traveler.name,
        travelerGender: traveler.gender || '',
        amount: traveler.amount || order.amount,
        subOrders,
      };
      const key = String(newRow.orderNo) + '_' + String(newRow.travelerName);
      const idx = allRows.findIndex(r => String(r.orderNo) + '_' + String(r.travelerName) === key);
      if (idx >= 0) { allRows[idx] = newRow; } else { allRows.push(newRow); }
    }

  } catch (e) {
    log(`  Order ${order.orderNo} failed: ${e.message}`, 'warn');
    allRows.push({ ...order, travelerName: '处理失败', subOrders: [] });
  } finally {
    // 确保详情页一定关闭
    if (detailPage) try { await detailPage.close(); } catch(e2) {}
  }
}

// ============================================================
// 提取出行人
// 结构：出行人区域 .resource-card-wrapper（title=出行人）
//   每行 tr.ant-table-row：
//     td[0] span[data-ignorechecktext] = 姓名
//     td[1] = 性别/类型
//     td最后 a.view-visa-dropdown = 查看签证子订单
// ============================================================
async function extractTravelers(page) {
  return await page.evaluate(() => {
    const travelers = [];

    // 找 title 为"出行人"的 resource-card-wrapper
    let travelerCard = null;
    document.querySelectorAll('.resource-card-wrapper').forEach(card => {
      const title = card.querySelector('.resource-head .title');
      if (title && title.innerText.trim() === '出行人') travelerCard = card;
    });

    // 如果没找到 resource-card-wrapper，尝试 Tab 布局（另一种详情页）
    if (!travelerCard) {
      // Tab 布局：点击"出行人"Tab 或直接找出行人表格
      const tabBtns = document.querySelectorAll('.ant-tabs-tab, [role="tab"]');
      for (const tab of tabBtns) {
        if (tab.innerText.trim() === '出行人') { tab.click(); break; }
      }
      // 等待渲染后找表格（在 evaluate 里不能 await，直接找现有内容）
      const travelerTable = document.querySelector('.ant-table-tbody');
      if (travelerTable) {
        const result = [];
        travelerTable.querySelectorAll('tr.ant-table-row').forEach(row => {
          const nameEl = row.querySelector('span[data-ignorechecktext]')
            || row.querySelector('td:first-child');
          if (!nameEl) return;
          const name = nameEl.innerText.trim();
          // 过滤：姓名格式 YUAN/HAOQIN 或中文名
          if (name && (name.includes('/') || /^[\u4e00-\u9fa5]{2,5}$/.test(name))) {
            result.push({ name, gender: '', amount: '' });
          }
        });
        if (result.length > 0) return result;
      }
      return [{ name: '(未找到出行人区域)', amount: '', gender: '' }];
    }

    const rows = travelerCard.querySelectorAll('.ant-table-tbody tr.ant-table-row');
    rows.forEach(row => {
      const cells = row.querySelectorAll('td.ant-table-cell');
      if (!cells.length) return;

      // 姓名：第1列第一个 span[data-ignorechecktext]
      const nameEl = cells[0] ? cells[0].querySelector('span[data-ignorechecktext]') : null;
      const name = nameEl ? nameEl.innerText.trim() : '';
      if (!name) return;

      // 性别：第2列
      const genderEl = cells[1];
      const gender = genderEl ? genderEl.innerText.trim().replace(/\s+/g, '/') : '';

      travelers.push({ name, gender, amount: '' });
    });

    // 提取结算金额
    const amtM = document.body.innerText.match(
      /(?:结算金额|实付金额|总金额|应付金额)[：:\s]*¥?\s*(\d+(?:\.\d{1,2})?)/
    );
    const total = amtM ? amtM[1] : '';

    if (travelers.length === 0) return [{ name: '(未识别)', amount: total, gender: '' }];

    if (travelers.length === 1) {
      travelers[0].amount = total;
    } else if (total) {
      const per = (parseFloat(total) / travelers.length).toFixed(2);
      travelers.forEach(t => { t.amount = per; });
    }

    return travelers;
  });
}
// ============================================================
// 提取子订单
// 流程：点击 a.view-visa-dropdown → dropdown li.drop-item → 
//       点击每个 li → 新标签页 visaOrderId=XXXX → 
//       页面顶部 span 内文字 = 状态
// ============================================================
async function extractSubOrders(page, travelerName) {
  const subOrders = [];
  try {
    // 找对应出行人行的 a.view-visa-dropdown 并点击
    const clicked = await page.evaluate((name) => {
      let travelerCard = null;
      document.querySelectorAll('.resource-card-wrapper').forEach(card => {
        const title = card.querySelector('.resource-head .title');
        if (title && title.innerText.trim() === '出行人') travelerCard = card;
      });
      if (!travelerCard) return false;

      const rows = travelerCard.querySelectorAll('.ant-table-tbody tr.ant-table-row');
      for (const row of rows) {
        const nameEl = row.querySelector('td:first-child span[data-ignorechecktext]');
        if (nameEl && nameEl.innerText.trim() === name) {
          const btn = row.querySelector('a.view-visa-dropdown');
          if (btn) { btn.click(); return true; }
        }
      }
      // 如果没精确匹配到，点第一个
      const firstBtn = travelerCard.querySelector('a.view-visa-dropdown');
      if (firstBtn) { firstBtn.click(); return true; }
      return false;
    }, travelerName);

    if (!clicked) {
      // 备用：Tab 布局里的"查看签证子订单"链接
      const altLinks = await page.evaluate((name) => {
        const links = Array.from(document.querySelectorAll('a'))
          .filter(a => a.innerText.includes('查看签证子订单'));
        return links.map(a => ({ href: a.href, text: a.innerText.trim() }))
          .filter(l => l.href.startsWith('http'));
      }, travelerName);

      if (altLinks.length > 0) {
        log(`    Using alt sub-order links: ${altLinks.length}`, 'default');
        for (const link of altLinks.slice(0, 6)) {
          if (shouldStop) break;
          let subPage = null;
          try {
            subPage = await browser.newPage();
            await subPage.goto(link.href, { waitUntil: 'domcontentloaded', timeout: 25000 });
            await sleep(1500);
            const subUrl = subPage.url();
            const idMatch = subUrl.match(/visaOrderId=(\d+)/);
            const visaOrderId = idMatch ? idMatch[1] : '';
            const { status, stepDetail } = await subPage.evaluate(() => {
              const keywords = ['已签约','待处理','已取消','进行中','已拒签','已出签','退款中','已退款','已归档','出签'];
              for (const span of document.querySelectorAll('span')) {
                if (keywords.includes(span.innerText.trim())) return { status: span.innerText.trim(), stepDetail: '' };
              }
              const doneSteps = [];
              document.querySelectorAll('.ant-timeline-item').forEach(item => {
                const c = item.querySelector('.ant-timeline-item-content');
                if (c && c.querySelector('.anticon-check')) doneSteps.push(c.innerText.trim().split('\n')[0]);
              });
              return { status: doneSteps.length > 0 ? doneSteps[doneSteps.length-1] : '状态未知', stepDetail: doneSteps.join('→') };
            });
            subOrders.push({ subNo: visaOrderId || link.text, status, stepDetail });
            log(`    Alt sub ${visaOrderId}: ${status}`, 'success');
          } catch(e) {
            subOrders.push({ subNo: link.text, status: '获取失败', stepDetail: '' });
          } finally {
            if (subPage) try { await subPage.close(); } catch(e) {}
          }
        }
        return subOrders;
      }

      log(`    No dropdown button for: ${travelerName}`, 'default');
      return subOrders;
    }

    await sleep(800);

    // 获取 dropdown 里所有 li，逐个点击获取新标签页
    const itemCount = await page.evaluate(() => {
      const items = document.querySelectorAll('.ant-dropdown:not(.ant-dropdown-hidden) li.drop-item');
      return items.length;
    });

    log(`    Found ${itemCount} sub-order items for ${travelerName}`, 'default');

    for (let i = 0; i < itemCount && !shouldStop; i++) {
      try {
        const prevCount = (await browser.pages()).length;

        // 重新点击按钮打开 dropdown（每次点完会关闭）
        if (i > 0) {
          await page.evaluate((name) => {
            let travelerCard = null;
            document.querySelectorAll('.resource-card-wrapper').forEach(card => {
              const title = card.querySelector('.resource-head .title');
              if (title && title.innerText.trim() === '出行人') travelerCard = card;
            });
            if (!travelerCard) return;
            // 精确匹配出行人姓名对应的按钮
            const rows = travelerCard.querySelectorAll('.ant-table-tbody tr.ant-table-row');
            for (const row of rows) {
              const nameEl = row.querySelector('td:first-child span[data-ignorechecktext]');
              if (nameEl && nameEl.innerText.trim() === name) {
                const btn = row.querySelector('a.view-visa-dropdown');
                if (btn) { btn.click(); return; }
              }
            }
          }, travelerName);
          await sleep(800);
        }

        // 点击第 i 个 li
        const itemText = await page.evaluate((idx) => {
          const items = document.querySelectorAll('.ant-dropdown:not(.ant-dropdown-hidden) li.drop-item');
          if (items[idx]) {
            const text = items[idx].innerText.trim();
            items[idx].click();
            return text;
          }
          return '';
        }, i);

        // 等待新标签页
        let subPage = null;
        let waited = 0;
        while (waited < 6000) {
          await sleep(400);
          waited += 400;
          const nowPages = await browser.pages();
          if (nowPages.length > prevCount) {
            subPage = nowPages[nowPages.length - 1];
            break;
          }
        }

        if (!subPage) {
          log(`    Sub-order ${i+1}: no new page opened`, 'warn');
          subOrders.push({ subNo: itemText, status: '页面未打开', stepDetail: '' });
          continue;
        }

        // 等待页面加载，超时后仍继续尝试提取
        try {
          await subPage.waitForFunction(() => document.readyState === 'complete', { timeout: 25000 });
        } catch(e) {
          log(`    Sub-order ${i+1}: load timeout, trying anyway...`, 'warn');
        }
        await sleep(1500);

        // 从 URL 提取 visaOrderId
        const subUrl = subPage.url();
        const idMatch = subUrl.match(/visaOrderId=(\d+)/);
        const visaOrderId = idMatch ? idMatch[1] : '';

        // 提取主状态 + 流程步骤（精确判断已完成步骤）
        const { status, stepDetail } = await subPage.evaluate(() => {
          // ── 主状态：顶部"签证"旁边的状态标签 ──
          // 结构：span>span 双层嵌套，内层无子元素
          const keywords = [
            '已签约','待处理','已取消','进行中','已拒签','已完成',
            '待收材料','材料审核中','已提交','签证中','已出签',
            '退款中','已退款','待审核','审核中','已寄出','已送签',
            '已归档','待归档','拒签','出签'
          ];
          let status = '';
          // 先用关键词精确匹配
          for (const span of document.querySelectorAll('span')) {
            if (keywords.includes(span.innerText.trim())) { status = span.innerText.trim(); break; }
          }
          // 备用：span>span 双层嵌套
          if (!status) {
            for (const span of document.querySelectorAll('span')) {
              if (span.childElementCount === 0) {
                const t = span.innerText.trim();
                const p = span.parentElement;
                if (t.length >= 2 && t.length <= 8 && p && p.tagName === 'SPAN' &&
                    p.childElementCount === 1 && /[\u4e00-\u9fa5]/.test(t) &&
                    !/首页|订单|管理|工作台|明细|添加|修改|编辑|删除|提交|操作/.test(t)) {
                  status = t; break;
                }
              }
            }
          }
          if (!status) status = '状态未知';

          // ── 流程步骤：只取有 anticon-check 图标的 timeline-item ──
          // 已完成：content 里有 <i class="anticon anticon-check">
          // 未完成：content 里是 <label class="ant-checkbox-wrapper">
          const doneSteps = [];
          document.querySelectorAll('.ant-timeline-item').forEach(item => {
            const contentEl = item.querySelector('.ant-timeline-item-content');
            if (!contentEl) return;
            const isChecked = !!contentEl.querySelector('.anticon-check');
            if (isChecked) {
              const text = contentEl.innerText.trim().split('\n')[0].trim();
              if (text) doneSteps.push(text);
            }
          });
          const stepDetail = doneSteps.join('→');

          return { status, stepDetail };
        });
        // 关闭子订单标签页
        try { await subPage.close(); } catch(e) {}
        subPage = null;


        log(`    Sub ${visaOrderId}: ${status} | ${stepDetail}`, 'success');
        subOrders.push({ subNo: visaOrderId || itemText, status, stepDetail });

      } catch(e) {
        log(`    Sub-order ${i+1} error: ${e.message}`, 'warn');
        // 尝试从已打开的页面补救提取
        if (subPage) {
          try {
            await sleep(2000);
            const salvageStatus = await subPage.evaluate(() => {
              const keywords = ['已签约','待处理','已取消','进行中','已拒签','已完成',
                '待收材料','材料审核中','已提交','签证中','已出签','退款中','已退款',
                '待审核','审核中','已寄出','已送签','已归档','拒签','出签'];
              for (const span of document.querySelectorAll('span')) {
                if (keywords.includes(span.innerText.trim())) return span.innerText.trim();
              }
              return '';
            });
            if (salvageStatus) {
              log(`    Salvaged status: ${salvageStatus}`, 'success');
              subOrders.push({ subNo: `子订单${i+1}`, status: salvageStatus, stepDetail: '' });
              await subPage.close();
              continue;
            }
          } catch(e2) {}
          try { await subPage.close(); } catch(e2) {} subPage = null;
        }
        subOrders.push({ subNo: `子订单${i+1}`, status: '获取失败', stepDetail: '' });
      }
    }

    // 关闭 dropdown
    await page.keyboard.press('Escape');
    await sleep(300);

  } catch(e) {
    log(`    extractSubOrders error: ${e.message}`, 'warn');
  }
  return subOrders;
}
// ============================================================
// 翻页
// ============================================================
async function nextPage(page) {
  try {
    const turned = await page.evaluate(() => {
      const btn = document.querySelector(
        'li.ant-pagination-next:not(.ant-pagination-disabled) button'
      );
      if (btn) { btn.click(); return true; }
      return false;
    });
    if (turned) await sleep(2500);
    return turned;
  } catch(e) { return false; }
}

// ============================================================
// 启动
// ============================================================
// 启动时加载本地数据
loadData();

const PORT = 3333;
server.listen(PORT, () => {
  console.log('');
  console.log('  ==========================================');
  console.log('   Ctrip Visa Scraper v2 - Ready!');
  console.log('  ==========================================');
  console.log(`   Open: http://localhost:${PORT}`);
  console.log('  ==========================================');
  console.log('');
  const { exec } = require('child_process');
  exec(`start http://localhost:${PORT}`);
});
