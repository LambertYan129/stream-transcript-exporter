// ==UserScript==
// @name         Microsoft Stream 转录导出
// @namespace    local.stream.transcript
// @version      1.0.0
// @description  按字幕序号收集当前视频的转录，校验后导出带时间戳的 TXT。
// @match        https://*.sharepoint.com/*
// @match        https://*.microsoftstream.com/*
// @match        https://*.cloud.microsoft/*
// @match        https://*.office.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==
(() => {
  'use strict';

  const indexed = '[aria-posinset]';
  const stamp = /^\[?((?:\d+:)?\d{1,3}:[0-5]\d)\]?$/;
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const compact = value => value.replace(/\s+/g, ' ').trim();

  function captionText(timestamp, text) {
    const match = text.match(/^(?:(\d+)\s*(?:小时|小時)\s*)?(?:(\d+)\s*(?:分钟|分鐘)\s*)?(?:(\d+)\s*秒\s*)?/);
    const seconds = timestamp.split(':').reduce((sum, part) => sum * 60 + Number(part), 0);
    const labelSeconds = Number(match[1] || 0) * 3600 + Number(match[2] || 0) * 60 + Number(match[3] || 0);
    return match[0] && seconds === labelSeconds ? text.slice(match[0].length).trimStart() : text;
  }
  function parseRow(row) {
    const lines = (row.innerText || '').split(/\n/).map(compact).filter(Boolean);
    const i = lines.findIndex(line => stamp.test(line));
    if (i >= 0 && i + 1 < lines.length) {
      const timestamp = lines[i].match(stamp)[1];
      const text = captionText(timestamp, lines.slice(i + 1).join(' '));
      return text ? '[' + timestamp + '] ' + text : null;
    }
    const inline = lines.join(' ').match(/^\[?((?:\d+:)?\d{1,3}:[0-5]\d)\]?\s+(.+)$/);
    const text = inline && captionText(inline[1], inline[2]);
    return text ? '[' + inline[1] + '] ' + text : null;
  }

  // Read time and text separately even when the browser joins the columns.
  function columnRows(root) {
    const doc = root.ownerDocument || root;
    const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const times = [];
    for (let node; (node = walker.nextNode());) {
      if (stamp.test(node.data.trim()) && node.parentElement?.getClientRects().length) times.push(node);
    }
    const rows = [];
    for (let i = 0; i < times.length; i++) {
      const time = times[i];
      for (let el = time.parentElement; el && el !== root && el !== doc.body; el = el.parentElement) {
        if (el.contains(times[i-1] || null) || el.contains(times[i+1] || null)) break;
        const range = doc.createRange();
        range.selectNodeContents(el);
        range.setStartAfter(time);
        const timestamp = time.data.trim().match(stamp)[1];
        const text = captionText(timestamp, compact(range.toString()));
        if (!text) continue;
        const marker = el.closest(indexed) || el.querySelector(indexed);
        rows.push({el, marker, value:'[' + timestamp + '] ' + text});
        break;
      }
    }
    return rows;
  }

  function readRows(root) {
    const columns = columnRows(root);
    if (columns.length) return columns;
    return [...root.querySelectorAll(indexed)].map(el => ({el, marker:el, value:parseRow(el)})).filter(r => r.value);
  }

  function findBox(root) {
    for (const {el} of readRows(root)) {
      let list = el.closest('[role="list"],[role="listbox"],[role="tree"]');
      for (let box = el.parentElement; box && box !== el.ownerDocument.body; box = box.parentElement) {
        if (box.scrollHeight > box.clientHeight && box.clientHeight &&
            /auto|scroll|overlay|hidden/.test(box.ownerDocument.defaultView.getComputedStyle(box).overflowY)) return box;
      }
      if (list) return list;
    }
    return null;
  }

  function panel(root = document) {
    const box = findBox(root);
    if (box) return box;
    for (const el of root.querySelectorAll('*')) {
      const inner = el.shadowRoot || (el.tagName === 'IFRAME' && el.contentDocument);
      if (inner) {
        const nested = panel(inner);
        if (nested) return nested;
      }
    }
    return null;
  }

  function appendWindow(lines, next) {
    if (!lines.length) { lines.push(...next); return; }
    for (let count = Math.min(lines.length, next.length); count > 0; count--) {
      if (next.slice(0, count).every((line, i) => line === lines[lines.length - count + i])) {
        lines.push(...next.slice(count));
        return;
      }
    }
    throw Error('滚动时字幕未连续衔接，请暂停视频后重新导出');
  }
  const bar = document.createElement('div');
  bar.style.cssText = 'position:fixed;z-index:2147483646;display:flex;align-items:center;gap:8px;padding:6px;background:white;color:#222;border:1px solid #aaa;border-radius:5px;font:13px system-ui';
  const button = document.createElement('button');
  bar.style.right = '20px';
  bar.style.bottom = '20px';
  button.style.cssText = 'background:#5031a5;color:white;border:0;border-radius:4px;padding:9px 14px;font:600 14px system-ui;cursor:pointer';
  button.textContent = '导出转录 TXT';
  const status = document.createElement('span');
  status.setAttribute('role', 'status');
  bar.append(button, status);
  let busy = false;
  function refresh() {
    const streamPage = /\/stream\.aspx$/i.test(location.pathname);
    if (!busy && !streamPage && !document.querySelector(indexed)) {
      bar.remove();
      return;
    }
    if (!bar.isConnected) document.body.append(bar);
  }
  function save(lines) {
    const title = document.querySelector('meta[property="og:title"]')?.content || document.title;
    const name = title.replace(/\s+[-|–]\s+Microsoft (?:Stream|SharePoint).*$/i, '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').trim().slice(0, 160).replace(/[. ]+$/, '') || '视频转录';
    const url = URL.createObjectURL(new Blob([lines.join('\r\n') + '\r\n'], {type:'text/plain;charset=utf-8'}));
    const a = document.createElement('a');
    a.href = url;
    a.download = name + '.txt';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  button.onclick = async () => {
    if (busy) return;
    const box = panel();
    if (!box) {
      status.textContent = '未识别到字幕行，请展开转录面板，等字幕出现后再点一次';
      return;
    }
    busy = true;
    button.disabled = true;

    const original = box.scrollTop;
    const startUrl = location.href;
    const records = new Map();
    let total = 0;
    let previous;
    try {
      for (let pass = 0; pass < 3; pass++) {
        box.scrollTop = 0;
        let bottomSamples = 0;
        const lines = [];
        while (bottomSamples < 4) {
          await wait(220 + pass * 100);
          if (!box.isConnected || location.href !== startUrl) throw Error('页面已切换，请重新导出');
          const rows = readRows(box);
          if (!rows.length) throw Error('字幕列表为空，请暂停视频并重试');
          for (const {marker, value} of rows) {
            if (!marker) continue;
            const pos = Number(marker.getAttribute('aria-posinset'));
            const size = Number(marker.closest('[aria-setsize]')?.getAttribute('aria-setsize'));
            if (!Number.isInteger(pos) || pos < 1) continue;
            if (Number.isInteger(size) && size > 0) {
              if (total && total !== size) throw Error('条目总数发生变化，请重新导出');
              total = size;
            }
            records.set(pos, value);
          }
          if (!total) appendWindow(lines, rows.map(r => r.value));
          status.textContent = total ? '已读取 ' + records.size + ' / ' + total : '第 ' + (pass+1) + ' 遍，已读取 ' + lines.length + ' 条';
          if (total && records.size === total && [...records.keys()].every(pos => pos <= total)) {
            save(Array.from({length:total}, (_, i) => records.get(i + 1)));
            status.textContent = '已校验并导出 ' + total + ' 条';
            return;
          }
          const end = Math.max(0, box.scrollHeight - box.clientHeight);
          if (box.scrollTop >= end - 1) bottomSamples++;
          else {
            bottomSamples = 0;
            const before = box.scrollTop;
            box.scrollTop = Math.min(end, before + Math.max(1, Math.floor(box.clientHeight * 0.65)));
            if (box.scrollTop === before) throw Error('转录面板无法滚动，请重新打开面板');
          }
        }
        if (!total && previous && lines.length === previous.length && lines.every((line,i) => line === previous[i])) {
          save(lines);
          status.textContent = '已导出 ' + lines.length + ' 条（两遍一致，页面未提供总数）';
          return;
        }
        previous = lines;
      }
      if (!total) throw Error('多次扫描结果不一致，请暂停视频并重试');
      throw Error('仅收集到 ' + records.size + ' / ' + total + ' 条，请重新导出');
    } catch (error) {
      status.textContent = error.message;
    } finally {
      if (box.isConnected && startUrl === location.href) box.scrollTop = original;
      busy = false;
      button.disabled = false;
      refresh();
    }
  };

  refresh();
  setInterval(refresh, 1200);
})();
