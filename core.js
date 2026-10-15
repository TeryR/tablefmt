/*
 * FlattenCore v0.1 — Excel 多级表头识别与展平（纯函数，浏览器/Node 双端可用）
 * 算法来源：格物 excel-agent headers.py 的思路移植 + 简化（v0.1 启发式）
 */
(function (global) {
  'use strict';

  function cellAt(rows, r, c) {
    const row = rows[r];
    return row ? row[c] : undefined;
  }

  // 把合并单元格锚点的值扩散到整个合并区（虚拟展开）
  function makeExpanded(rows, merges) {
    const map = new Map();
    (merges || []).forEach(m => {
      const v = cellAt(rows, m.s.r, m.s.c);
      for (let r = m.s.r; r <= m.e.r; r++)
        for (let c = m.s.c; c <= m.e.c; c++)
          if (r !== m.s.r || c !== m.s.c) map.set(r + ':' + c, v);
    });
    return (r, c) => (map.has(r + ':' + c) ? map.get(r + ':' + c) : cellAt(rows, r, c));
  }

  const isFilled = v => v != null && String(v).trim() !== '';
  function isNumLike(v) {
    if (v == null) return false;
    if (typeof v === 'number') return true;
    const t = String(v).trim();
    return t !== '' && !isNaN(Number(t));
  }

  /*
   * 探测表头深度 v0.3：
   * 主判据（数字游走）：表头行=文字占比高，首个数据行=数字占比≥0.3；
   * 回退（同构行游程，文字主导型表如费用台账）：数据行的特征是连续多行填充模式高度相似（Jaccard≥0.6 游程≥3），游程起点即表头深度。
   * v0.1 的“合并定深度”已废弃：数据区的纵向分类合并（科目列跨多行）会把深度投毒到 8+ （T6 实测）。
   */
  function fillMask(row, width) {
    const s = [];
    for (let c = 0; c < width; c++) if (isFilled(row ? row[c] : null)) s.push(c);
    return s;
  }
  function maskSim(a, b, width) {
    const ma = fillMask(a, width), mb = fillMask(b, width);
    const sa = new Set(ma);
    let inter = 0;
    mb.forEach(c => { if (sa.has(c)) inter++; });
    const union = ma.length + mb.length - inter;
    return union ? inter / union : 0;
  }
  function detectDepth(rows, merges, maxScan) {
    const val = makeExpanded(rows, merges);
    const scan = Math.min(rows.length, maxScan || 10);
    let numericPass = null;
    for (let d = 1; d < scan; d++) {
      const width = Math.max(1, (rows[d] || rows[0] || []).length);
      let filled = 0, text = 0;
      for (let r = 0; r < d; r++)
        for (let c = 0; c < width; c++) {
          const v = val(r, c);
          if (isFilled(v)) { filled++; if (!isNumLike(v)) text++; }
        }
      const textRatio = filled ? text / filled : 0;
      const drow = rows[d] || [];
      let dfilled = 0, dnum = 0;
      for (let c = 0; c < drow.length; c++) {
        const v = drow[c];
        if (isFilled(v)) { dfilled++; if (isNumLike(v)) dnum++; }
      }
      const numRatio = dfilled ? dnum / dfilled : 0;
      if (textRatio >= 0.6 && (dfilled === 0 || numRatio >= 0.3)) { numericPass = d; break; }
    }
    if (numericPass != null) return numericPass;
    const width = Math.max(1, ...rows.slice(0, 8).map(r => (r || []).length));
    for (let r = 0; r + 2 < rows.length; r++) {
      if (maskSim(rows[r], rows[r + 1], width) >= 0.6 && maskSim(rows[r + 1], rows[r + 2], width) >= 0.6) return r;
    }
    return 1;
  }

  /*
   * 展平表头：每列自上而下收集非空表头值，
   * 连续重复值折叠（纵向合并的锚值不重复），
   * 多级用 " / " 连接，重名自动加 (2)(3) 后缀
   */
  function flattenHeaders(rows, merges, depth) {
    const val = makeExpanded(rows, merges);
    const width = Math.max(1, ...rows.slice(0, Math.max(depth, 1)).map(r => (r || []).length));
    const names = [];
    const seen = {};
    for (let c = 0; c < width; c++) {
      const parts = [];
      for (let r = 0; r < depth; r++) {
        const v = val(r, c);
        if (isFilled(v)) {
          const s = String(v).trim();
          if (parts[parts.length - 1] !== s) parts.push(s);
        }
      }
      let name = parts.join(' / ') || ('列' + (c + 1));
      seen[name] = (seen[name] || 0) + 1;
      if (seen[name] > 1) name = name + ' (' + seen[name] + ')';
      names.push(name);
    }
    return names;
  }

  // 产出展平后的二维数组（新单行表头 + 数据）
  function buildFlattened(rows, merges, depth) {
    const names = flattenHeaders(rows, merges, depth);
    const val = makeExpanded(rows, merges);
    const out = [names];
    for (let r = depth; r < rows.length; r++) {
      const row = [];
      for (let c = 0; c < names.length; c++) {
        const v = val(r, c);
        row.push(v === undefined ? null : v);
      }
      out.push(row);
    }
    return out;
  }

  /*
   * 标题行识别：第0行存在跨度≥60%宽度的合并 → 是大标题行
   * 返回 { isTitle, title, startRow }
   */
  function detectTitleRow(rows, merges) {
    const width = Math.max(1, (rows[0] || []).length);
    let best = null;
    (merges || []).forEach(m => {
      if (m.s.r !== 0) return;
      const span = m.e.c - m.s.c + 1;
      if (span >= Math.max(2, Math.ceil(width * 0.6)) && (!best || span > best.span)) {
        best = { span, anchor: m.s.c, end: m.e.c };
      }
    });
    if (best) {
      return { isTitle: true, title: String(rows[0]?.[best.anchor] ?? '').trim(), startRow: 1 };
    }
    return { isTitle: false, title: null, startRow: 0 };
  }

  /*
   * 剥离标题行（含合并锚点值搬运）：
   * 纵向合并的锚点若在被剥离的标题行里，其值必须搬到新首行对应列——
   * 否则"区域/备注"这类跨标题行合并的文字会随标题一起丢失
   */
  function stripTitleRow(rows, merges, startRow) {
    const newMerges = [];
    const carried = [];
    (merges || []).forEach(m => {
      if (m.e.r < startRow) return; // 完全在剥离区内
      const nsr = Math.max(m.s.r - startRow, 0);
      const ner = m.e.r - startRow;
      newMerges.push({ s: { r: nsr, c: m.s.c }, e: { r: ner, c: m.e.c } });
      if (m.s.r < startRow) {
        carried.push({ r: nsr, c: m.s.c, v: cellAt(rows, m.s.r, m.s.c) });
      }
    });
    const newRows = rows.slice(startRow).map(r => (r || []).slice());
    carried.forEach(({ r, c, v }) => {
      while (newRows.length <= r) newRows.push([]);
      if (!isFilled(newRows[r][c])) newRows[r][c] = v;
    });
    return { rows: newRows, merges: newMerges };
  }

  /*
   * 空值统计：每列空值占比（数据区内），识别"周更/非日更"类字段
   */
  function emptyStats(rows, startRow, width) {
    const stats = [];
    const dataRows = rows.length - startRow;
    for (let c = 0; c < width; c++) {
      let empty = 0, filled = 0;
      for (let r = startRow; r < rows.length; r++) {
        const row = rows[r];
        const v = row ? row[c] : undefined;
        (v != null && String(v).trim() !== '') ? filled++ : empty++;
      }
      stats.push({ col: c, empty, filled, ratio: dataRows ? empty / dataRows : 0 });
    }
    return stats;
  }

  /*
   * 生成 LLM 可读上下文（v0.2 核心产出：不改原表，产出"理解包"）
   * 输入：展平后的 rows（含表头行）、原始表头层数、空值统计、标题
   */
  function toLLMContext(flat, origDepth, empties, title) {
    const names = flat[0];
    const dataRows = flat.length - 1;
    let out = '以下是一张 Excel 表格的结构化转写，请先理解表格结构，再回答我稍后提出的问题。\n\n';

    if (title) out += '## 表标题\n' + title + '\n\n';

    out += '## 表结构\n';
    out += `- 数据区 ${dataRows} 行 × ${names.length} 列；原表表头为 ${origDepth} 层合并表头，已展平（层级关系用 / 分隔）\n`;
    out += '- 列清单：\n';
    names.forEach((n, i) => out += `  ${i + 1}. ${n}\n`);

    const noisy = (empties || []).filter(e => e.ratio >= 0.4);
    if (noisy.length) {
      out += '- 空值提示：\n';
      noisy.forEach(e => out += `  - 列「${names[e.col]}」含 ${e.empty}/${dataRows} 个空值：疑似非日更汇总字段或占位列，空值不代表缺失\n`);
    }

    out += '\n## 数据\n\n';
    out += '| # | ' + names.join(' | ') + ' |\n';
    out += '|' + ('---|').repeat(names.length + 1) + '\n';
    for (let r = 1; r < flat.length; r++) {
      out += '| ' + r + ' | ' + flat[r].map(v => v == null ? '' : String(v)).join(' | ') + ' |\n';
    }
    return out;
  }

  global.FlattenCore = { detectDepth, flattenHeaders, buildFlattened, makeExpanded, isFilled, isNumLike, detectTitleRow, stripTitleRow, emptyStats, toLLMContext };
})(typeof window !== 'undefined' ? window : globalThis);
