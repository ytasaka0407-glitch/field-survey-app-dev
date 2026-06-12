// FieldSurveyApp/assets/js/modules/excel/export.js
import { dataMap, selectedCategories, ensureSingle, ensureMulti, getOrInitStationData } from '../state.js';
import { sanitizeSheetName, makeUniqueSheetName, getImageDim } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

export async function exportToExcel(projectTitle, projectDate, sharedStations) {
  const ExcelJSRef = window.ExcelJS;
  const wb = new ExcelJSRef.Workbook();
  wb.creator = '現地調査レポートツール';
  wb.created = new Date();
  wb.properties.title = '現地調査報告書';

  const titleStyle        = { font: { size: 20, bold: true, name: 'Meiryo UI' }, alignment: { horizontal: 'center', vertical: 'middle' } };
  const sectionTitleStyle = { font: { size: 14, bold: true, name: 'Meiryo UI' } };
  const labelStyle        = { font: { bold: true, name: 'Meiryo UI' } };
  const linkStyle         = { font: { color: { argb: 'FF1F4E79' }, underline: true, name: 'Meiryo UI' } };
  const borderThin        = { style: 'thin', color: { argb: 'FF999999' } };

  // 安全な文字列長（Excel 1セル=最大32767文字。余裕を見て32000に制限）
  const safeStr = (s, max = 32000) => {
    const t = (s ?? '').toString();
    return t.length > max ? t.slice(0, max) : t;
  };

  // 日付をローカル日付文字列へ正規化
  const formatLocalDate = (value) => {
    if (!value) return '';

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      const d = String(value.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    }

    if (typeof value === 'string') {
      const s = value.trim();
      const m = /^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/.exec(s);
      if (m) return `${m[1]}/${m[2]}/${m[3]}`;
      return s;
    }

    if (typeof value === 'number') {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(epoch.getTime() + value * 86400000);
      if (Number.isNaN(d.getTime())) return '';
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${y}/${m}/${day}`;
    }

    return String(value ?? '');
  };

  const projectDateStr = formatLocalDate(projectDate);
  const coverDateText = projectDateStr || '';

  // 表紙
  const wsCover = wb.addWorksheet('表紙', {
    pageSetup: {
      paperSize: 9,
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
    },
    headerFooter: { oddFooter: '&Rページ &P / &N' }
  });
  wsCover.views = [{ showGridLines: false }];

  for (let i = 1; i <= 12; i++) wsCover.getColumn(i).width = 16;
  wsCover.mergeCells('A3:H6');
  wsCover.getCell('A3').value = '現地調査報告書';
  wsCover.getCell('A3').style = titleStyle;
  wsCover.getCell('G8').value = '案件名'; wsCover.getCell('G8').style = labelStyle;
  wsCover.getCell('H8').value = safeStr(projectTitle || '-');
  wsCover.getCell('H8').font = { name: 'Meiryo UI' };
  wsCover.getCell('G10').value = '日付'; wsCover.getCell('G10').style = labelStyle;
  wsCover.getCell('H10').value = coverDateText || '-';
  wsCover.getCell('H10').font = { name: 'Meiryo UI' };
  wsCover.getCell('H10').alignment = { horizontal: 'left' };

  // 目次
  const wsToc = wb.addWorksheet('目次', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
    headerFooter: { oddFooter: '&Rページ &P / &N' }
  });
  wsToc.views = [{ showGridLines: false }];
  wsToc.getColumn(1).width = 50;
  wsToc.getColumn(2).width = 18;
  wsToc.getCell('A1').value = '目次';
  wsToc.getCell('A1').style = sectionTitleStyle;

  // エントリ生成（並び: 単一 → 基地局ごとマルチ）
  const cats = [...selectedCategories];
  const selectedSingles = cats.filter(c => (dataMap[c]?.mode || 'single') === 'single');
  const selectedMultis  = cats.filter(c => (dataMap[c]?.mode) === 'multi');

  const leadingNum = (name) => {
    const m = /^(\d+)\./.exec(name || '');
    return m ? Number(m[1]) : Number.POSITIVE_INFINITY;
  };

  const singleEntries = [];
  for (const cat of selectedSingles) {
    const v = ensureSingle(cat, projectDateStr);
    singleEntries.push({ type: 'single', cat, stationId: null, stationName: null, displayLabel: cat, model: v });
  }

  singleEntries.sort((a, b) => {
    const na = leadingNum(a.cat), nb = leadingNum(b.cat);
    if (na !== nb) return na - nb;
    return a.cat.localeCompare(b.cat, 'ja');
  });

  let stationLoop = [];
  if (Array.isArray(sharedStations) && sharedStations.length) {
    stationLoop = sharedStations.map(s => ({ id: s.id, name: s.name }));
  } else {
    const stationIdSet = new Map();
    for (const cat of selectedMultis) {
      const mv = ensureMulti(cat, projectDateStr);
      const sd = mv.stationData || {};
      Object.keys(sd).forEach(id => {
        if (!stationIdSet.has(id)) stationIdSet.set(id, id);
      });
    }
    stationLoop = Array.from(stationIdSet.entries()).map(([id, name]) => ({ id, name }));
    stationLoop.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  }

  const multiCatsSorted = selectedMultis.slice().sort((a, b) => {
    const na = leadingNum(a), nb = leadingNum(b);
    if (na !== nb) return na - nb;
    return a.localeCompare(b, 'ja');
  });

  const multiEntries = [];
  for (const st of stationLoop) {
    for (const cat of multiCatsSorted) {
      const mv = ensureMulti(cat, projectDateStr);
      const stData = getOrInitStationData(cat, st.id, projectDateStr);
      if (stData._excluded) continue;
      multiEntries.push({
        type: 'multi',
        cat,
        stationId: st.id,
        stationName: st.name,
        displayLabel: `${cat}（${st.name}）`,
        model: stData
      });
    }
  }

  const entries = [...singleEntries, ...multiEntries];

  const usedNames = new Set();
  for (const e of entries) {
    const base = e.type === 'single'
      ? sanitizeSheetName(e.cat)
      : sanitizeSheetName(`${e.cat} - ${e.stationName}`);
    e.sheetName = makeUniqueSheetName(base, usedNames);
    usedNames.add(e.sheetName);
  }

  const colLetterToIndex = (L) => L.charCodeAt(0) - 65;
  function applyBottomBorderRange(ws, startCol, endCol, row) {
    for (let c = startCol; c <= endCol; c++) {
      ws.getCell(row, c).border = { bottom: borderThin };
    }
  }

  const photoManifests = [];

  async function addOneEntrySheet(entry) {
    const ws = wb.addWorksheet(entry.sheetName, {
      pageSetup: {
        paperSize: 9,
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
      },
      headerFooter: { oddFooter: '&Rページ &P / &N' }
    });
    ws.views = [{ showGridLines: false }];

    const colWidths = [20, 16, 16, 18, 16, 16, 18, 16, 16, 18, 10, 10];
    colWidths.forEach((w, i) => ws.getColumn(i + 1).width = w);
    ws.getColumn(11).hidden = true;

    ws.getCell('A1').value = entry.displayLabel;
    ws.getCell('A1').style = sectionTitleStyle;
    ws.getCell('J1').value = { text: '目次へ戻る', hyperlink: `#'目次'!A1` };
    ws.getCell('J1').style = linkStyle;
    ws.getCell('J1').alignment = { horizontal: 'right', vertical: 'middle' };

    // 調査日
    ws.getCell('A3').value = '調査日';
    ws.getCell('A3').style = labelStyle;
    const dText = formatLocalDate(entry.model.date || projectDate);
    ws.getCell('B3').value = dText || '-';
    ws.getCell('B3').font = { name: 'Meiryo UI' };
    ws.getCell('B3').alignment = { horizontal: 'left' };

    // 設置場所
    ws.getCell('A4').value = '設置場所';
    ws.getCell('A4').style = labelStyle;
    ws.mergeCells('B4:J4');
    ws.getCell('B4').value = safeStr(entry.model.location || '-');
    ws.getCell('B4').font = { name: 'Meiryo UI' };

    applyBottomBorderRange(ws, 1, 1, 3);
    applyBottomBorderRange(ws, 2, 2, 3);
    applyBottomBorderRange(ws, 1, 1, 4);
    applyBottomBorderRange(ws, 2, 10, 4);

    let nextRow = 5;

    // 新設/既設流用
    ws.getCell(`A${nextRow}`).value = '新設/既設流用';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow}`);
    ws.getCell(`B${nextRow}`).value = (entry.model.installType === 'reuse') ? '既設流用' : '新設';
    ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    nextRow++;

    // 設置方法（新設時のみ）
    if (entry.model.installType === 'new') {
      ws.getCell(`A${nextRow}`).value = '設置方法';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      ws.mergeCells(`B${nextRow}:J${nextRow}`);
      ws.getCell(`B${nextRow}`).value = safeStr((entry.model.method ?? '').toString() || '-');
      ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };

      applyBottomBorderRange(ws, 1, 1, nextRow);
      applyBottomBorderRange(ws, 2, 10, nextRow);
      nextRow++;
    }

    // 系統図との整合性
    ws.getCell(`A${nextRow}`).value = '系統図との整合性';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow}`);
    ws.getCell(`B${nextRow}`).value = (entry.model.diagramStatus === 'ng') ? 'NG' : 'OK';
    ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    nextRow++;

    // NG理由
    if (entry.model.diagramStatus === 'ng') {
      ws.getCell(`A${nextRow}`).value = 'NG理由';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      ws.mergeCells(`B${nextRow}:J${nextRow}`);
      ws.getCell(`B${nextRow}`).value = safeStr((entry.model.diagramNgReason ?? '').toString() || '-');
      ws.getCell(`B${nextRow}`).alignment = { wrapText: true, vertical: 'top' };
      ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };

      applyBottomBorderRange(ws, 1, 1, nextRow);
      applyBottomBorderRange(ws, 2, 10, nextRow);
      nextRow++;
    }

    // その他調査内容
    ws.getCell(`A${nextRow}`).value = 'その他調査内容';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow + 1}`);
    ws.getCell(`B${nextRow}`).value = safeStr((entry.model.details || '').replace(/\r?\n/g, '\n'));
    ws.getCell(`B${nextRow}`).alignment = { wrapText: true, vertical: 'top' };
    ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    ws.getRow(nextRow).height = 22;
    ws.getRow(nextRow + 1).height = 22;
    applyBottomBorderRange(ws, 1, 1, nextRow + 1);
    applyBottomBorderRange(ws, 2, 10, nextRow + 1);
    nextRow += 2;

    // 追加項目
    const schema = getSchemaFor(entry.cat, entry.type === 'multi' ? 'multi' : 'single');
    const CORE_KEYS = new Set(['date', 'location', 'installType', 'method', 'diagramStatus', 'diagramNgReason', 'details', 'photos']);
    const extraFields = schema.filter(f => !CORE_KEYS.has(f.key));

    if (extraFields.length) {
      ws.getCell(`A${nextRow}`).value = '追加項目';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      let extraStartRow = nextRow + 1;

      for (let i = 0; i < extraFields.length; i++) {
        const f = extraFields[i];
        const row = extraStartRow + i;
        ws.getCell(`A${row}`).value = f.label || f.key;
        ws.getCell(`A${row}`).font = { name: 'Meiryo UI' };
        ws.mergeCells(`B${row}:J${row}`);
        const val = entry.model[f.key];
        ws.getCell(`B${row}`).value = safeStr((val ?? '').toString());
        ws.getCell(`B${row}`).alignment = { wrapText: true, vertical: 'top' };
        ws.getCell(`B${row}`).font = { name: 'Meiryo UI' };
        ws.getCell(`K${row}`).value = f.key;
      }
      nextRow = extraStartRow + extraFields.length;
    }

    // 写真
    let startRow = nextRow + 1;
    const photos = Array.isArray(entry.model.photos) ? entry.model.photos : [];
    const validPhotos = photos.filter(p => p?.dataUrl && String(p.dataUrl).startsWith('data:image/'));

    const IMAGE_COLS = ['B', 'C', 'D', 'E', 'F'];
    const DESC_COL_START = 'G';
    const DESC_COL_END = 'J';

    const ROW_HEIGHT_PT = 24;
    const BLOCK_ROWS = 11;
    const GAP_ROWS = 2;

    const colPixels = (colIdx) => (ws.getColumn(colIdx + 1).width || 10) * 7;
    const rowPixels = (rowIdx) => (ws.getRow(rowIdx).height || 18) * 1.333;
    const sumColPixels = (letters) =>
      letters.reduce((sum, L) => sum + colPixels(colLetterToIndex(L)), 0);

    const containerW = sumColPixels(IMAGE_COLS);

    for (let i = 0; i < validPhotos.length; i++) {
      const p = validPhotos[i];

      for (let r = startRow; r < startRow + BLOCK_ROWS; r++) {
        ws.getRow(r).height = ROW_HEIGHT_PT;
      }

      let containerH = 0;
      for (let rr = startRow; rr < startRow + BLOCK_ROWS; rr++) {
        containerH += rowPixels(rr);
      }

      const { w: imgW, h: imgH } = await getImageDim(p.dataUrl);
      const ratioW = containerW / imgW;
      const ratioH = containerH / imgH;
      const ratio = Math.min(ratioW, ratioH);
      const drawW = Math.max(1, Math.floor(imgW * ratio));
      const drawH = Math.max(1, Math.floor(imgH * ratio));

      const imgId = wb.addImage({
        base64: p.dataUrl.split(',')[1],
        extension: p.dataUrl.startsWith('data:image/png') ? 'png' : 'jpeg',
      });
      ws.addImage(imgId, {
        tl: { col: colLetterToIndex(IMAGE_COLS[0]), row: startRow - 1 },
        ext: { width: drawW, height: drawH },
      });

      const descRange = `${DESC_COL_START}${startRow}:${DESC_COL_END}${startRow + BLOCK_ROWS - 1}`;
      ws.mergeCells(descRange);
      const descCell = ws.getCell(`${DESC_COL_START}${startRow}`);
      descCell.value = safeStr((p.caption || '').replace(/\r\n?/g, '\n'));
      descCell.alignment = { wrapText: true, vertical: 'top' };
      descCell.font = { name: 'Meiryo UI' };
      descCell.border = {
        top: borderThin,
        left: borderThin,
        bottom: borderThin,
        right: borderThin,
      };

      const fileName = p.name || `photo_${i + 1}.jpg`;
      photoManifests.push({
        type: entry.type,
        sheetName: entry.sheetName,
        category: entry.cat,
        station: entry.type === 'multi' ? (entry.stationName || '') : '',
        fileName,
        caption: safeStr(p.caption || '', 32000),
        imgCol: IMAGE_COLS[0],
        imgRowStart: startRow,
      });

      startRow += (BLOCK_ROWS + GAP_ROWS);
    }
  }

  for (const e of entries) {
    await addOneEntrySheet(e);
  }

  let tocRow = 3;
  for (const e of entries) {
    wsToc.getCell(`A${tocRow}`).value = { text: e.displayLabel, hyperlink: `#'${e.sheetName}'!A1` };
    wsToc.getCell(`A${tocRow}`).style = linkStyle;
    wsToc.getCell(`A${tocRow}`).alignment = { horizontal: 'left', vertical: 'middle' };
    const photosCount = Array.isArray(e.model.photos) ? e.model.photos.length : 0;
    wsToc.getCell(`B${tocRow}`).value = `写真 ${photosCount}枚`;
    wsToc.getCell(`B${tocRow}`).font = { name: 'Meiryo UI' };
    tocRow++;
  }

  const wsPhotos = wb.addWorksheet('PHOTOS');
  wsPhotos.getColumn(1).width = 10;
  wsPhotos.getColumn(2).width = 28;
  wsPhotos.getColumn(3).width = 36;
  wsPhotos.getColumn(4).width = 24;
  wsPhotos.getColumn(5).width = 36;
  wsPhotos.getColumn(6).width = 60;
  wsPhotos.getColumn(7).width = 8;
  wsPhotos.getColumn(8).width = 10;
  wsPhotos.getRow(1).values = ['type', 'sheetName', 'category', 'station', 'fileName', 'caption', 'imgCol', 'imgRowStart'];

  let pr = 2;
  for (const rec of photoManifests) {
    wsPhotos.getRow(pr).values = [
      rec.type,
      rec.sheetName,
      rec.category,
      rec.station,
      rec.fileName,
      rec.caption,
      rec.imgCol,
      rec.imgRowStart
    ];
    pr++;
  }
  wsPhotos.state = 'hidden';

  const buf = await wb.xlsx.writeBuffer();
  const title = (projectTitle || '').trim();
  const suffix = '現地調査レポート';
  const namePart = title ? `${title}_${suffix}` : suffix;
  const safeNamePart = namePart.replace(/[\\/:*?"<>|]/g, '_');
  const fileName = `${safeNamePart}.xlsx`;
  window.saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
}
