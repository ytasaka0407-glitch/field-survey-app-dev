// 共通ユーティリティ
export function sanitizeSheetName(name) {
  const invalid = /[\\\/\?\*\[\]\:]/g;
  let n = (name || "").replace(invalid, " ").trim();
  if (!n) n = "シート";
  if (n.length > 31) n = n.slice(0, 31);
  return n;
}

export function makeUniqueSheetName(base, used) {
  let n = base;
  let i = 2;
  while (used.has(n)) {
    const suffix = ` (${i})`;
    const core = base.slice(0, 31 - suffix.length);
    n = core + suffix;
    i++;
  }
  return n;
}

export function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function fromInputDate(str) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str || '');
  if (!m) return '';
  return `${m[1]}-${m[2]}-${m[3]}`;
}

export function toInputDateString(val) {
  if (!val) return '';

  // JS Date / ExcelJS Date はローカル日付で文字列化
  if (val instanceof Date) {
    if (Number.isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const day = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  // ExcelJS の richText / text / result 対応
  if (val && typeof val === 'object' && 'result' in val) {
    return toInputDateString(val.result);
  }
  if (val && typeof val === 'object' && 'text' in val) {
    return toInputDateString(val.text);
  }

  // 文字列は yyyy-mm-dd / yyyy/mm/dd を正規化
  if (typeof val === 'string') {
    const s = val.trim();

    const m = /^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/.exec(s);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;

    // それ以外はそのまま返す
    return s;
  }

  // Excel serial date
  if (typeof val === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + val * 86400000);
    if (Number.isNaN(d.getTime())) return '';
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return '';
}

export function getImageDim(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || img.width, h: img.naturalHeight || img.height });
    img.src = dataUrl;
  });
}

export function hashId(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  return Math.abs(h).toString(36);
}

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[m]);
}

export function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

export function resizeImage(dataUrl, maxWidth = 1024, quality = 0.8) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxWidth / img.width);
      const w = Math.round(img.width * scale);
      const h = Math.round(img.height * scale);
      const canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      const out = canvas.toDataURL("image/jpeg", quality);
      resolve(out);
    };
    img.src = dataUrl;
  });
}

export function stationIdFromName(name) {
  return "st_" + hashId((name || "").trim().toLowerCase());
}
