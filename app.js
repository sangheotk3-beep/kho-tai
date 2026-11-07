'use strict';

/* Điền owner và repo nếu dùng tên miền riêng. Trên *.github.io trang tự nhận ra repo. */
const CAU_HINH = { owner: '', repo: '' };

const NHOM = { 'tai-lieu': 'Tài liệu', vba: 'VBA' };
const DUOI_VBA = ['bas', 'cls', 'frm', 'frx', 'vba', 'xlsm', 'xlsb', 'xlam', 'xla', 'docm', 'dotm', 'pptm', 'accdb', 'mdb'];
const DUOI_CODE = ['bas', 'cls', 'frm', 'vba'];
const DUOI_CHU = [...DUOI_CODE, 'txt', 'md', 'csv', 'json', 'xml'];
const DUOI_ANH = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'];
const DUOI_OFFICE = ['doc', 'docx', 'docm', 'dotx', 'xls', 'xlsx', 'xlsm', 'xlsb', 'ppt', 'pptx', 'pptm'];
const MIME_ANH = { svg: 'image/svg+xml', jpg: 'image/jpeg' };

function kieuXem(ten) {
  const d = duoi(ten);
  if (DUOI_CHU.includes(d)) return 'chu';
  if (d === 'pdf') return 'pdf';
  if (DUOI_ANH.includes(d)) return 'anh';
  if (DUOI_OFFICE.includes(d)) return 'office';
  return '';
}
const GIOI_HAN_MB = 25;
const TOKEN_KEY = 'kho-tai-lieu-token';

const KHOA_VBA = 'Sub|Function|End|Dim|As|Set|If|Then|Else|ElseIf|For|Each|In|To|Step|Next|Do|Loop|While|Wend|Until|Select|Case|With|Exit|Public|Private|Friend|Static|Const|Option|Explicit|Type|Enum|Property|Get|Let|ByVal|ByRef|Optional|ParamArray|New|Nothing|Not|And|Or|Xor|Is|Like|Mod|True|False|Call|GoTo|On|Error|Resume|ReDim|Preserve|Integer|Long|LongLong|Double|Single|Currency|String|Boolean|Date|Byte|Object|Variant|Attribute|Declare|PtrSafe|Lib|Alias|Implements|Me|Stop|Global';
const RE_VBA = new RegExp(`("(?:[^"\\n]|"")*")|('.*$|\\bRem\\b.*$)|\\b(${KHOA_VBA})\\b`, 'gim');

const $ = (id) => document.getElementById(id);
const st = { repo: null, branch: 'main', files: [], group: 'all', q: '', token: '', canWrite: false, queue: [], viewing: null };

function chiTiet(html) { return html.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }
function duoi(ten) { const i = ten.lastIndexOf('.'); return i > 0 ? ten.slice(i + 1).toLowerCase() : ''; }
function maPath(p) { return p.split('/').map(encodeURIComponent).join('/'); }
function dinhCo(b) {
  if (b < 1024) return `${b} B`;
  if (b < 1048576) return `${(b / 1024).toFixed(1).replace('.', ',')} KB`;
  return `${(b / 1048576).toFixed(1).replace('.', ',')} MB`;
}
function boDau(s) { return s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase(); }
function setStatus(t) { $('status').textContent = t; }

function nhanRepo() {
  /* Chỉ cho đổi repo qua link khi chạy thử trên máy, để link lạ không thể mở kho khác trên trang thật. */
  const q = location.hostname === 'localhost' ? new URLSearchParams(location.search).get('repo') : '';
  if (q && /^[\w.-]+\/[\w.-]+$/.test(q)) { const [owner, repo] = q.split('/'); return { owner, repo }; }
  if (CAU_HINH.owner && CAU_HINH.repo) return CAU_HINH;
  if (location.hostname.endsWith('.github.io')) {
    const owner = location.hostname.split('.')[0];
    const dau = location.pathname.split('/').filter(Boolean)[0];
    const repo = dau && !dau.includes('.') ? dau : `${owner}.github.io`;
    return { owner, repo };
  }
  return null;
}

async function api(duongDan, opt = {}) {
  const headers = { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28', ...(opt.headers || {}) };
  const token = opt.token ?? st.token;
  if (token) headers.Authorization = `Bearer ${token}`;
  const res = await fetch(`https://api.github.com${duongDan}`, { ...opt, headers });
  if (res.ok) return res.status === 204 ? null : res.json();
  const body = await res.json().catch(() => ({}));
  const err = new Error(body.message || `HTTP ${res.status}`);
  err.status = res.status;
  err.rateLimited = res.headers.get('x-ratelimit-remaining') === '0' || /rate limit/i.test(body.message || '');
  throw err;
}

function thongBaoLoi(e) {
  if (e.rateLimited) return 'GitHub đang giới hạn lượt xem từ mạng này. Chờ khoảng 1 giờ hoặc kết nối token rồi tải lại.';
  if (e.status === 401) return 'Token không hợp lệ hoặc đã hết hạn. Tạo token mới rồi kết nối lại.';
  if (e.status === 403) return 'Token chưa có quyền ghi. Bật quyền Contents: Read and write cho repo này.';
  if (e.status === 404) return 'Không tìm thấy repo. Kiểm tra repo còn công khai và đúng tên.';
  if (e.status === 409) return 'Repo vừa thay đổi. Tải lại trang rồi thử lại.';
  if (e.status === 422) return 'GitHub từ chối file này. Kiểm tra tên file và dung lượng.';
  if (e instanceof TypeError) return 'Không kết nối được GitHub. Kiểm tra mạng rồi thử lại.';
  return `GitHub báo lỗi: ${e.message}`;
}

async function taiDanhSach() {
  setStatus('Đang tải danh sách');
  $('state').hidden = true;
  const { owner, repo } = st.repo;
  const info = await api(`/repos/${owner}/${repo}`);
  st.branch = info.default_branch;
  st.canWrite = Boolean(st.token && info.permissions && info.permissions.push);
  const tree = await api(`/repos/${owner}/${repo}/git/trees/${encodeURIComponent(st.branch)}?recursive=1`);
  st.files = tree.tree
    .filter((n) => n.type === 'blob' && /^(tai-lieu|vba)\//.test(n.path) && !n.path.endsWith('/.gitkeep'))
    .map((n) => ({ path: n.path, sha: n.sha, size: n.size || 0, name: n.path.slice(n.path.indexOf('/') + 1), group: n.path.split('/')[0] }))
    .sort((a, b) => a.name.localeCompare(b.name, 'vi'));
  capNhatNutToken();
  ve();
  setStatus(tree.truncated ? 'Repo quá lớn nên danh sách có thể chưa đủ' : `Sẵn sàng, ${st.files.length} file trong kho`);
}

function hienState(tieuDe, noiDung, nut) {
  const box = $('state');
  box.replaceChildren();
  const h = document.createElement('h2'); h.textContent = tieuDe;
  const p = document.createElement('p'); p.textContent = noiDung;
  box.append(h, p);
  if (nut) {
    const b = document.createElement('button');
    b.className = 'btn btn-primary'; b.type = 'button'; b.textContent = nut.text; b.onclick = nut.fn;
    box.append(b);
  }
  box.hidden = false;
}

function ve() {
  const q = boDau(st.q.trim());
  const list = st.files.filter((f) => (st.group === 'all' || f.group === st.group) && (!q || boDau(f.name).includes(q)));
  const tb = $('rows');
  tb.replaceChildren();
  list.forEach((f, i) => tb.append(taoDong(f, i + 1)));
  $('countBox').textContent = `${list.length} file`;
  if (list.length) { $('state').hidden = true; return; }
  if (!st.files.length) hienState('Kho chưa có file', 'Bấm Thêm file hoặc kéo file vào đây để upload.', { text: 'Thêm file', fn: batDauChonFile });
  else if (q) hienState('Không có file khớp', 'Thử từ khóa ngắn hơn hoặc chọn tab Tất cả.');
  else hienState(`Tab ${NHOM[st.group]} chưa có file`, 'Bấm Thêm file để upload vào nhóm này.', { text: 'Thêm file', fn: batDauChonFile });
}

function taoDong(f, so) {
  const tr = document.createElement('tr');
  const td = (cls) => { const c = document.createElement('td'); c.className = cls; tr.append(c); return c; };
  td('rownum').textContent = so;
  const nut = document.createElement('button');
  nut.type = 'button'; nut.textContent = f.name; nut.title = f.name;
  const xemDuoc = Boolean(kieuXem(f.name));
  nut.onclick = () => (xemDuoc ? xemFile(f) : taiFile(f));
  td('name').append(nut);
  td('type').textContent = `${NHOM[f.group]}${duoi(f.name) ? ` .${duoi(f.name)}` : ''}`;
  td('size').textContent = dinhCo(f.size);
  const a = td('acts');
  const link = (text, fn, cls = '') => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = `link ${cls}`; b.textContent = text; b.onclick = fn;
    b.setAttribute('aria-label', `${text} ${f.name}`);
    a.append(b);
  };
  if (xemDuoc) link('Xem', () => xemFile(f));
  link('Tải', () => taiFile(f));
  if (st.canWrite) link('Xóa', () => xoaFile(f), 'link-danger');
  return tr;
}

function urlRaw(f) {
  const { owner, repo } = st.repo;
  return `https://raw.githubusercontent.com/${owner}/${repo}/${encodeURIComponent(st.branch)}/${maPath(f.path)}`;
}

async function layRaw(f) {
  const res = await fetch(urlRaw(f));
  if (!res.ok) { const e = new Error(`HTTP ${res.status}`); e.status = res.status; throw e; }
  return res.blob();
}

async function taiFile(f) {
  setStatus(`Đang tải ${f.name}`);
  try {
    const url = URL.createObjectURL(await layRaw(f));
    const a = Object.assign(document.createElement('a'), { href: url, download: f.name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    setStatus(`Đã tải ${f.name}`);
  } catch (e) { setStatus(thongBaoLoi(e)); }
}

function thuGiaiMa(buf, enc) {
  try { return new TextDecoder(enc, { fatal: true }).decode(buf); } catch { return null; }
}

/* VBA lưu bằng mã của Windows: máy Nhật dùng Shift-JIS, máy Việt dùng Windows-1258. */
function doanMa(buf) {
  const b = new Uint8Array(buf);
  if (b[0] === 0xff && b[1] === 0xfe) return 'utf-16le';
  if (b[0] === 0xfe && b[1] === 0xff) return 'utf-16be';
  if (thuGiaiMa(buf, 'utf-8') !== null) return 'utf-8';
  if (thuGiaiMa(buf, 'shift_jis') === null) return 'windows-1258';
  /* Byte đầu 0x81-0x9F là kana và kanji thường dùng; chữ Việt gần như chỉ nằm ở 0xA1-0xFF. */
  let nhat = 0; let viet = 0;
  for (let i = 0; i < b.length; i += 1) {
    const c = b[i];
    if (c < 0x80) continue;
    if (c >= 0x81 && c <= 0x9f) { nhat += 1; i += 1; }
    else if (c >= 0xe0) { viet += 1; i += 1; }
    else viet += 1;
  }
  return nhat > viet ? 'shift_jis' : 'windows-1258';
}

function hienChu() {
  const f = st.viewing;
  if (!f || !st.viewingBuf) return;
  const chon = $('encSelect').value;
  const enc = chon === 'auto' ? doanMa(st.viewingBuf) : chon;
  const tho = new TextDecoder(enc).decode(st.viewingBuf);
  /* Windows-1258 tách dấu thanh thành ký tự riêng, gộp lại để sao chép đúng. */
  const text = enc === 'windows-1258' ? tho.normalize('NFC') : tho;
  st.viewingText = text;
  $('codeBox').lang = enc === 'windows-1258' ? 'vi' : /[\u3040-\u30ff\u4e00-\u9fff]/.test(text) ? 'ja' : 'vi';
  if (DUOI_CODE.includes(duoi(f.name))) $('codeBox').innerHTML = toMau(text); else $('codeBox').textContent = text;
}

function toMau(text) {
  let out = ''; let cuoi = 0;
  for (const m of text.matchAll(RE_VBA)) {
    out += chiTiet(text.slice(cuoi, m.index));
    const cls = m[1] ? 's' : m[2] ? 'c' : 'k';
    out += `<span class="${cls}">${chiTiet(m[0])}</span>`;
    cuoi = m.index + m[0].length;
  }
  return out + chiTiet(text.slice(cuoi));
}

function donViewer() {
  st.viewing = null;
  st.viewingText = '';
  st.viewingBuf = null;
  $('frameBox').removeAttribute('src');
  $('imgBox').removeAttribute('src');
  if (st.blobUrl) { URL.revokeObjectURL(st.blobUrl); st.blobUrl = ''; }
}

function hienKhung(id) {
  ['codeWrap', 'frameBox', 'imgWrap'].forEach((k) => { $(k).hidden = k !== id; });
}

async function xemFile(f) {
  donViewer();
  st.viewing = f;
  const kieu = kieuXem(f.name);
  $('viewerTitle').textContent = f.name;
  $('codeBox').textContent = 'Đang tải';
  hienKhung('codeWrap');
  $('btnCopy').hidden = kieu !== 'chu';
  $('encSelect').hidden = kieu !== 'chu';
  $('encSelect').value = 'auto';
  $('btnCopy').textContent = 'Sao chép';
  $('viewer').showModal();
  try {
    if (kieu === 'office') {
      /* Microsoft đọc file công khai từ GitHub rồi hiển thị trong khung này. */
      $('frameBox').src = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(urlRaw(f))}`;
      hienKhung('frameBox');
      return;
    }
    const blob = await layRaw(f);
    if (st.viewing !== f) return;
    if (kieu === 'chu') {
      const buf = await blob.arrayBuffer();
      if (st.viewing !== f) return;
      st.viewingBuf = buf;
      hienChu();
      return;
    }
    const d = duoi(f.name);
    const type = kieu === 'pdf' ? 'application/pdf' : MIME_ANH[d] || `image/${d}`;
    st.blobUrl = URL.createObjectURL(new Blob([blob], { type }));
    if (kieu === 'pdf') { $('frameBox').src = st.blobUrl; hienKhung('frameBox'); }
    else { $('imgBox').src = st.blobUrl; $('imgBox').alt = f.name; hienKhung('imgWrap'); }
  } catch (e) { $('codeBox').textContent = thongBaoLoi(e); hienKhung('codeWrap'); }
}

function capNhatNutToken() {
  $('btnToken').textContent = st.canWrite ? 'Đã kết nối' : 'Kết nối để upload';
  $('btnForget').hidden = !st.token;
}

function batDauChonFile() {
  if (!st.canWrite) { moToken(); return; }
  $('filePicker').click();
}

function moToken() {
  $('tokenError').hidden = true;
  $('tokenInput').value = '';
  $('tokenInput').required = !st.token;
  $('tokenDlg').showModal();
}

function luuToken(t, nho) {
  st.token = t;
  try { if (nho) localStorage.setItem(TOKEN_KEY, t); else localStorage.removeItem(TOKEN_KEY); } catch { /* trình duyệt chặn lưu trữ */ }
}

function nhomMacDinh(ten) {
  if (DUOI_VBA.includes(duoi(ten))) return 'vba';
  return st.group === 'vba' ? 'vba' : 'tai-lieu';
}

function tenSach(ten) {
  return ten.normalize('NFC').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').replace(/^\.+/, '').trim();
}

function moUpload(files) {
  if (!files.length) return;
  if (!st.canWrite) { hienThongBaoCanToken(); return; }
  st.queue = files.map((file) => ({ file, name: tenSach(file.name), group: nhomMacDinh(file.name), state: 'cho' }));
  $('uploadError').hidden = true;
  $('btnUploadStart').disabled = false;
  $('btnUploadStart').hidden = false;
  $('btnUploadCancel').disabled = false;
  $('btnUploadCancel').textContent = 'Hủy';
  veQueue();
  $('uploadDlg').showModal();
}

function hienThongBaoCanToken() { setStatus('Cần kết nối token trước khi upload'); moToken(); }

function veQueue() {
  const ul = $('queue');
  ul.replaceChildren();
  st.queue.forEach((it) => {
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'q-name'; name.textContent = it.name; name.title = it.name;
    const sel = document.createElement('select');
    sel.setAttribute('aria-label', `Nhóm của ${it.name}`);
    Object.entries(NHOM).forEach(([v, t]) => sel.add(new Option(t, v, false, v === it.group)));
    sel.onchange = () => { it.group = sel.value; capNhatGhiChu(it, note); };
    sel.disabled = it.state !== 'cho';
    const note = document.createElement('span');
    li.append(name, sel, note);
    it.note = note;
    capNhatGhiChu(it, note);
    ul.append(li);
  });
}

function capNhatGhiChu(it, note) {
  const quaLon = it.file.size > GIOI_HAN_MB * 1048576;
  const trung = st.files.some((f) => f.path === `${it.group}/${it.name}`);
  let text = dinhCo(it.file.size); let cls = '';
  if (it.state === 'dang') text = 'Đang upload';
  else if (it.state === 'xong') { text = 'Đã upload'; cls = 'ok'; }
  else if (it.state === 'loi') { text = it.loi; cls = 'bad'; }
  else if (!it.name) { text = 'Tên file không hợp lệ, file này sẽ được bỏ qua'; cls = 'bad'; }
  else if (quaLon) { text = `${text}, quá ${GIOI_HAN_MB} MB nên sẽ được bỏ qua`; cls = 'bad'; }
  else if (trung) { text = `${text}, sẽ thay file cùng tên`; cls = 'bad'; }
  note.className = `q-note ${cls}`;
  note.textContent = text;
}

function docBase64(file) {
  return new Promise((ok, loi) => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(',')[1] || '');
    r.onerror = () => loi(new Error('Không đọc được file'));
    r.readAsDataURL(file);
  });
}

async function chayUpload() {
  const { owner, repo } = st.repo;
  $('btnUploadStart').disabled = true;
  $('btnUploadCancel').disabled = true;
  let xong = 0; let loi = 0;
  const daGap = new Set();
  /* Upload lần lượt vì GitHub báo xung đột khi ghi nhiều file cùng lúc. */
  for (const it of st.queue) {
    const path = `${it.group}/${it.name}`;
    if (!it.name || it.file.size > GIOI_HAN_MB * 1048576 || daGap.has(path)) {
      it.state = 'loi';
      it.loi = daGap.has(path) ? 'Trùng tên với file khác trong lượt này, đã bỏ qua' : 'Đã bỏ qua';
      capNhatGhiChu(it, it.note); loi += 1; continue;
    }
    daGap.add(path);
    it.state = 'dang'; capNhatGhiChu(it, it.note);
    try {
      const cu = st.files.find((f) => f.path === path);
      const body = { message: `${cu ? 'Cập nhật' : 'Thêm'} ${path}`, content: await docBase64(it.file), branch: st.branch };
      if (cu) body.sha = cu.sha;
      await api(`/repos/${owner}/${repo}/contents/${maPath(path)}`, { method: 'PUT', body: JSON.stringify(body) });
      it.state = 'xong'; xong += 1;
    } catch (e) { it.state = 'loi'; it.loi = thongBaoLoi(e); loi += 1; }
    capNhatGhiChu(it, it.note);
  }
  $('btnUploadCancel').disabled = false;
  $('btnUploadCancel').textContent = 'Đóng';
  $('btnUploadStart').hidden = true;
  await taiDanhSach().catch((e) => setStatus(thongBaoLoi(e)));
  setStatus(loi ? `Đã upload ${xong} file, ${loi} file chưa upload được` : `Đã upload ${xong} file`);
}

function xacNhan(ten) {
  return new Promise((ok) => {
    const d = $('confirmDlg');
    $('confirmText').textContent = `${ten} sẽ biến mất khỏi trang này. Bản cũ vẫn còn trong lịch sử commit của repo.`;
    d.returnValue = 'no';
    d.addEventListener('close', () => ok(d.returnValue === 'yes'), { once: true });
    d.showModal();
  });
}

async function xoaFile(f) {
  if (!(await xacNhan(f.name))) return;
  const { owner, repo } = st.repo;
  setStatus(`Đang xóa ${f.name}`);
  try {
    await api(`/repos/${owner}/${repo}/contents/${maPath(f.path)}`, {
      method: 'DELETE',
      body: JSON.stringify({ message: `Xóa ${f.path}`, sha: f.sha, branch: st.branch }),
    });
    await taiDanhSach();
    setStatus(`Đã xóa ${f.name}`);
  } catch (e) { setStatus(thongBaoLoi(e)); }
}

function ganSuKien() {
  $('search').addEventListener('input', (e) => { st.q = e.target.value; ve(); });
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => {
    st.group = t.dataset.group;
    document.querySelectorAll('.tab').forEach((x) => x.setAttribute('aria-pressed', String(x === t)));
    ve();
  }));
  document.querySelectorAll('dialog [data-close]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));
  $('btnPick').addEventListener('click', batDauChonFile);
  $('btnToken').addEventListener('click', moToken);
  $('filePicker').addEventListener('change', (e) => { moUpload([...e.target.files]); e.target.value = ''; });

  $('tokenForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const t = $('tokenInput').value.trim();
    if (!t) { $('tokenDlg').close(); return; }
    $('btnTokenSave').disabled = true;
    $('tokenError').hidden = true;
    try {
      const { owner, repo } = st.repo;
      const info = await api(`/repos/${owner}/${repo}`, { token: t });
      if (!info.permissions || !info.permissions.push) {
        const loi = new Error('no push'); loi.status = 403; throw loi;
      }
      luuToken(t, $('tokenRemember').checked);
      $('tokenDlg').close();
      await taiDanhSach();
      setStatus('Đã kết nối, có thể upload file');
    } catch (err) {
      $('tokenError').textContent = thongBaoLoi(err);
      $('tokenError').hidden = false;
    } finally { $('btnTokenSave').disabled = false; }
  });
  $('btnForget').addEventListener('click', async () => {
    luuToken('', false); st.canWrite = false; $('tokenDlg').close();
    capNhatNutToken(); ve(); setStatus('Đã ngắt kết nối');
  });

  $('uploadForm').addEventListener('submit', (e) => { e.preventDefault(); chayUpload(); });
  $('uploadDlg').addEventListener('cancel', (e) => { if ($('btnUploadCancel').disabled) e.preventDefault(); });

  $('btnCopy').addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(st.viewingText || ''); $('btnCopy').textContent = 'Đã sao chép'; }
    catch { $('btnCopy').textContent = 'Chưa sao chép được'; }
  });
  $('btnViewerDownload').addEventListener('click', () => st.viewing && taiFile(st.viewing));
  $('viewer').addEventListener('close', donViewer);
  $('encSelect').addEventListener('change', hienChu);

  const sheet = $('sheet');
  let dem = 0;
  const coFile = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  sheet.addEventListener('dragenter', (e) => { if (coFile(e)) { e.preventDefault(); dem += 1; sheet.classList.add('dragging'); } });
  sheet.addEventListener('dragover', (e) => { if (coFile(e)) e.preventDefault(); });
  sheet.addEventListener('dragleave', () => { dem = Math.max(0, dem - 1); if (!dem) sheet.classList.remove('dragging'); });
  sheet.addEventListener('drop', (e) => {
    e.preventDefault(); dem = 0; sheet.classList.remove('dragging');
    moUpload([...e.dataTransfer.files]);
  });
}

async function khoiDong() {
  ganSuKien();
  try { st.token = localStorage.getItem(TOKEN_KEY) || ''; } catch { /* không có localStorage */ }
  $('tokenRemember').checked = Boolean(st.token);
  st.repo = nhanRepo();
  if (!st.repo) {
    setStatus('Chưa nhận ra repo');
    hienState('Chưa nhận ra repo', 'Mở trang qua địa chỉ github.io, hoặc điền owner và repo trong CAU_HINH ở đầu file app.js.');
    return;
  }
  try { await taiDanhSach(); }
  catch (e) {
    /* Token đã lưu hết hạn không được chặn việc xem file. */
    if (e.status === 401 && st.token) {
      luuToken('', false);
      try { await taiDanhSach(); setStatus('Token đã hết hạn, hãy kết nối lại để upload'); return; } catch { /* báo lỗi bên dưới */ }
    }
    setStatus('Không tải được danh sách');
    hienState('Không tải được danh sách', thongBaoLoi(e), { text: 'Thử lại', fn: () => taiDanhSach().catch((x) => hienState('Không tải được danh sách', thongBaoLoi(x))) });
  }
}

khoiDong();
