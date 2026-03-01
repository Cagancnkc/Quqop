// --- State ---
let allData = [];
let headers = [];
let filteredData = [];
let currentPage = 1;
const PAGE_SIZE = 20;

// --- Upload ---
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');

uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

function handleFile(file) {
  const allowedTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'text/csv',
  ];
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('Lütfen .xlsx, .xls veya .csv dosyası seçin.');
    return;
  }

  fileName.textContent = file.name;
  fileInfo.style.display = 'flex';

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (json.length === 0) {
      alert('Dosya boş veya okunamadı.');
      return;
    }
    headers = Object.keys(json[0]);
    allData = json;
    filteredData = [...allData];
    currentPage = 1;
    renderTable();
    updateStats();
    addMessage('assistant', `"${file.name}" dosyası yüklendi. ${allData.length} satır, ${headers.length} sütun bulundu.`);
  };
  reader.readAsArrayBuffer(file);
}

function clearFile() {
  allData = [];
  headers = [];
  filteredData = [];
  fileInfo.style.display = 'none';
  fileInput.value = '';
  fileName.textContent = '';
  document.getElementById('tableWrapper').innerHTML = `
    <div class="empty-state">
      <span class="empty-icon">📋</span>
      <p>Henüz dosya yüklenmedi.</p>
    </div>`;
  document.getElementById('pagination').innerHTML = '';
  updateStats();
}

// --- Table ---
function renderTable() {
  const wrapper = document.getElementById('tableWrapper');
  if (filteredData.length === 0) {
    wrapper.innerHTML = `<div class="empty-state"><span class="empty-icon">🔍</span><p>Sonuç bulunamadı.</p></div>`;
    document.getElementById('pagination').innerHTML = '';
    return;
  }

  const start = (currentPage - 1) * PAGE_SIZE;
  const pageData = filteredData.slice(start, start + PAGE_SIZE);

  const table = document.createElement('table');
  table.className = 'data-table';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  pageData.forEach(row => {
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] ?? '';
      td.title = row[h] ?? '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  wrapper.innerHTML = '';
  wrapper.appendChild(table);

  renderPagination();
}

function renderPagination() {
  const total = Math.ceil(filteredData.length / PAGE_SIZE);
  const pagination = document.getElementById('pagination');
  pagination.innerHTML = '';
  if (total <= 1) return;

  for (let i = 1; i <= total; i++) {
    const btn = document.createElement('button');
    btn.className = 'page-btn' + (i === currentPage ? ' active' : '');
    btn.textContent = i;
    btn.onclick = () => { currentPage = i; renderTable(); };
    pagination.appendChild(btn);
  }
}

function filterTable() {
  const q = document.getElementById('searchInput').value.toLowerCase();
  if (!q) {
    filteredData = [...allData];
  } else {
    filteredData = allData.filter(row =>
      headers.some(h => String(row[h]).toLowerCase().includes(q))
    );
  }
  currentPage = 1;
  renderTable();
}

function exportCSV() {
  if (filteredData.length === 0) { alert('Dışa aktarılacak veri yok.'); return; }
  const ws = XLSX.utils.json_to_sheet(filteredData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  XLSX.writeFile(wb, 'export.csv');
}

// --- Stats ---
function updateStats() {
  const rows = allData.length;
  const cols = headers.length;
  let filled = 0, empty = 0;
  allData.forEach(row => {
    headers.forEach(h => {
      const v = row[h];
      (v !== '' && v !== null && v !== undefined) ? filled++ : empty++;
    });
  });

  document.getElementById('statRows').textContent = rows || '—';
  document.getElementById('statCols').textContent = cols || '—';
  document.getElementById('statFilled').textContent = rows ? filled : '—';
  document.getElementById('statEmpty').textContent = rows ? empty : '—';
}

// --- AI Chat ---
function addMessage(role, text) {
  const messages = document.getElementById('chatMessages');
  const bubble = document.createElement('div');
  bubble.className = `chat-bubble ${role}`;
  bubble.textContent = text;
  messages.appendChild(bubble);
  messages.scrollTop = messages.scrollHeight;
}

function sendMessage() {
  const input = document.getElementById('chatInput');
  const text = input.value.trim();
  if (!text) return;
  addMessage('user', text);
  input.value = '';

  if (allData.length === 0) {
    addMessage('assistant', 'Henüz bir dosya yüklenmedi. Lütfen önce Excel veya CSV dosyanızı yükleyin.');
    return;
  }

  setTimeout(() => {
    const reply = generateReply(text.toLowerCase());
    addMessage('assistant', reply);
  }, 400);
}

function generateReply(q) {
  if (q.includes('kaç satır') || q.includes('satır sayısı') || q.includes('satır var')) {
    return `Tabloda toplam ${allData.length} satır bulunuyor.`;
  }
  if (q.includes('kaç sütun') || q.includes('sütun sayısı') || q.includes('sütun var')) {
    return `Tabloda toplam ${headers.length} sütun bulunuyor: ${headers.join(', ')}.`;
  }
  if (q.includes('sütun') && q.includes('liste')) {
    return `Sütunlar: ${headers.join(', ')}.`;
  }
  if (q.includes('boş') || q.includes('eksik')) {
    let empty = 0;
    allData.forEach(row => headers.forEach(h => { if (row[h] === '' || row[h] == null) empty++; }));
    return `Tabloda ${empty} adet boş/eksik hücre bulunuyor.`;
  }

  // Column sum / average
  for (const h of headers) {
    if (q.includes(h.toLowerCase())) {
      const nums = allData.map(r => parseFloat(r[h])).filter(v => !isNaN(v));
      if (nums.length > 0) {
        const sum = nums.reduce((a, b) => a + b, 0);
        const avg = (sum / nums.length).toFixed(2);
        const max = Math.max(...nums);
        const min = Math.min(...nums);
        if (q.includes('toplam') || q.includes('sum')) return `"${h}" sütununun toplamı: ${sum.toLocaleString('tr-TR')}.`;
        if (q.includes('ortalama') || q.includes('ort') || q.includes('avg')) return `"${h}" sütununun ortalaması: ${avg}.`;
        if (q.includes('max') || q.includes('en büyük')) return `"${h}" sütununun en büyük değeri: ${max}.`;
        if (q.includes('min') || q.includes('en küçük')) return `"${h}" sütununun en küçük değeri: ${min}.`;
        return `"${h}" sütunu: toplam ${sum.toLocaleString('tr-TR')}, ortalama ${avg}, min ${min}, max ${max}.`;
      }
    }
  }

  return `Sorunuzu anlayamadım. "kaç satır", "kaç sütun", sütun adı + "toplam/ortalama/max/min" gibi sorular sorabilirsiniz.`;
}
