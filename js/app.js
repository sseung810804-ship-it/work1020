/**
 * 국방/공공 예산 인터랙티브 대시보드 메인 애플리케이션
 */

// 전역 상태 관리
const AppState = {
  rawData: [],          // 원본 데이터 배열
  filteredData: [],     // 필터링된 데이터
  activeTab: 'overview',// overview, domain, new-projects, table
  selectedDomain: 'all',// 분야별 탭에서 선택된 분야
  tableFilterDomain: '',// 테이블 분야 필터
  tableFilterNew: false,// 테이블 신규사업 전용 필터
  tableSearch: '',      // 테이블 검색어
  tableSortCol: 'no',   // 정렬 컬럼
  tableSortAsc: true,   // 오름차순 여부
  currentPage: 1,
  pageSize: 15,
  charts: {}            // Chart.js 인스턴스 저장소
};

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  initTabs();
  initUpload();
  initTableControls();
  
  // 기본 샘플 데이터 로드
  if (typeof DEFAULT_BUDGET_DATA !== 'undefined' && Array.isArray(DEFAULT_BUDGET_DATA)) {
    loadData(DEFAULT_BUDGET_DATA, '2027년 국방예산 정부안 (기본 데이터)');
  }
}

// ==========================================
// 1. 탭 네비게이션 제어
// ==========================================
function initTabs() {
  const tabButtons = document.querySelectorAll('.tab-btn');
  tabButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const tabName = btn.getAttribute('data-tab');
      switchTab(tabName);
    });
  });
}

function switchTab(tabName) {
  AppState.activeTab = tabName;
  
  // 버튼 활성화
  document.querySelectorAll('.tab-btn').forEach(btn => {
    if (btn.getAttribute('data-tab') === tabName) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  // 탭 컨텐츠 표시
  document.querySelectorAll('.tab-content').forEach(content => {
    if (content.id === `tab-${tabName}`) {
      content.classList.remove('hidden');
    } else {
      content.classList.add('hidden');
    }
  });

  // 차트 리사이즈/재렌더링
  setTimeout(() => {
    Object.values(AppState.charts).forEach(chart => {
      if (chart) chart.resize();
    });
  }, 100);

  if (tabName === 'domain') {
    renderDomainView();
  } else if (tabName === 'new-projects') {
    renderNewProjectsView();
  }
}

// ==========================================
// 2. 파일 업로드 및 파싱 (Excel / CSV)
// ==========================================
function initUpload() {
  const dropzone = document.getElementById('uploadDropzone');
  const fileInput = document.getElementById('fileInput');
  const resetBtn = document.getElementById('btnResetSample');
  const downloadCsvBtn = document.getElementById('btnDownloadCsv');
  const downloadTemplateBtn = document.getElementById('btnDownloadTemplate');

  if (dropzone) {
    dropzone.addEventListener('click', () => fileInput.click());

    // 드래그 앤 드롭
    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
      }
    });
  }

  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFile(e.target.files[0]);
      }
    });
  }

  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      if (confirm('기본 2027년 국방예산 샘플 데이터로 복원하시겠습니까?')) {
        loadData(DEFAULT_BUDGET_DATA, '2027년 국방예산 정부안 (기본 데이터)');
      }
    });
  }

  if (downloadCsvBtn) {
    downloadCsvBtn.addEventListener('click', exportToCsv);
  }

  if (downloadTemplateBtn) {
    downloadTemplateBtn.addEventListener('click', exportTemplate);
  }
}

function handleFile(file) {
  const fileName = file.name;
  const ext = fileName.split('.').pop().toLowerCase();

  const reader = new FileReader();

  if (ext === 'csv') {
    reader.onload = function(e) {
      try {
        const text = e.target.result;
        parseCsvData(text, fileName);
      } catch (err) {
        alert('CSV 파일 파싱 중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsText(file, 'UTF-8');
  } else if (ext === 'xlsx' || ext === 'xls') {
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        // 첫 번째 시트 사용
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        parseExcelRows(json, fileName);
      } catch (err) {
        alert('엑셀 파일 파싱 중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert('지원되지 않는 파일 형식입니다. .xlsx, .xls, .csv 파일을 선택해주세요.');
  }
}

// 엑셀 2차원 배열 파싱
function parseExcelRows(rows, fileName) {
  if (!rows || rows.length < 2) {
    alert('데이터가 충분하지 않습니다.');
    return;
  }

  // 헤더 행 탐색 (이미지 속 컬럼 키워드 탐색)
  let headerIndex = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const rowStr = rows[i].map(c => String(c).trim()).join(' ');
    if (rowStr.includes('분야') || rowStr.includes('사업') || rowStr.includes('2026') || rowStr.includes('2027') || rowStr.includes('증감')) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) {
    headerIndex = 0; // fallback
  }

  const rawHeaders = rows[headerIndex].map(c => String(c).trim());
  const headerMap = detectColumnMapping(rawHeaders);

  const parsedData = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    // 모든 셀이 비어있으면 스킵
    if (row.every(cell => !cell || String(cell).trim() === '')) continue;

    const item = extractRowItem(row, headerMap, parsedData.length + 1);
    // 사업명이나 분야가 유효한 경우만 추가
    if (item['세부 사업 / 주요 지표'] || item['분야 (중점영역)']) {
      parsedData.push(item);
    }
  }

  if (parsedData.length === 0) {
    alert('유효한 데이터 행을 찾을 수 없습니다. 파일 서식을 확인해주세요.');
    return;
  }

  loadData(parsedData, fileName);
}

// CSV 텍스트 파싱
function parseCsvData(csvText, fileName) {
  // 간단하고 안전한 CSV 분할
  const lines = parseCSVToRows(csvText);
  parseExcelRows(lines, fileName);
}

// CSV 파서 (따옴표 지원)
function parseCSVToRows(text) {
  const p = [];
  let row = [''];
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    const next = text[i + 1];

    if (c === '"') {
      if (inQuotes && next === '"') {
        row[row.length - 1] += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      row.push('');
    } else if ((c === '\r' || c === '\n') && !inQuotes) {
      if (c === '\r' && next === '\n') {
        i++;
      }
      p.push(row);
      row = [''];
    } else {
      row[row.length - 1] += c;
    }
  }
  if (row.length > 1 || row[0] !== '') {
    p.push(row);
  }
  return p;
}

// 컬럼 인덱스 매핑 (헤더 유연 감지)
function detectColumnMapping(headers) {
  const map = {
    no: -1,
    domain: -1,
    project: -1,
    unit: -1,
    y2026: -1,
    y2027: -1,
    diff: -1,
    rate: -1,
    evidence: -1
  };

  headers.forEach((h, idx) => {
    const clean = h.replace(/\s+/g, '').toLowerCase();

    if (map.no === -1 && (clean.includes('연번') || clean.includes('순번') || clean === 'no')) {
      map.no = idx;
    } else if (map.domain === -1 && (clean.includes('분야') || clean.includes('중점영역') || clean.includes('영역'))) {
      map.domain = idx;
    } else if (map.project === -1 && (clean.includes('세부사업') || clean.includes('사업명') || clean.includes('주요지표') || clean.includes('사업'))) {
      map.project = idx;
    } else if (map.unit === -1 && (clean.includes('구분') || clean.includes('단위'))) {
      map.unit = idx;
    } else if (map.y2026 === -1 && (clean.includes('2026') || clean.includes('전년') || clean.includes('(a)'))) {
      map.y2026 = idx;
    } else if (map.y2027 === -1 && (clean.includes('2027') || clean.includes('정부안') || clean.includes('당해') || clean.includes('(b)'))) {
      map.y2027 = idx;
    } else if (map.diff === -1 && (clean.includes('증감') && !clean.includes('증감율') && !clean.includes('증감률')) || clean.includes('(b-a)')) {
      map.diff = idx;
    } else if (map.rate === -1 && (clean.includes('증감율') || clean.includes('증감률') || clean.includes('%'))) {
      map.rate = idx;
    } else if (map.evidence === -1 && (clean.includes('핵심') || clean.includes('보도자료') || clean.includes('원문') || clean.includes('근거') || clean.includes('비고') || clean.includes('내용'))) {
      map.evidence = idx;
    }
  });

  // 누락된 기본 위치 fallback (이미지 기준 순서)
  if (map.no === -1 && headers.length > 0) map.no = 0;
  if (map.domain === -1 && headers.length > 1) map.domain = 1;
  if (map.project === -1 && headers.length > 2) map.project = 2;
  if (map.unit === -1 && headers.length > 3) map.unit = 3;
  if (map.y2026 === -1 && headers.length > 4) map.y2026 = 4;
  if (map.y2027 === -1 && headers.length > 5) map.y2027 = 5;
  if (map.diff === -1 && headers.length > 6) map.diff = 6;
  if (map.rate === -1 && headers.length > 7) map.rate = 7;
  if (map.evidence === -1 && headers.length > 8) map.evidence = 8;

  return map;
}

// 데이터 행에서 모델 객체 추출
function extractRowItem(row, map, fallbackNo) {
  const getVal = (idx) => (idx >= 0 && idx < row.length ? String(row[idx]).trim() : '');

  const noVal = parseInt(getVal(map.no)) || fallbackNo;
  const domain = getVal(map.domain) || '기타 / 미분류';
  const project = getVal(map.project) || '(사업명 미지정)';
  const unit = getVal(map.unit) || '억원';

  const cleanNum = (str) => {
    if (!str) return 0;
    // 쉼표, 공백, 따옴표, '억', '원' 제거, 음수 부호 처리 (△ -> -)
    const normalized = str.replace(/,/g, '').replace(/\s+/g, '').replace(/억원/g, '').replace(/원/g, '').replace(/△/g, '-');
    const num = parseFloat(normalized);
    return isNaN(num) ? 0 : num;
  };

  const y2026 = cleanNum(getVal(map.y2026));
  const y2027 = cleanNum(getVal(map.y2027));

  let diff = cleanNum(getVal(map.diff));
  if (diff === 0 && (y2027 !== 0 || y2026 !== 0)) {
    diff = y2027 - y2026;
  }

  let rateStr = getVal(map.rate);
  let rate = 0;
  if (rateStr) {
    if (rateStr.includes('신규') || rateStr.includes('신설')) {
      rate = 100.0;
    } else {
      rate = cleanNum(rateStr);
    }
  } else if (y2026 > 0) {
    rate = parseFloat(((diff / y2026) * 100).toFixed(1));
  } else if (y2027 > 0 && y2026 === 0) {
    rate = 100.0;
  }

  const evidence = getVal(map.evidence) || '';

  // 신규사업 판별
  const isNew = checkIfNewProject(unit, project, y2026, y2027, evidence, rateStr);

  return {
    "연번": noVal,
    "분야 (중점영역)": domain,
    "세부 사업 / 주요 지표": project,
    "구분/단위": unit,
    "2026년 (A)": y2026,
    "2027년 정부안 (B)": y2027,
    "증감 (B-A)": diff,
    "증감율 (%)": rate,
    "핵심 내용 및 보도자료 박스 원문 근거": evidence,
    "_isNew": isNew
  };
}

// 신규 사업 여부 판단
function checkIfNewProject(unit, project, y2026, y2027, evidence, rateStr) {
  const checkText = `${unit} ${project} ${evidence} ${rateStr}`.toLowerCase();
  if (checkText.includes('신규') || checkText.includes('신설') || checkText.includes('새로운') || checkText.includes('착수')) {
    return true;
  }
  if (y2026 === 0 && y2027 > 0) {
    return true;
  }
  return false;
}

// ==========================================
// 3. 데이터 로드 및 전체 화면 갱신
// ==========================================
function loadData(dataArray, sourceName) {
  AppState.rawData = dataArray.map(item => {
    // _isNew 프로퍼티 보장
    if (typeof item._isNew === 'undefined') {
      item._isNew = checkIfNewProject(
        item["구분/단위"] || '',
        item["세부 사업 / 주요 지표"] || '',
        item["2026년 (A)"] || 0,
        item["2027년 정부안 (B)"] || 0,
        item["핵심 내용 및 보도자료 박스 원문 근거"] || '',
        String(item["증감율 (%)"] || '')
      );
    }
    return item;
  });

  // 파일 정보 표시 업데이트
  const fileBadge = document.getElementById('loadedFileInfo');
  if (fileBadge) {
    fileBadge.textContent = `${sourceName} (총 ${AppState.rawData.length}개 사업)`;
  }

  // 1. KPI 카드 갱신
  renderKPICards();

  // 2. 종합 탭 차트 및 통계 갱신
  renderOverviewCharts();

  // 3. 분야별 탭 갱신
  renderDomainFilterOptions();
  renderDomainView();

  // 4. 신규사업 탭 갱신
  renderNewProjectsView();

  // 5. 테이블 필터 및 렌더링
  applyTableFilters();
}

// KPI 카드 계산 및 렌더링
function renderKPICards() {
  const data = AppState.rawData;

  let total2026 = 0;
  let total2027 = 0;
  let totalDiff = 0;
  let newProjectCount = 0;
  let newProjectBudget = 0;

  data.forEach(item => {
    const y26 = Number(item["2026년 (A)"]) || 0;
    const y27 = Number(item["2027년 정부안 (B)"]) || 0;
    const diff = Number(item["증감 (B-A)"]) || (y27 - y26);

    total2026 += y26;
    total2027 += y27;
    totalDiff += diff;

    if (item._isNew) {
      newProjectCount++;
      newProjectBudget += y27;
    }
  });

  const totalRate = total2026 > 0 ? ((totalDiff / total2026) * 100).toFixed(1) : (total2027 > 0 ? 100.0 : 0);

  // 통화/숫자 포맷터 (억 원 및 조 원 단위 환산)
  const formatMoney = (val) => {
    if (Math.abs(val) >= 10000) {
      const jo = (val / 10000).toFixed(2);
      return `${val.toLocaleString()}억 (${jo}조)`;
    }
    return `${val.toLocaleString()}억원`;
  };

  document.getElementById('kpiTotal2026').textContent = formatMoney(total2026);
  document.getElementById('kpiTotal2027').textContent = formatMoney(total2027);

  const diffEl = document.getElementById('kpiTotalDiff');
  const rateEl = document.getElementById('kpiTotalRate');
  const sign = totalDiff >= 0 ? '+' : '';

  diffEl.textContent = `${sign}${formatMoney(totalDiff)}`;
  diffEl.className = totalDiff >= 0 ? 'text-2xl font-bold text-emerald-600' : 'text-2xl font-bold text-rose-600';

  rateEl.textContent = `${sign}${totalRate}%`;
  rateEl.className = totalDiff >= 0 ? 'text-xs font-semibold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800' : 'text-xs font-semibold px-2 py-0.5 rounded-full bg-rose-100 text-rose-800';

  document.getElementById('kpiTotalProjects').textContent = `${data.length}개`;
  document.getElementById('kpiNewProjects').textContent = `${newProjectCount}개 (${formatMoney(newProjectBudget)})`;
}

// ==========================================
// 4. 종합 대시보드 뷰 및 차트
// ==========================================
function renderOverviewCharts() {
  const data = AppState.rawData;

  // 분야별 집계
  const domainMap = {};
  data.forEach(item => {
    const domain = item["분야 (중점영역)"] || '기타';
    if (!domainMap[domain]) {
      domainMap[domain] = { y2026: 0, y2027: 0, diff: 0, count: 0, newCount: 0 };
    }
    const y26 = Number(item["2026년 (A)"]) || 0;
    const y27 = Number(item["2027년 정부안 (B)"]) || 0;
    domainMap[domain].y2026 += y26;
    domainMap[domain].y2027 += y27;
    domainMap[domain].diff += (Number(item["증감 (B-A)"]) || (y27 - y26));
    domainMap[domain].count++;
    if (item._isNew) domainMap[domain].newCount++;
  });

  const domainLabels = Object.keys(domainMap);
  const domain2027Values = domainLabels.map(d => domainMap[d].y2027);
  const domain2026Values = domainLabels.map(d => domainMap[d].y2026);

  // 1. 도넛 차트: 분야별 2027년 예산 비중
  const ctxDonut = document.getElementById('chartDomainShare');
  if (ctxDonut) {
    if (AppState.charts.domainShare) AppState.charts.domainShare.destroy();

    const vibrantColors = [
      '#2563eb', '#0d9488', '#8b5cf6', '#f59e0b', '#ec4899',
      '#06b6d4', '#84cc16', '#6366f1', '#14b8a6', '#64748b'
    ];

    AppState.charts.domainShare = new Chart(ctxDonut, {
      type: 'doughnut',
      data: {
        labels: domainLabels,
        datasets: [{
          data: domain2027Values,
          backgroundColor: vibrantColors.slice(0, domainLabels.length),
          borderWidth: 2,
          borderColor: '#ffffff',
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: {
              boxWidth: 12,
              font: { size: 11, family: 'Pretendard' }
            }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const val = context.raw;
                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${val.toLocaleString()}억원 (${pct}%)`;
              }
            }
          }
        },
        cutout: '65%'
      }
    });
  }

  // 2. 바 차트: 분야별 2026 vs 2027 비교
  const ctxCompare = document.getElementById('chartDomainCompare');
  if (ctxCompare) {
    if (AppState.charts.domainCompare) AppState.charts.domainCompare.destroy();

    AppState.charts.domainCompare = new Chart(ctxCompare, {
      type: 'bar',
      data: {
        labels: domainLabels.map(d => d.replace(/^[①-⑳\d\.\s]+/, '')), // 번호 기호 정돈
        datasets: [
          {
            label: '2026년 예산 (A)',
            data: domain2026Values,
            backgroundColor: 'rgba(148, 163, 184, 0.75)',
            borderColor: '#94a3b8',
            borderWidth: 1,
            borderRadius: 4
          },
          {
            label: '2027년 정부안 (B)',
            data: domain2027Values,
            backgroundColor: 'rgba(37, 99, 235, 0.85)',
            borderColor: '#2563eb',
            borderWidth: 1,
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: { boxWidth: 12, font: { size: 12, family: 'Pretendard' } }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                return ` ${context.dataset.label}: ${context.raw.toLocaleString()}억원`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: 11, family: 'Pretendard' } }
          },
          y: {
            beginAtZero: true,
            ticks: {
              callback: function(val) {
                return val >= 10000 ? (val / 10000).toFixed(0) + '조' : val.toLocaleString();
              }
            }
          }
        }
      }
    });
  }

  // 3. 증감액 TOP 5 & 증감율 TOP 5 렌더링
  renderTopRankingCards(data);

  // 4. 분야별 요약 매트릭스 테이블
  renderDomainSummaryTable(domainMap);
}

function renderTopRankingCards(data) {
  // 증감액 TOP 5
  const sortedByDiff = [...data].sort((a, b) => (b["증감 (B-A)"] || 0) - (a["증감 (B-A)"] || 0)).slice(0, 5);
  const topDiffContainer = document.getElementById('topDiffList');
  if (topDiffContainer) {
    topDiffContainer.innerHTML = sortedByDiff.map((item, idx) => `
      <div class="flex items-center justify-between p-3 rounded-lg border border-slate-100 hover:bg-blue-50/50 transition cursor-pointer" onclick="openDetailModal(${item["연번"]})">
        <div class="flex items-center gap-3">
          <span class="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold ${idx === 0 ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-600'}">${idx + 1}</span>
          <div>
            <div class="text-sm font-semibold text-slate-900 flex items-center gap-1.5">
              ${escapeHtml(item["세부 사업 / 주요 지표"])}
              ${item._isNew ? '<span class="badge-new text-[10px] py-0 px-1.5">신규</span>' : ''}
            </div>
            <div class="text-xs text-slate-500">${escapeHtml(item["분야 (중점영역)"])}</div>
          </div>
        </div>
        <div class="text-right">
          <div class="text-sm font-bold text-emerald-600">+${Number(item["증감 (B-A)"]).toLocaleString()}억</div>
          <div class="text-xs text-slate-400">27년: ${Number(item["2027년 정부안 (B)"]).toLocaleString()}억</div>
        </div>
      </div>
    `).join('');
  }

  // 증감율 TOP 5 (신규 사업 제외 또는 포함)
  const sortedByRate = [...data]
    .filter(i => (i["2026년 (A)"] || 0) > 0) // 기존 사업 중 증가율 큰 순
    .sort((a, b) => (b["증감율 (%)"] || 0) - (a["증감율 (%)"] || 0))
    .slice(0, 5);

  const topRateContainer = document.getElementById('topRateList');
  if (topRateContainer) {
    topRateContainer.innerHTML = sortedByRate.map((item, idx) => `
      <div class="flex items-center justify-between p-3 rounded-lg border border-slate-100 hover:bg-emerald-50/50 transition cursor-pointer" onclick="openDetailModal(${item["연번"]})">
        <div class="flex items-center gap-3">
          <span class="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold ${idx === 0 ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}">${idx + 1}</span>
          <div>
            <div class="text-sm font-semibold text-slate-900">${escapeHtml(item["세부 사업 / 주요 지표"])}</div>
            <div class="text-xs text-slate-500">${escapeHtml(item["분야 (중점영역)"])}</div>
          </div>
        </div>
        <div class="text-right">
          <div class="text-sm font-bold text-blue-600">+${Number(item["증감율 (%)"]).toFixed(1)}%</div>
          <div class="text-xs text-slate-400">+${Number(item["증감 (B-A)"]).toLocaleString()}억</div>
        </div>
      </div>
    `).join('');
  }
}

function renderDomainSummaryTable(domainMap) {
  const tbody = document.getElementById('domainSummaryTableBody');
  if (!tbody) return;

  const total2027 = Object.values(domainMap).reduce((acc, cur) => acc + cur.y2027, 0);

  const html = Object.keys(domainMap).map((domain, idx) => {
    const item = domainMap[domain];
    const share = total2027 > 0 ? ((item.y2027 / total2027) * 100).toFixed(1) : '0.0';
    const rate = item.y2026 > 0 ? (((item.y2027 - item.y2026) / item.y2026) * 100).toFixed(1) : (item.y2027 > 0 ? '100.0' : '0.0');

    return `
      <tr class="hover:bg-slate-50 transition cursor-pointer" onclick="goToDomainDetail('${escapeHtml(domain)}')">
        <td class="px-4 py-3 font-semibold text-slate-900 flex items-center gap-2">
          <span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
          ${escapeHtml(domain)}
        </td>
        <td class="px-4 py-3 text-center">
          <span class="text-xs font-medium px-2 py-0.5 rounded bg-slate-100 text-slate-700">${item.count}개</span>
          ${item.newCount > 0 ? `<span class="badge-new text-[10px] ml-1">신규 ${item.newCount}</span>` : ''}
        </td>
        <td class="px-4 py-3 text-right text-slate-600">${item.y2026.toLocaleString()}</td>
        <td class="px-4 py-3 text-right font-bold text-slate-900">${item.y2027.toLocaleString()}</td>
        <td class="px-4 py-3 text-right font-medium ${item.diff >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
          ${item.diff >= 0 ? '+' : ''}${item.diff.toLocaleString()}
        </td>
        <td class="px-4 py-3 text-right font-bold ${item.diff >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
          ${item.diff >= 0 ? '+' : ''}${rate}%
        </td>
        <td class="px-4 py-3 text-right">
          <div class="flex items-center justify-end gap-2">
            <span class="text-xs text-slate-500">${share}%</span>
            <div class="w-16 bg-slate-100 rounded-full h-1.5 overflow-hidden">
              <div class="bg-blue-600 h-1.5 rounded-full" style="width: ${Math.min(share, 100)}%"></div>
            </div>
          </div>
        </td>
      </tr>
    `;
  }).join('');

  tbody.innerHTML = html;
}

// ==========================================
// 5. 분야별(중점영역별) 심층 분석 뷰
// ==========================================
function renderDomainFilterOptions() {
  const domainSelect = document.getElementById('domainSelect');
  if (!domainSelect) return;

  const domains = [...new Set(AppState.rawData.map(d => d["분야 (중점영역)"] || '기타'))];

  domainSelect.innerHTML = `
    <option value="all">전체 분야 한눈에 보기</option>
    ${domains.map(d => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join('')}
  `;

  domainSelect.onchange = (e) => {
    AppState.selectedDomain = e.target.value;
    renderDomainView();
  };
}

function goToDomainDetail(domainName) {
  AppState.selectedDomain = domainName;
  const domainSelect = document.getElementById('domainSelect');
  if (domainSelect) domainSelect.value = domainName;
  switchTab('domain');
}

function renderDomainView() {
  const selected = AppState.selectedDomain;
  const filtered = selected === 'all' 
    ? AppState.rawData 
    : AppState.rawData.filter(d => (d["분야 (중점영역)"] || '기타') === selected);

  // 분야 요약 메트릭
  let sum26 = 0, sum27 = 0, sumDiff = 0, newCnt = 0;
  filtered.forEach(i => {
    sum26 += Number(i["2026년 (A)"]) || 0;
    sum27 += Number(i["2027년 정부안 (B)"]) || 0;
    sumDiff += (Number(i["증감 (B-A)"]) || (Number(i["2027년 정부안 (B)"]) - Number(i["2026년 (A)"])));
    if (i._isNew) newCnt++;
  });
  const rate = sum26 > 0 ? ((sumDiff / sum26) * 100).toFixed(1) : (sum27 > 0 ? 100.0 : 0);

  document.getElementById('domainViewTitle').textContent = selected === 'all' ? '전체 분야 종합 분석' : selected;
  document.getElementById('domainStatProjects').textContent = `${filtered.length}개 사업 (신규 ${newCnt}개)`;
  document.getElementById('domainStat2027').textContent = `${sum27.toLocaleString()}억원`;
  document.getElementById('domainStatDiff').textContent = `${sumDiff >= 0 ? '+' : ''}${sumDiff.toLocaleString()}억원 (${sumDiff >= 0 ? '+' : ''}${rate}%)`;

  // 수평 막대 차트 (해당 분야 사업별 예산 비교)
  renderDomainDetailChart(filtered);

  // 카드 목록 렌더링
  const cardsContainer = document.getElementById('domainProjectsCards');
  if (cardsContainer) {
    cardsContainer.innerHTML = filtered.map(item => `
      <div class="dashboard-card p-5 hover:border-blue-400 transition cursor-pointer flex flex-col justify-between" onclick="openDetailModal(${item["연번"]})">
        <div>
          <div class="flex items-start justify-between gap-2 mb-2">
            <span class="badge-domain">${escapeHtml(item["분야 (중점영역)"])}</span>
            ${item._isNew ? '<span class="badge-new">신규사업</span>' : `<span class="${item["증감 (B-A)"] >= 0 ? 'badge-increase' : 'badge-decrease'}">${item["증감 (B-A)"] >= 0 ? '+' : ''}${item["증감율 (%)"]}%</span>`}
          </div>
          <h4 class="text-base font-bold text-slate-900 mb-1 hover:text-blue-600 transition">${escapeHtml(item["세부 사업 / 주요 지표"])}</h4>
          <p class="text-xs text-slate-500 line-clamp-2 mb-4 leading-relaxed">${escapeHtml(item["핵심 내용 및 보도자료 박스 원문 근거"]) || '원문 근거 정보 없음'}</p>
        </div>
        <div class="pt-3 border-t border-slate-100 flex items-center justify-between">
          <div>
            <span class="text-[11px] text-slate-400 block">2027년 정부안</span>
            <span class="text-base font-extrabold text-blue-900">${Number(item["2027년 정부안 (B)"]).toLocaleString()}억</span>
          </div>
          <div class="text-right">
            <span class="text-[11px] text-slate-400 block">전년대비 증감</span>
            <span class="text-xs font-bold ${item["증감 (B-A)"] >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
              ${item["증감 (B-A)"] >= 0 ? '+' : ''}${Number(item["증감 (B-A)"]).toLocaleString()}억
            </span>
          </div>
        </div>
      </div>
    `).join('');
  }
}

function renderDomainDetailChart(items) {
  const ctx = document.getElementById('chartDomainDetailBar');
  if (!ctx) return;

  if (AppState.charts.domainDetailBar) AppState.charts.domainDetailBar.destroy();

  // 상위 10개 사업
  const sorted = [...items].sort((a, b) => (b["2027년 정부안 (B)"] || 0) - (a["2027년 정부안 (B)"] || 0)).slice(0, 10);

  AppState.charts.domainDetailBar = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(i => i["세부 사업 / 주요 지표"]),
      datasets: [
        {
          label: '2026년 예산',
          data: sorted.map(i => i["2026년 (A)"] || 0),
          backgroundColor: '#cbd5e1',
          borderRadius: 4
        },
        {
          label: '2027년 정부안',
          data: sorted.map(i => i["2027년 정부안 (B)"] || 0),
          backgroundColor: '#3b82f6',
          borderRadius: 4
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { family: 'Pretendard' } } }
      },
      scales: {
        x: {
          ticks: {
            callback: (v) => v.toLocaleString() + '억'
          }
        }
      }
    }
  });
}

// ==========================================
// 6. 신규사업(New Projects) 집중 분석 뷰
// ==========================================
function renderNewProjectsView() {
  const newProjects = AppState.rawData.filter(i => i._isNew);

  let totalNewBudget = 0;
  newProjects.forEach(i => totalNewBudget += (Number(i["2027년 정부안 (B)"]) || 0));

  const totalAll2027 = AppState.rawData.reduce((acc, cur) => acc + (Number(cur["2027년 정부안 (B)"]) || 0), 0);
  const newShare = totalAll2027 > 0 ? ((totalNewBudget / totalAll2027) * 100).toFixed(1) : 0;

  document.getElementById('newProjCountStat').textContent = `${newProjects.length}개`;
  document.getElementById('newProjBudgetStat').textContent = `${totalNewBudget.toLocaleString()}억원`;
  document.getElementById('newProjShareStat').textContent = `${newShare}%`;

  // 분야별 신규사업 분포 차트
  const domainNewMap = {};
  newProjects.forEach(i => {
    const d = i["분야 (중점영역)"] || '기타';
    domainNewMap[d] = (domainNewMap[d] || 0) + (Number(i["2027년 정부안 (B)"]) || 0);
  });

  const ctxNew = document.getElementById('chartNewProjectsDist');
  if (ctxNew) {
    if (AppState.charts.newProjectsDist) AppState.charts.newProjectsDist.destroy();

    AppState.charts.newProjectsDist = new Chart(ctxNew, {
      type: 'pie',
      data: {
        labels: Object.keys(domainNewMap),
        datasets: [{
          data: Object.values(domainNewMap),
          backgroundColor: ['#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#3b82f6', '#06b6d4']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { boxWidth: 12, font: { family: 'Pretendard' } } }
        }
      }
    });
  }

  // 신규사업 집중 카드 그리드
  const container = document.getElementById('newProjectsGrid');
  if (container) {
    if (newProjects.length === 0) {
      container.innerHTML = `<div class="col-span-full text-center py-12 text-slate-400">신규 사업으로 분류된 항목이 없습니다.</div>`;
      return;
    }

    container.innerHTML = newProjects.map(item => `
      <div class="dashboard-card p-6 border-l-4 border-l-purple-500 hover:shadow-lg transition cursor-pointer" onclick="openDetailModal(${item["연번"]})">
        <div class="flex items-center justify-between mb-3">
          <span class="badge-new">신규 착수/도입</span>
          <span class="text-xs font-semibold px-2.5 py-1 rounded-md bg-purple-50 text-purple-700 border border-purple-100">
            ${escapeHtml(item["분야 (중점영역)"])}
          </span>
        </div>
        <h3 class="text-lg font-bold text-slate-900 mb-2">${escapeHtml(item["세부 사업 / 주요 지표"])}</h3>
        <p class="text-sm text-slate-600 mb-4 bg-slate-50 p-3 rounded-lg border border-slate-100 leading-relaxed">
          ${escapeHtml(item["핵심 내용 및 보도자료 박스 원문 근거"]) || '핵심 원문 근거가 등록되어 있지 않습니다.'}
        </p>
        <div class="flex items-end justify-between pt-3 border-t border-slate-100">
          <div>
            <span class="text-xs text-slate-400 block">2027년 정부안 편성액</span>
            <span class="text-2xl font-black text-purple-700">${Number(item["2027년 정부안 (B)"]).toLocaleString()}</span>
            <span class="text-sm font-semibold text-slate-600">억원</span>
          </div>
          <button class="text-xs font-semibold text-purple-600 hover:text-purple-800 flex items-center gap-1">
            원문 전체보기 <i class="fa-solid fa-chevron-right text-[10px]"></i>
          </button>
        </div>
      </div>
    `).join('');
  }
}

// ==========================================
// 7. 전체 데이터 테이블 & 검색/필터
// ==========================================
function initTableControls() {
  const searchInput = document.getElementById('tableSearchInput');
  const domainFilter = document.getElementById('tableDomainFilter');
  const newOnlyFilter = document.getElementById('tableNewOnlyFilter');
  const prevPageBtn = document.getElementById('prevPageBtn');
  const nextPageBtn = document.getElementById('nextPageBtn');

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      AppState.tableSearch = e.target.value.trim().toLowerCase();
      AppState.currentPage = 1;
      applyTableFilters();
    });
  }

  if (domainFilter) {
    domainFilter.addEventListener('change', (e) => {
      AppState.tableFilterDomain = e.target.value;
      AppState.currentPage = 1;
      applyTableFilters();
    });
  }

  if (newOnlyFilter) {
    newOnlyFilter.addEventListener('change', (e) => {
      AppState.tableFilterNew = e.target.checked;
      AppState.currentPage = 1;
      applyTableFilters();
    });
  }

  if (prevPageBtn) {
    prevPageBtn.addEventListener('click', () => {
      if (AppState.currentPage > 1) {
        AppState.currentPage--;
        renderDataTable();
      }
    });
  }

  if (nextPageBtn) {
    nextPageBtn.addEventListener('click', () => {
      const maxPage = Math.ceil(AppState.filteredData.length / AppState.pageSize);
      if (AppState.currentPage < maxPage) {
        AppState.currentPage++;
        renderDataTable();
      }
    });
  }

  // 테이블 헤더 정렬 클릭 이벤트
  document.querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (AppState.tableSortCol === col) {
        AppState.tableSortAsc = !AppState.tableSortAsc;
      } else {
        AppState.tableSortCol = col;
        AppState.tableSortAsc = true;
      }
      applyTableFilters();
    });
  });
}

function applyTableFilters() {
  // 도메인 필터 옵션 채우기
  const domainFilter = document.getElementById('tableDomainFilter');
  if (domainFilter && domainFilter.options.length <= 1) {
    const domains = [...new Set(AppState.rawData.map(d => d["분야 (중점영역)"] || '기타'))];
    domainFilter.innerHTML = '<option value="">전체 분야</option>' + 
      domains.map(d => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join('');
  }

  let list = [...AppState.rawData];

  // 1. 검색어 필터
  if (AppState.tableSearch) {
    const q = AppState.tableSearch;
    list = list.filter(item => {
      const fDomain = (item["분야 (중점영역)"] || '').toLowerCase();
      const fProject = (item["세부 사업 / 주요 지표"] || '').toLowerCase();
      const fEvidence = (item["핵심 내용 및 보도자료 박스 원문 근거"] || '').toLowerCase();
      return fDomain.includes(q) || fProject.includes(q) || fEvidence.includes(q);
    });
  }

  // 2. 분야 필터
  if (AppState.tableFilterDomain) {
    list = list.filter(item => (item["분야 (중점영역)"] || '') === AppState.tableFilterDomain);
  }

  // 3. 신규사업 전용 필터
  if (AppState.tableFilterNew) {
    list = list.filter(item => item._isNew);
  }

  // 4. 정렬
  const col = AppState.tableSortCol;
  const asc = AppState.tableSortAsc ? 1 : -1;
  list.sort((a, b) => {
    let valA, valB;
    if (col === 'no') {
      valA = Number(a["연번"]) || 0;
      valB = Number(b["연번"]) || 0;
    } else if (col === 'domain') {
      valA = a["분야 (중점영역)"] || '';
      valB = b["분야 (중점영역)"] || '';
    } else if (col === 'project') {
      valA = a["세부 사업 / 주요 지표"] || '';
      valB = b["세부 사업 / 주요 지표"] || '';
    } else if (col === 'y2026') {
      valA = Number(a["2026년 (A)"]) || 0;
      valB = Number(b["2026년 (A)"]) || 0;
    } else if (col === 'y2027') {
      valA = Number(a["2027년 정부안 (B)"]) || 0;
      valB = Number(b["2027년 정부안 (B)"]) || 0;
    } else if (col === 'diff') {
      valA = Number(a["증감 (B-A)"]) || 0;
      valB = Number(b["증감 (B-A)"]) || 0;
    } else if (col === 'rate') {
      valA = Number(a["증감율 (%)"]) || 0;
      valB = Number(b["증감율 (%)"]) || 0;
    }
    if (valA > valB) return 1 * asc;
    if (valA < valB) return -1 * asc;
    return 0;
  });

  AppState.filteredData = list;
  renderDataTable();
}

function renderDataTable() {
  const tbody = document.getElementById('mainDataTableBody');
  const countEl = document.getElementById('tableRowCount');
  const pageInfoEl = document.getElementById('pageInfo');
  if (!tbody) return;

  const total = AppState.filteredData.length;
  if (countEl) countEl.textContent = `총 ${total}건`;

  const totalPages = Math.max(1, Math.ceil(total / AppState.pageSize));
  if (AppState.currentPage > totalPages) AppState.currentPage = totalPages;

  if (pageInfoEl) pageInfoEl.textContent = `${AppState.currentPage} / ${totalPages} 페이지`;

  const start = (AppState.currentPage - 1) * AppState.pageSize;
  const pageItems = AppState.filteredData.slice(start, start + AppState.pageSize);

  if (pageItems.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center py-10 text-slate-400">조건에 부합하는 데이터가 없습니다.</td></tr>`;
    return;
  }

  tbody.innerHTML = pageItems.map(item => {
    const diff = Number(item["증감 (B-A)"]) || 0;
    const rate = Number(item["증감율 (%)"]) || 0;

    return `
      <tr class="hover:bg-blue-50/40 transition" onclick="openDetailModal(${item["연번"]})">
        <td class="text-center text-slate-500 font-mono text-xs">${item["연번"]}</td>
        <td>
          <span class="badge-domain text-xs">${escapeHtml(item["분야 (중점영역)"])}</span>
        </td>
        <td class="font-semibold text-slate-900">
          <div class="flex items-center gap-1.5">
            ${escapeHtml(item["세부 사업 / 주요 지표"])}
            ${item._isNew ? '<span class="badge-new text-[10px]">신규</span>' : ''}
          </div>
        </td>
        <td class="text-center text-xs text-slate-500">${escapeHtml(item["구분/단위"])}</td>
        <td class="text-right font-mono text-slate-600">${Number(item["2026년 (A)"]).toLocaleString()}</td>
        <td class="text-right font-mono font-bold text-slate-900">${Number(item["2027년 정부안 (B)"]).toLocaleString()}</td>
        <td class="text-right font-mono font-semibold ${diff >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
          ${diff >= 0 ? '+' : ''}${diff.toLocaleString()}
        </td>
        <td class="text-right font-mono font-bold ${diff >= 0 ? 'text-emerald-600' : 'text-rose-600'}">
          ${diff >= 0 ? '+' : ''}${rate.toFixed(1)}%
        </td>
        <td class="text-xs text-slate-600 max-w-xs truncate" title="${escapeHtml(item["핵심 내용 및 보도자료 박스 원문 근거"])}">
          ${escapeHtml(item["핵심 내용 및 보도자료 박스 원문 근거"]) || '-'}
        </td>
      </tr>
    `;
  }).join('');
}

// ==========================================
// 8. 상세 모달 (원문 근거 팝업)
// ==========================================
function openDetailModal(no) {
  const item = AppState.rawData.find(d => Number(d["연번"]) === Number(no));
  if (!item) return;

  const modal = document.getElementById('detailModal');
  if (!modal) return;

  document.getElementById('modalProjectName').textContent = item["세부 사업 / 주요 지표"];
  document.getElementById('modalDomainName').textContent = item["분야 (중점영역)"];
  document.getElementById('modalUnit').textContent = item["구분/단위"] || '억원';
  
  const badgeNew = document.getElementById('modalNewBadge');
  if (item._isNew) {
    badgeNew.classList.remove('hidden');
  } else {
    badgeNew.classList.add('hidden');
  }

  document.getElementById('modal2026').textContent = `${Number(item["2026년 (A)"]).toLocaleString()}억원`;
  document.getElementById('modal2027').textContent = `${Number(item["2027년 정부안 (B)"]).toLocaleString()}억원`;
  
  const diff = Number(item["증감 (B-A)"]) || 0;
  const rate = Number(item["증감율 (%)"]) || 0;
  const diffEl = document.getElementById('modalDiff');
  diffEl.textContent = `${diff >= 0 ? '+' : ''}${diff.toLocaleString()}억원 (${diff >= 0 ? '+' : ''}${rate.toFixed(1)}%)`;
  diffEl.className = diff >= 0 ? 'text-lg font-bold text-emerald-600' : 'text-lg font-bold text-rose-600';

  document.getElementById('modalEvidenceText').textContent = item["핵심 내용 및 보도자료 박스 원문 근거"] || '기재된 원문 및 근거 내용이 없습니다.';

  modal.classList.remove('hidden');
}

function closeDetailModal() {
  const modal = document.getElementById('detailModal');
  if (modal) modal.classList.add('hidden');
}

// ==========================================
// 9. 내보내기 & 템플릿 다운로드
// ==========================================
function exportToCsv() {
  if (AppState.filteredData.length === 0) {
    alert('내보낼 데이터가 없습니다.');
    return;
  }

  const headers = ['연번', '분야 (중점영역)', '세부 사업 / 주요 지표', '구분/단위', '2026년 (A)', '2027년 정부안 (B)', '증감 (B-A)', '증감율 (%)', '핵심 내용 및 보도자료 박스 원문 근거'];
  
  let csvContent = '\uFEFF'; // BOM for Excel UTF-8
  csvContent += headers.join(',') + '\r\n';

  AppState.filteredData.forEach(row => {
    const line = [
      row["연번"],
      `"${(row["분야 (중점영역)"] || '').replace(/"/g, '""')}"`,
      `"${(row["세부 사업 / 주요 지표"] || '').replace(/"/g, '""')}"`,
      `"${(row["구분/단위"] || '').replace(/"/g, '""')}"`,
      row["2026년 (A)"],
      row["2027년 정부안 (B)"],
      row["증감 (B-A)"],
      `"${row["증감율 (%)"]}%"`,
      `"${(row["핵심 내용 및 보도자료 박스 원문 근거"] || '').replace(/"/g, '""')}"`
    ];
    csvContent += line.join(',') + '\r\n';
  });

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `국방예산_대시보드_데이터_${new Date().toISOString().slice(0,10)}.csv`;
  link.click();
}

function exportTemplate() {
  const headers = ['연번', '분야 (중점영역)', '세부 사업 / 주요 지표', '구분/단위', '2026년 (A)', '2027년 정부안 (B)', '증감 (B-A)', '증감율 (%)', '핵심 내용 및 보도자료 박스 원문 근거'];
  const sampleRows = [
    [1, '① 자주국방 & 전작권 전환', '전작권 전환 및 자주국방 핵심전력 확보', '억원', 90222, 121744, 31522, '34.9%', '전작권 전환 및 자주국방 예산 34.9% 대폭 증액 (KF-21 양산, 핵추진 잠수함 착수 등)'],
    [2, '② 효율적 군 인력구조 전환', '상비예비군 운영 예산', '억원', 87, 256, 169, '194.3%', '평시 추가훈련을 통한 전시 즉각 전투력 발휘 보장 (+169억원 증액, 3,700명 -> 7,000명 확대)'],
    [3, '① 자주국방 & 전작권 전환', '핵추진 잠수함 사업 착수', '신규', 0, 1500, 1500, '신규', '40년 장기 국책 프로젝트 출범. 국내 1,800여 개 공급망 기업 참여 및 원자력·조선 4만 개 이상 고급 일자리 창출']
  ];

  // XLSX 사용 가능하면 XLSX로 내보내기
  if (typeof XLSX !== 'undefined') {
    const wsData = [headers, ...sampleRows];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // 열 너비 설정
    ws['!cols'] = [
      { wch: 8 },  // 연번
      { wch: 26 }, // 분야
      { wch: 36 }, // 세부사업
      { wch: 12 }, // 단위
      { wch: 14 }, // 2026
      { wch: 16 }, // 2027
      { wch: 14 }, // 증감
      { wch: 12 }, // 증감율
      { wch: 60 }  // 원문근거
    ];

    XLSX.utils.book_append_sheet(wb, ws, "국방예산_서식");
    XLSX.writeFile(wb, "국방예산_입력_양식_템플릿.xlsx");
  } else {
    // CSV fallback
    exportToCsv();
  }
}

// 유틸리티: HTML 이스케이프
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
