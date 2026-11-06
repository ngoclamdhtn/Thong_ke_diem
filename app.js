/**
 * HỆ THỐNG QUẢN LÝ & THỐNG KÊ ĐIỂM THI THEO LỚP
 * Tự động kết nối điểm từ file XLS/XLSX vào Danh sách & Thống kê theo chuẩn GD&ĐT
 */

(function () {
  'use strict';

  // State
  const state = {
    students: [],           // Danh sách học sinh đầy đủ
    scoreMap: {},           // Map SBD -> Điểm { "12000001": 6.5, ... }
    fileName: '',           // Tên file điểm vừa chọn
    statsData: [],          // Dữ liệu bảng thống kê từng lớp
    summaryRow: null,       // Dòng CỘNG tổng kết
    scoredClassSet: new Set(), // Tập hợp các lớp có ít nhất 1 con điểm
    
    // Bộ lọc danh sách học sinh
    filters: {
      lop: '',
      phong: '',
      status: 'all',        // all | scored | unscored | ge8 | ge5 | lt5 | lt35
      search: ''
    },
    
    // Phân trang danh sách học sinh
    pagination: {
      page: 1,
      pageSize: 50
    },
    
    // Tùy chọn hiển thị
    options: {
      roundToOneDec: true,  // Làm tròn 1 chữ số thập phân (khớp ảnh mẫu)
      useCommaDec: true,    // Dùng dấu phẩy phân tách thập phân (chuẩn VN: 7,5)
      includeColGe9: true   // Hiển thị cột Điểm >= 9 như ảnh mẫu
    },
    
    // Biểu đồ Chart.js
    charts: {
      barComparison: null,
      stackedDistribution: null,
      doughnutTotal: null
    }
  };

  // Khởi tạo khi trang tải xong
  document.addEventListener('DOMContentLoaded', initApp);

  function initApp() {
    // 1. Nạp danh sách học sinh mặc định từ Danh_sach.xlsx (qua window.DEFAULT_DANH_SACH)
    loadDefaultStudents();

    // 2. Lắng nghe các sự kiện UI
    setupEventListeners();

    // 3. Khởi tạo biểu đồ rỗng hoặc sẵn sàng
    initCharts();

    // 4. Render danh sách ban đầu
    applyFiltersAndRenderStudents();
    recalculateStatistics();
  }

  /**
   * Nạp danh sách ban đầu từ biến window.DEFAULT_DANH_SACH
   */
  function loadDefaultStudents() {
    if (window.DEFAULT_DANH_SACH && Array.isArray(window.DEFAULT_DANH_SACH)) {
      state.students = window.DEFAULT_DANH_SACH.map(s => ({
        stt: s.stt,
        sbd: normalizeSBD(s.sbd),
        name: s.name || '',
        lop: (s.lop || '').trim(),
        phong: (s.phong || '').trim(),
        diem: (s.diem !== null && s.diem !== undefined && s.diem !== '') ? parseFloat(s.diem) : null
      }));
      showBanner(`Đã sẵn sàng dữ liệu gốc từ sheet <strong>Danh_sach</strong> của file <strong>Danh_sach.xlsx</strong> (${state.students.length.toLocaleString()} học sinh). Hãy chọn file điểm để kết nối!`, 'info');
    } else {
      showBanner('Chưa tìm thấy dữ liệu danh sách học sinh mặc định.', 'warning');
    }
    updateKPIs();
    populateFilterDropdowns();
  }

  /**
   * Chuẩn hóa Số Báo Danh (loại bỏ .0, khoảng trắng thừa)
   */
  function normalizeSBD(val) {
    if (val === null || val === undefined) return '';
    let str = String(val).trim();
    if (str.endsWith('.0')) {
      str = str.slice(0, -2);
    }
    return str;
  }

  /**
   * Chuẩn hóa điểm số
   */
  function parseScore(val) {
    if (val === null || val === undefined || val === '') return null;
    if (typeof val === 'number') return val;
    let s = String(val).trim().replace(',', '.');
    let num = parseFloat(s);
    return isNaN(num) ? null : num;
  }

  /**
   * Định dạng số hiển thị (chấm hoặc phẩy)
   */
  function formatNum(val, decimals = 1) {
    if (val === null || val === undefined || isNaN(val)) return '-';
    let formatted = Number(val).toFixed(decimals);
    if (state.options.useCommaDec) {
      formatted = formatted.replace('.', ',');
    }
    return formatted;
  }

  /**
   * Định dạng phần trăm hiển thị (VD: 28,57%)
   */
  function formatPercent(val) {
    if (val === null || val === undefined || isNaN(val)) return '0,00%';
    let formatted = Number(val).toFixed(2);
    if (state.options.useCommaDec) {
      formatted = formatted.replace('.', ',');
    }
    return formatted + '%';
  }

  /**
   * Thiết lập các sự kiện trên trang
   */
  function setupEventListeners() {
    // 1. File Upload (Chọn file điểm)
    const scoreFileInput = document.getElementById('score-file-input');
    const dropzone = document.getElementById('file-dropzone');
    const btnSelectFile = document.getElementById('btn-select-file');

    if (btnSelectFile && scoreFileInput) {
      btnSelectFile.addEventListener('click', () => scoreFileInput.click());
    }

    if (dropzone && scoreFileInput) {
      dropzone.addEventListener('click', (e) => {
        if (e.target !== btnSelectFile) scoreFileInput.click();
      });

      // Kéo thả file
      ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.add('dragover');
        });
      });

      ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.remove('dragover');
        });
      });

      dropzone.addEventListener('drop', (e) => {
        if (e.dataTransfer && e.dataTransfer.files.length > 0) {
          handleScoreFile(e.dataTransfer.files[0]);
        }
      });
    }

    if (scoreFileInput) {
      scoreFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          handleScoreFile(e.target.files[0]);
        }
      });
    }

    // Nút dùng file mẫu có sẵn
    const btnSampleFile = document.getElementById('btn-load-sample');
    if (btnSampleFile) {
      btnSampleFile.addEventListener('click', loadSampleScoreFile);
    }

    // Nút tải danh sách học sinh khác nếu muốn
    const rosterFileInput = document.getElementById('roster-file-input');
    const btnChangeRoster = document.getElementById('btn-change-roster');
    if (btnChangeRoster && rosterFileInput) {
      btnChangeRoster.addEventListener('click', () => rosterFileInput.click());
      rosterFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          handleRosterFile(e.target.files[0]);
        }
      });
    }

    // 2. Chuyển Tabs
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const targetTab = btn.getAttribute('data-tab');
        switchTab(targetTab);
      });
    });

    // 3. Nút Thống Kê
    const btnThongKe = document.getElementById('btn-thong-ke');
    if (btnThongKe) {
      btnThongKe.addEventListener('click', () => {
        recalculateStatistics();
        switchTab('tab-stats');
        showBanner('Đã cập nhật bảng thống kê điểm thi theo lớp thành công!', 'success');
      });
    }

    // 4. Nút Lọc và các input lọc
    const btnLoc = document.getElementById('btn-loc');
    const btnResetFilter = document.getElementById('btn-reset-filter');
    const selectLop = document.getElementById('filter-lop');
    const selectPhong = document.getElementById('filter-phong');
    const selectStatus = document.getElementById('filter-status');
    const inputSearch = document.getElementById('filter-search');

    if (btnLoc) {
      btnLoc.addEventListener('click', () => {
        readFilterInputs();
        state.pagination.page = 1;
        applyFiltersAndRenderStudents();
      });
    }

    if (btnResetFilter) {
      btnResetFilter.addEventListener('click', () => {
        selectLop.value = '';
        selectPhong.value = '';
        selectStatus.value = 'all';
        inputSearch.value = '';
        readFilterInputs();
        state.pagination.page = 1;
        applyFiltersAndRenderStudents();
      });
    }

    [selectLop, selectPhong, selectStatus].forEach(el => {
      if (el) el.addEventListener('change', () => {
        readFilterInputs();
        state.pagination.page = 1;
        applyFiltersAndRenderStudents();
      });
    });

    if (inputSearch) {
      inputSearch.addEventListener('input', debounce(() => {
        readFilterInputs();
        state.pagination.page = 1;
        applyFiltersAndRenderStudents();
      }, 250));
    }

    // 5. Nút Xuất Excel
    const btnXuatExcel = document.getElementById('btn-xuat-excel');
    if (btnXuatExcel) {
      btnXuatExcel.addEventListener('click', exportFullExcelWorkbook);
    }

    const btnXuatExcelStats = document.getElementById('btn-xuat-stats-only');
    if (btnXuatExcelStats) {
      btnXuatExcelStats.addEventListener('click', exportStatsOnlyExcel);
    }

    const btnXuatExcelScores = document.getElementById('btn-xuat-scores-only');
    if (btnXuatExcelScores) {
      btnXuatExcelScores.addEventListener('click', exportScoresOnlyExcel);
    }

    // 6. Phân trang
    const selectPageSize = document.getElementById('page-size-select');
    if (selectPageSize) {
      selectPageSize.addEventListener('change', (e) => {
        state.pagination.pageSize = e.target.value === 'all' ? 999999 : parseInt(e.target.value);
        state.pagination.page = 1;
        applyFiltersAndRenderStudents();
      });
    }

    // 7. Tùy chọn làm tròn & định dạng số
    const chkRound = document.getElementById('opt-round-one');
    if (chkRound) {
      chkRound.addEventListener('change', (e) => {
        state.options.roundToOneDec = e.target.checked;
        recalculateStatistics();
      });
    }

    const chkColGe9 = document.getElementById('opt-col-ge9');
    if (chkColGe9) {
      chkColGe9.addEventListener('change', (e) => {
        state.options.includeColGe9 = e.target.checked;
        renderStatisticsTable();
      });
    }

    // 8. In báo cáo
    const btnPrint = document.getElementById('btn-print-stats');
    if (btnPrint) {
      btnPrint.addEventListener('click', () => window.print());
    }

    // 9. Nút sao chép bảng thống kê
    const btnCopyTable = document.getElementById('btn-copy-stats');
    if (btnCopyTable) {
      btnCopyTable.addEventListener('click', copyStatsTableToClipboard);
    }
  }

  function readFilterInputs() {
    state.filters.lop = document.getElementById('filter-lop')?.value || '';
    state.filters.phong = document.getElementById('filter-phong')?.value || '';
    state.filters.status = document.getElementById('filter-status')?.value || 'all';
    state.filters.search = (document.getElementById('filter-search')?.value || '').trim().toLowerCase();
  }

  function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-tab') === tabId);
    });
    document.querySelectorAll('.tab-content').forEach(tab => {
      tab.classList.toggle('active', tab.id === tabId);
    });
    if (tabId === 'tab-charts') {
      updateCharts();
    }
  }

  /**
   * Đọc file điểm từ người dùng chọn (mau_file.xls hoặc .xlsx)
   * Yêu cầu: Kết nối điểm từ cột "E" trên sheet "Bảng điểm tổng" vào cột "F" của "Danh_sach"
   */
  function handleScoreFile(file) {
    if (!file) return;
    state.fileName = file.name;
    const reader = new FileReader();

    showBanner(`Đang đọc và xử lý file điểm "${file.name}"...`, 'info');

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Tìm sheet "Bảng điểm tổng" (hỗ trợ không phân biệt hoa thường hoặc khoảng trắng)
        let sheetName = workbook.SheetNames.find(name => 
          name.trim().toLowerCase() === 'bảng điểm tổng' || 
          name.trim().toLowerCase() === 'bang diem tong'
        );

        if (!sheetName) {
          // Nếu không tìm thấy chính xác, thử tìm sheet có chữ 'điểm' hoặc lấy sheet đầu tiên
          sheetName = workbook.SheetNames.find(name => 
            name.toLowerCase().includes('điểm') || name.toLowerCase().includes('diem')
          ) || workbook.SheetNames[0];
          console.warn(`Không tìm thấy sheet "Bảng điểm tổng", dùng sheet "${sheetName}"`);
        }

        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
          throw new Error(`Không đọc được nội dung sheet "${sheetName}" trong file.`);
        }

        // Chuyển sheet sang dạng mảng hàng [[A, B, C, D, E, ...], ...]
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        if (!rows || rows.length < 2) {
          throw new Error('Sheet điểm không có đủ dữ liệu hàng để kết nối.');
        }

        // Xác định vị trí cột:
        // Cột E (cột thứ 5, index 4): "Tổng điểm"
        // Cột C (cột thứ 3, index 2): "Số báo danh"
        let sbdColIdx = 2; // Cột C mặc định
        let scoreColIdx = 4; // Cột E mặc định

        const headerRow = rows[0] || [];
        for (let i = 0; i < headerRow.length; i++) {
          const h = String(headerRow[i] || '').trim().toLowerCase();
          if (h.includes('số báo danh') || h.includes('so bao danh') || h === 'sbd') {
            sbdColIdx = i;
          }
          if (h.includes('tổng điểm') || h.includes('tong diem') || h === 'điểm' || h === 'diem') {
            scoreColIdx = i;
          }
        }

        const newScoreMap = {};
        let validScoresCount = 0;

        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row || row.length === 0) continue;

          const sbdRaw = row[sbdColIdx];
          const sbd = normalizeSBD(sbdRaw);
          if (!sbd) continue;

          const scoreVal = parseScore(row[scoreColIdx]);
          if (scoreVal !== null) {
            newScoreMap[sbd] = scoreVal;
            validScoresCount++;
          }
        }

        state.scoreMap = newScoreMap;

        // Tiến hành KẾT NỐI ĐIỂM vào cột F của Danh sách học sinh
        connectScoresToStudents();

        // Cập nhật lại giao diện
        populateFilterDropdowns();
        applyFiltersAndRenderStudents();
        recalculateStatistics();
        updateKPIs();

        // Hiển thị thông báo kết quả chi tiết
        const matchedCount = state.students.filter(s => s.diem !== null).length;
        const totalClassesWithScores = state.scoredClassSet.size;
        
        showBanner(
          `✅ <strong>Kết nối thành công!</strong> Đã đọc <strong>${validScoresCount}</strong> bản ghi từ sheet <em>"${sheetName}"</em> (cột E). Kết nối được <strong>${matchedCount}</strong> học sinh trong <strong>${totalClassesWithScores}</strong> lớp có điểm.`,
          'success'
        );

        // Đổi tên nhãn trên dropzone
        const dropzoneText = document.querySelector('.dropzone-text h4');
        const dropzoneSub = document.querySelector('.dropzone-text p');
        if (dropzoneText) dropzoneText.textContent = `File đang kết nối: ${file.name}`;
        if (dropzoneSub) dropzoneSub.textContent = `Đã kết nối ${matchedCount} học sinh (${(file.size / 1024).toFixed(1)} KB)`;

      } catch (err) {
        console.error('Lỗi khi đọc file điểm:', err);
        showBanner(`❌ Lỗi khi đọc file: ${err.message}`, 'warning');
      }
    };

    reader.onerror = function () {
      showBanner('❌ Không thể mở file do lỗi từ trình duyệt.', 'warning');
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * Kết nối điểm từ state.scoreMap vào từng học sinh
   */
  function connectScoresToStudents() {
    state.scoredClassSet.clear();

    state.students.forEach(student => {
      const sbd = student.sbd;
      if (sbd && state.scoreMap.hasOwnProperty(sbd)) {
        student.diem = state.scoreMap[sbd];
        if (student.lop) {
          state.scoredClassSet.add(student.lop);
        }
      } else {
        student.diem = null;
      }
    });
  }

  /**
   * Tính toán lại toàn bộ Bảng Thống Kê theo lớp
   * Yêu cầu nghiệp vụ:
   * "Lớp không có con điểm nào thì không xuất hiện trong bảng thống kê."
   * Các trường thông tin:
   * TT, Lớp, Số lượng trên danh sách, Số lượng có điểm, Điểm trung bình, Điểm trung vị,
   * Điểm cao nhất, Điểm thấp nhất, Điểm>=9 (theo ảnh), Tỉ lệ %, Điểm>=8, Tỉ lệ %,
   * Điểm>=5, Tỉ lệ %, 3.5<Điểm<5, Tỉ lệ %, Điểm<3.5, Tỉ lệ %
   */
  function recalculateStatistics() {
    // 1. Nhóm học sinh theo lớp
    const classMap = new Map();

    state.students.forEach(s => {
      const lop = s.lop || 'Chưa xếp lớp';
      if (!classMap.has(lop)) {
        classMap.set(lop, {
          lop: lop,
          totalOnList: 0,
          scores: []
        });
      }
      const item = classMap.get(lop);
      item.totalOnList++;
      if (s.diem !== null && s.diem !== undefined) {
        // Áp dụng làm tròn điểm theo chuẩn 1 chữ số thập phân (nếu được chọn)
        const effectiveScore = state.options.roundToOneDec ? Math.round((s.diem + 1e-9) * 10) / 10 : s.diem;
        item.scores.push(effectiveScore);
      }
    });

    // 2. Lọc chỉ lấy các lớp CÓ ÍT NHẤT 1 CON ĐIỂM
    const scoredClasses = [];
    classMap.forEach(item => {
      if (item.scores.length > 0) {
        scoredClasses.push(item);
      }
    });

    // Sắp xếp lớp theo thứ tự tự nhiên (12A1, 12A2, ...)
    scoredClasses.sort((a, b) => a.lop.localeCompare(b.lop, 'vi', { numeric: true }));

    // Cập nhật tập hợp các lớp có điểm
    state.scoredClassSet = new Set(scoredClasses.map(c => c.lop));

    // 3. Tính toán các chỉ số cho từng lớp
    const statsRows = [];
    let allScoredStudents = [];
    let grandTotalOnList = 0;
    let grandTotalWithScore = 0;

    scoredClasses.forEach((item, idx) => {
      const sc = item.scores.slice().sort((a, b) => a - b);
      const nTotal = item.totalOnList;
      const nScored = sc.length;

      grandTotalOnList += nTotal;
      grandTotalWithScore += nScored;
      allScoredStudents = allScoredStudents.concat(sc);

      // Điểm trung bình
      const sum = sc.reduce((acc, v) => acc + v, 0);
      const avg = sum / nScored;

      // Điểm trung vị
      const median = calculateMedian(sc);

      // Điểm cao nhất, thấp nhất
      const maxScore = sc[sc.length - 1];
      const minScore = sc[0];

      // Đếm theo các mức điểm
      const ge9 = sc.filter(x => x >= 9.0).length;
      const ge8 = sc.filter(x => x >= 8.0).length;
      const ge5 = sc.filter(x => x >= 5.0).length;
      // 3.5 <= Điểm < 5.0 (khớp chính xác bảng ảnh mẫu chuẩn GD&ĐT)
      const range35_5 = sc.filter(x => x >= 3.5 && x < 5.0).length;
      // Điểm < 3.5
      const lt35 = sc.filter(x => x < 3.5).length;

      statsRows.push({
        tt: idx + 1,
        lop: item.lop,
        totalOnList: nTotal,
        totalWithScore: nScored,
        avg: avg,
        median: median,
        maxScore: maxScore,
        minScore: minScore,
        ge9: ge9,
        rateGe9: (ge9 / nScored) * 100,
        ge8: ge8,
        rateGe8: (ge8 / nScored) * 100,
        ge5: ge5,
        rateGe5: (ge5 / nScored) * 100,
        range35_5: range35_5,
        rateRange35_5: (range35_5 / nScored) * 100,
        lt35: lt35,
        rateLt35: (lt35 / nScored) * 100
      });
    });

    state.statsData = statsRows;

    // 4. Tính toán dòng CỘNG (Tổng kết toàn bộ các lớp có điểm)
    if (allScoredStudents.length > 0) {
      allScoredStudents.sort((a, b) => a - b);
      const grandSum = allScoredStudents.reduce((acc, v) => acc + v, 0);
      const grandAvg = grandSum / grandTotalWithScore;
      const grandMedian = calculateMedian(allScoredStudents);
      const grandMax = allScoredStudents[allScoredStudents.length - 1];
      const grandMin = allScoredStudents[0];

      const grandGe9 = allScoredStudents.filter(x => x >= 9.0).length;
      const grandGe8 = allScoredStudents.filter(x => x >= 8.0).length;
      const grandGe5 = allScoredStudents.filter(x => x >= 5.0).length;
      const grandRange35_5 = allScoredStudents.filter(x => x >= 3.5 && x < 5.0).length;
      const grandLt35 = allScoredStudents.filter(x => x < 3.5).length;

      state.summaryRow = {
        tt: 'CỘNG',
        lop: '',
        totalOnList: grandTotalOnList,
        totalWithScore: grandTotalWithScore,
        avg: grandAvg,
        median: grandMedian,
        maxScore: grandMax,
        minScore: grandMin,
        ge9: grandGe9,
        rateGe9: (grandGe9 / grandTotalWithScore) * 100,
        ge8: grandGe8,
        rateGe8: (grandGe8 / grandTotalWithScore) * 100,
        ge5: grandGe5,
        rateGe5: (grandGe5 / grandTotalWithScore) * 100,
        range35_5: grandRange35_5,
        rateRange35_5: (grandRange35_5 / grandTotalWithScore) * 100,
        lt35: grandLt35,
        rateLt35: (grandLt35 / grandTotalWithScore) * 100
      };
    } else {
      state.summaryRow = null;
    }

    renderStatisticsTable();
    updateKPIs();
    updateCharts();
  }

  /**
   * Tính điểm trung vị (Median)
   */
  function calculateMedian(sortedArr) {
    if (!sortedArr || sortedArr.length === 0) return 0;
    const mid = Math.floor(sortedArr.length / 2);
    if (sortedArr.length % 2 !== 0) {
      return sortedArr[mid];
    } else {
      return (sortedArr[mid - 1] + sortedArr[mid]) / 2;
    }
  }

  /**
   * Render Bảng Thống Kê Theo Lớp theo đúng phong cách ảnh mẫu
   */
  function renderStatisticsTable() {
    const tbody = document.getElementById('stats-table-body');
    const tfoot = document.getElementById('stats-table-foot');
    const theadRow = document.getElementById('stats-thead-row');
    if (!tbody) return;

    tbody.innerHTML = '';
    if (tfoot) tfoot.innerHTML = '';

    // Cập nhật header nếu có/ẩn cột Điểm >= 9
    const colGe9Visible = state.options.includeColGe9;
    const ge9HeaderCells = document.querySelectorAll('.th-ge9');
    ge9HeaderCells.forEach(el => el.style.display = colGe9Visible ? '' : 'none');

    if (state.statsData.length === 0) {
      const colSpan = colGe9Visible ? 18 : 16;
      tbody.innerHTML = `<tr><td colspan="${colSpan}" style="padding: 30px; color: #64748b; font-style: italic;">Chưa có dữ liệu điểm thi. Vui lòng chọn file điểm để xem thống kê.</td></tr>`;
      return;
    }

    // Render từng hàng lớp
    state.statsData.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${row.tt}</td>
        <td class="col-lop">${escapeHtml(row.lop)}</td>
        <td>${row.totalOnList}</td>
        <td>${row.totalWithScore}</td>
        <td>${formatNum(row.avg, 1)}</td>
        <td>${formatNum(row.median, row.median % 1 === 0 ? 0 : (row.median * 10) % 1 === 0 ? 1 : 2)}</td>
        <td>${formatNum(row.maxScore, row.maxScore % 1 === 0 ? 0 : 1)}</td>
        <td>${formatNum(row.minScore, row.minScore % 1 === 0 ? 0 : 1)}</td>
        ${colGe9Visible ? `
          <td class="text-red">${row.ge9}</td>
          <td class="text-red">${formatPercent(row.rateGe9)}</td>
        ` : ''}
        <td>${row.ge8}</td>
        <td>${formatPercent(row.rateGe8)}</td>
        <td class="text-blue">${row.ge5}</td>
        <td class="text-blue">${formatPercent(row.rateGe5)}</td>
        <td>${row.range35_5}</td>
        <td>${formatPercent(row.rateRange35_5)}</td>
        <td class="text-red">${row.lt35}</td>
        <td class="text-red">${formatPercent(row.rateLt35)}</td>
      `;
      tbody.appendChild(tr);
    });

    // Render dòng CỘNG ở footer
    if (tfoot && state.summaryRow) {
      const s = state.summaryRow;
      const trFoot = document.createElement('tr');
      trFoot.className = 'row-cong';
      trFoot.innerHTML = `
        <td style="font-weight: 800;">${s.tt}</td>
        <td></td>
        <td style="font-weight: 800;">${s.totalOnList}</td>
        <td style="font-weight: 800;">${s.totalWithScore}</td>
        <td style="font-weight: 800;">${formatNum(s.avg, 1)}</td>
        <td style="font-weight: 800;">${formatNum(s.median, 1)}</td>
        <td style="font-weight: 800;">${formatNum(s.maxScore, s.maxScore % 1 === 0 ? 0 : 1)}</td>
        <td style="font-weight: 800;">${formatNum(s.minScore, s.minScore % 1 === 0 ? 0 : 1)}</td>
        ${colGe9Visible ? `
          <td class="text-red" style="font-weight: 800;">${s.ge9}</td>
          <td class="text-red" style="font-weight: 800;">${formatPercent(s.rateGe9)}</td>
        ` : ''}
        <td style="font-weight: 800;">${s.ge8}</td>
        <td style="font-weight: 800;">${formatPercent(s.rateGe8)}</td>
        <td class="text-blue" style="font-weight: 800;">${s.ge5}</td>
        <td class="highlight-yellow">${formatPercent(s.rateGe5)}</td>
        <td style="font-weight: 800;">${s.range35_5}</td>
        <td style="font-weight: 800;">${formatPercent(s.rateRange35_5)}</td>
        <td class="text-red" style="font-weight: 800;">${s.lt35}</td>
        <td class="text-red" style="font-weight: 800;">${formatPercent(s.rateLt35)}</td>
      `;
      tfoot.appendChild(trFoot);
    }
  }

  /**
   * Lọc và Render Danh sách học sinh chi tiết
   */
  function applyFiltersAndRenderStudents() {
    const { lop, phong, status, search } = state.filters;

    const filtered = state.students.filter(s => {
      // Lọc theo Lớp
      if (lop && s.lop !== lop) return false;

      // Lọc theo Phòng
      if (phong && s.phong !== phong) return false;

      // Lọc theo Tình trạng điểm
      if (status === 'scored' && s.diem === null) return false;
      if (status === 'unscored' && s.diem !== null) return false;
      if (status === 'ge8' && (s.diem === null || s.diem < 8.0)) return false;
      if (status === 'ge5' && (s.diem === null || s.diem < 5.0)) return false;
      if (status === 'lt5' && (s.diem === null || s.diem >= 5.0)) return false;
      if (status === 'lt35' && (s.diem === null || s.diem >= 3.5)) return false;

      // Tìm kiếm theo SBD hoặc Họ và tên
      if (search) {
        const sbdMatch = s.sbd.toLowerCase().includes(search);
        const nameMatch = s.name.toLowerCase().includes(search);
        if (!sbdMatch && !nameMatch) return false;
      }

      return true;
    });

    renderStudentsTable(filtered);
    renderPagination(filtered.length);
  }

  /**
   * Render bảng danh sách học sinh
   */
  function renderStudentsTable(filteredList) {
    const tbody = document.getElementById('students-table-body');
    const badgeCount = document.getElementById('badge-students-count');
    if (!tbody) return;

    if (badgeCount) {
      badgeCount.textContent = `${filteredList.length.toLocaleString()} học sinh`;
    }

    const { page, pageSize } = state.pagination;
    const startIdx = (page - 1) * pageSize;
    const pageItems = filteredList.slice(startIdx, startIdx + pageSize);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 30px; color: #64748b; font-style: italic;">Không tìm thấy học sinh nào phù hợp với bộ lọc.</td></tr>`;
      return;
    }

    tbody.innerHTML = '';
    pageItems.forEach(s => {
      const tr = document.createElement('tr');
      const scoreBadge = getScoreBadge(s.diem);
      tr.innerHTML = `
        <td style="color: var(--text-muted); font-weight: 500;">${s.stt || '-'}</td>
        <td style="font-family: monospace; font-weight: 600; color: #1e3a8a;">${escapeHtml(s.sbd)}</td>
        <td style="font-weight: 600;">${escapeHtml(s.name)}</td>
        <td><span class="tag" style="font-weight: 600; background: #e0f2fe; color: #0369a1;">${escapeHtml(s.lop)}</span></td>
        <td><span class="tag">${escapeHtml(s.phong)}</span></td>
        <td>${scoreBadge}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  /**
   * Tạo Badge màu sắc theo thang điểm
   */
  function getScoreBadge(score) {
    if (score === null || score === undefined) {
      return `<span class="score-badge badge-none">Chưa có</span>`;
    }
    const val = Number(score);
    const formatted = formatNum(val, 1);

    if (val >= 9.0) {
      return `<span class="score-badge badge-excellent" title="Xuất sắc/Giỏi (>=9.0)">${formatted}</span>`;
    } else if (val >= 8.0) {
      return `<span class="score-badge badge-good" title="Giỏi (>=8.0)">${formatted}</span>`;
    } else if (val >= 5.0) {
      return `<span class="score-badge badge-average" title="Đạt (>=5.0)">${formatted}</span>`;
    } else if (val >= 3.5) {
      return `<span class="score-badge badge-weak" title="Yếu (3.5 - 4.9)">${formatted}</span>`;
    } else {
      return `<span class="score-badge badge-poor" title="Kém (<3.5)">${formatted}</span>`;
    }
  }

  /**
   * Render phân trang
   */
  function renderPagination(totalItems) {
    const container = document.getElementById('students-pagination');
    const info = document.getElementById('pagination-info');
    if (!container) return;

    const { page, pageSize } = state.pagination;
    const totalPages = Math.ceil(totalItems / pageSize) || 1;

    const startItem = totalItems === 0 ? 0 : (page - 1) * pageSize + 1;
    const endItem = Math.min(page * pageSize, totalItems);

    if (info) {
      info.textContent = `Hiển thị ${startItem} - ${endItem} trên tổng số ${totalItems.toLocaleString()} học sinh`;
    }

    container.innerHTML = '';
    if (totalPages <= 1) return;

    // Nút Trước
    const prevBtn = document.createElement('button');
    prevBtn.className = 'page-btn';
    prevBtn.innerHTML = '&laquo; Trước';
    prevBtn.disabled = page === 1;
    prevBtn.onclick = () => {
      state.pagination.page--;
      applyFiltersAndRenderStudents();
    };
    container.appendChild(prevBtn);

    // Các trang
    const maxVisiblePages = 5;
    let startPage = Math.max(1, page - 2);
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
    if (endPage - startPage < maxVisiblePages - 1) {
      startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    for (let p = startPage; p <= endPage; p++) {
      const pBtn = document.createElement('button');
      pBtn.className = `page-btn ${p === page ? 'active' : ''}`;
      pBtn.textContent = p;
      pBtn.onclick = () => {
        state.pagination.page = p;
        applyFiltersAndRenderStudents();
      };
      container.appendChild(pBtn);
    }

    // Nút Tiếp
    const nextBtn = document.createElement('button');
    nextBtn.className = 'page-btn';
    nextBtn.innerHTML = 'Sau &raquo;';
    nextBtn.disabled = page === totalPages;
    nextBtn.onclick = () => {
      state.pagination.page++;
      applyFiltersAndRenderStudents();
    };
    container.appendChild(nextBtn);
  }

  /**
   * Cập nhật danh sách dropdown cho Bộ lọc Lớp và Phòng
   */
  function populateFilterDropdowns() {
    const selectLop = document.getElementById('filter-lop');
    const selectPhong = document.getElementById('filter-phong');

    const classSet = new Set();
    const roomSet = new Set();
    const classScoreCounts = {};

    state.students.forEach(s => {
      if (s.lop) {
        classSet.add(s.lop);
        if (!classScoreCounts[s.lop]) {
          classScoreCounts[s.lop] = { total: 0, scored: 0 };
        }
        classScoreCounts[s.lop].total++;
        if (s.diem !== null) classScoreCounts[s.lop].scored++;
      }
      if (s.phong) roomSet.add(s.phong);
    });

    const sortedClasses = Array.from(classSet).sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }));
    const sortedRooms = Array.from(roomSet).sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }));

    if (selectLop) {
      const curVal = selectLop.value;
      selectLop.innerHTML = '<option value="">-- Tất cả các lớp --</option>';
      sortedClasses.forEach(c => {
        const info = classScoreCounts[c];
        const label = info && info.scored > 0 
          ? `${c} (${info.scored}/${info.total} có điểm)` 
          : `${c} (${info ? info.total : 0} HS)`;
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = label;
        if (c === curVal) opt.selected = true;
        selectLop.appendChild(opt);
      });
    }

    if (selectPhong) {
      const curVal = selectPhong.value;
      selectPhong.innerHTML = '<option value="">-- Tất cả các phòng --</option>';
      sortedRooms.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p;
        opt.textContent = p;
        if (p === curVal) opt.selected = true;
        selectPhong.appendChild(opt);
      });
    }
  }

  /**
   * Cập nhật các ô Thẻ chỉ số tổng quan (KPI)
   */
  function updateKPIs() {
    const totalStudents = state.students.length;
    const scoredStudents = state.students.filter(s => s.diem !== null).length;
    const rateScored = totalStudents > 0 ? (scoredStudents / totalStudents) * 100 : 0;
    const totalClassesWithScores = state.scoredClassSet.size;

    const elTotal = document.getElementById('kpi-total-students');
    const elScored = document.getElementById('kpi-scored-students');
    const elClasses = document.getElementById('kpi-scored-classes');
    const elAvg = document.getElementById('kpi-avg-score');
    const elPassRate = document.getElementById('kpi-pass-rate');

    if (elTotal) elTotal.textContent = totalStudents.toLocaleString();
    if (elScored) elScored.textContent = `${scoredStudents.toLocaleString()} (${formatPercent(rateScored)})`;
    if (elClasses) elClasses.textContent = `${totalClassesWithScores} lớp`;

    if (state.summaryRow) {
      if (elAvg) elAvg.textContent = formatNum(state.summaryRow.avg, 1);
      if (elPassRate) elPassRate.textContent = formatPercent(state.summaryRow.rateGe5);
    } else {
      if (elAvg) elAvg.textContent = '-';
      if (elPassRate) elPassRate.textContent = '-';
    }
  }

  /**
   * XUẤT FILE EXCEL:
   * Yêu cầu nghiệp vụ:
   * "Có nút 'Xuất Excel' để xuất ra danh sách điểm đã kết nối và bảng thống kê bằng excel.
   * Lớp không có con điểm nào thì không xuất hiện trong danh sách điểm và bảng thống kê."
   */
  function exportFullExcelWorkbook() {
    if (state.statsData.length === 0) {
      alert('Chưa có dữ liệu thống kê để xuất Excel. Vui lòng kết nối file điểm trước.');
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      // ==========================================
      // SHEET 1: DANH SÁCH ĐIỂM
      // "Lớp không có con điểm nào thì không xuất hiện trong danh sách điểm"
      // ==========================================
      const scoredStudents = state.students.filter(s => state.scoredClassSet.has(s.lop));

      // Sắp xếp theo Lớp và STT
      scoredStudents.sort((a, b) => {
        if (a.lop !== b.lop) return a.lop.localeCompare(b.lop, 'vi', { numeric: true });
        return (a.stt || 0) - (b.stt || 0);
      });

      const scoreSheetData = [
        ['DANH SÁCH ĐIỂM THI ĐÃ KẾT NỐI THEO LỚP'],
        [`(Chỉ bao gồm các lớp có điểm thi - Tổng cộng ${scoredStudents.length} học sinh)`],
        [],
        ['STT', 'SBD', 'Họ và tên', 'Lớp', 'Phòng', 'Điểm']
      ];

      scoredStudents.forEach((s, idx) => {
        scoreSheetData.push([
          idx + 1,
          s.sbd,
          s.name,
          s.lop,
          s.phong,
          s.diem !== null ? s.diem : ''
        ]);
      });

      const wsScores = XLSX.utils.aoa_to_sheet(scoreSheetData);
      wsScores['!cols'] = [
        { wch: 8 },  // STT
        { wch: 15 }, // SBD
        { wch: 28 }, // Họ và tên
        { wch: 12 }, // Lớp
        { wch: 12 }, // Phòng
        { wch: 12 }  // Điểm
      ];
      XLSX.utils.book_append_sheet(wb, wsScores, 'Danh_sach_diem');

      // ==========================================
      // SHEET 2: BẢNG THỐNG KÊ THEO LỚP
      // Khớp chuẩn theo cấu trúc ảnh mẫu
      // ==========================================
      const statsHeaders = [
        'TT', 'Lớp', 'Số lượng trên danh sách', 'Số lượng có điểm',
        'Điểm trung bình', 'Điểm trung vị', 'Điểm cao nhất', 'Điểm thấp nhất',
        'Điểm>=9', 'Tỉ lệ %',
        'Điểm>=8', 'Tỉ lệ %',
        'Điểm>=5', 'Tỉ lệ %',
        '3.5<Điểm<5', 'Tỉ lệ %',
        'Điểm<3.5', 'Tỉ lệ %'
      ];

      const statsSheetData = [
        ['BẢNG THỐNG KÊ KẾT QUẢ ĐIỂM THI THEO LỚP'],
        [],
        statsHeaders
      ];

      state.statsData.forEach(r => {
        statsSheetData.push([
          r.tt,
          r.lop,
          r.totalOnList,
          r.totalWithScore,
          Number(r.avg.toFixed(1)),
          Number(r.median.toFixed(2)),
          r.maxScore,
          r.minScore,
          r.ge9,
          Number((r.rateGe9 / 100).toFixed(4)),
          r.ge8,
          Number((r.rateGe8 / 100).toFixed(4)),
          r.ge5,
          Number((r.rateGe5 / 100).toFixed(4)),
          r.range35_5,
          Number((r.rateRange35_5 / 100).toFixed(4)),
          r.lt35,
          Number((r.rateLt35 / 100).toFixed(4))
        ]);
      });

      // Thêm dòng CỘNG
      if (state.summaryRow) {
        const s = state.summaryRow;
        statsSheetData.push([
          s.tt,
          '',
          s.totalOnList,
          s.totalWithScore,
          Number(s.avg.toFixed(1)),
          Number(s.median.toFixed(2)),
          s.maxScore,
          s.minScore,
          s.ge9,
          Number((s.rateGe9 / 100).toFixed(4)),
          s.ge8,
          Number((s.rateGe8 / 100).toFixed(4)),
          s.ge5,
          Number((s.rateGe5 / 100).toFixed(4)),
          s.range35_5,
          Number((s.rateRange35_5 / 100).toFixed(4)),
          s.lt35,
          Number((s.rateLt35 / 100).toFixed(4))
        ]);
      }

      const wsStats = XLSX.utils.aoa_to_sheet(statsSheetData);

      // Đặt định dạng phần trăm cho các cột tỉ lệ
      const range = XLSX.utils.decode_range(wsStats['!ref']);
      const percentCols = [9, 11, 13, 15, 17]; // 0-based col indices for Tỉ lệ %
      for (let R = 3; R <= range.e.r; ++R) {
        percentCols.forEach(C => {
          const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
          if (wsStats[cellAddr] && typeof wsStats[cellAddr].v === 'number') {
            wsStats[cellAddr].z = '0.00%';
          }
        });
      }

      wsStats['!cols'] = [
        { wch: 6 },  // TT
        { wch: 10 }, // Lớp
        { wch: 14 }, // SL trên DS
        { wch: 14 }, // SL có điểm
        { wch: 12 }, // ĐTB
        { wch: 12 }, // Trung vị
        { wch: 12 }, // Cao nhất
        { wch: 12 }, // Thấp nhất
        { wch: 10 }, { wch: 11 }, // >=9, %
        { wch: 10 }, { wch: 11 }, // >=8, %
        { wch: 10 }, { wch: 11 }, // >=5, %
        { wch: 12 }, { wch: 11 }, // 3.5-5, %
        { wch: 10 }, { wch: 11 }  // <3.5, %
      ];

      XLSX.utils.book_append_sheet(wb, wsStats, 'Thong_ke_theo_lop');

      // Tải file về
      const today = new Date().toISOString().slice(0, 10);
      downloadWorkbook(wb, `Thong_Ke_Va_Diem_Thi_${today}.xlsx`);
      showBanner(`Đã xuất trọn bộ Excel thành công (2 sheet: Danh sách điểm & Thống kê)!`, 'success');

    } catch (err) {
      console.error('Lỗi xuất Excel:', err);
      alert('Không thể xuất Excel: ' + err.message);
    }
  }

  function exportStatsOnlyExcel() {
    if (state.statsData.length === 0) {
      alert('Chưa có dữ liệu thống kê.');
      return;
    }
    const wb = XLSX.utils.book_new();
    const statsHeaders = [
      'TT', 'Lớp', 'Số lượng trên danh sách', 'Số lượng có điểm',
      'Điểm trung bình', 'Điểm trung vị', 'Điểm cao nhất', 'Điểm thấp nhất',
      'Điểm>=9', 'Tỉ lệ %', 'Điểm>=8', 'Tỉ lệ %', 'Điểm>=5', 'Tỉ lệ %',
      '3.5<Điểm<5', 'Tỉ lệ %', 'Điểm<3.5', 'Tỉ lệ %'
    ];
    const data = [['BẢNG THỐNG KÊ KẾT QUẢ ĐIỂM THI THEO LỚP'], [], statsHeaders];
    state.statsData.forEach(r => {
      data.push([
        r.tt, r.lop, r.totalOnList, r.totalWithScore,
        Number(r.avg.toFixed(1)), Number(r.median.toFixed(2)), r.maxScore, r.minScore,
        r.ge9, Number((r.rateGe9 / 100).toFixed(4)),
        r.ge8, Number((r.rateGe8 / 100).toFixed(4)),
        r.ge5, Number((r.rateGe5 / 100).toFixed(4)),
        r.range35_5, Number((r.rateRange35_5 / 100).toFixed(4)),
        r.lt35, Number((r.rateLt35 / 100).toFixed(4))
      ]);
    });
    if (state.summaryRow) {
      const s = state.summaryRow;
      data.push([
        s.tt, '', s.totalOnList, s.totalWithScore,
        Number(s.avg.toFixed(1)), Number(s.median.toFixed(2)), s.maxScore, s.minScore,
        s.ge9, Number((s.rateGe9 / 100).toFixed(4)),
        s.ge8, Number((s.rateGe8 / 100).toFixed(4)),
        s.ge5, Number((s.rateGe5 / 100).toFixed(4)),
        s.range35_5, Number((s.rateRange35_5 / 100).toFixed(4)),
        s.lt35, Number((s.rateLt35 / 100).toFixed(4))
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Thong_ke');
    downloadWorkbook(wb, `Bang_Thong_Ke_Theo_Lop_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function exportScoresOnlyExcel() {
    const scoredStudents = state.students.filter(s => state.scoredClassSet.has(s.lop));
    if (scoredStudents.length === 0) {
      alert('Chưa có học sinh nào có điểm để xuất.');
      return;
    }
    const wb = XLSX.utils.book_new();
    const data = [
      ['DANH SÁCH ĐIỂM THI HỌC SINH'],
      [],
      ['STT', 'SBD', 'Họ và tên', 'Lớp', 'Phòng', 'Điểm']
    ];
    scoredStudents.forEach((s, idx) => {
      data.push([idx + 1, s.sbd, s.name, s.lop, s.phong, s.diem]);
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Danh_sach_diem');
    downloadWorkbook(wb, `Danh_Sach_Diem_Hoc_Sinh_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  /**
   * Tải workbook xuống trình duyệt (hỗ trợ offline qua Blob)
   */
  function downloadWorkbook(wb, filename) {
    try {
      XLSX.writeFile(wb, filename);
    } catch (e) {
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);
    }
  }

  /**
   * Sao chép bảng thống kê vào Clipboard (dễ dán vào Excel / Word)
   */
  function copyStatsTableToClipboard() {
    const table = document.getElementById('stats-table');
    if (!table) return;

    let text = '';
    const rows = table.querySelectorAll('tr');
    rows.forEach(r => {
      const cells = r.querySelectorAll('th, td');
      const rowData = [];
      cells.forEach(c => {
        if (c.style.display !== 'none') {
          rowData.push(c.innerText.trim());
        }
      });
      text += rowData.join('\t') + '\n';
    });

    navigator.clipboard.writeText(text).then(() => {
      showBanner('📋 Đã sao chép toàn bộ bảng thống kê vào bộ nhớ tạm (Clipboard)! Bạn có thể dán trực tiếp vào Excel hoặc Word.', 'success');
    }).catch(err => {
      alert('Không thể sao chép: ' + err);
    });
  }

  /**
   * Tải file điểm mẫu có sẵn mau_file.xls (hoạt động ngay cả khi mở file:/// offline)
   */
  function loadSampleScoreFile() {
    if (window.SAMPLE_SCORE_MAP && Object.keys(window.SAMPLE_SCORE_MAP).length > 0) {
      state.fileName = 'mau_file.xls (Mẫu thử nghiệm)';
      state.scoreMap = Object.assign({}, window.SAMPLE_SCORE_MAP);
      connectScoresToStudents();
      populateFilterDropdowns();
      applyFiltersAndRenderStudents();
      recalculateStatistics();
      updateKPIs();

      const matchedCount = state.students.filter(s => s.diem !== null).length;
      showBanner(
        `✅ <strong>Đã tải dữ liệu mẫu từ "mau_file.xls"!</strong> Đã kết nối thành công <strong>${matchedCount}</strong> học sinh trong <strong>${state.scoredClassSet.size}</strong> lớp (12A1 &rarr; 12A9).`,
        'success'
      );

      const dropzoneText = document.querySelector('.dropzone-text h4');
      const dropzoneSub = document.querySelector('.dropzone-text p');
      if (dropzoneText) dropzoneText.textContent = 'File đang kết nối: mau_file.xls';
      if (dropzoneSub) dropzoneSub.textContent = `Đã kết nối ${matchedCount} học sinh có điểm`;
      return;
    }

    // Dự phòng qua fetch nếu có server
    fetch('mau_file.xls')
      .then(res => res.arrayBuffer())
      .then(buf => {
        const file = new File([buf], 'mau_file.xls');
        handleScoreFile(file);
      })
      .catch(() => {
        document.getElementById('score-file-input')?.click();
      });
  }

  /**
   * Xử lý khi người dùng muốn đổi file Danh sách học sinh khác
   */
  function handleRosterFile(file) {
    if (!file) return;
    const reader = new FileReader();
    showBanner(`Đang nạp danh sách học sinh từ file "${file.name}"...`, 'info');

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('danh_sach') || n.toLowerCase().includes('danh sach')) || wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!rows || rows.length < 2) throw new Error('File danh sách không đủ dữ liệu.');

        let sttIdx = 0, sbdIdx = 1, nameIdx = 2, lopIdx = 3, phongIdx = 4;
        const header = rows[0] || [];
        for (let i = 0; i < header.length; i++) {
          const h = String(header[i] || '').toLowerCase();
          if (h === 'stt') sttIdx = i;
          if (h.includes('sbd') || h.includes('báo danh')) sbdIdx = i;
          if (h.includes('họ') || h.includes('tên') || h.includes('name')) nameIdx = i;
          if (h.includes('lớp') || h.includes('lop')) lopIdx = i;
          if (h.includes('phòng') || h.includes('phong')) phongIdx = i;
        }

        const newStudents = [];
        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row || row.length === 0) continue;
          const sbd = normalizeSBD(row[sbdIdx]);
          if (!sbd && !row[nameIdx]) continue;
          newStudents.push({
            stt: row[sttIdx] || r,
            sbd: sbd,
            name: String(row[nameIdx] || '').trim(),
            lop: String(row[lopIdx] || '').trim(),
            phong: String(row[phongIdx] || '').trim(),
            diem: null
          });
        }

        state.students = newStudents;
        connectScoresToStudents();
        populateFilterDropdowns();
        applyFiltersAndRenderStudents();
        recalculateStatistics();
        updateKPIs();
        showBanner(`Đã nạp thành công ${newStudents.length} học sinh từ file "${file.name}"!`, 'success');

      } catch (err) {
        showBanner(`❌ Lỗi nạp danh sách: ${err.message}`, 'warning');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /**
   * Biểu đồ trực quan (Chart.js)
   */
  function initCharts() {
    if (typeof Chart === 'undefined') return;

    // Thiết lập font chữ chung cho Chart.js
    Chart.defaults.font.family = "'Segoe UI', -apple-system, system-ui, sans-serif";
    Chart.defaults.font.size = 12;
    Chart.defaults.color = '#475569';
  }

  function updateCharts() {
    if (typeof Chart === 'undefined') return;
    if (state.statsData.length === 0) return;

    const labels = state.statsData.map(r => r.lop);
    const avgScores = state.statsData.map(r => Number(r.avg.toFixed(1)));
    const medianScores = state.statsData.map(r => Number(r.median.toFixed(2)));

    // 1. Biểu đồ so sánh Điểm TB và Trung Vị
    const ctxBar = document.getElementById('chart-comparison')?.getContext('2d');
    if (ctxBar) {
      if (state.charts.barComparison) state.charts.barComparison.destroy();

      state.charts.barComparison = new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Điểm Trung Bình (ĐTB)',
              data: avgScores,
              backgroundColor: 'rgba(37, 99, 235, 0.85)',
              borderColor: '#1d4ed8',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: 'Điểm Trung Vị',
              data: medianScores,
              backgroundColor: 'rgba(16, 185, 129, 0.85)',
              borderColor: '#047857',
              borderWidth: 1,
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top' },
            tooltip: {
              callbacks: {
                label: (ctx) => `${ctx.dataset.label}: ${formatNum(ctx.parsed.y, 2)}`
              }
            }
          },
          scales: {
            y: {
              beginAtZero: true,
              max: 10,
              title: { display: true, text: 'Thang điểm 10' }
            }
          }
        }
      });
    }

    // 2. Biểu đồ xếp chồng phân bố phổ điểm theo lớp
    const ctxStacked = document.getElementById('chart-distribution')?.getContext('2d');
    if (ctxStacked) {
      if (state.charts.stackedDistribution) state.charts.stackedDistribution.destroy();

      const rateGe8 = state.statsData.map(r => Number(r.rateGe8.toFixed(1)));
      const rate5to8 = state.statsData.map(r => Number((r.rateGe5 - r.rateGe8).toFixed(1)));
      const rate35to5 = state.statsData.map(r => Number(r.rateRange35_5.toFixed(1)));
      const rateLt35 = state.statsData.map(r => Number(r.rateLt35.toFixed(1)));

      state.charts.stackedDistribution = new Chart(ctxStacked, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Điểm >= 8 (Giỏi)',
              data: rateGe8,
              backgroundColor: '#3b82f6'
            },
            {
              label: '5 <= Điểm < 8 (Khá/TB)',
              data: rate5to8,
              backgroundColor: '#10b981'
            },
            {
              label: '3.5 <= Điểm < 5 (Yếu)',
              data: rate35to5,
              backgroundColor: '#f59e0b'
            },
            {
              label: 'Điểm < 3.5 (Kém)',
              data: rateLt35,
              backgroundColor: '#ef4444'
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top' },
            tooltip: {
              callbacks: {
                label: (ctx) => `${ctx.dataset.label}: ${formatPercent(ctx.parsed.y)}`
              }
            }
          },
          scales: {
            x: { stacked: true },
            y: {
              stacked: true,
              max: 100,
              title: { display: true, text: 'Tỉ lệ %' }
            }
          }
        }
      });
    }

    // 3. Biểu đồ cơ cấu phổ điểm toàn khối (Doughnut)
    const ctxDoughnut = document.getElementById('chart-doughnut')?.getContext('2d');
    if (ctxDoughnut && state.summaryRow) {
      if (state.charts.doughnutTotal) state.charts.doughnutTotal.destroy();

      const s = state.summaryRow;
      const countGe8 = s.ge8;
      const count5to8 = s.ge5 - s.ge8;
      const count35to5 = s.range35_5;
      const countLt35 = s.lt35;

      state.charts.doughnutTotal = new Chart(ctxDoughnut, {
        type: 'doughnut',
        data: {
          labels: ['Điểm >= 8', '5 <= Điểm < 8', '3.5 <= Điểm < 5', 'Điểm < 3.5'],
          datasets: [{
            data: [countGe8, count5to8, count35to5, countLt35],
            backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444'],
            hoverOffset: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom' }
          }
        }
      });
    }
  }

  /**
   * Hiển thị thông báo banner trên đầu trang
   */
  function showBanner(message, type = 'info') {
    const container = document.getElementById('alert-container');
    if (!container) return;

    const alertClass = type === 'success' ? 'alert-success' : type === 'warning' ? 'alert-warning' : 'alert-info';
    container.innerHTML = `
      <div class="alert-banner ${alertClass}">
        <div>${message}</div>
        <button style="background:none; border:none; cursor:pointer; font-size:16px; color:inherit; font-weight:bold;" onclick="this.parentElement.remove()">&times;</button>
      </div>
    `;
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function debounce(fn, delay) {
    let timer = null;
    return function (...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), delay);
    };
  }

})();
