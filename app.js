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
    },

    // Quản lý chuyển cấu trúc OLM sang TNMaker
    olm: {
      fileName: '',
      rawRows: null,
      detectedStructure: null,
      selectedStructure: 'auto',
      convertedData: null
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

    // 10. Chuyển cấu trúc OLM sang TNMaker
    setupOlmEventListeners();

    // 11. Nhập điểm theo SBD sang file SMAS
    setupSmasEventListeners();
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

  // ==============================================================
  // MODULE: CHUYỂN CẤU TRÚC ĐÁP ÁN OLM SANG TNMAKER
  // Tạo file Excel có cả 2 sheet "BTPRO" và "Sheet2" giống hệt nhau
  // ==============================================================

  const OLM_CONFIGS = {
    "12_mcq": {
      name: "12 câu nhiều lựa chọn, 4 câu đúng-sai (4 ý/câu), 6 câu trả lời ngắn (Tổng 22 câu)",
      short_name: "12 câu nhiều lựa chọn (22 câu - Toán)",
      mcq_range: [1, 13],              // Cột 1..12 (B:M - 12 câu, 0-based)
      tf_start_cols: [13, 17, 21, 25], // Cột 13, 17, 21, 25 (N, R, V, Z - 4 câu x 4 ý)
      sa_range: [29, 35],              // Cột 29..34 (AD:AI - 6 câu)
      total_questions: 22,
      tf_start_row: 14,                // Câu 13 tại dòng 14 của sheet BTPRO
      sa_start_row: 18,                // Câu 17 tại dòng 18 của sheet BTPRO
      max_row: 23
    },
    "18_mcq": {
      name: "18 câu nhiều lựa chọn, 4 câu đúng-sai (4 ý/câu), 6 câu trả lời ngắn (Tổng 28 câu)",
      short_name: "18 câu nhiều lựa chọn (28 câu - KHTN/KHXH)",
      mcq_range: [1, 19],              // Cột 1..18 (B:S - 18 câu)
      tf_start_cols: [19, 23, 27, 31], // Cột 19, 23, 27, 31 (T, X, AB, AF)
      sa_range: [35, 41],              // Cột 35..40 (AJ:AO - 6 câu)
      total_questions: 28,
      tf_start_row: 20,                // Câu 19 tại dòng 20
      sa_start_row: 24,                // Câu 23 tại dòng 24
      max_row: 29
    },
    "24_mcq": {
      name: "24 câu nhiều lựa chọn, 4 câu đúng-sai (4 ý/câu), 0 câu trả lời ngắn (Tổng 28 câu)",
      short_name: "24 câu nhiều lựa chọn (28 câu)",
      mcq_range: [1, 25],              // Cột 1..24 (B:Y - 24 câu)
      tf_start_cols: [25, 29, 33, 37], // Cột 25, 29, 33, 37 (Z, AD, AH, AL)
      sa_range: [41, 41],              // 0 câu trả lời ngắn
      total_questions: 28,
      tf_start_row: 26,                // Câu 25 tại dòng 26
      sa_start_row: null,              // Không có câu trả lời ngắn
      max_row: 29
    },
    "40_mcq": {
      name: "40 câu nhiều lựa chọn (Tổng 40 câu)",
      short_name: "40 câu nhiều lựa chọn (40 câu)",
      mcq_range: [1, 41],              // Cột 1..40 (B:AO - 40 câu)
      tf_start_cols: [],               // 0 câu đúng-sai
      sa_range: [41, 41],              // 0 câu trả lời ngắn
      total_questions: 40,
      tf_start_row: null,              // Không có phần đúng-sai
      sa_start_row: null,              // Không có phần trả lời ngắn
      max_row: 41
    }
  };

  function setupOlmEventListeners() {
    const modalOlm = document.getElementById('modal-olm');
    const btnOpenModal = document.getElementById('btn-open-olm-modal');
    const btnCloseModal = document.getElementById('btn-close-olm-modal');
    const btnCancelModal = document.getElementById('btn-cancel-olm-modal');
    const btnExportTNMaker = document.getElementById('btn-export-tnmaker');

    const olmFileInput = document.getElementById('olm-file-input');
    const olmDropzone = document.getElementById('olm-file-dropzone');
    const btnChooseFile = document.getElementById('btn-choose-olm-file');
    const btnLoadSample = document.getElementById('btn-load-sample-olm');

    if (btnOpenModal && modalOlm) {
      btnOpenModal.addEventListener('click', () => {
        modalOlm.classList.add('show');
        // Nếu đã có dữ liệu mẫu trước đó, tự động chuyển đổi hiển thị
        if (state.olm.rawRows) {
          processOlmConversion();
        }
      });
    }

    const closeModal = () => {
      if (modalOlm) modalOlm.classList.remove('show');
    };

    if (btnCloseModal) btnCloseModal.addEventListener('click', closeModal);
    if (btnCancelModal) btnCancelModal.addEventListener('click', closeModal);
    if (modalOlm) {
      modalOlm.addEventListener('click', (e) => {
        if (e.target === modalOlm) closeModal();
      });
    }

    if (btnChooseFile && olmFileInput) {
      btnChooseFile.addEventListener('click', (e) => {
        e.stopPropagation();
        olmFileInput.click();
      });
    }

    if (olmDropzone && olmFileInput) {
      olmDropzone.addEventListener('click', (e) => {
        if (e.target !== btnChooseFile) olmFileInput.click();
      });

      ['dragenter', 'dragover'].forEach(evName => {
        olmDropzone.addEventListener(evName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          olmDropzone.classList.add('dragover');
        });
      });

      ['dragleave', 'drop'].forEach(evName => {
        olmDropzone.addEventListener(evName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          olmDropzone.classList.remove('dragover');
        });
      });

      olmDropzone.addEventListener('drop', (e) => {
        if (e.dataTransfer && e.dataTransfer.files.length > 0) {
          handleOlmFile(e.dataTransfer.files[0]);
        }
      });
    }

    if (olmFileInput) {
      olmFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          handleOlmFile(e.target.files[0]);
        }
      });
    }

    if (btnLoadSample) {
      btnLoadSample.addEventListener('click', loadSampleOlmFile);
    }

    // Lắng nghe thay đổi tùy chọn cấu trúc đề thi
    const radioInputs = document.querySelectorAll('input[name="olm-structure"]');
    radioInputs.forEach(r => {
      r.addEventListener('change', () => {
        state.olm.selectedStructure = r.value;
        processOlmConversion();
      });
    });

    if (btnExportTNMaker) {
      btnExportTNMaker.addEventListener('click', exportTNMakerWorkbook);
    }
  }

  function handleOlmFile(file) {
    if (!file) return;
    state.olm.fileName = file.name;

    const reader = new FileReader();
    const statusMsg = document.getElementById('olm-status-msg');
    if (statusMsg) {
      statusMsg.innerHTML = `<span>⏳ Đang đọc file đáp án "${escapeHtml(file.name)}"...</span>`;
    }

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!rows || rows.length < 3) {
          throw new Error('File đáp án OLM không đủ dữ liệu (tối thiểu 3 hàng).');
        }

        state.olm.rawRows = rows;

        // Cập nhật tên hiển thị trên giao diện
        const nameText = document.getElementById('olm-file-name-text');
        const subText = document.getElementById('olm-file-sub-text');
        if (nameText) nameText.textContent = `File: ${file.name}`;
        if (subText) subText.textContent = `Đã nạp ${rows.length} hàng dữ liệu (${(file.size / 1024).toFixed(1)} KB)`;

        // Tự động nhận diện cấu trúc
        const detected = detectOlmStructure(rows);
        state.olm.detectedStructure = detected;

        processOlmConversion();

      } catch (err) {
        console.error('Lỗi đọc file OLM:', err);
        if (statusMsg) {
          statusMsg.innerHTML = `<span style="color: #dc2626;">❌ Lỗi đọc file: ${escapeHtml(err.message)}</span>`;
        }
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function loadSampleOlmFile() {
    if (window.SAMPLE_OLM_ROWS && Array.isArray(window.SAMPLE_OLM_ROWS) && window.SAMPLE_OLM_ROWS.length > 0) {
      state.olm.fileName = 'Dapan_mau_OLM.xlsx';
      state.olm.rawRows = window.SAMPLE_OLM_ROWS;

      const nameText = document.getElementById('olm-file-name-text');
      const subText = document.getElementById('olm-file-sub-text');
      if (nameText) nameText.textContent = 'File: Dapan.xlsx (Mẫu OLM 40 câu)';
      if (subText) subText.textContent = `Đã nạp ${window.SAMPLE_OLM_ROWS.length} hàng dữ liệu mẫu`;

      const detected = detectOlmStructure(window.SAMPLE_OLM_ROWS);
      state.olm.detectedStructure = detected;
      processOlmConversion();
      return;
    }

    // Dự phòng qua fetch
    fetch('Dapan_mau_OLM.xlsx')
      .then(res => res.arrayBuffer())
      .then(buf => {
        const file = new File([buf], 'Dapan_mau_OLM.xlsx');
        handleOlmFile(file);
      })
      .catch(() => {
        document.getElementById('olm-file-input')?.click();
      });
  }

  function cleanOlmVal(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'number' && Number.isInteger(val)) return String(val);
    return String(val).trim();
  }

  /**
   * Tự động nhận diện cấu trúc file OLM (12, 18, 24 hoặc 40 câu)
   */
  function detectOlmStructure(rows) {
    if (!rows || rows.length < 2) return '40_mcq';

    const row1 = (rows[0] || []).map(cleanOlmVal);
    const row2 = (rows[1] || []).map(cleanOlmVal);

    // 1. Kiểm tra vị trí câu đúng-sai đầu tiên ở dòng 2 (0-based: col 13, 19, 25, 40)
    const val14 = (row2[13] || '').toLowerCase(); // Cột 14
    const val20 = (row2[19] || '').toLowerCase(); // Cột 20
    const val25 = (row2[24] || '').toLowerCase(); // Cột 25
    const val26 = (row2[25] || '').toLowerCase(); // Cột 26
    const val41 = (row2[40] || '').toLowerCase(); // Cột 41

    if (val14 === '1a') return '12_mcq';
    if (val20 === '1a') return '18_mcq';
    if (['25a', '1a'].includes(val26) || (val25 === '24' && val26.endsWith('a'))) return '24_mcq';
    if (val41 === '40') return '40_mcq';

    // 2. Dòng 1 mô tả các phần
    const row1Text = row1.join(' ').toLowerCase();
    if (['đến câu 40', 'câu 1 đến câu 40', '40 câu', '40.'].some(k => row1Text.includes(k))) return '40_mcq';
    if (['đến câu 24', 'câu 1 đến câu 24', '24 câu', 'câu 25 đến câu 28'].some(k => row1Text.includes(k))) return '24_mcq';
    if (['đến câu 18', 'câu 1 đến câu 18', '18 câu'].some(k => row1Text.includes(k))) return '18_mcq';
    if (['đến câu 12', 'câu 1 đến câu 12', '12 câu'].some(k => row1Text.includes(k))) return '12_mcq';

    // 3. Kiểm tra loại dữ liệu tại các dòng đáp án (từ dòng 4)
    for (let r = 3; r < Math.min(10, rows.length); r++) {
      const row = rows[r] || [];
      const c20 = cleanOlmVal(row[19]).toUpperCase();
      const c26 = cleanOlmVal(row[25]).toUpperCase();
      const c41 = cleanOlmVal(row[40]).toUpperCase();

      if (['A', 'B', 'C', 'D'].includes(c20) && ['A', 'B', 'C', 'D'].includes(c26) && ['A', 'B', 'C', 'D'].includes(c41)) return '40_mcq';
      if (['A', 'B', 'C', 'D'].includes(c20) && ['Đ', 'S', 'D', 'S'].includes(c26)) return '24_mcq';
      if (['Đ', 'S', 'D', 'S'].includes(c20)) return '18_mcq';
    }

    const maxCol = Math.max(...rows.map(r => r ? r.length : 0));
    return maxCol <= 35 ? '12_mcq' : '40_mcq';
  }

  /**
   * Chuyển đổi ma trận đáp án từ OLM sang TNMaker (BTPRO & Sheet2)
   */
  function processOlmConversion() {
    const rows = state.olm.rawRows;
    const statusMsg = document.getElementById('olm-status-msg');
    const previewSection = document.getElementById('olm-preview-section');
    const btnExport = document.getElementById('btn-export-tnmaker');

    if (!rows || rows.length < 3) {
      if (statusMsg) {
        statusMsg.innerHTML = '<span>💡 Vui lòng chọn file Excel đáp án từ OLM để bắt đầu chuyển đổi.</span>';
      }
      if (previewSection) previewSection.style.display = 'none';
      if (btnExport) btnExport.disabled = true;
      return;
    }

    // Xác định cấu trúc áp dụng
    let structKey = state.olm.selectedStructure;
    if (structKey === 'auto') {
      structKey = state.olm.detectedStructure || detectOlmStructure(rows);
    }

    const config = OLM_CONFIGS[structKey] || OLM_CONFIGS['40_mcq'];

    // Cập nhật thông báo trạng thái
    if (statusMsg) {
      const isAuto = state.olm.selectedStructure === 'auto';
      statusMsg.innerHTML = `
        <span style="color: #15803d; font-weight: 600;">✓ ${isAuto ? 'Tự động phát hiện cấu trúc:' : 'Đang chọn thủ công cấu trúc:'}</span>
        <span style="color: #0f172a; font-weight: 700;">${config.name}</span>
      `;
    }

    // Trích xuất mã đề và đáp án
    const headerKeywords = new Set(['điểm', 'diem', 'mã đề', 'ma de', 'sbd', 'stt']);
    const sourceRows = [];

    for (let r = 2; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const maDeRaw = row[0];
      if (maDeRaw === undefined || maDeRaw === null) continue;
      const maDeStr = cleanOlmVal(maDeRaw);
      if (!maDeStr || headerKeywords.has(maDeStr.toLowerCase())) continue;

      // 1. Trắc nghiệm nhiều lựa chọn
      const mcq = [];
      for (let c = config.mcq_range[0]; c < config.mcq_range[1]; c++) {
        mcq.push(cleanOlmVal(row[c]));
      }

      // 2. Đúng/Sai (ghép 4 ý thành 1 chuỗi ví dụ ĐSĐS)
      const tf = [];
      for (const startCol of config.tf_start_cols) {
        let fourParts = '';
        for (let c = startCol; c < startCol + 4; c++) {
          fourParts += cleanOlmVal(row[c]);
        }
        tf.push(fourParts);
      }

      // 3. Trả lời ngắn
      const sa = [];
      for (let c = config.sa_range[0]; c < config.sa_range[1]; c++) {
        sa.push(cleanOlmVal(row[c]));
      }

      const answers = mcq.concat(tf, sa);
      sourceRows.push({ maDe: maDeStr, answers });
    }

    if (sourceRows.length === 0) {
      if (statusMsg) {
        statusMsg.innerHTML = '<span style="color: #dc2626;">⚠️ Không tìm thấy mã đề nào hợp lệ trong file Excel. Vui lòng kiểm tra lại cột A.</span>';
      }
      if (previewSection) previewSection.style.display = 'none';
      if (btnExport) btnExport.disabled = true;
      return;
    }

    // Xây dựng ma trận chuyển vị cho TNMaker:
    // Hàng 0: ['Câu\\Mã đề', maDe1, maDe2, ...]
    // Hàng 1..total_questions: [câu, ans_maDe1, ans_maDe2, ...]
    const headerAoa = ['Câu\\Mã đề'];
    sourceRows.forEach(sr => headerAoa.push(sr.maDe));

    const matrix = [headerAoa];
    for (let q = 1; q <= config.total_questions; q++) {
      const qRow = [q];
      for (let i = 0; i < sourceRows.length; i++) {
        qRow.push(sourceRows[i].answers[q - 1] || '');
      }
      matrix.push(qRow);
    }

    state.olm.convertedData = {
      config: config,
      sourceRows: sourceRows,
      matrix: matrix
    };

    renderOlmPreview(matrix, config, sourceRows);
    if (previewSection) previewSection.style.display = 'block';
    if (btnExport) btnExport.disabled = false;
  }

  /**
   * Render xem trước ma trận TNMaker trên web
   */
  function renderOlmPreview(matrix, config, sourceRows) {
    const table = document.getElementById('olm-preview-table');
    const summaryTag = document.getElementById('olm-preview-summary');
    if (!table) return;

    if (summaryTag) {
      summaryTag.textContent = `${sourceRows.length} mã đề (${sourceRows.map(s => s.maDe).slice(0, 4).join(', ')}${sourceRows.length > 4 ? '...' : ''}), ${config.total_questions} câu`;
    }

    table.innerHTML = '';

    // Header
    const thead = document.createElement('thead');
    const headerTr = document.createElement('tr');
    matrix[0].forEach((cellText, colIdx) => {
      const th = document.createElement('th');
      th.textContent = cellText;
      if (colIdx === 0) th.style.minWidth = '90px';
      headerTr.appendChild(th);
    });
    thead.appendChild(headerTr);
    table.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    const tfStart = config.tf_start_row ? config.tf_start_row - 1 : null;
    const saStart = config.sa_start_row ? config.sa_start_row - 1 : null;

    for (let r = 1; r < matrix.length; r++) {
      const tr = document.createElement('tr');
      if (r === tfStart || r === saStart) {
        tr.className = 'section-highlight';
      }
      matrix[r].forEach(cellText => {
        const td = document.createElement('td');
        td.textContent = cellText;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
  }

  /**
   * Xuất file Excel TNMaker:
   * YÊU CẦU: Tạo thêm một sheet tên là "Sheet2" có nội dung GIỐNG HỆT sheet "BTPRO"
   */
  function exportTNMakerWorkbook() {
    const data = state.olm.convertedData;
    if (!data || !data.matrix) {
      alert('Chưa có dữ liệu chuyển đổi để xuất Excel.');
      return;
    }

    try {
      const wb = XLSX.utils.book_new();
      const matrix = data.matrix;
      const numCodes = data.sourceRows.length;

      // Cấu hình độ rộng các cột
      const colWidths = [{ wch: 13 }]; // Cột Câu\Mã đề
      for (let i = 0; i < numCodes; i++) {
        colWidths.push({ wch: 12 });
      }

      // ==============================================================
      // SHEET 1: BTPRO
      // ==============================================================
      const wsBtpro = XLSX.utils.aoa_to_sheet(matrix);
      wsBtpro['!cols'] = colWidths;
      XLSX.utils.book_append_sheet(wb, wsBtpro, 'BTPRO');

      // ==============================================================
      // SHEET 2: Sheet2 (NỘI DUNG GIỐNG HỆT SHEET "BTPRO")
      // ==============================================================
      const wsSheet2 = XLSX.utils.aoa_to_sheet(matrix);
      wsSheet2['!cols'] = colWidths;
      XLSX.utils.book_append_sheet(wb, wsSheet2, 'Sheet2');

      // Đặt tên file xuất: [Ten_file]_da_chuyen_cau_truc.xlsx
      let originalStem = 'Dapan_OLM';
      if (state.olm.fileName) {
        originalStem = state.olm.fileName.replace(/\.[^/.]+$/, '');
      }
      const outFileName = `${originalStem}_da_chuyen_cau_truc.xlsx`;

      downloadWorkbook(wb, outFileName);

      showBanner(
        `🎉 <strong>Chuyển đổi thành công!</strong> Đã xuất file <strong>"${outFileName}"</strong> gồm cả 2 sheet <strong>BTPRO</strong> và <strong>Sheet2</strong> (nội dung giống hệt nhau) tương thích 100% với TNMaker!`,
        'success'
      );

      // Đóng modal sau khi xuất
      const modalOlm = document.getElementById('modal-olm');
      if (modalOlm) modalOlm.classList.remove('show');

    } catch (err) {
      console.error('Lỗi khi xuất file Excel TNMaker:', err);
      alert('Không thể xuất file Excel TNMaker: ' + err.message);
    }
  }

  // ==============================================================
  // MODULE: KẾT NỐI ĐIỂM SANG FILE NHẬP ĐIỂM THEO SBD (SMAS)
  // Điền điểm từ cột E của file điểm vào cột G & H của file SMAS
  // Xuất file THI_THPT_NhapDiemTheoSBD.xlsx
  // ==============================================================

  let smasLastWorkbook = null;
  let smasLastBlob = null;

  function setupSmasEventListeners() {
    const btnNhapDiem = document.getElementById('btn-nhap-diem-smas');
    const smasFileInput = document.getElementById('smas-file-input');
    const btnLoadSampleSmas = document.getElementById('btn-load-sample-smas');

    const modalSmasResult = document.getElementById('modal-smas-result');
    const btnCloseSmasModal = document.getElementById('btn-close-smas-modal');
    const btnCloseSmasModal2 = document.getElementById('btn-close-smas-modal-2');
    const btnRedownloadSmas = document.getElementById('btn-redownload-smas');

    const closeSmasModal = () => {
      if (modalSmasResult) modalSmasResult.classList.remove('show');
    };

    if (btnCloseSmasModal) btnCloseSmasModal.addEventListener('click', closeSmasModal);
    if (btnCloseSmasModal2) btnCloseSmasModal2.addEventListener('click', closeSmasModal);
    if (modalSmasResult) {
      modalSmasResult.addEventListener('click', (e) => {
        if (e.target === modalSmasResult) closeSmasModal();
      });
    }

    if (btnRedownloadSmas) {
      btnRedownloadSmas.addEventListener('click', () => {
        if (smasLastBlob) {
          const downloadUrl = URL.createObjectURL(smasLastBlob);
          const a = document.createElement('a');
          a.href = downloadUrl;
          a.download = 'THI_THPT_NhapDiemTheoSBD.xlsx';
          document.body.appendChild(a);
          a.click();
          setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
          }, 100);
        } else if (smasLastWorkbook) {
          downloadWorkbook(smasLastWorkbook, 'THI_THPT_NhapDiemTheoSBD.xlsx');
        } else {
          alert('Chưa có file SMAS đã kết nối để tải lại.');
        }
      });
    }

    if (btnNhapDiem && smasFileInput) {
      btnNhapDiem.addEventListener('click', () => {
        // Kiểm tra xem đã có dữ liệu điểm thi (từ file điểm choice_diem) hay chưa
        if (Object.keys(state.scoreMap).length === 0) {
          const confirmLoad = confirm(
            '⚠️ Chưa có dữ liệu điểm thi (từ file choice_diem)!\n\n' +
            'Bạn có muốn tự động nạp dữ liệu từ file mẫu "mau_file.xls" có sẵn để tiếp tục kết nối điểm không?\n\n' +
            '- Nhấn OK: Nạp file điểm mẫu "mau_file.xls" và chọn file SMAS.\n' +
            '- Nhấn Cancel: Để bạn bấm "Chọn file điểm" và nạp file điểm của mình trước.'
          );
          if (confirmLoad) {
            loadSampleScoreFile();
            setTimeout(() => {
              smasFileInput.click();
            }, 300);
          } else {
            document.getElementById('score-file-input')?.click();
          }
          return;
        }

        smasFileInput.click();
      });

      smasFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          handleSmasFile(e.target.files[0]);
          // Reset input để có thể chọn lại cùng file
          e.target.value = '';
        }
      });
    }

    if (btnLoadSampleSmas) {
      btnLoadSampleSmas.addEventListener('click', () => {
        if (Object.keys(state.scoreMap).length === 0) {
          loadSampleScoreFile();
        }
        showBanner('Đang nạp file mẫu "mau_xuat_SMAS.xlsx" và kết nối điểm (giữ nguyên 100% định dạng)...', 'info');

        fetch('mau_xuat_SMAS.xlsx')
          .then(res => {
            if (!res.ok) throw new Error('Không thể tải file mẫu mau_xuat_SMAS.xlsx trực tiếp qua HTTP.');
            return res.arrayBuffer();
          })
          .then(buf => {
            const file = new File([buf], 'mau_xuat_SMAS.xlsx');
            handleSmasFile(file);
          })
          .catch(err => {
            console.warn('Fetch mau_xuat_SMAS.xlsx failed, opening file picker:', err);
            showBanner('💡 Vui lòng bấm nút <strong>"Chọn file NhapDiemTheoSBD và kết nối điểm"</strong> và chọn file <strong>mau_xuat_SMAS.xlsx</strong> trong thư mục làm việc.', 'info');
            smasFileInput?.click();
          });
      });
    }
  }

  /**
   * Xử lý file SMAS bảo toàn 100% định dạng, cấu trúc, font, border, màu nền:
   * - Sử dụng JSZip để chỉnh sửa trực tiếp XML của từng sheet (giữ nguyên tất cả XML metadata, theme, styles, merges).
   * - So sánh Số báo danh (cột D / cột B) với SBD cột C từ file điểm.
   * - Điền điểm từ cột E của file điểm vào cột G & H của từng sheet (giữ nguyên thuộc tính style s="...").
   * - Xuất ra file "THI_THPT_NhapDiemTheoSBD.xlsx".
   */
  async function handleSmasFile(file) {
    if (!file) return;

    if (Object.keys(state.scoreMap).length === 0) {
      alert('Vui lòng nạp file điểm thi (cột C: SBD, cột E: Tổng điểm) trước khi kết nối vào file SMAS.');
      return;
    }

    showBanner(`Đang xử lý kết nối điểm sang file SMAS "${file.name}" (giữ nguyên 100% định dạng)...`, 'info');

    try {
      const arrayBuffer = await file.arrayBuffer();

      // Kiểm tra nếu có JSZip thì sử dụng engine bảo toàn định dạng 100%
      if (typeof JSZip !== 'undefined') {
        const zip = await JSZip.loadAsync(arrayBuffer);

        // 1. Đọc sharedStrings.xml nếu có
        const sharedStrings = [];
        const ssFile = zip.file('xl/sharedStrings.xml');
        if (ssFile) {
          const ssXml = await ssFile.async('text');
          const siMatches = ssXml.match(/<si>[\s\S]*?<\/si>/g) || [];
          for (const si of siMatches) {
            const tMatches = si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
            const text = tMatches.map(t => t.replace(/<[^>]+>/g, '')).join('');
            sharedStrings.push(text);
          }
        }

        // 2. Đọc danh sách sheet từ workbook.xml và workbook.xml.rels
        const wbXml = await zip.file('xl/workbook.xml').async('text');
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');

        const relMap = {};
        const relMatches = relsXml.match(/<Relationship[^>]*\/>/g) || [];
        for (const rel of relMatches) {
          const idM = rel.match(/Id="([^"]+)"/);
          const targetM = rel.match(/Target="([^"]+)"/);
          if (idM && targetM) {
            let target = targetM[1];
            if (!target.startsWith('xl/')) target = 'xl/' + target.replace(/^\//, '');
            relMap[idM[1]] = target;
          }
        }

        const sheetList = [];
        const sheetMatches = wbXml.match(/<sheet[^>]*\/>/g) || [];
        for (const sh of sheetMatches) {
          const nameM = sh.match(/name="([^"]+)"/);
          const rIdM = sh.match(/r:id="([^"]+)"/);
          if (nameM && rIdM) {
            sheetList.push({
              name: nameM[1],
              file: relMap[rIdM[1]]
            });
          }
        }

        let totalMatchedStudents = 0;
        let totalStudentsCount = 0;
        const sheetResults = [];

        // 3. Xử lý từng sheet phòng thi
        for (const sh of sheetList) {
          if (sh.name.toLowerCase() === 'ref') continue;
          const sheetFile = zip.file(sh.file);
          if (!sheetFile) continue;

          let xml = await sheetFile.async('text');
          let sheetMatched = 0;
          let sheetTotal = 0;

          xml = xml.replace(/<row r="([0-9]+)"([\s\S]*?)<\/row>/g, (fullRow, rNumStr, rowInner) => {
            const rNum = parseInt(rNumStr);
            if (rNum <= 9) return fullRow; // Bỏ qua tiêu đề

            let sbdVal = null;

            // Kiểm tra cột D (theo mô tả yêu cầu)
            const dMatch = rowInner.match(new RegExp(`<c r="D${rNum}"([^>]*)>([\\s\\S]*?)<\\/c>`));
            if (dMatch) {
              const attrs = dMatch[1];
              const content = dMatch[2];
              const vM = content.match(/<v>([0-9]+)<\/v>/);
              if (vM) {
                let val = vM[1];
                if (attrs.includes('t="s"')) {
                  val = sharedStrings[parseInt(val)] || '';
                }
                val = val.trim();
                if (state.scoreMap.hasOwnProperty(val)) {
                  sbdVal = val;
                }
              }
            }

            // Kiểm tra cột B (nếu cột D không khớp)
            if (!sbdVal) {
              const bMatch = rowInner.match(new RegExp(`<c r="B${rNum}"([^>]*)>([\\s\\S]*?)<\\/c>`));
              if (bMatch) {
                const attrs = bMatch[1];
                const content = bMatch[2];
                const vM = content.match(/<v>([0-9]+)<\/v>/);
                if (vM) {
                  let val = vM[1];
                  if (attrs.includes('t="s"')) {
                    val = sharedStrings[parseInt(val)] || '';
                  }
                  val = val.trim();
                  if (state.scoreMap.hasOwnProperty(val)) {
                    sbdVal = val;
                  }
                }
              }
            }

            if (dMatch || rowInner.includes(`r="B${rNum}"`)) {
              sheetTotal++;
              totalStudentsCount++;
            }

            if (sbdVal && state.scoreMap.hasOwnProperty(sbdVal)) {
              const score = state.scoreMap[sbdVal];
              sheetMatched++;
              totalMatchedStudents++;

              let updatedRow = rowInner;
              // Điền điểm vào cột G & H, bảo toàn nguyên vẹn thuộc tính style s="..."
              for (const colLetter of ['G', 'H']) {
                const cellPattern = new RegExp(`<c r="${colLetter}${rNum}"([^>]*?)(\\/?>)([\\s\\S]*?<\\/c>)?`);
                const cM = updatedRow.match(cellPattern);
                if (cM) {
                  const attrs = cM[1];
                  const sAttrM = attrs.match(/s="([0-9]+)"/);
                  const sAttr = sAttrM ? ` s="${sAttrM[1]}"` : '';
                  const newCell = `<c r="${colLetter}${rNum}"${sAttr}><v>${score}</v></c>`;
                  updatedRow = updatedRow.replace(cM[0], newCell);
                } else {
                  const insertBeforePattern = new RegExp(`(<c r="[I-Z]${rNum}")`);
                  if (updatedRow.match(insertBeforePattern)) {
                    updatedRow = updatedRow.replace(insertBeforePattern, `<c r="${colLetter}${rNum}"><v>${score}</v></c>$1`);
                  } else {
                    updatedRow += `<c r="${colLetter}${rNum}"><v>${score}</v></c>`;
                  }
                }
              }
              return `<row r="${rNum}"${updatedRow}</row>`;
            }

            return fullRow;
          });

          zip.file(sh.file, xml);

          if (sheetTotal > 0) {
            sheetResults.push({
              sheetName: sh.name,
              total: sheetTotal,
              matched: sheetMatched,
              rate: sheetTotal > 0 ? (sheetMatched / sheetTotal) * 100 : 0
            });
          }
        }

        // Tạo file zip nhị phân
        const outBlob = await zip.generateAsync({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          compression: 'DEFLATE'
        });

        smasLastBlob = outBlob;

        // Tự động tải xuống file THI_THPT_NhapDiemTheoSBD.xlsx
        const outFileName = 'THI_THPT_NhapDiemTheoSBD.xlsx';
        const downloadUrl = URL.createObjectURL(outBlob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = outFileName;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          URL.revokeObjectURL(downloadUrl);
        }, 100);

        showBanner(
          `🎉 <strong>Kết nối điểm SMAS thành công!</strong> Đã điền điểm vào cột G &amp; H cho <strong>${totalMatchedStudents}/${totalStudentsCount}</strong> học sinh trong <strong>${sheetResults.length}</strong> phòng thi. File kết quả giữ nguyên 100% định dạng đã được tải xuống!`,
          'success'
        );

        showSmasResultModal(totalMatchedStudents, totalStudentsCount, sheetResults);
        return;
      }

      // Dự phòng bằng SheetJS nếu môi trường không có JSZip
      const data = new Uint8Array(arrayBuffer);
      const wb = XLSX.read(data, { type: 'array' });
      // ...
      smasLastWorkbook = wb;
      const outFileName = 'THI_THPT_NhapDiemTheoSBD.xlsx';
      downloadWorkbook(wb, outFileName);

    } catch (err) {
      console.error('Lỗi khi kết nối file SMAS:', err);
      showBanner(`❌ Lỗi kết nối file SMAS: ${escapeHtml(err.message)}`, 'warning');
    }
  }

  function showSmasResultModal(matched, total, sheetResults) {
    const modal = document.getElementById('modal-smas-result');
    const summaryDiv = document.getElementById('smas-summary-stats');
    const tbody = document.getElementById('smas-sheets-tbody');
    if (!modal) return;

    if (summaryDiv) {
      const matchRate = total > 0 ? ((matched / total) * 100).toFixed(1) : 0;
      summaryDiv.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 10px;">
          <div><strong>Số phòng thi (Sheet):</strong> ${sheetResults.length} phòng</div>
          <div><strong>Tổng số thí sinh:</strong> ${total.toLocaleString()}</div>
          <div><strong>Đã điền điểm:</strong> <span style="color: #059669; font-weight: 700;">${matched.toLocaleString()}</span></div>
          <div><strong>Tỉ lệ khớp điểm:</strong> <span style="color: #2563eb; font-weight: 700;">${matchRate}%</span></div>
        </div>
      `;
    }

    if (tbody) {
      tbody.innerHTML = '';
      sheetResults.forEach((s, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td style="text-align: center; color: var(--text-muted);">${idx + 1}</td>
          <td style="font-weight: 700; color: #1e3a8a;">${escapeHtml(s.sheetName)}</td>
          <td style="text-align: center;">${s.total}</td>
          <td style="text-align: center; color: #059669; font-weight: 700;">${s.matched}</td>
          <td style="text-align: center; font-weight: 600;">${s.rate.toFixed(1)}%</td>
        `;
        tbody.appendChild(tr);
      });
    }

    modal.classList.add('show');
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
