/* ==========================================================================
   Lao Student Data Collection System - Application Logic (Trends & Comprehensive KPI)
   ========================================================================== */

// 🟢 Google Apps Script Web App URL ສຳລັບບັນທຶກຂໍ້ມູນລົງ Google Sheet ໂດຍອັດໂນມັດ
const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwu0jUyKngnRIHHDGkKfQScouZT9i96hoFETg3hU1nVbzMSpQNtpa81etSHsDaN05yZEg/exec"; 

document.addEventListener('DOMContentLoaded', () => {
  // State variables
  let records = [];
  let districtChart = null;
  let genderChart = null;
  let monthlyTrendChart = null;
  let movementChart = null;
  let parsedExcelData = null;

  // DOM Elements (Safe Retrieval)
  const themeToggleBtn = document.getElementById('themeToggleBtn');
  const studentDataForm = document.getElementById('studentDataForm');
  const recordIdInput = document.getElementById('recordId');
  const formModeBadge = document.getElementById('formModeBadge');

  const districtSelect = document.getElementById('districtSelect');
  const schoolSelect = document.getElementById('schoolSelect');
  const collectorName = document.getElementById('collectorName');
  const educationLevelSelect = document.getElementById('educationLevelSelect');
  const gradeLevelSelect = document.getElementById('gradeLevelSelect');
  
  // Lao Month & Academic Year Selector elements
  const entryLaoMonthSelect = document.getElementById('entryLaoMonthSelect');
  const entryAcademicYearSelect = document.getElementById('entryAcademicYearSelect');

  const passedRegisteredTotal = document.getElementById('passedRegisteredTotal');
  const passedRegisteredFemale = document.getElementById('passedRegisteredFemale');
  const transferInTotal = document.getElementById('transferInTotal');
  const transferInFemale = document.getElementById('transferInFemale');
  const transferOutTotal = document.getElementById('transferOutTotal');
  const transferOutFemale = document.getElementById('transferOutFemale');
  const transferOutReason = document.getElementById('transferOutReason');
  const dropoutTotal = document.getElementById('dropoutTotal');
  const dropoutFemale = document.getElementById('dropoutFemale');
  const dropoutReason = document.getElementById('dropoutReason');
  const repeaterTotal = document.getElementById('repeaterTotal');
  const repeaterFemale = document.getElementById('repeaterFemale');
  const actualAttendingTotal = document.getElementById('actualAttendingTotal');
  const actualAttendingFemale = document.getElementById('actualAttendingFemale');

  const recalculateBtn = document.getElementById('recalculateBtn');
  const resetFormBtn = document.getElementById('resetFormBtn');
  const submitFormBtn = document.getElementById('submitFormBtn');

  // Search & Filter Elements
  const tableBody = document.getElementById('tableBody');
  const tableRecordCount = document.getElementById('tableRecordCount');
  const searchInput = document.getElementById('searchInput');
  const filterDistrict = document.getElementById('filterDistrict');
  const filterLevel = document.getElementById('filterLevel');
  const filterGrade = document.getElementById('filterGrade');
  const filterMonth = document.getElementById('filterMonth');
  const filterAcademicYear = document.getElementById('filterAcademicYear');
  const resetFiltersBtn = document.getElementById('resetFiltersBtn');

  const exportExcelBtn = document.getElementById('exportExcelBtn');
  const printReportBtn = document.getElementById('printReportBtn');

  // Excel Modal Elements
  const importExcelModalBtn = document.getElementById('importExcelModalBtn');
  const excelImportModal = document.getElementById('excelImportModal');
  const closeExcelModalBtn = document.getElementById('closeExcelModalBtn');
  const cancelExcelModalBtn = document.getElementById('cancelExcelModalBtn');
  const excelDropzone = document.getElementById('excelDropzone');
  const excelFileInput = document.getElementById('excelFileInput');
  const browseExcelFileBtn = document.getElementById('browseExcelFileBtn');
  const confirmExcelImportBtn = document.getElementById('confirmExcelImportBtn');
  const downloadExcelTemplateBtn = document.getElementById('downloadExcelTemplateBtn');

  const excelPreviewContainer = document.getElementById('excelPreviewContainer');
  const excelPreviewHead = document.getElementById('excelPreviewHead');
  const excelPreviewBody = document.getElementById('excelPreviewBody');
  const excelRowCount = document.getElementById('excelRowCount');
  const excelFileNameBadge = document.getElementById('excelFileNameBadge');

  // Stats KPI Elements
  const statRegisteredTotal = document.getElementById('statRegisteredTotal');
  const statRegisteredFemale = document.getElementById('statRegisteredFemale');
  const statRegisteredPercent = document.getElementById('statRegisteredPercent');
  
  const statTransferInTotal = document.getElementById('statTransferInTotal');
  const statTransferInFemale = document.getElementById('statTransferInFemale');
  const statTransferInPercent = document.getElementById('statTransferInPercent');

  const statTransferOutTotal = document.getElementById('statTransferOutTotal');
  const statTransferOutFemale = document.getElementById('statTransferOutFemale');
  const statTransferOutPercent = document.getElementById('statTransferOutPercent');

  const statTotalDropout = document.getElementById('statTotalDropout');
  const statDropoutFemale = document.getElementById('statDropoutFemale');
  const statDropoutPercent = document.getElementById('statDropoutPercent');

  const statRepeaterTotal = document.getElementById('statRepeaterTotal');
  const statRepeaterFemale = document.getElementById('statRepeaterFemale');
  const statRepeaterPercent = document.getElementById('statRepeaterPercent');

  const statTotalStudents = document.getElementById('statTotalStudents');
  const statFemaleStudents = document.getElementById('statFemaleStudents');
  const statFemalePercent = document.getElementById('statFemalePercent');

  // Modal elements
  const detailModal = document.getElementById('detailModal');
  const closeModalBtn = document.getElementById('closeModalBtn');
  const modalContent = document.getElementById('modalContent');
  const modalTitle = document.getElementById('modalTitle');

  // Initialize Application
  initApp();

  function initApp() {
    loadRecordsFromStorage();
    setupLaoMonthDropdowns();
    setupSearchFilterDropdowns();
    populateDistrictDropdowns();
    setupEventListeners();
    renderApp();
  }

  // Populate Lao Month and Academic Year (ສົກຮຽນ) Selectors
  function setupLaoMonthDropdowns() {
    if (!entryLaoMonthSelect || !entryAcademicYearSelect) return;

    entryLaoMonthSelect.innerHTML = '<option value="">-- ເລືອກເດືອນ --</option>';
    LAO_MONTHS.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.value;
      opt.textContent = m.name;
      entryLaoMonthSelect.appendChild(opt);
    });

    entryAcademicYearSelect.innerHTML = '';
    ACADEMIC_YEARS.forEach(ay => {
      const opt = document.createElement('option');
      opt.value = ay;
      opt.textContent = `ສົກຮຽນ ${ay}`;
      if (ay === "2026-2027") opt.selected = true;
      entryAcademicYearSelect.appendChild(opt);
    });

    const currentMonth = String(new Date().getMonth() + 1).padStart(2, '0');
    entryLaoMonthSelect.value = currentMonth;
  }

  // Populate Filter Month & Academic Year Dropdowns
  function setupSearchFilterDropdowns() {
    if (filterMonth) {
      filterMonth.innerHTML = '<option value="">ທຸກໆເດືອນ</option>';
      LAO_MONTHS.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.value;
        opt.textContent = m.name;
        filterMonth.appendChild(opt);
      });
    }

    if (filterAcademicYear) {
      filterAcademicYear.innerHTML = '<option value="">ທຸກໆສົກຮຽນ</option>';
      ACADEMIC_YEARS.forEach(ay => {
        const opt = document.createElement('option');
        opt.value = ay;
        opt.textContent = `ສົກຮຽນ ${ay}`;
        filterAcademicYear.appendChild(opt);
      });
    }
  }

  // Convert raw entry_month string into formatted Lao Month + Academic Year
  function getLaoMonthText(rawMonthStr) {
    if (!rawMonthStr) return "-";
    
    let monthVal = "";
    let yearVal = "";

    if (rawMonthStr.includes(':')) {
      const parts = rawMonthStr.split(':');
      yearVal = parts[0];
      monthVal = parts[1];
    } else if (rawMonthStr.includes('-')) {
      const parts = rawMonthStr.split('-');
      if (parts.length === 3) {
        yearVal = `${parts[0]}-${parts[1]}`;
        monthVal = parts[2];
      } else {
        yearVal = parts[0];
        monthVal = parts[1];
      }
    } else {
      monthVal = rawMonthStr;
    }

    const mObj = LAO_MONTHS.find(m => m.value === monthVal || m.name.includes(monthVal) || m.short.includes(monthVal));
    const monthName = mObj ? `ເດືອນ${mObj.short}` : `ເດືອນ ${monthVal}`;

    if (yearVal) {
      return yearVal.includes('-') ? `${monthName} (${yearVal})` : `${monthName} ${yearVal}`;
    }
    return monthName;
  }

  function getRecordAcademicYear(rawMonthStr) {
    if (!rawMonthStr) return "";
    if (rawMonthStr.includes(':')) return rawMonthStr.split(':')[0];
    if (rawMonthStr.includes('-')) {
      const parts = rawMonthStr.split('-');
      if (parts.length === 3) return `${parts[0]}-${parts[1]}`;
      return parts[0];
    }
    return "";
  }

  function getRecordMonthCode(rawMonthStr) {
    if (!rawMonthStr) return "";
    if (rawMonthStr.includes(':')) return rawMonthStr.split(':')[1];
    if (rawMonthStr.includes('-')) {
      const parts = rawMonthStr.split('-');
      if (parts.length === 3) return parts[2];
      return parts[1];
    }
    return rawMonthStr;
  }

  function loadRecordsFromStorage() {
    const saved = localStorage.getItem('lao_school_records');
    if (saved) {
      try {
        records = JSON.parse(saved);
      } catch (e) {
        records = [];
      }
    } else {
      records = [];
    }
  }

  function saveRecordsToStorage() {
    localStorage.setItem('lao_school_records', JSON.stringify(records));
  }

  // Populate District Dropdowns from Pre-loaded Presets + Imported Excel
  function populateDistrictDropdowns() {
    const importedDistricts = JSON.parse(localStorage.getItem('lao_excel_custom_districts') || '{}');
    const allDistrictsSet = new Set([
      ...Object.keys(DEFAULT_LAO_DISTRICTS),
      ...Object.keys(importedDistricts)
    ]);

    if (districtSelect) {
      districtSelect.innerHTML = '<option value="">-- ເລືອກເມືອງ --</option>';
      allDistrictsSet.forEach(distName => {
        const opt = document.createElement('option');
        opt.value = distName;
        opt.textContent = distName;
        districtSelect.appendChild(opt);
      });
    }

    if (filterDistrict) {
      filterDistrict.innerHTML = '<option value="">ທຸກໆເມືອງ (7 ເມືອງ)</option>';
      allDistrictsSet.forEach(distName => {
        const filterOpt = document.createElement('option');
        filterOpt.value = distName;
        filterOpt.textContent = distName;
        filterDistrict.appendChild(filterOpt);
      });
    }
  }

  // Populate School Dropdown based on selected district (Pre-loaded + Excel)
  function populateSchools(selectedDistrict) {
    if (!schoolSelect) return;

    schoolSelect.innerHTML = '<option value="">-- ເລືອກໂຮງຮຽນ --</option>';
    
    if (!selectedDistrict) {
      schoolSelect.disabled = true;
      schoolSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກເມືອງກ່ອນ --</option>';
      return;
    }

    let schoolsList = [];

    if (DEFAULT_LAO_DISTRICTS[selectedDistrict]) {
      DEFAULT_LAO_DISTRICTS[selectedDistrict].forEach(s => {
        if (!schoolsList.includes(s)) schoolsList.push(s);
      });
    }

    const importedDistricts = JSON.parse(localStorage.getItem('lao_excel_custom_districts') || '{}');
    if (importedDistricts[selectedDistrict]) {
      importedDistricts[selectedDistrict].forEach(s => {
        if (!schoolsList.includes(s)) schoolsList.push(s);
      });
    }

    if (schoolsList.length > 0) {
      schoolsList.forEach(sch => {
        const opt = document.createElement('option');
        opt.value = sch;
        opt.textContent = sch;
        schoolSelect.appendChild(opt);
      });
      schoolSelect.disabled = false;
    } else {
      schoolSelect.disabled = true;
      schoolSelect.innerHTML = '<option value="">-- ບໍ່ມີໂຮງຮຽນໃນເມືອງນີ້ --</option>';
    }
  }

  // Auto-filter Education Level options based on selected School Prefix (ປ.ຖ / ມ.ຕ / ມ.ສ)
  function handleSchoolSelectChange(schName) {
    if (!schName) {
      resetEducationLevelOptions();
      return;
    }

    const lowerSch = schName.toLowerCase();

    if (lowerSch.includes('ປ.ຖ') || lowerSch.includes('ປະຖົມ')) {
      setEducationLevelOptions(["ປະຖົມສຶກສາ"], "ປະຖົມສຶກສາ");
    } else if (lowerSch.includes('ມ.ຕ')) {
      setEducationLevelOptions(["ມັດທະຍົມຕອນຕົ້ນ"], "ມັດທະຍົມຕອນຕົ້ນ");
    } else if (lowerSch.includes('ມ.ສ') || lowerSch.includes('ມັດທະຍົມສົມບູນ')) {
      setEducationLevelOptions(["ມັດທະຍົມສົມບູນ"], "ມັດທະຍົມສົມບູນ");
    } else {
      resetEducationLevelOptions();
    }
  }

  function setEducationLevelOptions(allowedLevels, defaultSelected = "") {
    if (!educationLevelSelect) return;
    educationLevelSelect.innerHTML = '<option value="">-- ເລືອກຊັ້ນຮຽນ --</option>';
    allowedLevels.forEach(lvl => {
      const opt = document.createElement('option');
      opt.value = lvl;
      opt.textContent = lvl;
      if (lvl === defaultSelected) opt.selected = true;
      educationLevelSelect.appendChild(opt);
    });

    if (defaultSelected) {
      populateGrades(defaultSelected);
    }
  }

  function resetEducationLevelOptions() {
    if (!educationLevelSelect || !gradeLevelSelect) return;

    educationLevelSelect.innerHTML = `
      <option value="">-- ເລືອກຊັ້ນຮຽນ --</option>
      <option value="ປະຖົມສຶກສາ">ປະຖົມສຶກສາ</option>
      <option value="ມັດທະຍົມຕອນຕົ້ນ">ມັດທະຍົມຕອນຕົ້ນ</option>
      <option value="ມັດທະຍົມສົມບູນ">ມັດທະຍົມສົມບູນ</option>
    `;
    gradeLevelSelect.disabled = true;
    gradeLevelSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກຊັ້ນຮຽນກ່ອນ --</option>';
  }

  // Populate Grade Levels based on selected Education Level (Cascading)
  function populateGrades(selectedLevel) {
    if (!gradeLevelSelect) return;
    gradeLevelSelect.innerHTML = '<option value="">-- ເລືອກຂັ້ນຮຽນ --</option>';
    
    if (!selectedLevel) {
      gradeLevelSelect.disabled = true;
      gradeLevelSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກຊັ້ນຮຽນກ່ອນ --</option>';
      return;
    }

    let gradeList = EDUCATION_LEVELS[selectedLevel] || [];

    const customGrades = JSON.parse(localStorage.getItem('lao_excel_custom_grades') || '{}');
    if (customGrades[selectedLevel]) {
      customGrades[selectedLevel].forEach(cg => {
        if (!gradeList.includes(cg)) gradeList.push(cg);
      });
    }

    gradeList.forEach(grade => {
      const opt = document.createElement('option');
      opt.value = grade;
      opt.textContent = grade;
      gradeLevelSelect.appendChild(opt);
    });

    gradeLevelSelect.disabled = false;
  }

  // Setup Event Listeners
  function setupEventListeners() {
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);

    if (districtSelect) {
      districtSelect.addEventListener('change', (e) => {
        populateSchools(e.target.value);
      });
    }

    if (schoolSelect) {
      schoolSelect.addEventListener('change', (e) => {
        handleSchoolSelectChange(e.target.value);
      });
    }

    if (educationLevelSelect) {
      educationLevelSelect.addEventListener('change', (e) => {
        populateGrades(e.target.value);
      });
    }

    document.querySelectorAll('.calc-trigger').forEach(input => {
      input.addEventListener('input', calculateActualAttending);
    });
    
    if (recalculateBtn) recalculateBtn.addEventListener('click', calculateActualAttending);

    setupValidationWatchers();

    if (studentDataForm) studentDataForm.addEventListener('submit', handleFormSubmit);
    if (resetFormBtn) resetFormBtn.addEventListener('click', resetForm);

    if (searchInput) searchInput.addEventListener('input', renderTable);
    if (filterDistrict) filterDistrict.addEventListener('change', renderTable);
    if (filterLevel) filterLevel.addEventListener('change', renderTable);
    if (filterGrade) filterGrade.addEventListener('change', renderTable);
    if (filterMonth) filterMonth.addEventListener('change', renderTable);
    if (filterAcademicYear) filterAcademicYear.addEventListener('change', renderTable);

    if (resetFiltersBtn) resetFiltersBtn.addEventListener('click', resetSearchFilters);

    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
    if (printReportBtn) printReportBtn.addEventListener('click', () => window.print());

    if (importExcelModalBtn) importExcelModalBtn.addEventListener('click', () => excelImportModal.classList.add('active'));
    if (closeExcelModalBtn) closeExcelModalBtn.addEventListener('click', closeExcelModal);
    if (cancelExcelModalBtn) cancelExcelModalBtn.addEventListener('click', closeExcelModal);

    if (browseExcelFileBtn) browseExcelFileBtn.addEventListener('click', () => excelFileInput.click());
    if (excelDropzone) {
      excelDropzone.addEventListener('click', (e) => {
        if (e.target !== browseExcelFileBtn) excelFileInput.click();
      });

      excelDropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        excelDropzone.classList.add('dragover');
      });

      excelDropzone.addEventListener('dragleave', () => excelDropzone.classList.remove('dragover'));
      excelDropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        excelDropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
          handleExcelFileSelect(e.dataTransfer.files[0]);
        }
      });
    }

    if (excelFileInput) {
      excelFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          handleExcelFileSelect(e.target.files[0]);
        }
      });
    }

    if (confirmExcelImportBtn) confirmExcelImportBtn.addEventListener('click', processExcelImport);
    if (downloadExcelTemplateBtn) downloadExcelTemplateBtn.addEventListener('click', downloadExcelTemplate);

    if (closeModalBtn) closeModalBtn.addEventListener('click', () => detailModal.classList.remove('active'));
    if (detailModal) {
      detailModal.addEventListener('click', (e) => {
        if (e.target === detailModal) detailModal.classList.remove('active');
      });
    }
  }

  function resetSearchFilters() {
    if (searchInput) searchInput.value = '';
    if (filterDistrict) filterDistrict.value = '';
    if (filterLevel) filterLevel.value = '';
    if (filterGrade) filterGrade.value = '';
    if (filterMonth) filterMonth.value = '';
    if (filterAcademicYear) filterAcademicYear.value = '';
    renderTable();
    showToast('ລ້າງຕົວຈຳແນກ (Filter) ຮຽບຮ້ອຍແລ້ວ', 'info');
  }

  function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    if (themeToggleBtn) themeToggleBtn.innerHTML = newTheme === 'dark' ? '<i class="fa-solid fa-moon"></i>' : '<i class="fa-solid fa-sun"></i>';
    showToast(`ສະຫຼັບເປັນ ${newTheme === 'dark' ? 'Dark Mode' : 'Light Mode'} ແລ້ວ`, 'info');
  }

  function calculateActualAttending() {
    if (!passedRegisteredTotal || !actualAttendingTotal) return;

    const passed = parseInt(passedRegisteredTotal.value) || 0;
    const passedFem = parseInt(passedRegisteredFemale.value) || 0;

    const tIn = parseInt(transferInTotal.value) || 0;
    const tInFem = parseInt(transferInFemale.value) || 0;

    const tOut = parseInt(transferOutTotal.value) || 0;
    const tOutFem = parseInt(transferOutFemale.value) || 0;

    const drop = parseInt(dropoutTotal.value) || 0;
    const dropFem = parseInt(dropoutFemale.value) || 0;

    let calcTotal = passed + tIn - tOut - drop;
    let calcFemale = passedFem + tInFem - tOutFem - dropFem;

    if (calcTotal < 0) calcTotal = 0;
    if (calcFemale < 0) calcFemale = 0;

    actualAttendingTotal.value = calcTotal;
    actualAttendingFemale.value = calcFemale;
  }

  function setupValidationWatchers() {
    if (!passedRegisteredTotal) return;

    const pairs = [
      { total: passedRegisteredTotal, female: passedRegisteredFemale },
      { total: transferInTotal, female: transferInFemale },
      { total: transferOutTotal, female: transferOutFemale },
      { total: dropoutTotal, female: dropoutFemale },
      { total: repeaterTotal, female: repeaterFemale },
      { total: actualAttendingTotal, female: actualAttendingFemale }
    ];

    pairs.forEach(pair => {
      if (pair.total && pair.female) {
        const check = () => {
          const tot = parseInt(pair.total.value) || 0;
          const fem = parseInt(pair.female.value) || 0;

          if (fem > tot) {
            pair.female.style.borderColor = 'var(--accent-danger)';
            pair.female.style.boxShadow = '0 0 0 3px rgba(239, 68, 68, 0.3)';
          } else {
            pair.female.style.borderColor = '';
            pair.female.style.boxShadow = '';
          }
        };

        pair.total.addEventListener('input', check);
        pair.female.addEventListener('input', check);
      }
    });

    const checkReasons = () => {
      const tOutCount = parseInt(transferOutTotal.value) || 0;
      const dropCount = parseInt(dropoutTotal.value) || 0;

      if (transferOutReason) {
        if (tOutCount > 0 && !transferOutReason.value.trim()) {
          transferOutReason.style.borderColor = 'var(--accent-danger)';
        } else {
          transferOutReason.style.borderColor = '';
        }
      }

      if (dropoutReason) {
        if (dropCount > 0 && !dropoutReason.value.trim()) {
          dropoutReason.style.borderColor = 'var(--accent-danger)';
        } else {
          dropoutReason.style.borderColor = '';
        }
      }
    };

    if (transferOutTotal) transferOutTotal.addEventListener('input', checkReasons);
    if (transferOutReason) transferOutReason.addEventListener('input', checkReasons);
    if (dropoutTotal) dropoutTotal.addEventListener('input', checkReasons);
    if (dropoutReason) dropoutReason.addEventListener('input', checkReasons);
  }

  function validateForm() {
    if (!studentDataForm) return true;

    const pairs = [
      { total: passedRegisteredTotal, female: passedRegisteredFemale, name: "ເສັງຜ່ານລົງທະບຽນ" },
      { total: transferInTotal, female: transferInFemale, name: "ຍ້າຍເຂົ້າມາຮຽນໃໝ່" },
      { total: transferOutTotal, female: transferOutFemale, name: "ຍ້າຍອອກ" },
      { total: dropoutTotal, female: dropoutFemale, name: "ປະລະການຮຽນ" },
      { total: repeaterTotal, female: repeaterFemale, name: "ຄ້າງຫ້ອງ" },
      { total: actualAttendingTotal, female: actualAttendingFemale, name: "ໜ້າຮຽນຕົວຈິງ" }
    ];

    for (let p of pairs) {
      if (p.total && p.female) {
        const tot = parseInt(p.total.value) || 0;
        const fem = parseInt(p.female.value) || 0;
        if (fem > tot) {
          showToast(`ຂໍ້ຜິດພາດ: ຈຳນວນນັກຮຽນຍິງ ຕ້ອງບໍ່ກາຍຈຳນວນນັກຮຽນທັງໝົດ ໃນຊ່ອງ "${p.name}"`, 'error');
          p.female.focus();
          return false;
        }
      }
    }

    const tOutVal = parseInt(transferOutTotal.value) || 0;
    const tOutReasonVal = transferOutReason.value.trim();
    if (tOutVal > 0 && !tOutReasonVal) {
      showToast('ຂໍ້ຜິດພາດ: ກະລຸນາປ້ອນ "ເຫດຜົນໃນການຍ້າຍອອກ" (ຍ້ອນມີຈຳນວນນັກຮຽນຍ້າຍອອກ)', 'error');
      transferOutReason.focus();
      transferOutReason.style.borderColor = 'var(--accent-danger)';
      return false;
    }

    const dropVal = parseInt(dropoutTotal.value) || 0;
    const dropReasonVal = dropoutReason.value.trim();
    if (dropVal > 0 && !dropReasonVal) {
      showToast('ຂໍ້ຜິດພາດ: ກະລຸນາປ້ອນ "ເຫດຜົນໃນການປະລະການຮຽນ" (ຍ້ອນມີຈຳນວນນັກຮຽນປະລະ)', 'error');
      dropoutReason.focus();
      dropoutReason.style.borderColor = 'var(--accent-danger)';
      return false;
    }

    if (!tOutReasonVal) {
      showToast('ກະລຸນາປ້ອນ "ເຫດຜົນໃນການຍ້າຍອອກ" (ຖ້າບໍ່ມີໃຫ້ປ້ອນ "-")', 'error');
      transferOutReason.focus();
      return false;
    }

    if (!dropReasonVal) {
      showToast('ກະລຸນາປ້ອນ "ເຫດຜົນໃນການປະລະການຮຽນ" (ຖ້າບໍ່ມີໃຫ້ປ້ອນ "-")', 'error');
      dropoutReason.focus();
      return false;
    }

    return true;
  }

  function handleFormSubmit(e) {
    e.preventDefault();

    if (!districtSelect.value) {
      showToast('ກະລຸນາເລືອກເມືອງ', 'error');
      return;
    }

    if (!validateForm()) return;

    const id = recordIdInput.value || `REC-${Date.now()}`;
    const isEdit = !!recordIdInput.value;

    const formattedLaoMonthStr = `${entryAcademicYearSelect.value}:${entryLaoMonthSelect.value}`;

    const formData = {
      id: id,
      district: districtSelect.value,
      school: schoolSelect.value,
      collector_name: collectorName.value.trim(),
      education_level: educationLevelSelect.value,
      grade_level: gradeLevelSelect.value,
      entry_month: formattedLaoMonthStr,
      entry_month_text: getLaoMonthText(formattedLaoMonthStr),

      passed_registered_total: parseInt(passedRegisteredTotal.value) || 0,
      passed_registered_female: parseInt(passedRegisteredFemale.value) || 0,

      transfer_in_total: parseInt(transferInTotal.value) || 0,
      transfer_in_female: parseInt(transferInFemale.value) || 0,

      transfer_out_total: parseInt(transferOutTotal.value) || 0,
      transfer_out_female: parseInt(transferOutFemale.value) || 0,
      transfer_out_reason: transferOutReason.value.trim(),

      dropout_total: parseInt(dropoutTotal.value) || 0,
      dropout_female: parseInt(dropoutFemale.value) || 0,
      dropout_reason: dropoutReason.value.trim(),

      repeater_total: parseInt(repeaterTotal.value) || 0,
      repeater_female: parseInt(repeaterFemale.value) || 0,

      actual_attending_total: parseInt(actualAttendingTotal.value) || 0,
      actual_attending_female: parseInt(actualAttendingFemale.value) || 0,

      created_at: isEdit ? (records.find(r => r.id === id)?.created_at || new Date().toLocaleString()) : new Date().toISOString().replace('T', ' ').substring(0, 16)
    };

    if (isEdit) {
      const idx = records.findIndex(r => r.id === id);
      if (idx !== -1) records[idx] = formData;
      showToast('ອັບເດດຂໍ້ມູນສຳເລັດແລ້ວ!', 'success');
    } else {
      records.unshift(formData);
      showToast('ບັນທຶກຂໍ້ມູນໃໝ່ສຳເລັດແລ້ວ!', 'success');
    }

    // 🟢 Auto-Sync to Google Sheets via Google Apps Script Web App
    if (GOOGLE_APPS_SCRIPT_URL) {
      fetch(GOOGLE_APPS_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData)
      }).then(() => {
        showToast('🟢 ບັນທຶກລົງ Google Sheet ສຳເລັດແລ້ວ!', 'success');
      }).catch(err => {
        console.error("Google Sheets Sync Error:", err);
      });
    }

    saveRecordsToStorage();
    resetForm();
    renderApp();
  }

  function resetForm() {
    if (!studentDataForm) return;

    studentDataForm.reset();
    if (recordIdInput) recordIdInput.value = '';
    if (formModeBadge) {
      formModeBadge.textContent = 'ບັນທຶກຂໍ້ມູນໃໝ່';
      formModeBadge.style.background = 'rgba(99, 102, 241, 0.15)';
      formModeBadge.style.color = 'var(--accent-primary)';
    }
    if (submitFormBtn) submitFormBtn.innerHTML = '<i class="fa-solid fa-paper-plane"></i> ບັນທຶກ ແລະ ສົ່ງຂໍ້ມູນ';

    if (schoolSelect) {
      schoolSelect.disabled = true;
      schoolSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກເມືອງກ່ອນ --</option>';
    }
    
    if (transferOutReason) transferOutReason.style.borderColor = '';
    if (dropoutReason) dropoutReason.style.borderColor = '';

    resetEducationLevelOptions();
    setupLaoMonthDropdowns();
  }

  window.editRecord = function(id) {
    const rec = records.find(r => r.id === id);
    if (!rec) return;

    if (!studentDataForm) {
      window.location.href = `index.html?edit=${id}`;
      return;
    }

    recordIdInput.value = rec.id;
    formModeBadge.textContent = `ແກ້ໄຂຂໍ້ມູນ: ${rec.id}`;
    formModeBadge.style.background = 'rgba(245, 158, 11, 0.15)';
    formModeBadge.style.color = 'var(--accent-warning)';
    submitFormBtn.innerHTML = '<i class="fa-solid fa-pen-to-square"></i> ອັບເດດຂໍ້ມູນ';

    districtSelect.value = rec.district;
    populateSchools(rec.district);
    schoolSelect.value = rec.school;

    collectorName.value = rec.collector_name;
    
    handleSchoolSelectChange(rec.school);
    educationLevelSelect.value = rec.education_level;
    populateGrades(rec.education_level);
    gradeLevelSelect.value = rec.grade_level;

    if (rec.entry_month) {
      const ay = getRecordAcademicYear(rec.entry_month);
      const mc = getRecordMonthCode(rec.entry_month);
      if (ay) entryAcademicYearSelect.value = ay;
      if (mc) entryLaoMonthSelect.value = mc;
    }

    passedRegisteredTotal.value = rec.passed_registered_total;
    passedRegisteredFemale.value = rec.passed_registered_female;

    transferInTotal.value = rec.transfer_in_total;
    transferInFemale.value = rec.transfer_in_female;

    transferOutTotal.value = rec.transfer_out_total;
    transferOutFemale.value = rec.transfer_out_female;
    transferOutReason.value = rec.transfer_out_reason || '';

    dropoutTotal.value = rec.dropout_total;
    dropoutFemale.value = rec.dropout_female;
    dropoutReason.value = rec.dropout_reason || '';

    repeaterTotal.value = rec.repeater_total;
    repeaterFemale.value = rec.repeater_female;

    actualAttendingTotal.value = rec.actual_attending_total;
    actualAttendingFemale.value = rec.actual_attending_female;

    document.getElementById('dataFormCard').scrollIntoView({ behavior: 'smooth' });
  };

  window.deleteRecord = function(id) {
    if (confirm('ທ່ານແນ່ໃຈບໍ່ວ່າຕ້ອງການລົບຂໍ້ມູນລາຍການນີ້?')) {
      records = records.filter(r => r.id !== id);
      saveRecordsToStorage();
      renderApp();
      showToast('ລົບຂໍ້ມູນສຳເລັດແລ້ວ', 'info');
    }
  };

  window.viewRecordDetail = function(id) {
    const rec = records.find(r => r.id === id);
    if (!rec || !detailModal) return;

    modalTitle.textContent = `ລາຍລະອຽດຂໍ້ມູນ: ${rec.school}`;
    modalContent.innerHTML = `
      <div style="display: flex; flex-direction: column; gap: 14px;">
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; background: rgba(255,255,255,0.04); padding: 14px; border-radius: 8px;">
          <div><strong>ເມືອງ:</strong> ${rec.district}</div>
          <div><strong>ໂຮງຮຽນ:</strong> ${rec.school}</div>
          <div><strong>ຊັ້ນ/ຂັ້ນ:</strong> ${rec.education_level} (${rec.grade_level})</div>
          <div><strong>ເດືອນ / ສົກຮຽນ:</strong> ${getLaoMonthText(rec.entry_month)}</div>
          <div><strong>ຜູ້ປ້ອນຂໍ້ມູນ:</strong> ${rec.collector_name}</div>
          <div><strong>ວັນທີບັນທຶກ:</strong> ${rec.created_at}</div>
        </div>

        <h4 style="color: var(--accent-primary); border-bottom: 1px solid var(--border-color); padding-bottom: 4px; margin-top: 8px;">
          <i class="fa-solid fa-list-check"></i> ຕົວເລກສະຖິຕິ:
        </h4>

        <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
          <thead>
            <tr style="background: rgba(0,0,0,0.2);">
              <th style="padding: 8px; text-align: left;">ລາຍການ</th>
              <th style="padding: 8px; text-align: center;">ທັງໝົດ</th>
              <th style="padding: 8px; text-align: center;">ຍິງ</th>
              <th style="padding: 8px; text-align: left;">ເຫດຜົນ</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="padding: 8px;">7. ເສັງຜ່ານມາລົງທະບຽນ</td>
              <td style="padding: 8px; text-align: center;">${rec.passed_registered_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.passed_registered_female}</td>
              <td style="padding: 8px;">-</td>
            </tr>
            <tr>
              <td style="padding: 8px;">8. ຍ້າຍເຂົ້າມາຮຽນໃໝ່</td>
              <td style="padding: 8px; text-align: center;">${rec.transfer_in_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.transfer_in_female}</td>
              <td style="padding: 8px;">-</td>
            </tr>
            <tr>
              <td style="padding: 8px;">9. ຍ້າຍອອກ</td>
              <td style="padding: 8px; text-align: center;">${rec.transfer_out_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.transfer_out_female}</td>
              <td style="padding: 8px;">${rec.transfer_out_reason || '-'}</td>
            </tr>
            <tr>
              <td style="padding: 8px;">10. ປະລະການຮຽນ</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-danger);">${rec.dropout_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.dropout_female}</td>
              <td style="padding: 8px;">${rec.dropout_reason || '-'}</td>
            </tr>
            <tr>
              <td style="padding: 8px;">11. ນັກຮຽນຄ້າງຫ້ອງ</td>
              <td style="padding: 8px; text-align: center;">${rec.repeater_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.repeater_female}</td>
              <td style="padding: 8px;">-</td>
            </tr>
            <tr style="font-weight: bold; background: rgba(16, 185, 129, 0.1);">
              <td style="padding: 8px; color: var(--accent-success);">12. ໜ້າຮຽນຕົວຈິງທັງໝົດ</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-success);">${rec.actual_attending_total}</td>
              <td style="padding: 8px; text-align: center; color: var(--accent-pink);">${rec.actual_attending_female}</td>
              <td style="padding: 8px;">-</td>
            </tr>
          </tbody>
        </table>
      </div>
    `;

    detailModal.classList.add('active');
  };

  // EXCEL IMPORT LOGIC
  function closeExcelModal() {
    if (!excelImportModal) return;
    excelImportModal.classList.remove('active');
    parsedExcelData = null;
    if (excelFileInput) excelFileInput.value = '';
    if (excelPreviewContainer) excelPreviewContainer.style.display = 'none';
    if (confirmExcelImportBtn) confirmExcelImportBtn.disabled = true;
  }

  function handleExcelFileSelect(file) {
    if (!file) return;

    if (excelFileNameBadge) excelFileNameBadge.textContent = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        if (!jsonRows || jsonRows.length === 0) {
          showToast('ບໍ່ພົບຂໍ້ມູນໃນຟາຍ Excel ທີ່ເລືອກ', 'error');
          return;
        }

        parsedExcelData = jsonRows;
        renderExcelPreview(jsonRows);
        if (confirmExcelImportBtn) confirmExcelImportBtn.disabled = false;
        showToast(`ອ່ານຟາຍ Excel ສຳເລັດ (${jsonRows.length} ແຖວ)`, 'success');

      } catch (err) {
        console.error(err);
        showToast('ບໍ່ສາມາດອ່ານຟາຍ Excel ນີ້ໄດ້. ກະລຸນາກວດສອບຟໍແມັດຟາຍ', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function renderExcelPreview(rows) {
    if (!rows || rows.length === 0 || !excelPreviewHead) return;

    excelRowCount.textContent = rows.length;
    const headers = Object.keys(rows[0]);

    excelPreviewHead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

    const previewRows = rows.slice(0, 10);
    excelPreviewBody.innerHTML = previewRows.map(row => {
      return `<tr>${headers.map(h => `<td>${row[h] || ''}</td>`).join('')}</tr>`;
    }).join('');

    excelPreviewContainer.style.display = 'block';
  }

  function processExcelImport() {
    if (!parsedExcelData || parsedExcelData.length === 0) return;

    const mode = document.querySelector('input[name="importTarget"]:checked')?.value || 'PRESETS';

    if (mode === 'PRESETS') {
      let addedDistCount = 0;
      let addedSchoolCount = 0;

      const customImported = JSON.parse(localStorage.getItem('lao_excel_custom_districts') || '{}');

      parsedExcelData.forEach(row => {
        const distVal = (row['ເມືອງ'] || row['ຊື່ເມືອງ'] || row['District'] || row['district'] || '').toString().trim();
        const schoolVal = (row['ໂຮງຮຽນ'] || row['ຊື່ໂຮງຮຽນ'] || row['School'] || row['school'] || '').toString().trim();

        if (distVal) {
          if (!customImported[distVal]) {
            customImported[distVal] = [];
            addedDistCount++;
          }
          if (schoolVal && !customImported[distVal].includes(schoolVal)) {
            customImported[distVal].push(schoolVal);
            addedSchoolCount++;
          }
        }
      });

      localStorage.setItem('lao_excel_custom_districts', JSON.stringify(customImported));

      populateDistrictDropdowns();
      showToast(`ນຳເຂົ້າສຳເລັດ! ເພີ່ມ ${addedDistCount} ເມືອງ ແລະ ${addedSchoolCount} ໂຮງຮຽນ ໃສ່ Dropdown ແລ້ວ`, 'success');

    } else {
      let importedCount = 0;

      parsedExcelData.forEach(row => {
        const getVal = (keys, defaultVal = "") => {
          for (let k of keys) {
            if (row[k] !== undefined && row[k] !== "") return row[k];
          }
          return defaultVal;
        };

        const dist = getVal(['ເມືອງ', 'ຊື່ເມືອງ', 'District', 'district']);
        const sch = getVal(['ໂຮງຮຽນ', 'ຊື່ໂຮງຮຽນ', 'School', 'school']);

        if (dist && sch) {
          const defaultAY = entryAcademicYearSelect ? entryAcademicYearSelect.value : "2026-2027";
          const defaultM = entryLaoMonthSelect ? entryLaoMonthSelect.value : "08";
          const rawMonth = getVal(['ເດືອນ', 'ເດືອນປ້ອນຂໍ້ມູນ', 'Month', 'entry_month'], `${defaultAY}:${defaultM}`);

          const rec = {
            id: `REC-EXCEL-${Date.now()}-${Math.floor(Math.random()*1000)}`,
            district: dist,
            school: sch,
            collector_name: getVal(['ຜູ້ປ້ອນຂໍ້ມູນ', 'ຜູ້ປ້ອນ', 'Collector', 'collector_name'], 'Excel Import'),
            education_level: getVal(['ຊັ້ນຮຽນ', 'Level', 'education_level'], 'ມັດທະຍົມສົມບູນ'),
            grade_level: getVal(['ຂັ້ນຮຽນ', 'Grade', 'grade_level'], 'ມ.1'),
            entry_month: rawMonth,

            passed_registered_total: parseInt(getVal(['ເສັງຜ່ານລົງທະບຽນ_ຮວມ', 'ລົງທະບຽນ_ຮວມ', 'PassedTotal'], 0)) || 0,
            passed_registered_female: parseInt(getVal(['ເສັງຜ່ານລົງທະບຽນ_ຍິງ', 'ລົງທະບຽນ_ຍິງ', 'PassedFemale'], 0)) || 0,

            transfer_in_total: parseInt(getVal(['ຍ້າຍເຂົ້າ_ຮວມ', 'TransferInTotal'], 0)) || 0,
            transfer_in_female: parseInt(getVal(['ຍ້າຍເຂົ້າ_ຍິງ', 'TransferInFemale'], 0)) || 0,

            transfer_out_total: parseInt(getVal(['ຍ້າຍອອກ_ຮວມ', 'TransferOutTotal'], 0)) || 0,
            transfer_out_female: parseInt(getVal(['ຍ້າຍອອກ_ຍິງ', 'TransferOutFemale'], 0)) || 0,
            transfer_out_reason: getVal(['ເຫດຜົນຍ້າຍອອກ', 'TransferOutReason'], '-'),

            dropout_total: parseInt(getVal(['ປະລະ_ຮວມ', 'DropoutTotal'], 0)) || 0,
            dropout_female: parseInt(getVal(['ປະລະ_ຍິງ', 'DropoutFemale'], 0)) || 0,
            dropout_reason: getVal(['ເຫດຜົນປະລະ', 'DropoutReason'], '-'),

            repeater_total: parseInt(getVal(['ຄ້າງຫ້ອງ_ຮວມ', 'RepeaterTotal'], 0)) || 0,
            repeater_female: parseInt(getVal(['ຄ້າງຫ້ອງ_ຍິງ', 'RepeaterFemale'], 0)) || 0,

            actual_attending_total: parseInt(getVal(['ໜ້າຕົວຈິງ_ຮວມ', 'ActualTotal'], 0)) || 0,
            actual_attending_female: parseInt(getVal(['ໜ້າຕົວຈິງ_ຍິງ', 'ActualFemale'], 0)) || 0,

            created_at: new Date().toISOString().replace('T', ' ').substring(0, 16)
          };

          if (rec.actual_attending_total === 0) {
            rec.actual_attending_total = rec.passed_registered_total + rec.transfer_in_total - rec.transfer_out_total - rec.dropout_total;
            rec.actual_attending_female = rec.passed_registered_female + rec.transfer_in_female - rec.transfer_out_female - rec.dropout_female;
            if (rec.actual_attending_total < 0) rec.actual_attending_total = 0;
            if (rec.actual_attending_female < 0) rec.actual_attending_female = 0;
          }

          records.unshift(rec);
          importedCount++;

          if (GOOGLE_APPS_SCRIPT_URL) {
            fetch(GOOGLE_APPS_SCRIPT_URL, {
              method: 'POST',
              mode: 'no-cors',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(rec)
            });
          }
        }
      });

      saveRecordsToStorage();
      renderApp();
      showToast(`ນຳເຂົ້າ ${importedCount} ລາຍການສະຖິຕິໃສ່ຕາຕະລາງ ແລະ Google Sheet ສຳເລັດ!`, 'success');
    }

    closeExcelModal();
  }

  function downloadExcelTemplate() {
    const templateData = [
      {
        "ເມືອງ": "ເມືອງໄຊ",
        "ໂຮງຮຽນ": "ໂຮງຮຽນ ມ.ສ ເທດສະບານແຂວງ",
        "ຊັ້ນຮຽນ": "ມັດທະຍົມສົມບູນ",
        "ຂັ້ນຮຽນ": "ມ.1",
        "ຜູ້ປ້ອນຂໍ້ມູນ": "ສົມຈິດ ວົງສະຫວັດ",
        "ເດືອນ/ສົກຮຽນ": "2026-2027:08",
        "ເສັງຜ່ານລົງທະບຽນ_ຮວມ": 45,
        "ເສັງຜ່ານລົງທະບຽນ_ຍິງ": 23,
        "ຍ້າຍເຂົ້າ_ຮວມ": 3,
        "ຍ້າຍເຂົ້າ_ຍິງ": 2,
        "ຍ້າຍອອກ_ຮວມ": 1,
        "ຍ້າຍອອກ_ຍິງ": 1,
        "ເຫດຜົນຍ້າຍອອກ": "ຍ້າຍຕາມຄອບຄົວ",
        "ປະລະ_ຮວມ": 1,
        "ປະລະ_ຍິງ": 0,
        "ເຫດຜົນປະລະ": "ຊ່ວຍວຽກຄອບຄົວ",
        "ຄ້າງຫ້ອງ_ຮວມ": 2,
        "ຄ້າງຫ້ອງ_ຍິງ": 1,
        "ໜ້າຕົວຈິງ_ຮວມ": 46,
        "ໜ້າຕົວຈິງ_ຍິງ": 24
      }
    ];

    const worksheet = XLSX.utils.json_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "LaoSchoolDataTemplate");

    XLSX.writeFile(workbook, "Lao_School_Data_Template.xlsx");
    showToast('ດາວໂຫຼດແມ່ແບບ Excel (Template) ສຳເລັດແລ້ວ!', 'success');
  }

  // EXPORT TO EXCEL (.xlsx) FUNCTIONALITY
  function exportToExcel() {
    if (records.length === 0) {
      showToast('ບໍ່ມີຂໍ້ມູນທີ່ຈະສົ່ງອອກ', 'error');
      return;
    }

    const excelRows = records.map(r => {
      return {
        "ID": r.id,
        "ເດືອນ ແລະ ສົກຮຽນ": getLaoMonthText(r.entry_month),
        "ເມືອງ": r.district,
        "ໂຮງຮຽນ": r.school,
        "ຊັ້ນຮຽນ": r.education_level,
        "ຂັ້ນຮຽນ": r.grade_level,
        "ຜູ້ປ້ອນຂໍ້ມູນ": r.collector_name,

        "ເສັງຜ່ານລົງທະບຽນ_ທັງໝົດ": r.passed_registered_total,
        "ເສັງຜ່ານລົງທະບຽນ_ຍິງ": r.passed_registered_female,

        "ຍ້າຍເຂົ້າ_ທັງໝົດ": r.transfer_in_total,
        "ຍ້າຍເຂົ້າ_ຍິງ": r.transfer_in_female,

        "ຍ້າຍອອກ_ທັງໝົດ": r.transfer_out_total,
        "ຍ້າຍອອກ_ຍິງ": r.transfer_out_female,
        "ເຫດຜົນຍ້າຍອອກ": r.transfer_out_reason || "",

        "ປະລະ_ທັງໝົດ": r.dropout_total,
        "ປະລະ_ຍິງ": r.dropout_female,
        "ເຫດຜົນປະລະ": r.dropout_reason || "",

        "ຄ້າງຫ້ອງ_ທັງໝົດ": r.repeater_total,
        "ຄ້າງຫ້ອງ_ຍິງ": r.repeater_female,

        "ໜ້າຕົວຈິງ_ທັງໝົດ": r.actual_attending_total,
        "ໜ້າຕົວຈິງ_ຍິງ": r.actual_attending_female,

        "ວັນທີບັນທຶກ": r.created_at
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(excelRows);

    const colWidths = [
      { wch: 18 }, { wch: 28 }, { wch: 22 }, { wch: 28 }, { wch: 18 }, { wch: 12 }, { wch: 22 },
      { wch: 24 }, { wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 18 }, { wch: 16 }, { wch: 26 },
      { wch: 16 }, { wch: 14 }, { wch: 26 }, { wch: 18 }, { wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 20 }
    ];
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Student_Data");

    const todayStr = new Date().toISOString().substring(0, 10);
    XLSX.writeFile(workbook, `Lao_School_Data_Report_${todayStr}.xlsx`);
    showToast('ສົ່ງອອກຟາຍ Excel (.xlsx) ສຳເລັດແລ້ວ!', 'success');
  }

  // Comprehensive Multi-Criteria Search Engine
  function renderTable() {
    if (!tableBody) return;

    const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const selectedDist = filterDistrict ? filterDistrict.value : '';
    const selectedLvl = filterLevel ? filterLevel.value : '';
    const selectedGrade = filterGrade ? filterGrade.value : '';
    const selectedMonth = filterMonth ? filterMonth.value : '';
    const selectedAY = filterAcademicYear ? filterAcademicYear.value : '';

    const filtered = records.filter(r => {
      const laoMonthText = getLaoMonthText(r.entry_month).toLowerCase();
      const recAY = getRecordAcademicYear(r.entry_month);
      const recMonthCode = getRecordMonthCode(r.entry_month);

      const matchQuery = !query || 
        r.district.toLowerCase().includes(query) || 
        r.school.toLowerCase().includes(query) || 
        r.education_level.toLowerCase().includes(query) ||
        r.grade_level.toLowerCase().includes(query) ||
        r.collector_name.toLowerCase().includes(query) ||
        laoMonthText.includes(query) ||
        recAY.includes(query) ||
        recMonthCode.includes(query);

      const matchDist = !selectedDist || r.district === selectedDist;
      const matchLvl = !selectedLvl || r.education_level === selectedLvl;
      const matchGrade = !selectedGrade || r.grade_level === selectedGrade;
      const matchMonth = !selectedMonth || recMonthCode === selectedMonth;
      const matchAY = !selectedAY || recAY === selectedAY;

      return matchQuery && matchDist && matchLvl && matchGrade && matchMonth && matchAY;
    });

    if (tableRecordCount) tableRecordCount.textContent = `${filtered.length} ລາຍການ`;
    tableBody.innerHTML = '';

    if (filtered.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="11" style="text-align: center; padding: 28px; color: var(--text-muted);">
            <i class="fa-solid fa-folder-open" style="font-size: 2rem; margin-bottom: 8px; display: block; color: var(--accent-primary);"></i>
            ບໍ່ພົບຂໍ້ມູນທີ່ກົງກັບເງື່ອນໄຂການຄົ້ນຫາ
          </td>
        </tr>
      `;
      return;
    }

    filtered.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${getLaoMonthText(r.entry_month)}</strong></td>
        <td>${r.district}</td>
        <td><strong>${r.school}</strong></td>
        <td><span class="chip chip-primary">${r.education_level} (${r.grade_level})</span></td>
        <td>${r.passed_registered_total} <span class="female-highlight">(${r.passed_registered_female})</span></td>
        <td><span class="chip chip-success">+${r.transfer_in_total}</span></td>
        <td><span class="chip chip-warning">-${r.transfer_out_total}</span></td>
        <td><span style="color: var(--accent-danger); font-weight: bold;">-${r.dropout_total}</span></td>
        <td>
          <strong style="color: var(--accent-success);">${r.actual_attending_total}</strong>
          <span class="female-highlight">(${r.actual_attending_female})</span>
        </td>
        <td>${r.collector_name}</td>
        <td style="text-align: center;">
          <div class="table-actions">
            <button class="btn btn-secondary btn-sm" onclick="viewRecordDetail('${r.id}')" title="ເບິ່ງລາຍລະອຽດ">
              <i class="fa-solid fa-eye"></i>
            </button>
            <button class="btn btn-secondary btn-sm" onclick="editRecord('${r.id}')" title="ແກ້ໄຂ">
              <i class="fa-solid fa-pen-to-square"></i>
            </button>
            <button class="btn btn-secondary btn-sm" onclick="deleteRecord('${r.id}')" title="ລົບ" style="color: var(--accent-danger);">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </td>
      `;
      tableBody.appendChild(tr);
    });
  }

  // RENDER COMPREHENSIVE STAT KPI CARDS (WITH EXACT WORDING AND FEMALE PERCENTAGE FOR ALL CARDS)
  function renderStats() {
    const regTotal = records.reduce((acc, r) => acc + (r.passed_registered_total || 0), 0);
    const regFemale = records.reduce((acc, r) => acc + (r.passed_registered_female || 0), 0);
    const regPct = regTotal > 0 ? ((regFemale / regTotal) * 100).toFixed(1) : 0;

    const tInTotal = records.reduce((acc, r) => acc + (r.transfer_in_total || 0), 0);
    const tInFemale = records.reduce((acc, r) => acc + (r.transfer_in_female || 0), 0);
    const tInPct = tInTotal > 0 ? ((tInFemale / tInTotal) * 100).toFixed(1) : 0;

    const tOutTotal = records.reduce((acc, r) => acc + (r.transfer_out_total || 0), 0);
    const tOutFemale = records.reduce((acc, r) => acc + (r.transfer_out_female || 0), 0);
    const tOutPct = tOutTotal > 0 ? ((tOutFemale / tOutTotal) * 100).toFixed(1) : 0;

    const dropTotal = records.reduce((acc, r) => acc + (r.dropout_total || 0), 0);
    const dropFemale = records.reduce((acc, r) => acc + (r.dropout_female || 0), 0);
    const dropPct = dropTotal > 0 ? ((dropFemale / dropTotal) * 100).toFixed(1) : 0;

    const repTotal = records.reduce((acc, r) => acc + (r.repeater_total || 0), 0);
    const repFemale = records.reduce((acc, r) => acc + (r.repeater_female || 0), 0);
    const repPct = repTotal > 0 ? ((repFemale / repTotal) * 100).toFixed(1) : 0;

    const actualTotal = records.reduce((acc, r) => acc + (r.actual_attending_total || 0), 0);
    const actualFemale = records.reduce((acc, r) => acc + (r.actual_attending_female || 0), 0);
    const actualPct = actualTotal > 0 ? ((actualFemale / actualTotal) * 100).toFixed(1) : 0;

    if (statRegisteredTotal) statRegisteredTotal.textContent = regTotal.toLocaleString();
    if (statRegisteredFemale) statRegisteredFemale.textContent = regFemale.toLocaleString();
    if (statRegisteredPercent) statRegisteredPercent.textContent = regPct;

    if (statTransferInTotal) statTransferInTotal.textContent = tInTotal.toLocaleString();
    if (statTransferInFemale) statTransferInFemale.textContent = tInFemale.toLocaleString();
    if (statTransferInPercent) statTransferInPercent.textContent = tInPct;

    if (statTransferOutTotal) statTransferOutTotal.textContent = tOutTotal.toLocaleString();
    if (statTransferOutFemale) statTransferOutFemale.textContent = tOutFemale.toLocaleString();
    if (statTransferOutPercent) statTransferOutPercent.textContent = tOutPct;

    if (statTotalDropout) statTotalDropout.textContent = dropTotal.toLocaleString();
    if (statDropoutFemale) statDropoutFemale.textContent = dropFemale.toLocaleString();
    if (statDropoutPercent) statDropoutPercent.textContent = dropPct;

    if (statRepeaterTotal) statRepeaterTotal.textContent = repTotal.toLocaleString();
    if (statRepeaterFemale) statRepeaterFemale.textContent = repFemale.toLocaleString();
    if (statRepeaterPercent) statRepeaterPercent.textContent = repPct;

    if (statTotalStudents) statTotalStudents.textContent = actualTotal.toLocaleString();
    if (statFemaleStudents) statFemaleStudents.textContent = actualFemale.toLocaleString();
    if (statFemalePercent) statFemalePercent.textContent = actualPct;
  }

  // RENDER ADVANCED ANALYTICS CHARTS (MONTHLY TRENDS, DISTRICTS, GENDER & MOVEMENTS)
  function renderCharts() {
    const canvasTrend = document.getElementById('monthlyTrendChart');
    const canvasDist = document.getElementById('districtChart');
    const canvasGender = document.getElementById('genderChart');
    const canvasMove = document.getElementById('movementChart');

    // 📈 CHART 1: MONTHLY TREND LINE CHART
    if (canvasTrend) {
      const monthDataMap = {};
      LAO_MONTHS.forEach(m => {
        monthDataMap[m.value] = { name: m.short, total: 0, female: 0 };
      });

      records.forEach(r => {
        const mc = getRecordMonthCode(r.entry_month);
        if (monthDataMap[mc]) {
          monthDataMap[mc].total += (r.actual_attending_total || 0);
          monthDataMap[mc].female += (r.actual_attending_female || 0);
        }
      });

      const trendLabels = LAO_MONTHS.map(m => m.short);
      const trendTotals = LAO_MONTHS.map(m => monthDataMap[m.value].total);
      const trendFemales = LAO_MONTHS.map(m => monthDataMap[m.value].female);

      const ctxTrend = canvasTrend.getContext('2d');
      if (monthlyTrendChart) monthlyTrendChart.destroy();

      monthlyTrendChart = new Chart(ctxTrend, {
        type: 'line',
        data: {
          labels: trendLabels,
          datasets: [
            {
              label: 'ໜ້າຮຽນຕົວຈິງທັງໝົດ',
              data: trendTotals,
              borderColor: '#6366f1',
              backgroundColor: 'rgba(99, 102, 241, 0.15)',
              borderWidth: 3,
              fill: true,
              tension: 0.35,
              pointBackgroundColor: '#6366f1',
              pointRadius: 4
            },
            {
              label: 'ໜ້າຮຽນຕົວຈິງ (ຍິງ)',
              data: trendFemales,
              borderColor: '#ec4899',
              backgroundColor: 'rgba(236, 72, 153, 0.08)',
              borderWidth: 2,
              fill: true,
              tension: 0.35,
              pointBackgroundColor: '#ec4899',
              pointRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { position: 'top', labels: { color: '#94a3b8' } }
          },
          scales: {
            y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
            x: { grid: { display: false } }
          }
        }
      });
    }

    // 📊 CHART 2: DISTRICT COMPARISON BAR CHART
    if (canvasDist) {
      const districtMap = {};
      Object.keys(DEFAULT_LAO_DISTRICTS).forEach(d => districtMap[d] = 0);

      records.forEach(r => {
        const dist = r.district || 'ອື່ນໆ';
        districtMap[dist] = (districtMap[dist] || 0) + (r.actual_attending_total || 0);
      });

      const distLabels = Object.keys(districtMap);
      const distValues = Object.values(districtMap);

      const ctxDist = canvasDist.getContext('2d');
      if (districtChart) districtChart.destroy();

      districtChart = new Chart(ctxDist, {
        type: 'bar',
        data: {
          labels: distLabels,
          datasets: [{
            label: 'ຈຳນວນນັກຮຽນ (ຄົນ)',
            data: distValues,
            backgroundColor: '#06b6d4',
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false } },
          scales: {
            y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
            x: { grid: { display: false } }
          }
        }
      });
    }

    // 🍩 CHART 3: GENDER RATIO DOUGHNUT CHART
    if (canvasGender) {
      let totalMale = 0;
      let totalFemale = 0;

      records.forEach(r => {
        totalFemale += (r.actual_attending_female || 0);
        totalMale += ((r.actual_attending_total || 0) - (r.actual_attending_female || 0));
      });

      const ctxGender = canvasGender.getContext('2d');
      if (genderChart) genderChart.destroy();

      genderChart = new Chart(ctxGender, {
        type: 'doughnut',
        data: {
          labels: ['ນັກຮຽນຊາຍ', 'ນັກຮຽນຍິງ'],
          datasets: [{
            data: [totalMale < 0 ? 0 : totalMale, totalFemale],
            backgroundColor: ['#6366f1', '#ec4899'],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { position: 'bottom', labels: { color: '#94a3b8' } }
          }
        }
      });
    }

    // 🔁 CHART 4: MOVEMENT COMPARISON BAR CHART
    if (canvasMove) {
      const reg = records.reduce((a, r) => a + (r.passed_registered_total || 0), 0);
      const tIn = records.reduce((a, r) => a + (r.transfer_in_total || 0), 0);
      const tOut = records.reduce((a, r) => a + (r.transfer_out_total || 0), 0);
      const drop = records.reduce((a, r) => a + (r.dropout_total || 0), 0);
      const rep = records.reduce((a, r) => a + (r.repeater_total || 0), 0);

      const ctxMove = canvasMove.getContext('2d');
      if (movementChart) movementChart.destroy();

      movementChart = new Chart(ctxMove, {
        type: 'bar',
        data: {
          labels: ['ລົງທະບຽນ', 'ຍ້າຍເຂົ້າ', 'ຍ້າຍອອກ', 'ປະລະ', 'ຄ້າງຫ້ອງ'],
          datasets: [{
            label: 'ຈຳນວນນັກຮຽນ (ຄົນ)',
            data: [reg, tIn, tOut, drop, rep],
            backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#ef4444', '#a855f7'],
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false } },
          scales: {
            y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
            x: { grid: { display: false } }
          }
        }
      });
    }
  }

  function showToast(message, type = 'info') {
    const toastContainer = document.getElementById('toastContainer');
    if (!toastContainer) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let iconClass = 'fa-info-circle';
    if (type === 'success') iconClass = 'fa-circle-check';
    if (type === 'error') iconClass = 'fa-triangle-exclamation';

    toast.innerHTML = `<i class="fa-solid ${iconClass}"></i> <span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(100%)';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  function renderApp() {
    renderTable();
    renderStats();
    renderCharts();
  }
});
