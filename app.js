/* ==========================================================================
   Lao Student Data Collection System - Application Logic
   ========================================================================== */

// 🟢 Google Apps Script Web App URL
const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbydJFn8REuZ0zgFQNHPEY49oSsbO660_mD3z71EJ7s3E4kFYZH8kQcCowH3ai93OCuSDQ/exec";

document.addEventListener('DOMContentLoaded', () => {
  // Application State
  let records = [];

  // DOM Elements - Form
  const studentDataForm = document.getElementById('studentDataForm');
  const recordIdInput = document.getElementById('recordId');
  const formModeBadge = document.getElementById('formModeBadge');

  const districtSelect = document.getElementById('districtSelect');
  const schoolSelect = document.getElementById('schoolSelect');
  const schoolCodeInput = document.getElementById('schoolCodeInput');
  const collectorNameInput = document.getElementById('collectorName');
  const educationLevelSelect = document.getElementById('educationLevelSelect');
  const gradeLevelSelect = document.getElementById('gradeLevelSelect');
  const entryLaoMonthSelect = document.getElementById('entryLaoMonthSelect');
  const entryAcademicYearSelect = document.getElementById('entryAcademicYearSelect');

  // DOM Elements - Numbers
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
  const themeToggleBtn = document.getElementById('themeToggleBtn');

  // DOM Elements - Dashboard / Table / Filters
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

  // Submission Status Elements
  const kpiTotalMasterSchools = document.getElementById('kpiTotalMasterSchools');
  const kpiSubmittedSchools = document.getElementById('kpiSubmittedSchools');
  const kpiUnsubmittedSchools = document.getElementById('kpiUnsubmittedSchools');
  const kpiSubmissionRate = document.getElementById('kpiSubmissionRate');
  const overallProgressBar = document.getElementById('overallProgressBar');
  const unsubmittedTableBody = document.getElementById('unsubmittedTableBody');
  const unsubmittedFilterDistrict = document.getElementById('unsubmittedFilterDistrict');
  const unsubmittedCountBadge = document.getElementById('unsubmittedCountBadge');

  const passedRegisteredLabel = document.getElementById('passedRegisteredLabel');
  const augustPrevYearContainer = document.getElementById('augustPrevYearContainer');
  const augustPrevYearLabel = document.getElementById('augustPrevYearLabel');
  const augustPrevYearTotal = document.getElementById('augustPrevYearTotal');
  const augustPrevYearFemale = document.getElementById('augustPrevYearFemale');
  const augustPrevYearError = document.getElementById('augustPrevYearError');
  const repeaterContainer = document.getElementById('repeaterContainer');
  const repeaterLabel = document.getElementById('repeaterLabel');
  const fieldTransferInLabel = document.getElementById('fieldTransferInLabel');
  const fieldTransferOutLabel = document.getElementById('fieldTransferOutLabel');
  const fieldDropoutLabel = document.getElementById('fieldDropoutLabel');
  const fieldActualLabel = document.getElementById('fieldActualLabel');
  const calcFormulaText = document.getElementById('calcFormulaText');

  // Initialize Application
  initApp();

  function initApp() {
    loadRecordsFromStorage();
    setupTheme();
    populateAllDropdowns();
    setupEventListeners();
    updateMonthDynamicLabels();
    calculateActualAttending();
    renderApp();
    syncSchoolsFromGoogleSheet();
  }

  function updateMonthDynamicLabels() {
    if (!entryLaoMonthSelect) return;
    const val = (entryLaoMonthSelect.value || "").trim();
    if (val === "08" || val === "ສິງຫາ") {
      if (passedRegisteredLabel) passedRegisteredLabel.innerHTML = '7. ນັກຮຽນທ້າຍປີຜ່ານມາ <span class="req">*</span>';
      if (augustPrevYearLabel) augustPrevYearLabel.innerHTML = '8. ນັກຮຽນທ້າຍປີຜ່ານມາ-ມາລົງທະບຽນຮຽນຕົວຈິງ <span class="req">*</span>';
      if (repeaterLabel) repeaterLabel.innerHTML = '9. ຈຳນວນນັກຮຽນຄ້າງຫ້ອງ <span class="req">*</span>';
      if (fieldTransferInLabel) fieldTransferInLabel.innerHTML = '10. ຈຳນວນນັກຮຽນຍ້າຍເຂົ້າມາຮຽນໃໝ່ <span class="req">*</span>';
      if (fieldTransferOutLabel) fieldTransferOutLabel.innerHTML = '11. ຈຳນວນນັກຮຽນຍ້າຍອອກ <span class="req">*</span>';
      if (fieldDropoutLabel) fieldDropoutLabel.innerHTML = '12. ຈຳນວນນັກຮຽນທີ່ປະລະການຮຽນ <span class="req">*</span>';
      if (fieldActualLabel) fieldActualLabel.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> 13. ຈຳນວນນັກຮຽນທັງໝົດທີ່ມີໜ້າຮຽນຕົວຈິງ (Auto Calculated)';
      if (calcFormulaText) calcFormulaText.innerHTML = '<i class="fa-solid fa-info-circle"></i> ສູດຄິດໄລ່: (ມາລົງທະບຽນ) + (ຄ້າງຫ້ອງ) + (ຍ້າຍເຂົ້າ) - (ຍ້າຍອອກ) - (ປະລະ)';

      if (augustPrevYearContainer) augustPrevYearContainer.style.display = 'block';
      if (repeaterContainer) repeaterContainer.style.display = 'block';
    } else {
      if (passedRegisteredLabel) passedRegisteredLabel.innerHTML = '7. ນັກຮຽນທ້າຍເດືອນ <span class="req">*</span>';
      if (fieldTransferInLabel) fieldTransferInLabel.innerHTML = '8. ຈຳນວນນັກຮຽນຍ້າຍເຂົ້າມາຮຽນໃໝ່ <span class="req">*</span>';
      if (fieldTransferOutLabel) fieldTransferOutLabel.innerHTML = '9. ຈຳນວນນັກຮຽນຍ້າຍອອກ <span class="req">*</span>';
      if (fieldDropoutLabel) fieldDropoutLabel.innerHTML = '10. ຈຳນວນນັກຮຽນທີ່ປະລະການຮຽນ <span class="req">*</span>';
      if (fieldActualLabel) fieldActualLabel.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> 11. ຈຳນວນນັກຮຽນທັງໝົດທີ່ມີໜ້າຮຽນຕົວຈິງ (Auto Calculated)';
      if (calcFormulaText) calcFormulaText.innerHTML = '<i class="fa-solid fa-info-circle"></i> ສູດຄິດໄລ່: (ນັກຮຽນທ້າຍເດືອນ) + (ຍ້າຍເຂົ້າ) - (ຍ້າຍອອກ) - (ປະລະ)';

      if (augustPrevYearContainer) augustPrevYearContainer.style.display = 'none';
      if (augustPrevYearTotal) augustPrevYearTotal.value = 0;
      if (augustPrevYearFemale) augustPrevYearFemale.value = 0;
      if (augustPrevYearError) augustPrevYearError.style.display = 'none';
      if (repeaterContainer) repeaterContainer.style.display = 'none';
      if (repeaterTotal) repeaterTotal.value = 0;
      if (repeaterFemale) repeaterFemale.value = 0;
    }
    validateAugustCounts();
    calculateActualAttending();
  }

  function validateAugustCounts() {
    if (!entryLaoMonthSelect) return true;
    const val = (entryLaoMonthSelect.value || "").trim();
    if (val !== "08" && val !== "ສິງຫາ") {
      if (augustPrevYearError) augustPrevYearError.style.display = 'none';
      return true;
    }

    const maxTotal = parseInt(passedRegisteredTotal?.value) || 0;
    const maxFemale = parseInt(passedRegisteredFemale?.value) || 0;
    const curTotal = parseInt(augustPrevYearTotal?.value) || 0;
    const curFemale = parseInt(augustPrevYearFemale?.value) || 0;

    let isValid = true;
    let errorMsgs = [];

    if (curTotal > maxTotal) {
      isValid = false;
      errorMsgs.push(`ທັງໝົດ (${curTotal}) ເກີນ ${maxTotal}`);
    }
    if (curFemale > maxFemale) {
      isValid = false;
      errorMsgs.push(`ຍິງ (${curFemale}) ເກີນ ${maxFemale}`);
    }

    if (!isValid && augustPrevYearError) {
      augustPrevYearError.style.display = 'block';
      augustPrevYearError.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> ຈຳນວນປ້ອນໄດ້ຕ້ອງບໍ່ເກີນ ນັກເສັງຜ່ານປີຜ່ານມາ (ຊ່ອງ 7: ທັງໝົດ ${maxTotal}, ຍິງ ${maxFemale}) - ${errorMsgs.join(', ')}`;
    } else if (augustPrevYearError) {
      augustPrevYearError.style.display = 'none';
    }

    return isValid;
  }

  // Validate female ≤ total for all pairs
  function validateFemaleCounts() {
    const pairs = [
      { tot: passedRegisteredTotal,  fem: passedRegisteredFemale,  errId: 'err_passedRegistered' },
      { tot: augustPrevYearTotal,    fem: augustPrevYearFemale,    errId: 'err_augustPrevYear' },
      { tot: transferInTotal,        fem: transferInFemale,        errId: 'err_transferIn' },
      { tot: transferOutTotal,       fem: transferOutFemale,       errId: 'err_transferOut' },
      { tot: dropoutTotal,           fem: dropoutFemale,           errId: 'err_dropout' },
      { tot: repeaterTotal,          fem: repeaterFemale,          errId: 'err_repeater' },
    ];

    let allValid = true;
    pairs.forEach(({ tot, fem, errId }) => {
      const errEl = document.getElementById(errId);
      if (!tot || !fem || !errEl) return;
      const t = parseInt(tot.value) || 0;
      const f = parseInt(fem.value) || 0;
      if (f > t) {
        errEl.style.display = 'block';
        allValid = false;
      } else {
        errEl.style.display = 'none';
      }
    });
    return allValid;
  }

  // Theme setup
  function setupTheme() {
    const savedTheme = localStorage.getItem('lao_app_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    if (themeToggleBtn) {
      themeToggleBtn.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('lao_app_theme', newTheme);
        updateThemeIcon(newTheme);
      });
    }
  }

  function updateThemeIcon(theme) {
    if (!themeToggleBtn) return;
    const icon = themeToggleBtn.querySelector('i');
    if (icon) {
      icon.className = theme === 'dark' ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
    }
  }

  // Load records strictly from LocalStorage (User records only)
  function loadRecordsFromStorage() {
    try {
      const saved = localStorage.getItem('lao_school_records');
      if (saved !== null) {
        records = JSON.parse(saved) || [];
      } else if (typeof SAMPLE_INITIAL_RECORDS !== 'undefined' && Array.isArray(SAMPLE_INITIAL_RECORDS)) {
        records = [...SAMPLE_INITIAL_RECORDS];
        localStorage.setItem('lao_school_records', JSON.stringify(records));
      } else {
        records = [];
      }
    } catch (e) {
      records = [];
    }

    // ✅ One-time migration: fix old education_level values
    const LEVEL_MAP = {
      'ມັດທະຍົມຕອນຕົ້ນ': 'ມັດທະຍົມສຶກສາຕອນຕົ້ນ'
    };
    let migrated = false;
    records.forEach(r => {
      if (LEVEL_MAP[r.education_level]) {
        r.education_level = LEVEL_MAP[r.education_level];
        migrated = true;
      }
    });
    if (migrated) {
      localStorage.setItem('lao_school_records', JSON.stringify(records));
    }
  }

  function saveRecordsToStorage() {
    localStorage.setItem('lao_school_records', JSON.stringify(records));
    localStorage.setItem('lao_school_records_updated_at', Date.now().toString());
  }

  // Populate All Form Dropdowns cleanly
  function populateAllDropdowns() {
    populateDistrictDropdowns();
    populateSchoolDropdown();
    populateEducationLevelDropdown();
    populateGradeLevelDropdown();
    populateLaoMonthDropdowns();
    populateAcademicYearDropdowns();
  }

  function populateDistrictDropdowns() {
    const districts = ["ເມືອງໄຊ", "ເມືອງຫຼາ", "ເມືອງນາໝໍ້", "ເມືອງງາ", "ເມືອງແບງ", "ເມືອງຮຸນ", "ເມືອງປາກແບງ"];

    if (districtSelect) {
      districtSelect.disabled = false;
      const curVal = districtSelect.value;
      districtSelect.innerHTML = '<option value="">-- ເລືອກເມືອງ --</option>';
      districts.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        if (d === curVal) opt.selected = true;
        districtSelect.appendChild(opt);
      });
    }

    if (filterDistrict) {
      filterDistrict.innerHTML = '<option value="">ທຸກໆເມືອງ (7 ເມືອງ)</option>';
      districts.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        filterDistrict.appendChild(opt);
      });
    }

    if (unsubmittedFilterDistrict) {
      unsubmittedFilterDistrict.innerHTML = '<option value="">ທຸກໆເມືອງ (7 ເມືອງ)</option>';
      districts.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        unsubmittedFilterDistrict.appendChild(opt);
      });
    }
  }

  function populateSchoolDropdown(selectedDistrict = "") {
    if (!schoolSelect) return;

    if (!selectedDistrict) {
      schoolSelect.disabled = true;
      schoolSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກເມືອງກ່ອນ --</option>';
      if (schoolCodeInput) schoolCodeInput.value = "";
      resetLevelAndGradeDropdowns();
      return;
    }

    schoolSelect.disabled = false;
    const curVal = schoolSelect.value;
    schoolSelect.innerHTML = '<option value="">-- ເລືອກໂຮງຮຽນ --</option>';

    let list = [];
    if (typeof DEFAULT_LAO_DISTRICTS !== 'undefined' && DEFAULT_LAO_DISTRICTS[selectedDistrict]) {
      list = [...DEFAULT_LAO_DISTRICTS[selectedDistrict]];
    }

    list = Array.from(new Set(list));

    if (list.length === 0) {
      schoolSelect.disabled = true;
      schoolSelect.innerHTML = '<option value="">-- ບໍ່ມີໂຮງຮຽນໃນເມືອງນີ້ --</option>';
      resetLevelAndGradeDropdowns();
      return;
    }

    list.forEach(sch => {
      const opt = document.createElement('option');
      opt.value = sch;
      opt.textContent = sch;
      if (sch === curVal) opt.selected = true;
      schoolSelect.appendChild(opt);
    });

    if (!curVal) {
      resetLevelAndGradeDropdowns();
    }
  }

  function resetLevelAndGradeDropdowns() {
    if (educationLevelSelect) {
      educationLevelSelect.disabled = true;
      educationLevelSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກໂຮງຮຽນກ່ອນ --</option>';
    }
    if (gradeLevelSelect) {
      gradeLevelSelect.disabled = true;
      gradeLevelSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກໂຮງຮຽນກ່ອນ --</option>';
    }
  }

  function populateEducationLevelDropdown() {
    if (!educationLevelSelect) return;
    if (educationLevelSelect.disabled) return;
    educationLevelSelect.innerHTML = `
      <option value="">-- ເລືອກຊັ້ນຮຽນ --</option>
      <option value="ປະຖົມສຶກສາ">ປະຖົມສຶກສາ</option>
      <option value="ມັດທະຍົມສຶກສາຕອນຕົ້ນ">ມັດທະຍົມສຶກສາຕອນຕົ້ນ</option>
      <option value="ມັດທະຍົມສົມບູນ">ມັດທະຍົມສົມບູນ</option>
    `;
  }

  function populateGradeLevelDropdown(selectedLevel = "") {
    if (!gradeLevelSelect) return;

    if (!selectedLevel) {
      gradeLevelSelect.disabled = true;
      gradeLevelSelect.innerHTML = '<option value="">-- ກະລຸນາເລືອກຊັ້ນຮຽນກ່ອນ --</option>';
      return;
    }

    gradeLevelSelect.disabled = false;
    const curVal = gradeLevelSelect.value;
    gradeLevelSelect.innerHTML = '<option value="">-- ເລືອກຂັ້ນຮຽນ --</option>';

    let grades = [];
    if (typeof EDUCATION_LEVELS !== 'undefined' && EDUCATION_LEVELS[selectedLevel]) {
      grades = [...EDUCATION_LEVELS[selectedLevel]];
    } else {
      if (selectedLevel === "ປະຖົມສຶກສາ") grades = ["ປ.1", "ປ.2", "ປ.3", "ປ.4", "ປ.5"];
      else if (selectedLevel === "ມັດທະຍົມສຶກສາຕອນຕົ້ນ") grades = ["ມ.1", "ມ.2", "ມ.3", "ມ.4"];
      else if (selectedLevel === "ມັດທະຍົມສົມບູນ") grades = ["ມ.1", "ມ.2", "ມ.3", "ມ.4", "ມ.5", "ມ.6", "ມ.7"];
    }

    grades = Array.from(new Set(grades));

    grades.forEach(g => {
      const opt = document.createElement('option');
      opt.value = g;
      opt.textContent = g;
      if (g === curVal) opt.selected = true;
      gradeLevelSelect.appendChild(opt);
    });
  }

  function populateLaoMonthDropdowns() {
    const months = [
      { value: "01", name: "ມັງກອນ" },
      { value: "02", name: "ກຸມພາ" },
      { value: "03", name: "ມີນາ" },
      { value: "04", name: "ເມສາ" },
      { value: "05", name: "ພຶດສະພາ" },
      { value: "06", name: "ມິຖຸນາ" },
      { value: "07", name: "ກໍລະກົດ" },
      { value: "08", name: "ສິງຫາ" },
      { value: "09", name: "ກັນຍາ" },
      { value: "10", name: "ຕຸລາ" },
      { value: "11", name: "ພະຈິກ" },
      { value: "12", name: "ທັນວາ" }
    ];

    if (entryLaoMonthSelect) {
      const cur = entryLaoMonthSelect.value || String(new Date().getMonth() + 1).padStart(2, '0');
      entryLaoMonthSelect.innerHTML = '<option value="">-- ເລືອກເດືອນ --</option>';
      months.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.name;
        opt.textContent = m.name;
        if (m.value === cur || m.name === cur) opt.selected = true;
        entryLaoMonthSelect.appendChild(opt);
      });
    }

    if (filterMonth) {
      filterMonth.innerHTML = '<option value="">ທຸກໆເດືອນ</option>';
      months.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.name;
        opt.textContent = m.name;
        filterMonth.appendChild(opt);
      });
    }
  }

  function populateAcademicYearDropdowns() {
    const years = ["2026-2027", "2027-2028", "2028-2029", "2029-2030", "2030-2031"];

    if (entryAcademicYearSelect) {
      const cur = entryAcademicYearSelect.value || "2026-2027";
      entryAcademicYearSelect.innerHTML = '';
      years.forEach(y => {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        if (y === cur) opt.selected = true;
        entryAcademicYearSelect.appendChild(opt);
      });
    }

    if (filterAcademicYear) {
      filterAcademicYear.innerHTML = '<option value="">ທຸກໆສົກຮຽນ</option>';
      years.forEach(y => {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        filterAcademicYear.appendChild(opt);
      });
    }
  }

  // School Code Lookup Helper
  function lookupSchoolCode(schName, district = '') {
    if (!schName) return '';
    const clean = schName.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();

    if (typeof DEFAULT_SCHOOL_CODES !== 'undefined') {
      if (district && DEFAULT_SCHOOL_CODES[district.trim() + '_' + clean]) {
        return DEFAULT_SCHOOL_CODES[district.trim() + '_' + clean];
      }
      if (DEFAULT_SCHOOL_CODES[clean]) {
        return DEFAULT_SCHOOL_CODES[clean];
      }
    }

    const noPrefix = clean.replace(/^(ປະຖົມ\s*ປະຖົມ|ປະຖົມ|ປ\.ຖ|ມ\.ສ|ມ\.ຕ)\s*/g, '').trim();
    if (typeof DEFAULT_SCHOOL_CODES !== 'undefined') {
      if (district) {
        const distPrefix = district.trim() + '_';
        for (const [key, code] of Object.entries(DEFAULT_SCHOOL_CODES)) {
          if (key.startsWith(distPrefix)) {
            const kName = key.slice(distPrefix.length).replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/^(ປະຖົມ\s*ປະຖົມ|ປະຖົມ|ປ\.ຖ|ມ\.ສ|ມ\.ຕ)\s*/g, '').trim();
            if (kName === noPrefix && code) return code;
          }
        }
      }
      for (const [key, code] of Object.entries(DEFAULT_SCHOOL_CODES)) {
        const keyClean = key.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/^(ປະຖົມ\s*ປະຖົມ|ປະຖົມ|ປ\.ຖ|ມ\.ສ|ມ\.ຕ)\s*/g, '').trim();
        if (keyClean === noPrefix && code) return code;
      }
    }
    return '';
  }

  // Event Listeners setup
  function setupEventListeners() {
    if (districtSelect) {
      districtSelect.addEventListener('change', (e) => {
        populateSchoolDropdown(e.target.value);
      });
    }

    if (schoolSelect) {
      schoolSelect.addEventListener('change', (e) => {
        const schName = e.target.value;
        if (!schName) {
          if (schoolCodeInput) schoolCodeInput.value = "";
          resetLevelAndGradeDropdowns();
          return;
        }

        if (schoolCodeInput) schoolCodeInput.value = lookupSchoolCode(schName);

        // Auto select & restrict Education Level & Grade Level based on selected school
        if (educationLevelSelect) {
          educationLevelSelect.disabled = false;

          if (schName.includes('ປ.ຖ') || schName.includes('ປະຖົມ')) {
            educationLevelSelect.innerHTML = '<option value="ປະຖົມສຶກສາ">ປະຖົມສຶກສາ</option>';
            educationLevelSelect.value = "ປະຖົມສຶກສາ";
            populateGradeLevelDropdown("ປະຖົມສຶກສາ");
          } else if (schName.includes('ມ.ຕ')) {
            educationLevelSelect.innerHTML = '<option value="ມັດທະຍົມສຶກສາຕອນຕົ້ນ">ມັດທະຍົມສຶກສາຕອນຕົ້ນ</option>';
            educationLevelSelect.value = "ມັດທະຍົມສຶກສາຕອນຕົ້ນ";
            populateGradeLevelDropdown("ມັດທະຍົມສຶກສາຕອນຕົ້ນ");
          } else if (schName.includes('ມ.ສ') || schName.includes('ມັດທະຍົມສົມບູນ')) {
            educationLevelSelect.innerHTML = '<option value="ມັດທະຍົມສົມບູນ">ມັດທະຍົມສົມບູນ</option>';
            educationLevelSelect.value = "ມັດທະຍົມສົມບູນ";
            populateGradeLevelDropdown("ມັດທະຍົມສົມບູນ");
          } else {
            educationLevelSelect.innerHTML = `
              <option value="">-- ເລືອກຊັ້ນຮຽນ --</option>
              <option value="ປະຖົມສຶກສາ">ປະຖົມສຶກສາ</option>
              <option value="ມັດທະຍົມສຶກສາຕອນຕົ້ນ">ມັດທະຍົມສຶກສາຕອນຕົ້ນ</option>
              <option value="ມັດທະຍົມສົມບູນ">ມັດທະຍົມສົມບູນ</option>
            `;
            populateGradeLevelDropdown();
          }
        }
      });
    }

    if (educationLevelSelect) {
      educationLevelSelect.addEventListener('change', (e) => {
        populateGradeLevelDropdown(e.target.value);
      });
    }

    if (entryLaoMonthSelect) {
      entryLaoMonthSelect.addEventListener('change', updateMonthDynamicLabels);
    }

    // Number Inputs Auto Calculation & Validation
    const numInputs = [
      passedRegisteredTotal, passedRegisteredFemale,
      augustPrevYearTotal, augustPrevYearFemale,
      transferInTotal, transferInFemale,
      transferOutTotal, transferOutFemale,
      dropoutTotal, dropoutFemale,
      repeaterTotal, repeaterFemale
    ];

    numInputs.forEach(inp => {
      if (inp) {
        inp.addEventListener('input', () => {
          calculateActualAttending();
          validateAugustCounts();
          validateFemaleCounts();
        });
      }
    });

    if (recalculateBtn) recalculateBtn.addEventListener('click', calculateActualAttending);

    if (studentDataForm) studentDataForm.addEventListener('submit', handleFormSubmit);
    if (resetFormBtn) resetFormBtn.addEventListener('click', resetForm);

    // Dashboard Filters
    if (searchInput) searchInput.addEventListener('input', renderTable);
    if (filterDistrict) filterDistrict.addEventListener('change', renderTable);
    if (filterLevel) filterLevel.addEventListener('change', renderTable);
    if (filterGrade) filterGrade.addEventListener('change', renderTable);
    if (filterMonth) filterMonth.addEventListener('change', renderTable);
    if (filterAcademicYear) filterAcademicYear.addEventListener('change', renderTable);
    if (resetFiltersBtn) resetFiltersBtn.addEventListener('click', resetSearchFilters);
    if (unsubmittedFilterDistrict) unsubmittedFilterDistrict.addEventListener('change', renderSubmissionStats);
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
  }

  // Automatic Formula Calculation
  function calculateActualAttending() {
    const isAugust = entryLaoMonthSelect && ((entryLaoMonthSelect.value || "").trim() === "08" || (entryLaoMonthSelect.value || "").trim() === "ສິງຫາ");

    const trInTot = parseInt(transferInTotal?.value) || 0;
    const trInFem = parseInt(transferInFemale?.value) || 0;
    const trOutTot = parseInt(transferOutTotal?.value) || 0;
    const trOutFem = parseInt(transferOutFemale?.value) || 0;
    const dropTot  = parseInt(dropoutTotal?.value)    || 0;
    const dropFem  = parseInt(dropoutFemale?.value)   || 0;

    let calcTot = 0;
    let calcFem = 0;

    if (isAugust) {
      // ສູດ: ລົງທະບຽນ (8) + ຄ້າງຫ້ອງ (9) + ຍ້າຍເຂົ້າ (10) − ຍ້າຍອອກ (11) − ປະລະ (12)
      const regTot = parseInt(augustPrevYearTotal?.value) || 0;
      const regFem = parseInt(augustPrevYearFemale?.value) || 0;
      const repTot = parseInt(repeaterTotal?.value) || 0;
      const repFem = parseInt(repeaterFemale?.value) || 0;
      calcTot = Math.max(0, regTot + repTot + trInTot - trOutTot - dropTot);
      calcFem = Math.max(0, regFem + repFem + trInFem - trOutFem - dropFem);
    } else {
      // ສູດ: ທ້າຍເດືອນ + ຍ້າຍເຂົ້າ − ຍ້າຍອອກ − ປະລະ
      const regTot = parseInt(passedRegisteredTotal?.value)  || 0;
      const regFem = parseInt(passedRegisteredFemale?.value) || 0;
      calcTot = Math.max(0, regTot + trInTot - trOutTot - dropTot);
      calcFem = Math.max(0, regFem + trInFem - trOutFem - dropFem);
    }

    if (actualAttendingTotal)  actualAttendingTotal.value  = calcTot;
    if (actualAttendingFemale) actualAttendingFemale.value = calcFem;
  }

  // Form Submit Handler
  function handleFormSubmit(e) {
    e.preventDefault();

    if (!districtSelect?.value || !schoolSelect?.value || !collectorNameInput?.value || !educationLevelSelect?.value || !gradeLevelSelect?.value) {
      showToast("ກະລຸນາ ປ້ອນຂໍ້ມູນທີ່ມີເຄື່ອງໝາຍ (*) ໃຫ້ຄົບຖ້ວນ!", "error");
      return;
    }

    if (!validateAugustCounts()) {
      showToast("ຢິງທ້າຍປີຜ່ານມາ ຕ້ອງບໍ່ເກີນ ນັກຮຽນທ້າຍປີຜ່ານມາ (ຊ່ອງ 7)!", "error");
      return;
    }

    if (!validateFemaleCounts()) {
      showToast("ຈຳນວນຍິງຕ້ອງບໍ່ເກີນຈຳນວນລວມທຸກຊ່ອງ!", "error");
      return;
    }

    calculateActualAttending();

    const recordId = recordIdInput?.value || ("REC-" + Date.now());

    const recordData = {
      id: recordId,
      entry_month_only: entryLaoMonthSelect?.value || "",
      entry_academic_year: entryAcademicYearSelect?.value || "",
      district: districtSelect.value,
      school: schoolSelect.value,
      school_code: schoolCodeInput?.value || lookupSchoolCode(schoolSelect.value),
      collector_name: collectorNameInput.value,
      education_level: educationLevelSelect.value,
      grade_level: gradeLevelSelect.value,
      passed_registered_total: parseInt(passedRegisteredTotal?.value) || 0,
      passed_registered_female: parseInt(passedRegisteredFemale?.value) || 0,
      august_prev_year_total: parseInt(augustPrevYearTotal?.value) || 0,
      august_prev_year_female: parseInt(augustPrevYearFemale?.value) || 0,
      transfer_in_total: parseInt(transferInTotal?.value) || 0,
      transfer_in_female: parseInt(transferInFemale?.value) || 0,
      transfer_out_total: parseInt(transferOutTotal?.value) || 0,
      transfer_out_female: parseInt(transferOutFemale?.value) || 0,
      transfer_out_reason: transferOutReason?.value || "",
      dropout_total: parseInt(dropoutTotal?.value) || 0,
      dropout_female: parseInt(dropoutFemale?.value) || 0,
      dropout_reason: dropoutReason?.value || "",
      repeater_total: parseInt(repeaterTotal?.value) || 0,
      repeater_female: parseInt(repeaterFemale?.value) || 0,
      actual_attending_total: parseInt(actualAttendingTotal?.value) || 0,
      actual_attending_female: parseInt(actualAttendingFemale?.value) || 0,
      created_at: new Date().toISOString().replace('T', ' ').substring(0, 16)
    };

    const existingIdx = records.findIndex(r => r.id === recordId);
    if (existingIdx >= 0) {
      records[existingIdx] = recordData;
      showToast("ອັບເດດຂໍ້ມູນນັກຮຽນ ສຳເລັດແລ້ວ!", "success");
    } else {
      records.unshift(recordData);
      showToast("ບັນທຶກຂໍ້ມູນນັກຮຽນໃໝ່ ສຳເລັດແລ້ວ!", "success");
    }

    saveRecordsToStorage();
    postRecordToGoogleSheet(recordData);
    resetForm();
    renderApp();
  }

  // Reset Form
  function resetForm() {
    if (studentDataForm) studentDataForm.reset();
    if (recordIdInput) recordIdInput.value = "";
    if (formModeBadge) {
      formModeBadge.textContent = "ບັນທຶກຂໍ້ມູນໃໝ່";
      formModeBadge.className = "badge-tag";
    }
    populateAllDropdowns();
    calculateActualAttending();
  }

  // Render Dashboard & Tables
  function renderApp() {
    renderKPI();
    renderTable();
    renderSubmissionStats();
    renderReasonSummaries();
    renderCharts();
    renderAdvancedCharts();
  }

  function renderKPI() {
    let totRegistered = 0, femRegistered = 0;
    let totStudents   = 0, femStudents   = 0;
    let totDrop = 0, femDrop = 0;
    let totTrOut = 0, femTrOut = 0;
    let totTrIn  = 0, femTrIn  = 0;
    let totRep   = 0, femRep   = 0;

    // Filter records according to active search/filters if selected
    const query = (searchInput?.value || "").toLowerCase();
    const fDist = filterDistrict?.value || "";
    const fLvl  = filterLevel?.value || "";
    const fGrd  = filterGrade?.value || "";
    const fMth  = filterMonth?.value || "";
    const fYr   = filterAcademicYear?.value || "";

    const activeRecs = records.filter(r => {
      if (fDist && r.district !== fDist) return false;
      if (fLvl  && r.education_level !== fLvl) return false;
      if (fGrd  && r.grade_level !== fGrd) return false;
      if (fMth  && r.entry_month_only !== fMth) return false;
      if (fYr   && r.entry_academic_year !== fYr) return false;
      if (query) {
        const text = `${r.school} ${r.school_code} ${r.collector_name} ${r.district}`.toLowerCase();
        if (!text.includes(query)) return false;
      }
      return true;
    });

    activeRecs.forEach(r => {
      const isAug = (r.entry_month_only || "").trim() === "ສິງຫາ" || (r.entry_month_only || "").trim() === "08";

      // 1. Registered: Fallback across all possible registered fields
      const regTot = isAug
        ? (parseInt(r.august_prev_year_total) || parseInt(r.passed_registered_total) || parseInt(r.registered_total) || 0)
        : (parseInt(r.passed_registered_total) || parseInt(r.august_prev_year_total) || parseInt(r.registered_total) || 0);

      const regFem = isAug
        ? (parseInt(r.august_prev_year_female) || parseInt(r.passed_registered_female) || parseInt(r.registered_female) || 0)
        : (parseInt(r.passed_registered_female) || parseInt(r.august_prev_year_female) || parseInt(r.registered_female) || 0);

      totRegistered += regTot;
      femRegistered += regFem;

      // 2. Transfer In
      totTrIn  += parseInt(r.transfer_in_total)  || 0;
      femTrIn  += parseInt(r.transfer_in_female) || 0;

      // 3. Transfer Out
      totTrOut += parseInt(r.transfer_out_total)  || 0;
      femTrOut += parseInt(r.transfer_out_female) || 0;

      // 4. Dropout
      totDrop  += parseInt(r.dropout_total)       || 0;
      femDrop  += parseInt(r.dropout_female)      || 0;

      // 5. Repeater
      totRep   += parseInt(r.repeater_total)       || 0;
      femRep   += parseInt(r.repeater_female)      || 0;

      // 6. Actual Attending
      totStudents   += parseInt(r.actual_attending_total)  || 0;
      femStudents   += parseInt(r.actual_attending_female) || 0;
    });

    const pct = (f, t) => t > 0 ? ((f / t) * 100).toFixed(2) : '0.00';
    const setEl = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };

    // Card 1: ລົງທະບຽນ
    setEl('statRegisteredTotal',   totRegistered.toLocaleString());
    setEl('statRegisteredFemale',  femRegistered.toLocaleString());
    setEl('statRegisteredPercent', pct(femRegistered, totRegistered));

    // Card 2: ຍ້າຍເຂົ້າ
    setEl('statTransferInTotal',   totTrIn.toLocaleString());
    setEl('statTransferInFemale',  femTrIn.toLocaleString());
    setEl('statTransferInPercent', pct(femTrIn, totTrIn));

    // Card 3: ຍ້າຍອອກ
    setEl('statTransferOutTotal',   totTrOut.toLocaleString());
    setEl('statTransferOutFemale',  femTrOut.toLocaleString());
    setEl('statTransferOutPercent', pct(femTrOut, totTrOut));

    // Card 4: ປະລະ
    setEl('statTotalDropout',   totDrop.toLocaleString());
    setEl('statDropoutFemale',  femDrop.toLocaleString());
    setEl('statDropoutPercent', pct(femDrop, totDrop));

    // Card 5: ຄ້າງຫ້ອງ
    setEl('statRepeaterTotal',   totRep.toLocaleString());
    setEl('statRepeaterFemale',  femRep.toLocaleString());
    setEl('statRepeaterPercent', pct(femRep, totRep));

    // Card 6: ໜ້າຮຽນຕົວຈິງ
    setEl('statTotalStudents',  totStudents.toLocaleString());
    setEl('statFemaleStudents', femStudents.toLocaleString());
    setEl('statFemalePercent',  pct(femStudents, totStudents));
  }

  function renderTable() {
    if (!tableBody) return;

    const query = (searchInput?.value || "").toLowerCase();
    const fDist = filterDistrict?.value || "";
    const fLvl  = filterLevel?.value || "";
    const fGrd  = filterGrade?.value || "";
    const fMth  = filterMonth?.value || "";
    const fYr   = filterAcademicYear?.value || "";

    const filtered = records.filter(r => {
      if (fDist && r.district !== fDist) return false;
      if (fLvl  && r.education_level !== fLvl) return false;
      if (fGrd  && r.grade_level !== fGrd) return false;
      if (fMth  && r.entry_month_only !== fMth) return false;
      if (fYr   && r.entry_academic_year !== fYr) return false;
      if (query) {
        const text = `${r.school} ${r.school_code} ${r.collector_name} ${r.district}`.toLowerCase();
        if (!text.includes(query)) return false;
      }
      return true;
    });

    if (tableRecordCount) tableRecordCount.textContent = filtered.length;

    tableBody.innerHTML = '';
    if (filtered.length === 0) {
      tableBody.innerHTML = '<tr><td colspan="13" style="text-align:center; padding: 2rem; color: var(--text-muted);">ບໍ່ພົບຂໍ້ມູນຕາມເງື່ອນໄຂທີ່ຄົ້ນຫາ</td></tr>';
      return;
    }

    const isAugustRec = r => (r.entry_month_only || "").trim() === "ສິງຫາ" || (r.entry_month_only || "").trim() === "08";

    filtered.forEach((r, idx) => {
      // ຄຳນວນ "ນັກຮຽນຕົ້ນ" ຕາມເດືອນ
      const baseLabel = isAugustRec(r) ? "ລົງທະບຽນ" : "ທ້າຍເດືອນ";
      const baseTotal = isAugustRec(r)
        ? (r.august_prev_year_total || 0)
        : (r.passed_registered_total || 0);
      const baseFemale = isAugustRec(r)
        ? (r.august_prev_year_female || 0)
        : (r.passed_registered_female || 0);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><strong>${r.district}</strong><br><small style="color:var(--text-muted);">${r.entry_month_only || ''} ${r.entry_academic_year || ''}</small></td>
        <td>${r.school}<br><small style="color:var(--text-muted); font-family:monospace;">${r.school_code || ''}</small></td>
        <td><strong>${r.grade_level}</strong></td>
        <td>${r.collector_name}</td>
        <td title="${baseLabel}"><strong>${baseTotal.toLocaleString()}</strong><br><small style="color:var(--accent-pink);">ຍິງ: ${baseFemale}</small></td>
        <td title="ຍ້າຍເຂົ້າ">${r.transfer_in_total || 0}<br><small style="color:var(--accent-pink);">ຍິງ: ${r.transfer_in_female || 0}</small></td>
        <td title="ຍ້າຍອອກ">${r.transfer_out_total || 0}<br><small style="color:var(--accent-pink);">ຍິງ: ${r.transfer_out_female || 0}</small></td>
        <td title="ປະລະ">${r.dropout_total || 0}<br><small style="color:var(--accent-pink);">ຍິງ: ${r.dropout_female || 0}</small></td>
        <td title="ຄ້າງຫ້ອງ">${isAugustRec(r) ? (r.repeater_total || 0) : '-'}<br>${isAugustRec(r) ? `<small style="color:var(--accent-pink);">ຍິງ: ${r.repeater_female || 0}</small>` : ''}</td>
        <td><strong style="color:var(--accent-success);">${(r.actual_attending_total || 0).toLocaleString()}</strong><br><small style="color:var(--accent-pink);">ຍິງ: ${r.actual_attending_female || 0}</small></td>
        <td><small>${r.created_at || ''}</small></td>
      `;
      tableBody.appendChild(tr);
    });
  }

  function editRecord(id) {
    const rec = records.find(r => r.id === id);
    if (!rec) return;

    if (recordIdInput) recordIdInput.value = rec.id;
    if (formModeBadge) {
      formModeBadge.textContent = "ແກ້ໄຂຂໍ້ມູນ";
      formModeBadge.className = "badge-tag badge-warning";
    }

    if (districtSelect) districtSelect.value = rec.district;
    populateSchoolDropdown(rec.district);
    if (schoolSelect) schoolSelect.value = rec.school;
    if (schoolCodeInput) schoolCodeInput.value = rec.school_code || lookupSchoolCode(rec.school);
    if (collectorNameInput) collectorNameInput.value = rec.collector_name;
    if (educationLevelSelect) educationLevelSelect.value = rec.education_level;
    populateGradeLevelDropdown(rec.education_level);
    if (gradeLevelSelect) gradeLevelSelect.value = rec.grade_level;
    if (entryLaoMonthSelect) entryLaoMonthSelect.value = rec.entry_month_only;
    if (entryAcademicYearSelect) entryAcademicYearSelect.value = rec.entry_academic_year;

    updateMonthDynamicLabels();

    if (passedRegisteredTotal) passedRegisteredTotal.value = rec.passed_registered_total;
    if (passedRegisteredFemale) passedRegisteredFemale.value = rec.passed_registered_female;
    if (augustPrevYearTotal) augustPrevYearTotal.value = rec.august_prev_year_total || 0;
    if (augustPrevYearFemale) augustPrevYearFemale.value = rec.august_prev_year_female || 0;
    if (transferInTotal) transferInTotal.value = rec.transfer_in_total;
    if (transferInFemale) transferInFemale.value = rec.transfer_in_female;
    if (transferOutTotal) transferOutTotal.value = rec.transfer_out_total;
    if (transferOutFemale) transferOutFemale.value = rec.transfer_out_female;
    if (transferOutReason) transferOutReason.value = rec.transfer_out_reason || "";
    if (dropoutTotal) dropoutTotal.value = rec.dropout_total;
    if (dropoutFemale) dropoutFemale.value = rec.dropout_female;
    if (dropoutReason) dropoutReason.value = rec.dropout_reason || "";
    if (repeaterTotal) repeaterTotal.value = rec.repeater_total;
    if (repeaterFemale) repeaterFemale.value = rec.repeater_female;

    calculateActualAttending();

    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast("ດຶງຂໍ້ມູນມາແກ້ໄຂແລ້ວ!", "info");
  }

  function deleteRecord(id) {
    if (!confirm("ທ່ານແນ່ໃຈບໍ່ວ່າຕ້ອງການລົບຂໍ້ມູນນີ້?")) return;
    records = records.filter(r => r.id !== id);
    saveRecordsToStorage();
    renderApp();
    showToast("ລົບຂໍ້ມູນສຳເລັດແລ້ວ!", "warning");
  }

  function resetSearchFilters() {
    if (searchInput) searchInput.value = "";
    if (filterDistrict) filterDistrict.value = "";
    if (filterLevel) filterLevel.value = "";
    if (filterGrade) filterGrade.value = "";
    if (filterMonth) filterMonth.value = "";
    if (filterAcademicYear) filterAcademicYear.value = "";
    renderTable();
  }

  // Normalize school name: keep level prefix, only strip zero-width chars and standardize ປະຖົມ→ປ.ຖ
  function normalizeSchoolName(name) {
    if (!name) return '';
    return name
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/^ປະຖົມ\s*/gi, 'ປ.ຖ ')
      .trim()
      .toLowerCase();
  }

  // School Submission Progress Tracking
  function renderSubmissionStats() {
    if (typeof DEFAULT_LAO_DISTRICTS === 'undefined') return;

    const fDist = filterDistrict?.value || unsubmittedFilterDistrict?.value || "";
    const fMth  = filterMonth?.value || "";
    const fYr   = filterAcademicYear?.value || "";

    // 1. Filter records by selected Month & Academic Year if specified
    const activeRecords = records.filter(r => {
      if (fMth && r.entry_month_only !== fMth) return false;
      if (fYr  && r.entry_academic_year !== fYr) return false;
      return true;
    });

    // 2. Build set of submitted school keys: "District_NormalizedName"
    const submittedKeys = new Set();
    activeRecords.forEach(r => {
      if (r.district && r.school) {
        const key = r.district.trim() + "_" + normalizeSchoolName(r.school);
        submittedKeys.add(key);
      }
    });

    // 3. Categorize all master schools into: fullySubmitted, partiallySubmitted, unsubmitted
    let totalMasterSchools = 0;
    let submittedCount = 0;
    const districtStats = {};
    const fullySubmittedList = [];
    const partiallySubmittedList = [];
    const unsubmittedList = [];

    Object.entries(DEFAULT_LAO_DISTRICTS).forEach(([dist, schs]) => {
      districtStats[dist] = { total: schs.length, submitted: 0 };
      schs.forEach(sch => {
        totalMasterSchools++;
        const normSch = normalizeSchoolName(sch);
        
        // Find submitted records for this specific school
        const schRecs = activeRecords.filter(r => r.district === dist && normalizeSchoolName(r.school) === normSch);
        const submittedGrades = new Set(schRecs.map(r => r.grade_level || 'General'));
        const gradeCount = submittedGrades.size;

        // Determine max grades for school type
        let maxGrades = 5;
        if (sch.startsWith('ມ.ສ')) maxGrades = 7;
        else if (sch.startsWith('ມ.ຕ')) maxGrades = 4;
        else if (sch.startsWith('ປະຖົມ') || sch.startsWith('ປ.ຖ')) maxGrades = 5;

        const schoolItem = {
          district: dist,
          name: sch,
          code: lookupSchoolCode(sch, dist),
          gradeCount: gradeCount,
          maxGrades: maxGrades
        };

        const isFilteredOut = fDist && dist !== fDist;

        if (gradeCount >= maxGrades || (gradeCount > 0 && schRecs.length >= maxGrades)) {
          submittedCount++;
          districtStats[dist].submitted++;
          if (!isFilteredOut) fullySubmittedList.push(schoolItem);
        } else if (gradeCount > 0) {
          submittedCount++;
          districtStats[dist].submitted++;
          if (!isFilteredOut) partiallySubmittedList.push(schoolItem);
        } else {
          if (!isFilteredOut) unsubmittedList.push(schoolItem);
        }
      });
    });

    const unsubmittedCount = Math.max(0, totalMasterSchools - submittedCount);
    const rate = totalMasterSchools > 0 ? ((submittedCount / totalMasterSchools) * 100).toFixed(2) : '0.00';

    // Update KPI Card Elements
    const setEl = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    setEl('statTotalSchoolsCount',   totalMasterSchools);
    setEl('statSchoolsSubmitted',    submittedCount);
    setEl('statSchoolsSubmittedPct', rate);
    setEl('statSchoolsNotSubmitted', unsubmittedCount);

    if (kpiTotalMasterSchools) kpiTotalMasterSchools.textContent = totalMasterSchools;
    if (kpiSubmittedSchools)   kpiSubmittedSchools.textContent   = submittedCount;
    if (kpiUnsubmittedSchools) kpiUnsubmittedSchools.textContent = unsubmittedCount;
    if (kpiSubmissionRate)     kpiSubmissionRate.textContent     = `${rate}%`;
    if (overallProgressBar)    overallProgressBar.style.width    = `${parseFloat(rate)}%`;

    // 4. District Breakdown Progress Bars
    const breakdownEl = document.getElementById('schoolSubmittedBreakdown');
    if (breakdownEl) {
      breakdownEl.innerHTML = Object.entries(districtStats)
        .filter(([dist]) => !fDist || dist === fDist)
        .map(([dist, s]) => {
          const pct = s.total > 0 ? ((s.submitted / s.total) * 100).toFixed(2) : '0.00';
          const pctNum = parseFloat(pct);
          const color = pctNum >= 80 ? 'var(--accent-success)' : pctNum >= 50 ? 'var(--accent-warning)' : 'var(--accent-danger)';
          return `<div style="margin-bottom: 12px;">
            <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:4px;">
              <span><strong>${dist}</strong></span>
              <span style="color:${color}; font-weight:600;">${s.submitted}/${s.total} ໂຮງຮຽນ (${pct}%)</span>
            </div>
            <div style="background:var(--border-color); border-radius:4px; height:8px; overflow:hidden;">
              <div style="width:${pctNum}%; height:100%; background:${color}; border-radius:4px; transition:width 0.4s ease;"></div>
            </div>
          </div>`;
        }).join('');
    }

    // Helper to render school status card lists matching user requested screenshot UI
    function renderSchoolStatusCardList(containerEl, badgeEl, list, emptyMsg, statusType) {
      if (badgeEl) badgeEl.textContent = list.length + ' ໂຮງຮຽນ';
      if (!containerEl) return;

      if (list.length === 0) {
        containerEl.innerHTML = `<p style="color: var(--text-muted); font-size: 0.85rem; text-align: center; padding: 25px 10px;">${emptyMsg}</p>`;
        return;
      }

      const isFull = statusType === 'full';
      const isPartial = statusType === 'partial';

      const borderColor = isFull ? '#10b981' : isPartial ? '#f59e0b' : '#ef4444';
      const badgeBg = isFull ? 'rgba(16, 185, 129, 0.12)' : isPartial ? 'rgba(245, 158, 11, 0.12)' : 'rgba(239, 68, 68, 0.12)';
      const badgeColor = isFull ? '#10b981' : isPartial ? '#d97706' : '#dc2626';
      const statusLabel = isFull ? 'ປ້ອນຄົບ' : isPartial ? 'ປ້ອນບາງສ່ວນ' : 'ຍັງບໍ່ປ້ອນ';

      containerEl.innerHTML = list.map((item, idx) => `
        <div style="
          background: var(--bg-card);
          border: 1px solid var(--border-color);
          border-left: 3px solid ${borderColor};
          border-radius: var(--radius-md);
          padding: 10px 14px;
          margin-bottom: 8px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        ">
          <div style="display: flex; align-items: center; gap: 12px; overflow: hidden;">
            <div style="
              width: 30px;
              height: 30px;
              border-radius: 50%;
              background: rgba(128,128,128,0.12);
              color: var(--text-main);
              display: flex;
              align-items: center;
              justify-content: center;
              font-weight: 600;
              font-size: 0.82rem;
              flex-shrink: 0;
            ">${idx + 1}</div>
            <div style="overflow: hidden;">
              <h4 style="font-size: 0.92rem; font-weight: 600; color: var(--text-main); margin: 0 0 2px 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${item.name}</h4>
              <p style="font-size: 0.78rem; color: var(--text-muted); margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                ${item.district} • ${item.code || '-'} • ${item.gradeCount}/${item.maxGrades} ຂັ້ນ
              </p>
            </div>
          </div>
          <span style="
            background: ${badgeBg};
            color: ${badgeColor};
            border-radius: 20px;
            padding: 3px 12px;
            font-size: 0.78rem;
            font-weight: 600;
            white-space: nowrap;
            flex-shrink: 0;
          ">${statusLabel}</span>
        </div>
      `).join('');
    }

    // 5. Render 3 Panels matching user requested UI
    renderSchoolStatusCardList(
      document.getElementById('fullySubmittedList'),
      document.getElementById('fullySubmittedCountBadge'),
      fullySubmittedList,
      'ບໍ່ມີໂຮງຮຽນທີ່ປ້ອນຄົບຕາມ Filter ນີ້',
      'full'
    );

    renderSchoolStatusCardList(
      document.getElementById('partiallySubmittedList'),
      document.getElementById('partiallySubmittedCountBadge'),
      partiallySubmittedList,
      'ບໍ່ມີໂຮງຮຽນທີ່ປ້ອນບາງສ່ວນ',
      'partial'
    );

    renderSchoolStatusCardList(
      document.getElementById('schoolNotSubmittedBreakdown'),
      document.getElementById('unsubmittedCountBadge'),
      unsubmittedList,
      'ບໍ່ມີໂຮງຮຽນທີ່ຍັງບໍ່ປ້ອນ',
      'unsubmitted'
    );

    // Unsubmitted Table (index.html sidebar fallback)
    if (unsubmittedTableBody) {
      unsubmittedTableBody.innerHTML = '';
      if (unsubmittedCountBadge) unsubmittedCountBadge.textContent = unsubmittedList.length;
      if (unsubmittedList.length === 0) {
        unsubmittedTableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 1.5rem; color: var(--accent-success);">✅ ປ້ອນຂໍ້ມູນຄົບທຸກໂຮງຮຽນແລ້ວ!</td></tr>';
      } else {
        unsubmittedList.slice(0, 100).forEach((item, i) => {
          const tr = document.createElement('tr');
          tr.innerHTML = `<td>${i+1}</td><td><strong>${item.district}</strong></td><td>${item.name}</td><td><code style="color:var(--accent-warning);">${item.code || '-'}</code></td>`;
          unsubmittedTableBody.appendChild(tr);
        });
      }
    }
  }

  // Render Transfer Out & Dropout Reasons Summary Breakdown
  function renderReasonSummaries() {
    const trOutContainer = document.getElementById('transferOutReasonsList');
    const dropContainer  = document.getElementById('dropoutReasonsList');

    if (!trOutContainer && !dropContainer) return;

    // 1. Aggregate Transfer Out Reasons
    const trOutMap = {};
    let totalTrOutCount = 0;

    records.forEach(r => {
      const reason = (r.transfer_out_reason || "").trim() || "ບໍ່ໄດ້ລະບຸສາເຫດ";
      const count  = parseInt(r.transfer_out_total) || 0;
      if (count > 0 || (r.transfer_out_reason && r.transfer_out_reason.trim())) {
        const cnt = count > 0 ? count : 1;
        trOutMap[reason] = (trOutMap[reason] || 0) + cnt;
        totalTrOutCount += cnt;
      }
    });

    if (trOutContainer) {
      const entries = Object.entries(trOutMap).sort((a, b) => b[1] - a[1]);
      if (entries.length === 0) {
        trOutContainer.innerHTML = '<p style="color:var(--text-muted); text-align:center; padding:15px; font-size:0.85rem;">ບໍ່ມີຂໍ້ມູນການຍ້າຍອອກ</p>';
      } else {
        trOutContainer.innerHTML = entries.map(([reason, cnt]) => {
          const pct = totalTrOutCount > 0 ? ((cnt / totalTrOutCount) * 100).toFixed(2) : '0.00';
          return `
            <div style="margin-bottom: 12px;">
              <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:4px;">
                <span><i class="fa-solid fa-circle-dot" style="color:var(--accent-warning); font-size:0.75rem;"></i> <strong>${reason}</strong></span>
                <span style="color:var(--accent-warning); font-weight:600;">${cnt} ຄົນ (${pct}%)</span>
              </div>
              <div style="background:var(--border-color); border-radius:4px; height:8px; overflow:hidden;">
                <div style="width:${parseFloat(pct)}%; height:100%; background:var(--accent-warning); border-radius:4px; transition:width 0.4s ease;"></div>
              </div>
            </div>
          `;
        }).join('');
      }
    }

    // 2. Aggregate Dropout Reasons
    const dropMap = {};
    let totalDropCount = 0;

    records.forEach(r => {
      const reason = (r.dropout_reason || "").trim() || "ບໍ່ໄດ້ລະບຸສາເຫດ";
      const count  = parseInt(r.dropout_total) || 0;
      if (count > 0 || (r.dropout_reason && r.dropout_reason.trim())) {
        const cnt = count > 0 ? count : 1;
        dropMap[reason] = (dropMap[reason] || 0) + cnt;
        totalDropCount += cnt;
      }
    });

    if (dropContainer) {
      const entries = Object.entries(dropMap).sort((a, b) => b[1] - a[1]);
      if (entries.length === 0) {
        dropContainer.innerHTML = '<p style="color:var(--text-muted); text-align:center; padding:15px; font-size:0.85rem;">ບໍ່ມີຂໍ້ມູນການປະລະ</p>';
      } else {
        dropContainer.innerHTML = entries.map(([reason, cnt]) => {
          const pct = totalDropCount > 0 ? ((cnt / totalDropCount) * 100).toFixed(2) : '0.00';
          return `
            <div style="margin-bottom: 12px;">
              <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:4px;">
                <span><i class="fa-solid fa-circle-dot" style="color:var(--accent-danger); font-size:0.75rem;"></i> <strong>${reason}</strong></span>
                <span style="color:var(--accent-danger); font-weight:600;">${cnt} ຄົນ (${pct}%)</span>
              </div>
              <div style="background:var(--border-color); border-radius:4px; height:8px; overflow:hidden;">
                <div style="width:${parseFloat(pct)}%; height:100%; background:var(--accent-danger); border-radius:4px; transition:width 0.4s ease;"></div>
              </div>
            </div>
          `;
        }).join('');
      }
    }
  }

  // Dynamic Chart Rendering for Dashboard (Chart.js)
  let chartInstances = {};

  function renderCharts() {
    if (typeof Chart === 'undefined') {
      if (!window._chartLoading) {
        window._chartLoading = true;
        const script = document.createElement('script');
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js";
        script.onload = () => { renderCharts(); };
        document.head.appendChild(script);
      } else {
        setTimeout(renderCharts, 400);
      }
      return;
    }

    const monthMap = {
      "01": 0, "1": 0, "ມັງກອນ": 0, "january": 0,
      "02": 1, "2": 1, "ກຸມພາ": 1, "february": 1,
      "03": 2, "3": 2, "ມີນາ": 2, "march": 2,
      "04": 3, "4": 3, "ເມສາ": 3, "april": 3,
      "05": 4, "5": 4, "ພຶດສະພາ": 4, "may": 4,
      "06": 5, "6": 5, "ມິຖຸນາ": 5, "june": 5,
      "07": 6, "7": 6, "ກໍລະກົດ": 6, "july": 6,
      "08": 7, "8": 7, "ສິງຫາ": 7, "august": 7,
      "09": 8, "9": 8, "ກັນຍາ": 8, "september": 8,
      "10": 9, "ຕຸລາ": 9, "october": 9,
      "11": 10, "ພະຈິກ": 10, "november": 10,
      "12": 11, "ທັນວາ": 11, "december": 11
    };

    const safeDestroy = (canvas, key) => {
      try {
        const existing = Chart.getChart(canvas);
        if (existing) existing.destroy();
      } catch (e) {}
      if (chartInstances[key]) {
        try { chartInstances[key].destroy(); } catch (e) {}
        chartInstances[key] = null;
      }
    };

    // 1. Monthly Trend Chart
    try {
      const monthlyCanvas = document.getElementById('monthlyTrendChart');
      if (monthlyCanvas) {
        safeDestroy(monthlyCanvas, 'monthly');
        const months = ["ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມິຖຸນາ", "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"];
        const monthlyTotals = new Array(12).fill(0);
        const monthlyFemales = new Array(12).fill(0);

        records.forEach(r => {
          const rawM = (r.entry_month_only || "").trim().toLowerCase();
          const mIdx = monthMap[rawM] !== undefined ? monthMap[rawM] : -1;
          if (mIdx >= 0 && mIdx < 12) {
            monthlyTotals[mIdx] += parseInt(r.actual_attending_total) || 0;
            monthlyFemales[mIdx] += parseInt(r.actual_attending_female) || 0;
          }
        });

        chartInstances.monthly = new Chart(monthlyCanvas, {
          type: 'line',
          data: {
            labels: months,
            datasets: [
              {
                label: 'ໜ້າຮຽນຕົວຈິງທັງໝົດ',
                data: monthlyTotals,
                borderColor: '#3b82f6',
                backgroundColor: 'rgba(59, 130, 246, 0.15)',
                fill: true,
                tension: 0.35,
                borderWidth: 3
              },
              {
                label: 'ຍິງ',
                data: monthlyFemales,
                borderColor: '#ec4899',
                backgroundColor: 'rgba(236, 72, 153, 0.15)',
                fill: true,
                tension: 0.35,
                borderWidth: 2
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8' } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
              y: { beginAtZero: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
          }
        });
      }
    } catch (err) {
      console.warn('Monthly chart error:', err);
    }

    // 2. District Comparison Bar Chart
    try {
      const districtCanvas = document.getElementById('districtChart');
      if (districtCanvas) {
        safeDestroy(districtCanvas, 'district');
        const districts = ["ເມືອງໄຊ", "ເມືອງຫຼາ", "ເມືອງນາໝໍ້", "ເມືອງງາ", "ເມືອງແບງ", "ເມືອງຮຸນ", "ເມືອງປາກແບງ"];
        const distTotals = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.actual_attending_total) || 0), 0));
        const distFemales = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.actual_attending_female) || 0), 0));

        chartInstances.district = new Chart(districtCanvas, {
          type: 'bar',
          data: {
            labels: districts,
            datasets: [
              { label: 'ທັງໝົດ', data: distTotals, backgroundColor: '#38bdf8', borderRadius: 4 },
              { label: 'ຍິງ', data: distFemales, backgroundColor: '#f472b6', borderRadius: 4 }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8' } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
              y: { beginAtZero: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
          }
        });
      }
    } catch (err) {
      console.warn('District chart error:', err);
    }

    // 3. Gender Ratio Doughnut Chart
    try {
      const genderCanvas = document.getElementById('genderChart');
      if (genderCanvas) {
        safeDestroy(genderCanvas, 'gender');
        let totFem = 0, totMale = 0;
        records.forEach(r => {
          const fem = parseInt(r.actual_attending_female) || 0;
          const tot = parseInt(r.actual_attending_total) || 0;
          totFem += fem;
          totMale += Math.max(0, tot - fem);
        });

        // Fallback placeholder if no records
        const dataArr = (totFem === 0 && totMale === 0) ? [1, 1] : [totFem, totMale];
        const colors  = (totFem === 0 && totMale === 0) ? ['rgba(236,72,153,0.3)', 'rgba(59,130,246,0.3)'] : ['#ec4899', '#3b82f6'];

        chartInstances.gender = new Chart(genderCanvas, {
          type: 'doughnut',
          data: {
            labels: ['ຍິງ', 'ຊາຍ'],
            datasets: [{
              data: dataArr,
              backgroundColor: colors,
              borderWidth: 0
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8' } } }
          }
        });
      }
    } catch (err) {
      console.warn('Gender chart error:', err);
    }

    // 4. Student Movement Comparison Bar Chart
    try {
      const movementCanvas = document.getElementById('movementChart');
      if (movementCanvas) {
        safeDestroy(movementCanvas, 'movement');
        const districts = ["ເມືອງໄຊ", "ເມືອງຫຼາ", "ເມືອງນາໝໍ້", "ເມືອງງາ", "ເມືອງແບງ", "ເມືອງຮຸນ", "ເມືອງປາກແບງ"];
        const attending = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.actual_attending_total) || 0), 0));
        const trIn = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.transfer_in_total) || 0), 0));
        const trOut = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.transfer_out_total) || 0), 0));
        const drop = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.dropout_total) || 0), 0));
        const rep = districts.map(d => records.filter(r => r.district === d).reduce((acc, r) => acc + (parseInt(r.repeater_total) || 0), 0));

        chartInstances.movement = new Chart(movementCanvas, {
          type: 'bar',
          data: {
            labels: districts,
            datasets: [
              { label: 'ມີໜ້າຮຽນຕົວຈິງ', data: attending, backgroundColor: '#3b82f6', borderRadius: 4 },
              { label: 'ຍ້າຍເຂົ້າ', data: trIn, backgroundColor: '#10b981', borderRadius: 4 },
              { label: 'ຍ້າຍອອກ', data: trOut, backgroundColor: '#f59e0b', borderRadius: 4 },
              { label: 'ປະລະ', data: drop, backgroundColor: '#ef4444', borderRadius: 4 },
              { label: 'ຄ້າງຫ້ອງ', data: rep, backgroundColor: '#8b5cf6', borderRadius: 4 }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8' } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
              y: { beginAtZero: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
          }
        });
      }
    } catch (err) {
      console.warn('Movement chart error:', err);
    }
  }

  // ===================================================
  // ADVANCED STATISTICS CHARTS (6 new charts)
  // ===================================================
  let advChartInstances = {};

  function renderAdvancedCharts() {
    if (typeof Chart === 'undefined') {
      setTimeout(renderAdvancedCharts, 500);
      return;
    }

    const DISTRICTS = ['ເມືອງໄຊ', 'ເມືອງຫຼາ', 'ເມືອງນາໝໍ້', 'ເມືອງງາ', 'ເມືອງແບງ', 'ເມືອງຮຸນ', 'ເມືອງປາກແບງ'];
    const DIST_SHORT = ['ໄຊ', 'ຫຼາ', 'ນາໝໍ້', 'ງາ', 'ແບງ', 'ຮຸນ', 'ປາກແບງ'];
    const MONTH_KEYS = ['ມັງກອນ','ກຸມພາ','ມີນາ','ເມສາ','ພຶດສະພາ','ມິຖຸນາ','ກໍລະກົດ','ສິງຫາ','ກັນຍາ','ຕຸລາ','ພະຈິກ','ທັນວາ'];
    const MONTH_MAP  = { 'ມັງກອນ':0,'ກຸມພາ':1,'ມີນາ':2,'ເມສາ':3,'ພຶດສະພາ':4,'ມິຖຸນາ':5,'ກໍລະກົດ':6,'ສິງຫາ':7,'ກັນຍາ':8,'ຕຸລາ':9,'ພະຈິກ':10,'ທັນວາ':11,'01':0,'02':1,'03':2,'04':3,'05':4,'06':5,'07':6,'08':7,'09':8,'10':9,'11':10,'12':11,'1':0,'2':1,'3':2,'4':3,'5':4,'6':5,'7':6,'8':7,'9':8 };

    const adv_safeDestroy = (canvas, key) => {
      try { const ex = Chart.getChart(canvas); if (ex) ex.destroy(); } catch(e) {}
      if (advChartInstances[key]) { try { advChartInstances[key].destroy(); } catch(e) {} advChartInstances[key] = null; }
    };

    // Pre-aggregate data by district
    const byDist = DISTRICTS.map(d => {
      const recs = records.filter(r => r.district === d);
      return {
        total:    recs.reduce((s,r) => s + (parseInt(r.actual_attending_total) || 0), 0),
        female:   recs.reduce((s,r) => s + (parseInt(r.actual_attending_female) || 0), 0),
        dropout:  recs.reduce((s,r) => s + (parseInt(r.dropout_total) || 0), 0),
        trIn:     recs.reduce((s,r) => s + (parseInt(r.transfer_in_total) || 0), 0),
        trOut:    recs.reduce((s,r) => s + (parseInt(r.transfer_out_total) || 0), 0),
        repeater: recs.reduce((s,r) => s + (parseInt(r.repeater_total) || 0), 0),
        count:    recs.length
      };
    });

    const GRID_COLOR  = 'rgba(255,255,255,0.05)';
    const TICK_COLOR  = '#94a3b8';
    const baseOpts = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: TICK_COLOR, font: { size: 11 } } } },
      scales: {
        x: { ticks: { color: TICK_COLOR, font: { size: 10 } }, grid: { color: GRID_COLOR } },
        y: { beginAtZero: true, ticks: { color: TICK_COLOR, font: { size: 10 } }, grid: { color: GRID_COLOR } }
      }
    };

    // ── Chart A: Stacked Gender Bar by District ──────────────
    try {
      const cvs = document.getElementById('stackedGenderChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'stackedGender');
        const males   = byDist.map(d => Math.max(0, d.total - d.female));
        const females = byDist.map(d => d.female);
        advChartInstances.stackedGender = new Chart(cvs, {
          type: 'bar',
          data: {
            labels: DIST_SHORT,
            datasets: [
              { label: 'ຊາຍ',  data: males,   backgroundColor: '#3b82f6', borderRadius: 4, stack: 's1' },
              { label: 'ຍິງ',  data: females, backgroundColor: '#ec4899', borderRadius: 4, stack: 's1' }
            ]
          },
          options: { ...baseOpts, plugins: { ...baseOpts.plugins }, scales: { x: { ...baseOpts.scales.x, stacked: true, grid: { display: false } }, y: { ...baseOpts.scales.y, stacked: true } } }
        });
      }
    } catch(e) { console.warn('stackedGenderChart error:', e); }

    // ── Chart B: Dropout Rate % Horizontal Bar ───────────────
    try {
      const cvs = document.getElementById('dropoutRateChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'dropoutRate');
        const rates = byDist.map(d => d.total > 0 ? +((d.dropout / d.total) * 100).toFixed(2) : 0);
        const colors = rates.map(r => r > 5 ? '#ef4444' : r > 2 ? '#f59e0b' : '#10b981');
        advChartInstances.dropoutRate = new Chart(cvs, {
          type: 'bar',
          data: {
            labels: DIST_SHORT,
            datasets: [{ label: 'ອັດຕາປະລະ (%)', data: rates, backgroundColor: colors, borderRadius: 4 }]
          },
          options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.x}%` } } },
            scales: {
              x: { beginAtZero: true, max: 15, ticks: { color: TICK_COLOR, callback: v => v+'%' }, grid: { color: GRID_COLOR } },
              y: { ticks: { color: TICK_COLOR, font: { size: 11 } }, grid: { display: false } }
            }
          }
        });
      }
    } catch(e) { console.warn('dropoutRateChart error:', e); }

    // ── Chart C: Radar KPI by District ───────────────────────
    try {
      const cvs = document.getElementById('radarKpiChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'radarKpi');
        const maxT  = Math.max(...byDist.map(d => d.total),  1);
        const maxDr = Math.max(...byDist.map(d => d.dropout), 1);
        const maxIn = Math.max(...byDist.map(d => d.trIn),    1);
        const maxRp = Math.max(...byDist.map(d => d.repeater),1);
        const norm  = (v, mx) => mx > 0 ? +((v / mx) * 100).toFixed(2) : 0;
        const colors7 = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899','#06b6d4'];
        advChartInstances.radarKpi = new Chart(cvs, {
          type: 'radar',
          data: {
            labels: ['ໜ້າຮຽນ', 'ຍ້າຍເຂົ້າ', 'ປະລະ', 'ຄ້າງຫ້ອງ', 'ຍິງ%'],
            datasets: DISTRICTS.map((d, i) => ({
              label: DIST_SHORT[i],
              data: [
                norm(byDist[i].total, maxT),
                norm(byDist[i].trIn, maxIn),
                norm(byDist[i].dropout, maxDr),
                norm(byDist[i].repeater, maxRp),
                byDist[i].total > 0 ? +((byDist[i].female / byDist[i].total) * 100).toFixed(2) : 0
              ],
              borderColor: colors7[i],
              backgroundColor: colors7[i].replace(')', ',0.12)').replace('rgb', 'rgba').replace('#', '').length > 10
                ? colors7[i] + '20'
                : colors7[i] + '20',
              borderWidth: 2,
              pointRadius: 3
            }))
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'right', labels: { color: TICK_COLOR, boxWidth: 12, font: { size: 10 } } } },
            scales: { r: { ticks: { display: false }, grid: { color: GRID_COLOR }, pointLabels: { color: TICK_COLOR, font: { size: 10 } }, suggestedMin: 0, suggestedMax: 100 } }
          }
        });
      }
    } catch(e) { console.warn('radarKpiChart error:', e); }

    // ── Chart D: School Level Pie (ປ.ຖ / ມ.ຕ / ມ.ສ) ─────────
    try {
      const cvs = document.getElementById('schoolLevelPieChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'schoolLevelPie');
        // Count schools by level from DEFAULT_LAO_DISTRICTS
        let ptTotal = 0, mtTotal = 0, msTotal = 0;
        if (typeof DEFAULT_LAO_DISTRICTS !== 'undefined') {
          Object.values(DEFAULT_LAO_DISTRICTS).forEach(dObj => {
            if (dObj['ມ.ສ']) msTotal += dObj['ມ.ສ'].length;
            if (dObj['ມ.ຕ']) mtTotal += dObj['ມ.ຕ'].length;
            if (dObj['ປ.ຖ']) ptTotal += dObj['ປ.ຖ'].length;
          });
        }
        // Fallback counts if no data
        if (ptTotal + mtTotal + msTotal === 0) { ptTotal = 470; mtTotal = 54; msTotal = 44; }
        advChartInstances.schoolLevelPie = new Chart(cvs, {
          type: 'doughnut',
          data: {
            labels: ['ປ.ຖ (ປະຖົມ)', 'ມ.ຕ (ມ.ຕ)', 'ມ.ສ (ມ.ສ)'],
            datasets: [{ data: [ptTotal, mtTotal, msTotal], backgroundColor: ['#06b6d4','#8b5cf6','#f59e0b'], borderWidth: 0, hoverOffset: 6 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { position: 'bottom', labels: { color: TICK_COLOR, font: { size: 11 } } },
              tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} ໂຮງຮຽນ` } }
            },
            cutout: '60%'
          }
        });
      }
    } catch(e) { console.warn('schoolLevelPieChart error:', e); }

    // ── Chart E: Net Flow by Month Line ──────────────────────
    try {
      const cvs = document.getElementById('netFlowChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'netFlow');
        const trIn   = new Array(12).fill(0);
        const trOut  = new Array(12).fill(0);
        const drop   = new Array(12).fill(0);
        const netFlow= new Array(12).fill(0);
        records.forEach(r => {
          const m = MONTH_MAP[(r.entry_month_only || '').trim().toLowerCase()];
          if (m !== undefined) {
            trIn[m]  += parseInt(r.transfer_in_total)  || 0;
            trOut[m] += parseInt(r.transfer_out_total) || 0;
            drop[m]  += parseInt(r.dropout_total)      || 0;
            netFlow[m] = trIn[m] - trOut[m] - drop[m];
          }
        });
        advChartInstances.netFlow = new Chart(cvs, {
          type: 'line',
          data: {
            labels: MONTH_KEYS,
            datasets: [
              { label: 'ຍ້າຍເຂົ້າ', data: trIn,    borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)', fill: false, tension: 0.3, borderWidth: 2 },
              { label: 'ຍ້າຍອອກ',   data: trOut,   borderColor: '#f59e0b', backgroundColor: 'rgba(245,158,11,0.1)',  fill: false, tension: 0.3, borderWidth: 2 },
              { label: 'ປະລະ',       data: drop,    borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.1)',   fill: false, tension: 0.3, borderWidth: 2 },
              { label: 'Net Flow',   data: netFlow, borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.15)', fill: true,  tension: 0.3, borderWidth: 3, borderDash: [] }
            ]
          },
          options: { ...baseOpts }
        });
      }
    } catch(e) { console.warn('netFlowChart error:', e); }

    // ── Chart F: Submission Progress % by District ───────────
    try {
      const cvs = document.getElementById('submissionProgressChart');
      if (cvs) {
        adv_safeDestroy(cvs, 'submissionProgress');
        let totalByDist = {}, submittedByDist = {};
        DISTRICTS.forEach(d => { totalByDist[d] = 0; submittedByDist[d] = 0; });
        if (typeof DEFAULT_LAO_DISTRICTS !== 'undefined') {
          DISTRICTS.forEach(d => {
            const dObj = DEFAULT_LAO_DISTRICTS[d] || {};
            totalByDist[d] = (dObj['ມ.ສ']||[]).length + (dObj['ມ.ຕ']||[]).length + (dObj['ປ.ຖ']||[]).length;
          });
        }
        records.forEach(r => {
          if (submittedByDist[r.district] !== undefined) submittedByDist[r.district]++;
        });
        const progressPct = DISTRICTS.map(d => totalByDist[d] > 0 ? +Math.min(100, ((submittedByDist[d] / totalByDist[d]) * 100)).toFixed(2) : 0);
        const pgColors = progressPct.map(p => p >= 80 ? '#10b981' : p >= 40 ? '#f59e0b' : '#ef4444');
        advChartInstances.submissionProgress = new Chart(cvs, {
          type: 'bar',
          data: {
            labels: DIST_SHORT,
            datasets: [{ label: 'ຄວາມຄືບໜ້າ (%)', data: progressPct, backgroundColor: pgColors, borderRadius: 6 }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y}% (${submittedByDist[DISTRICTS[ctx.dataIndex]]} / ${totalByDist[DISTRICTS[ctx.dataIndex]]} ໂຮງ)` } }
            },
            scales: {
              x: { ticks: { color: TICK_COLOR, font: { size: 11 } }, grid: { display: false } },
              y: { beginAtZero: true, max: 100, ticks: { color: TICK_COLOR, callback: v => v+'%' }, grid: { color: GRID_COLOR } }
            }
          }
        });
      }
    } catch(e) { console.warn('submissionProgressChart error:', e); }
  }

  // Google Apps Script Live Sync
  async function syncSchoolsFromGoogleSheet() {
    if (typeof GOOGLE_APPS_SCRIPT_URL === 'undefined' || !GOOGLE_APPS_SCRIPT_URL) return;
    try {
      const res = await fetch(`${GOOGLE_APPS_SCRIPT_URL}?action=getSchools`);
      if (!res.ok) return;
      const data = await res.json();
      if (data && data.status === 'success' && Array.isArray(data.data)) {
        const freshDistricts = {};
        const freshCodes = {};
        data.data.forEach(s => {
          if (!s.name || !s.district) return;
          const dist = s.district.trim();
          const name = s.name.trim();
          if (!freshDistricts[dist]) freshDistricts[dist] = [];
          if (!freshDistricts[dist].includes(name)) freshDistricts[dist].push(name);
          if (s.code) freshCodes[name] = s.code.trim();
        });
        if (Object.keys(freshDistricts).length > 0) {
          Object.assign(DEFAULT_LAO_DISTRICTS, freshDistricts);
          Object.assign(DEFAULT_SCHOOL_CODES, freshCodes);
          // Only repopulate dropdowns if user is not actively selecting form fields
          if (!studentDataForm || (!districtSelect?.value && !schoolSelect?.value)) {
            populateAllDropdowns();
          }
          renderSubmissionStats();
        }
      }
    } catch (err) {
      console.warn('Sync schools fallback:', err);
    }
  }

  async function postRecordToGoogleSheet(record) {
    if (typeof GOOGLE_APPS_SCRIPT_URL === 'undefined' || !GOOGLE_APPS_SCRIPT_URL) return;
    try {
      await fetch(GOOGLE_APPS_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record)
      });
      console.log('⚡ Record posted to Google Sheets');
    } catch (e) {
      console.warn('Google Sheets post error:', e);
    }
  }

  // Excel Export
  function exportToExcel() {
    if (records.length === 0) {
      showToast("ບໍ່ມີຂໍ້ມູນສຳລັບສົ່ງອອກ!", "warning");
      return;
    }

    let csvContent = "\uFEFF" + "ID,ເດືອນ,ສົກຮຽນ,ເມືອງ,ໂຮງຮຽນ,ລະຫັດ,ຊັ້ນຮຽນ,ຂັ້ນຮຽນ,ຜູ້ປ້ອນ,ລົງທະບຽນທັງໝົດ,ລົງທະບຽນຍິງ,ຍ້າຍເຂົ້າທັງໝົດ,ຍ້າຍເຂົ້າຍິງ,ຍ້າຍອອກທັງໝົດ,ຍ້າຍອອກຍິງ,ສາເຫດຍ້າຍອອກ,ປະລະທັງໝົດ,ປະລະຍິງ,ສາເຫດປະລະ,ຄືນຮຽນທັງໝົດ,ຄືນຮຽນຍິງ,ມາຮຽນຕົວຈິງທັງໝົດ,ມາຮຽນຕົວຈິງຍິງ,ວັນທີບັນທຶກ\n";

    records.forEach(r => {
      const row = [
        `"${r.id}"`, `"${r.entry_month_only}"`, `"${r.entry_academic_year}"`,
        `"${r.district}"`, `"${r.school}"`, `"${r.school_code}"`,
        `"${r.education_level}"`, `"${r.grade_level}"`, `"${r.collector_name}"`,
        r.passed_registered_total, r.passed_registered_female,
        r.transfer_in_total, r.transfer_in_female,
        r.transfer_out_total, r.transfer_out_female, `"${r.transfer_out_reason || ''}"`,
        r.dropout_total, r.dropout_female, `"${r.dropout_reason || ''}"`,
        r.repeater_total, r.repeater_female,
        r.actual_attending_total, r.actual_attending_female,
        `"${r.created_at}"`
      ];
      csvContent += row.join(",") + "\n";
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `lao_student_data_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    showToast("ສົ່ງອອກ CSV ສຳເລັດແລ້ວ!", "success");
  }

  // Toast Notification Helper
  function showToast(message, type = "info") {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.style.cssText = `
      position: fixed; bottom: 20px; right: 20px; z-index: 9999;
      background: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : type === 'warning' ? '#f59e0b' : '#3b82f6'};
      color: #fff; padding: 12px 20px; border-radius: 8px; font-weight: 500;
      box-shadow: 0 10px 25px rgba(0,0,0,0.3); transition: all 0.3s ease;
    `;
    toast.innerHTML = `<i class="fa-solid fa-circle-info"></i> ${message}`;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  // ===================================================
  // REAL-TIME AUTO-SYNC SYSTEM (Live Dashboard Updates & Google Sheet Sync)
  // ===================================================
  async function syncRecordsFromGoogleSheet(silent = true) {
    const gasUrl = (typeof GOOGLE_APPS_SCRIPT_URL !== 'undefined' && GOOGLE_APPS_SCRIPT_URL)
      ? GOOGLE_APPS_SCRIPT_URL
      : localStorage.getItem('gas_web_app_url');

    if (!gasUrl) return;

    const badge = document.getElementById('liveSyncStatusBadge');
    if (badge && !silent) {
      badge.innerHTML = `<i class="fa-solid fa-spin fa-spinner"></i> ກຳລັງ Sync ຈາກ Google Sheet...`;
      badge.style.color = '#f59e0b';
    }

    try {
      const res = await fetch(`${gasUrl}?action=getRecords`);
      if (!res.ok) throw new Error("HTTP error " + res.status);
      const json = await res.json();
      if (json && json.status === 'success' && Array.isArray(json.data)) {
        if (json.data.length > 0) {
          records = json.data;
          saveRecordsToStorage();
          renderApp();
          if (!silent) showToast(`Sync ຂໍ້ມູນຈາກ Google Sheet ສຳເລັດ! (${json.data.length} ລາຍການ)`, "success");
        } else if (!silent) {
          showToast("DataRecords ໃນ Google Sheet ຍັງຫວ່າງເປົ່າ!", "warning");
        }
      }
    } catch (err) {
      console.warn('Sync Google Sheet records error:', err);
      if (!silent) showToast("ບໍ່ສາມາດເຊື່ອມຕໍ່ Google Sheet ໄດ້: " + err.message, "error");
    } finally {
      if (badge) {
        badge.innerHTML = `<i class="fa-solid fa-rotate fa-spin" style="animation-duration: 4s;"></i> ປັບຂໍ້ມູນອັດຕະໂນມັດ (Live Auto-Sync)`;
        badge.style.color = 'var(--accent-success)';
      }
    }
  }

  function triggerLiveAutoSync() {
    // If user is on form page and actively filling out form inputs, DO NOT auto refresh
    if (studentDataForm) {
      const active = document.activeElement;
      if (active && (studentDataForm.contains(active) || active.tagName === 'INPUT' || active.tagName === 'SELECT')) {
        return;
      }
      // If user typed numbers into form, do not interrupt
      if ((passedRegisteredTotal && passedRegisteredTotal.value) || (transferInTotal && transferInTotal.value)) {
        return;
      }
    }

    loadRecordsFromStorage();
    renderApp();

    const isDashboardPage = !!document.getElementById('liveSyncStatusBadge') || !!document.getElementById('monthlyTrendChart');
    if (isDashboardPage) {
      syncSchoolsFromGoogleSheet();
      syncRecordsFromGoogleSheet(true);
    }

    const badge = document.getElementById('liveSyncStatusBadge');
    if (badge) {
      badge.style.transform = 'scale(1.08)';
      setTimeout(() => badge.style.transform = 'scale(1)', 300);
    }
  }

  // Sync Google Sheet Button Listener
  const syncGasBtn = document.getElementById('syncGasBtn');
  if (syncGasBtn) {
    syncGasBtn.addEventListener('click', () => {
      let url = (typeof GOOGLE_APPS_SCRIPT_URL !== 'undefined' && GOOGLE_APPS_SCRIPT_URL)
        ? GOOGLE_APPS_SCRIPT_URL
        : localStorage.getItem('gas_web_app_url');

      if (!url) {
        url = prompt("ກະລຸນາໃສ່ Web App URL ຂອງ Google Apps Script (https://script.google.com/macros/s/.../exec):");
        if (url && url.trim()) {
          localStorage.setItem('gas_web_app_url', url.trim());
        } else {
          return;
        }
      }
      syncRecordsFromGoogleSheet(false);
    });
  }



  const isDashboard = !!document.getElementById('liveSyncStatusBadge') || !!document.getElementById('monthlyTrendChart');

  // Initial load sync from Google Sheet (Only on Dashboard page)
  if (isDashboard) {
    syncRecordsFromGoogleSheet(true);
  }

  // 1. Cross-Tab Live Storage Event Listener (Only trigger if not typing)
  window.addEventListener('storage', (e) => {
    if (e.key === 'lao_school_records' || e.key === 'lao_school_records_updated_at') {
      triggerLiveAutoSync();
    }
  });

  // 2. Window Focus & Tab Visibility Change Auto-Sync (Only on Dashboard)
  if (isDashboard) {
    window.addEventListener('focus', triggerLiveAutoSync);
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden) triggerLiveAutoSync();
    });

    // 3. Periodic Background Auto-Sync (Every 10 seconds - Dashboard only)
    setInterval(triggerLiveAutoSync, 10000);
  }
});
