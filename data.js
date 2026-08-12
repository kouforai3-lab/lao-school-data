// Lao Education Data Presets & Configuration

// Empty default districts (Will ONLY be populated from imported Excel or manually added)
const LAO_DISTRICTS_DATA = {};

// Lao Months Configuration
const LAO_MONTHS = [
  { value: "01", name: "ເດືອນ 1 (ມັງກອນ)", short: "ມັງກອນ" },
  { value: "02", name: "ເດືອນ 2 (ກຸມພາ)", short: "ກຸມພາ" },
  { value: "03", name: "ເດືອນ 3 (ມີນາ)", short: "ມີນາ" },
  { value: "04", name: "ເດືອນ 4 (ເມສາ)", short: "ເມສາ" },
  { value: "05", name: "ເດືອນ 5 (ພຶດສະພາ)", short: "ພຶດສະພາ" },
  { value: "06", name: "ເດືອນ 6 (ມິຖຸນາ)", short: "ມິຖຸນາ" },
  { value: "07", name: "ເດືອນ 7 (ກໍລະກົດ)", short: "ກໍລະກົດ" },
  { value: "08", name: "ເດືອນ 8 (ສິງຫາ)", short: "ສິງຫາ" },
  { value: "09", name: "ເດືອນ 9 (ກັນຍາ)", short: "ກັນຍາ" },
  { value: "10", name: "ເດືອນ 10 (ຕຸລາ)", short: "ຕຸລາ" },
  { value: "11", name: "ເດືອນ 11 (ພະຈິກ)", short: "ພະຈິກ" },
  { value: "12", name: "ເດືອນ 12 (ທັນວາ)", short: "ທັນວາ" }
];

// Academic Years Configuration (ເລີ່ມແຕ່ ສົກຮຽນ 2026-2027 ເປັນຕົ້ນໄປ)
const ACADEMIC_YEARS = [
  "2026-2027",
  "2027-2028",
  "2028-2029",
  "2029-2030",
  "2030-2031",
  "2031-2032"
];

// Education Levels & Grade Level Mappings (Removed ມັດທະຍົມປາຍ)
const EDUCATION_LEVELS = {
  "ປະຖົມສຶກສາ": ["ປ.1", "ປ.2", "ປ.3", "ປ.4", "ປ.5"],
  "ມັດທະຍົມຕອນຕົ້ນ": ["ມ.1", "ມ.2", "ມ.3", "ມ.4"],
  "ມັດທະຍົມສົມບູນ": ["ມ.1", "ມ.2", "ມ.3", "ມ.4", "ມ.5", "ມ.6", "ມ.7"]
};

// Initial seed data if LocalStorage is empty
const INITIAL_SAMPLE_DATA = [];
