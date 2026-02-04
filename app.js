/**
 * LifeWealth 100 Simulator
 * 
 * 会計処理に関する注記：
 * 
 * 資産カテゴリー分類：
 * - cash：現金・預金（現金等価物 + ポイント）
 * - stocks：株式・現物株式（当初原価ベース）
 * - funds：投資信託・ファンド（複利計算により含み益を認識）
 * - bonds：債券・固定利付証券（評価額ベース、満期時に額面現金化）
 * - insurance：保険商品（複利計算により含み益を認識）
 * - dc：確定拠出年金・企業年金（拠出フェーズと給付フェーズで区分）
 * - usd：ドル積立（円換算・当初原価ベース）
 * - other：その他資産（当初原価ベース）
 * 
 * 運用益の認識：
 * - 複利計算対象（funds、insurance）：期中の利息・配当を含み益として計上
 * - その他：期中の時価変動を反映しない（初期評価額/原価ベース）
 * - 注：実現益と含み益を区分していない
 * 
 * 投資拠出の処理：
 * - 拠出（contribution）は資産の増加であり、損益計算書上の支出ではない
 * - 現金から各資産カテゴリーへの配分変更として処理
 * 
 * 年金・DC処理：
 * - 拠出期：開始年齢未満の期間に毎月拠出
 * - 給付期：開始年齢以上の期間に一括または分割で給付
 */

const birthDateInput = document.getElementById("birthDate");
const currentAssetsInput = document.getElementById("currentAssets");
const retirementAgeInput = document.getElementById("retirementAge");
const assetDataInput = document.getElementById("assetData");
const summaryDataInput = document.getElementById("summaryData");
const importButton = document.getElementById("applyImport");
const exportButton = document.getElementById("exportCsv");
const openCsvButton = document.getElementById("openCsv");
const statementYearFromSelect = document.getElementById("statementYearFrom");
const statementYearToSelect = document.getElementById("statementYearTo");
const exportBalanceSheetDecadeButton = document.getElementById(
  "exportBalanceSheetDecade"
);
const openBalanceSheetDecadeButton = document.getElementById(
  "openBalanceSheetDecade"
);
const exportProfitLossDecadeButton = document.getElementById(
  "exportProfitLossDecade"
);
const openProfitLossDecadeButton = document.getElementById(
  "openProfitLossDecade"
);
const bondTableBody = document.getElementById("bondTableBody");
const bondMaturedBody = document.getElementById("bondMaturedBody");
const otherAssetTableBody = document.getElementById("otherAssetTableBody");
const addBondRowButton = document.getElementById("addBondRow");
const sortBondRowsButton = document.getElementById("sortBondRows");
const addOtherAssetRowButton = document.getElementById("addOtherAssetRow");
const bondAverageRate = document.getElementById("bondAverageRate");
const bondUsdRateInput = document.getElementById("bondUsdRate");
const bondTotalAmount = document.getElementById("bondTotalAmount");
const otherAssetTotalAmount = document.getElementById("otherAssetTotalAmount");
const bondCashReclassTotal = document.getElementById("bondCashReclassTotal");
const bondCombinedTotal = document.getElementById("bondCombinedTotal");
const importStatus = document.getElementById("importStatus");
const reclassStatus = document.getElementById("reclassStatus");
const exportSyncFolderButton = document.getElementById("exportSyncFolder");
const importSyncFileButton = document.getElementById("importSyncFile");
const syncFileInput = document.getElementById("syncFileInput");
const syncStatus = document.getElementById("syncStatus");
const expenseInputs = Array.from(document.querySelectorAll(".expense-input"));
const monthlyExpense = document.getElementById("monthlyExpense");
const incomeInputs = Array.from(document.querySelectorAll(".income-input"));
const dividendYieldInput = document.getElementById("dividendYield");
const monthlyIncome = document.getElementById("monthlyIncome");
const retireExpenseInputs = Array.from(
  document.querySelectorAll(".retire-expense-input")
);
const monthlyRetireExpense = document.getElementById("monthlyRetireExpense");
const retireIncomeInputs = Array.from(
  document.querySelectorAll(".retire-income-input")
);
const monthlyRetireIncome = document.getElementById("monthlyRetireIncome");
const retirementIncomeEndAgeInput = document.getElementById(
  "retirementIncomeEndAge"
);
const pensionIncomeInputs = Array.from(
  document.querySelectorAll(".pension-income-input")
);
const monthlyPensionIncome = document.getElementById("monthlyPensionIncome");
const balanceCashInput = document.getElementById("balanceCash");
const adjustCashInput = document.getElementById("adjustCash");
const balanceStocksInput = document.getElementById("balanceStocks");
const balanceFundsInput = document.getElementById("balanceFunds");
const balanceBondsInput = document.getElementById("balanceBonds");
const balanceInsuranceInput = document.getElementById("balanceInsurance");
const balanceUsdInput = document.getElementById("balanceUsd");
const balanceDcInput = document.getElementById("balanceDc");
const adjustStocksInput = document.getElementById("adjustStocks");
const adjustFundsInput = document.getElementById("adjustFunds");
const adjustBondsInput = document.getElementById("adjustBonds");
const adjustInsuranceInput = document.getElementById("adjustInsurance");
const adjustUsdInput = document.getElementById("adjustUsd");
const adjustDcInput = document.getElementById("adjustDc");
const manualAdjustmentPairs = [
  { balance: balanceStocksInput, adjust: adjustStocksInput },
  { balance: balanceFundsInput, adjust: adjustFundsInput },
  { balance: balanceInsuranceInput, adjust: adjustInsuranceInput },
  { balance: balanceUsdInput, adjust: adjustUsdInput },
  { balance: balanceDcInput, adjust: adjustDcInput },
];
const bondAdjustmentPair = { balance: balanceBondsInput, adjust: adjustBondsInput };
const rateStocksInput = document.getElementById("rateStocks");
const rateFundsInput = document.getElementById("rateFunds");
const rateBondsInput = document.getElementById("rateBonds");
const rateInsuranceInput = document.getElementById("rateInsurance");
const contribStocksInput = document.getElementById("contribStocks");
const contribFundsInput = document.getElementById("contribFunds");
const contribBondsInput = document.getElementById("contribBonds");
const contribInsuranceInput = document.getElementById("contribInsurance");
const contribUsdInput = document.getElementById("contribUsd");
const contribDcInput = document.getElementById("contribDc");
const endAgeStocksInput = document.getElementById("endAgeStocks");
const endAgeFundsInput = document.getElementById("endAgeFunds");
const endAgeBondsInput = document.getElementById("endAgeBonds");
const endAgeInsuranceInput = document.getElementById("endAgeInsurance");
const endAgeUsdInput = document.getElementById("endAgeUsd");
const endAgeDcInput = document.getElementById("endAgeDc");
if (contribDcInput) {
  contribDcInput.disabled = true;
  contribDcInput.title = "年金詳細で設定してください";
}
if (contribInsuranceInput) {
  contribInsuranceInput.disabled = true;
  contribInsuranceInput.title = "保険料詳細で設定してください";
}
if (balanceBondsInput) {
  balanceBondsInput.disabled = true;
  balanceBondsInput.title = "債券入力の評価額から自動計算されます";
}
if (rateBondsInput) {
  rateBondsInput.disabled = true;
  rateBondsInput.title = "債券入力の利率平均が自動表示されます";
}
const investmentTotal = document.getElementById("investmentTotal");
const cashBalance = document.getElementById("cashBalance");
const cashDeduction = document.getElementById("cashDeduction");
const cashBaseDisplay = document.getElementById("cashBaseDisplay");
const cashDetailButton = document.getElementById("cashDetailButton");
const cashDetailBase = document.getElementById("cashDetailBase");
const cashDetailPoints = document.getElementById("cashDetailPoints");
const cashDetailReclass = document.getElementById("cashDetailReclass");
const cashDetailAdjustment = document.getElementById("cashDetailAdjustment");
const cashDetailFinal = document.getElementById("cashDetailFinal");
const cashDetailBackButton = document.getElementById("cashDetailBack");
const investmentContribTotal = document.getElementById("investmentContribTotal");
const investmentAfter = document.getElementById("investmentAfter");
const cashAfter = document.getElementById("cashAfter");
const investmentAlert = document.getElementById("investmentAlert");
const assetYearSelects = Array.from(
  document.querySelectorAll(".asset-year-select")
);
const assetDetailButtons = Array.from(
  document.querySelectorAll(".asset-detail-button")
);
const assetDetailTitle = document.getElementById("assetDetailTitle");
const assetDetailSubtitle = document.getElementById("assetDetailSubtitle");
const assetDetailYearSelect = document.getElementById("assetDetailYear");
const assetDetailTableBody = document.getElementById("assetDetailTableBody");
const assetDetailYearDelta = document.getElementById("assetDetailYearDelta");
const assetDetailBackButton = document.getElementById("assetDetailBack");
const bondDetailButton = document.getElementById("bondDetailButton");
const bondDetailBackButton = document.getElementById("bondDetailBack");
const bondRateAverageDisplay = document.getElementById("bondRateAverageDisplay");
const insuranceDetailButton = document.getElementById("insuranceDetailButton");
const insuranceDetailBackButton = document.getElementById("insuranceDetailBack");
const insuranceCurrentAmount = document.getElementById("insuranceCurrentAmount");
const insurancePlanBody = document.getElementById("insurancePlanBody");
const addInsurancePlanRowButton = document.getElementById("addInsurancePlanRow");
const pensionDetailButton = document.getElementById("pensionDetailButton");
const pensionDetailBackButton = document.getElementById("pensionDetailBack");
const pensionCurrentAmount = document.getElementById("pensionCurrentAmount");
const pensionPlanBody = document.getElementById("pensionPlanBody");
const addPensionPlanRowButton = document.getElementById("addPensionPlanRow");
const pensionChangeBody = document.getElementById("pensionChangeBody");
const addPensionChangeRowButton = document.getElementById("addPensionChangeRow");
const tabButtons = Array.from(document.querySelectorAll(".tab-button"));
const pages = Array.from(document.querySelectorAll(".page"));
const resultValue = document.getElementById("resultValue");
const resultMeta = document.getElementById("resultMeta");
const lastUpdated = document.getElementById("lastUpdated");
let importDirty = false;
let lastInvestmentBalanceTotal = null;
let assetDetailState = { key: "cash", year: null };
let cashInputManual = false;
let bondUsdRateListenerBound = false;

const ASSET_LABELS = {
  cash: "現金",
  stocks: "株式投資",
  funds: "投資信託",
  bonds: "債券",
  insurance: "積立保険",
  usd: "ドル積立",
  dc: "年金",
};

const ASSET_CATEGORY_KEYS = {
  cash: "cash",
  stocks: "stocks",
  funds: "funds",
  bonds: "bonds",
  insurance: "insurance",
  usd: "usd",
  dc: "dc",
};

const STORAGE_KEY = "lifewealth100.inputs.v1";
const BOND_STORAGE_KEY = "lifewealth100.bonds.v1";
const OTHER_ASSETS_KEY = "lifewealth100.otherAssets.v1";
const CASH_MANUAL_KEY = "lifewealth100.cash.manual.v1";
const INSURANCE_PLANS_KEY = "lifewealth100.insurance.plans.v1";
const INSURANCE_SCHEDULE_LEGACY_KEY = "lifewealth100.insurance.schedule.v1";
const PENSION_PLANS_KEY = "lifewealth100.pension.plans.v1";
const PENSION_CHANGES_KEY = "lifewealth100.pension.changes.v1";
const PUSH_TIMESTAMP_KEY = "lifewealth100.lastPushTimestamp.v1";
const ADJUSTMENT_APPLIED_KEY = "lifewealth100.adjustmentsApplied.v1";
const STORAGE_PREFIX = "lifewealth100.";
const GITHUB_REPO_FALLBACK = {
  owner: "yuuitihtnk167-art",
  repo: "LifeWealth100Simulator",
};
const persistInputs = Array.from(document.querySelectorAll("input, textarea")).filter(
  (el) =>
    el.type !== "button" &&
    el.type !== "submit" &&
    !el.dataset.noPersist
);

const yenFormatter = new Intl.NumberFormat("ja-JP", {
  style: "currency",
  currency: "JPY",
  maximumFractionDigits: 0,
});

const percentFormatter = new Intl.NumberFormat("ja-JP", {
  maximumFractionDigits: 2,
});

// 会計処理：複利計算対象カテゴリーの定義
// 投資信託（funds）と保険（insurance）のみ期中利息/配当を含む（複利計算）
// 他のカテゴリー（現金、株式、債券等）は期中の時価変動を反映しない（当初原価ベース）
// 会計処理：複利計算対象カテゴリーの定義
// 投資信託（funds）と保険（insurance）のみ期中利息/配当を含む（複利計算）
// 他のカテゴリー（現金、株式、債券等）は期中の時価変動を反映しない（当初原価ベース）
function isCompoundingCategory(key) {
  return key === "stocks" || key === "funds" || key === "insurance";
}

function toYenAmount(value) {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return Math.floor(value);
}

function normalizeBondStorage(raw) {
  if (!raw) {
    return { active: [], matured: [], usdRate: "" };
  }
  if (Array.isArray(raw)) {
    return { active: raw, matured: [], usdRate: "" };
  }
  if (raw && typeof raw === "object") {
    return {
      active: Array.isArray(raw.active) ? raw.active : [],
      matured: Array.isArray(raw.matured) ? raw.matured : [],
      usdRate: raw.usdRate ?? "",
    };
  }
  return { active: [], matured: [], usdRate: "" };
}

function readBondStorage() {
  const raw = safeParseJson(localStorage.getItem(BOND_STORAGE_KEY), null);
  return normalizeBondStorage(raw);
}

function writeBondStorage(data) {
  try {
    localStorage.setItem(
      BOND_STORAGE_KEY,
      JSON.stringify({
        active: data.active || [],
        matured: data.matured || [],
        usdRate: data.usdRate ?? "",
      })
    );
  } catch {
    // Ignore storage failures.
  }
}

function readOtherAssetsStorage() {
  const raw = safeParseJson(localStorage.getItem(OTHER_ASSETS_KEY), []);
  return Array.isArray(raw) ? raw : [];
}

function writeOtherAssetsStorage(rows) {
  try {
    localStorage.setItem(OTHER_ASSETS_KEY, JSON.stringify(rows || []));
  } catch {
    // Ignore storage failures.
  }
}

function readInsurancePlans() {
  const plans = safeParseJson(localStorage.getItem(INSURANCE_PLANS_KEY), []);
  if (Array.isArray(plans) && plans.length > 0) {
    return plans;
  }
  const legacy = safeParseJson(
    localStorage.getItem(INSURANCE_SCHEDULE_LEGACY_KEY),
    []
  );
  if (!Array.isArray(legacy) || legacy.length === 0) {
    return [];
  }
  const fallbackAmount = parseNumber(contribInsuranceInput?.value);
  if (Number.isFinite(fallbackAmount) && fallbackAmount > 0) {
    return [
      {
        id: makeRowId(),
        name: "移行",
        amount: String(Math.round(fallbackAmount)),
        endMonth: "",
      },
    ];
  }
  const latest = legacy
    .map((row) => ({
      age: parseNumber(row?.age),
      amount: parseNumber(row?.amount),
    }))
    .filter((row) => row.age !== null && row.amount !== null)
    .sort((a, b) => b.age - a.age)[0];
  if (!latest) {
    return [];
  }
  return [
    {
      id: makeRowId(),
      name: "移行",
      amount: String(Math.round(latest.amount)),
      endMonth: "",
    },
  ];
}

function writeInsurancePlans(rows) {
  try {
    localStorage.setItem(INSURANCE_PLANS_KEY, JSON.stringify(rows));
  } catch {
    // Ignore storage failures.
  }
}

function readPensionPlans() {
  return safeParseJson(localStorage.getItem(PENSION_PLANS_KEY), []);
}

function writePensionPlans(rows) {
  try {
    localStorage.setItem(PENSION_PLANS_KEY, JSON.stringify(rows));
  } catch {
    // Ignore storage failures.
  }
}

function readPensionChanges() {
  return safeParseJson(localStorage.getItem(PENSION_CHANGES_KEY), []);
}

function writePensionChanges(rows) {
  try {
    localStorage.setItem(PENSION_CHANGES_KEY, JSON.stringify(rows));
  } catch {
    // Ignore storage failures.
  }
}

function parseNumber(value) {
  const text = String(value ?? "");
  const normalized = text
    .replace(/[０-９]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/[－−]/g, "-")
    .replace(/[．]/g, ".")
    .replace(/[,\uFF0C]/g, "")
    .replace(/[¥￥]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function readAdjustmentsAppliedFlag() {
  try {
    return localStorage.getItem(ADJUSTMENT_APPLIED_KEY) === "1";
  } catch {
    return false;
  }
}

function writeAdjustmentsAppliedFlag(value) {
  try {
    localStorage.setItem(ADJUSTMENT_APPLIED_KEY, value ? "1" : "0");
  } catch {
    // Ignore storage failures.
  }
}

function readAdjustmentValue(input) {
  const value = parseNumber(input?.value);
  return Number.isFinite(value) ? value : 0;
}

function readPrevAdjustmentValue(input) {
  const value = parseNumber(input?.dataset?.prevAdjust);
  return Number.isFinite(value) ? value : 0;
}

function writePrevAdjustmentValue(input, value) {
  if (input) {
    input.dataset.prevAdjust = String(value);
  }
}

function applyAdjustmentToBase(balanceInput, adjustInput, baseValue) {
  if (!balanceInput || !adjustInput) {
    return false;
  }
  const adjustment = readAdjustmentValue(adjustInput);
  const base = Number.isFinite(baseValue) ? baseValue : 0;
  balanceInput.value = String(Math.round(base + adjustment));
  writePrevAdjustmentValue(adjustInput, adjustment);
  return true;
}

function applyAdjustmentChange(balanceInput, adjustInput) {
  if (!balanceInput || !adjustInput) {
    return false;
  }
  const prevAdjustment = readPrevAdjustmentValue(adjustInput);
  const nextAdjustment = readAdjustmentValue(adjustInput);
  const currentBalance = parseNumber(balanceInput.value);
  const base = Number.isFinite(currentBalance)
    ? currentBalance - prevAdjustment
    : 0;
  balanceInput.value = String(Math.round(base + nextAdjustment));
  writePrevAdjustmentValue(adjustInput, nextAdjustment);
  return true;
}

function initializeAdjustments({ applyToBalance = false, includeBonds = true } = {}) {
  manualAdjustmentPairs.forEach(({ balance, adjust }) => {
    if (!balance || !adjust) {
      return;
    }
    if (applyToBalance) {
      const base = parseNumber(balance.value) || 0;
      applyAdjustmentToBase(balance, adjust, base);
      return;
    }
    writePrevAdjustmentValue(adjust, readAdjustmentValue(adjust));
  });
  if (includeBonds && bondAdjustmentPair.balance && bondAdjustmentPair.adjust) {
    if (applyToBalance) {
      const base = parseNumber(bondAdjustmentPair.balance.value) || 0;
      applyAdjustmentToBase(bondAdjustmentPair.balance, bondAdjustmentPair.adjust, base);
      return;
    }
    writePrevAdjustmentValue(
      bondAdjustmentPair.adjust,
      readAdjustmentValue(bondAdjustmentPair.adjust)
    );
  }
}

function readCashManualFlag() {
  try {
    return localStorage.getItem(CASH_MANUAL_KEY) === "1";
  } catch {
    return false;
  }
}

function writeCashManualFlag(value) {
  try {
    localStorage.setItem(CASH_MANUAL_KEY, value ? "1" : "0");
  } catch {
    // Ignore storage failures.
  }
}

function getCashInputValue() {
  if (!balanceCashInput || !cashInputManual) {
    return null;
  }
  const raw = balanceCashInput.value;
  if (raw === "" || raw === null || raw === undefined) {
    return null;
  }
  const value = parseNumber(raw);
  return Number.isFinite(value) ? value : null;
}

function getCashAdjustment() {
  return parseNumber(adjustCashInput?.value) || 0;
}

function getSummaryBreakdownSafe() {
  if (importDirty) {
    return null;
  }
  return getSummaryBreakdown(summaryDataInput.value);
}

function resolveCashBaseWithoutPoints({
  cashInputValue,
  summaryBreakdown,
  currentAssetsValue,
  investmentTotal,
}) {
  if (Number.isFinite(cashInputValue)) {
    return cashInputValue;
  }
  if (summaryBreakdown && Number.isFinite(summaryBreakdown.cash)) {
    return summaryBreakdown.cash;
  }
  const points = summaryBreakdown?.points || 0;
  const totalValue = Number.isFinite(currentAssetsValue)
    ? currentAssetsValue
    : summaryBreakdown && Number.isFinite(summaryBreakdown.total)
      ? summaryBreakdown.total
      : null;
  if (!Number.isFinite(totalValue) || !Number.isFinite(investmentTotal)) {
    return null;
  }
  return totalValue - investmentTotal - points;
}

function buildCashBreakdown({ summaryBreakdown, currentAssetsValue, investmentTotal }) {
  const cashInputValue = getCashInputValue();
  const points = summaryBreakdown?.points || 0;
  const baseWithoutPoints = resolveCashBaseWithoutPoints({
    cashInputValue,
    summaryBreakdown,
    currentAssetsValue,
    investmentTotal,
  });
  const cashReclassTotal = getCashReclassTotalYen();
  const cashAdjustment = getCashAdjustment();
  const baseWithPoints = Number.isFinite(baseWithoutPoints)
    ? baseWithoutPoints + points
    : null;
  const cashFinal = Number.isFinite(baseWithPoints)
    ? baseWithPoints - cashReclassTotal + cashAdjustment
    : null;
  return {
    baseWithoutPoints,
    points,
    cashReclassTotal,
    cashAdjustment,
    baseWithPoints,
    cashFinal,
  };
}

function updateCashDetailDisplay(breakdown) {
  if (!cashDetailBase || !cashDetailPoints || !cashDetailReclass ||
    !cashDetailAdjustment || !cashDetailFinal) {
    return;
  }
  const formatOrDash = (value) =>
    Number.isFinite(value) ? yenFormatter.format(Math.round(value)) : "-";
  if (!breakdown || !Number.isFinite(breakdown.baseWithoutPoints)) {
    cashDetailBase.textContent = "-";
    cashDetailPoints.textContent = "-";
    cashDetailReclass.textContent = "-";
    cashDetailAdjustment.textContent = "-";
    cashDetailFinal.textContent = "-";
    return;
  }
  cashDetailBase.textContent = formatOrDash(breakdown.baseWithoutPoints);
  cashDetailPoints.textContent = formatOrDash(breakdown.points);
  cashDetailReclass.textContent = formatOrDash(breakdown.cashReclassTotal);
  cashDetailAdjustment.textContent = formatOrDash(breakdown.cashAdjustment);
  cashDetailFinal.textContent = formatOrDash(breakdown.cashFinal);
}

function getAssetLabel(key) {
  return ASSET_LABELS[key] || key || "-";
}

function parseYearMonth(value) {
  if (!value) {
    return null;
  }
  const [yearText, monthText] = String(value).split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (!Number.isFinite(year) || !Number.isFinite(month)) {
    return null;
  }
  if (month < 1 || month > 12) {
    return null;
  }
  return { year, month };
}

function getInsurancePlanEndMonthIndex(plan) {
  const parsed = parseYearMonth(plan?.endMonth);
  if (!parsed) {
    return Number.POSITIVE_INFINITY;
  }
  const endDate = new Date(parsed.year, parsed.month - 1, 1);
  return monthIndex(endDate) + 1;
}

function getInsurancePlanTotalForDate(plans, atDate) {
  const monthIndexValue = monthIndex(atDate || new Date());
  return (plans || []).reduce((sum, plan) => {
    const amount = parseNumber(plan?.amount);
    if (!Number.isFinite(amount) || amount <= 0) {
      return sum;
    }
    const endMonthIndex = getInsurancePlanEndMonthIndex(plan);
    if (monthIndexValue < endMonthIndex) {
      return sum + amount;
    }
    return sum;
  }, 0);
}

function syncInsuranceContributionFromDetail(plans, atDate) {
  if (!contribInsuranceInput) {
    return false;
  }
  const amount = getInsurancePlanTotalForDate(plans, atDate);
  const nextValue = String(Math.round(amount));
  if (contribInsuranceInput.value === nextValue) {
    return false;
  }
  contribInsuranceInput.value = nextValue;
  return true;
}

function getActivePensionPlans(plans) {
  return (plans || []).filter((plan) => {
    if (!plan || !plan.id) {
      return false;
    }
    return parseNumber(plan.startAge) !== null;
  });
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function makeRowId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function addYears(date, years) {
  return new Date(date.getFullYear() + years, date.getMonth(), date.getDate());
}

function addMonths(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, date.getDate());
}

function monthIndex(date) {
  return date.getFullYear() * 12 + date.getMonth();
}

function fullMonthsBetween(startDate, endDate) {
  let months =
    (endDate.getFullYear() - startDate.getFullYear()) * 12 +
    (endDate.getMonth() - startDate.getMonth());
  if (endDate.getDate() < startDate.getDate()) {
    months -= 1;
  }
  return Math.max(0, months);
}

function getDividendMultiplier(startMonthIndex, monthIndexValue, annualRate) {
  if (!Number.isFinite(annualRate) || annualRate === 0) {
    return 1;
  }
  const startYear = Math.floor(startMonthIndex / 12);
  const currentYear = Math.floor(monthIndexValue / 12);
  const years = currentYear - startYear;
  if (years <= 0) {
    return 1;
  }
  return Math.pow(1 + annualRate, years);
}

function getDividendDelta(
  baseDividend,
  startMonthIndex,
  monthIndexValue,
  annualRate
) {
  if (!Number.isFinite(baseDividend) || baseDividend === 0) {
    return 0;
  }
  const multiplier = getDividendMultiplier(
    startMonthIndex,
    monthIndexValue,
    annualRate
  );
  if (multiplier === 1) {
    return 0;
  }
  return baseDividend * (multiplier - 1);
}

function simulateToAge100({
  currentAssets,
  annualRate,
  retirementAge,
  retirementIncomeEndAge,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  baseDividendIncome,
  dividendYieldRate,
  monthsRemaining,
  startMonthIndex,
}) {
  const months = Math.max(0, monthsRemaining);
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  let assets = currentAssets;

  for (let i = 0; i < months; i += 1) {
    const startAssets = assets;
    const monthIndex = startMonthIndex + i;
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndex,
      dividendYieldRate
    );
    let cashFlow = monthlyNetCash + dividendDelta;
    if (monthIndex >= retirementAge) {
      cashFlow = retirementMonthlyNetCash + dividendDelta;
    }
    if (monthIndex >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash + dividendDelta;
    }
    const investmentIncome = startAssets * monthlyRate;
    assets = startAssets + cashFlow + investmentIncome;
  }

  return { months, assets };
}

function simulateToAge100Detailed({
  startDate,
  monthsRemaining,
  retirementAge,
  retirementIncomeEndAge,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  baseDividendIncome,
  dividendYieldRate,
  contributionSchedule,
  categories,
  categoryRates,
  bondMaturities,
  usdRate,
  pensionPlanState,
}) {
  const months = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const maturitySchedule = buildBondMaturitySchedule(bondMaturities, usdRate);
  const planState = clonePensionPlanState(pensionPlanState);
  const investKeys = [
    "stocks",
    "funds",
    "bonds",
    "insurance",
    "dc",
    "other",
  ];

  for (let i = 0; i < months; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let cashFlow = monthlyNetCash + dividendDelta;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash + dividendDelta;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash + dividendDelta;
    }

    investKeys.forEach((key) => {
      const rate = categoryRates[key] ?? 0;
      if (isCompoundingCategory(key)) {
        data[key] += data[key] * rate;
      }
    });

    data.cash += cashFlow;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        const amount = item.amount;
        if (item.category !== "cash") {
          data.cash -= amount;
        }
        data[item.category] += amount;
      }
    });

    applyBondMaturities(data, maturitySchedule, monthIndexValue);
    applyPensionPlanFlow(data, monthIndexValue, planState);
  }

  return { months, assets: sumCategoryTotal(data) };
}

function getContributionForMonth(monthIndexValue, schedule) {
  if (!schedule || schedule.length === 0) {
    return 0;
  }
  return schedule.reduce((sum, item) => {
    if (monthIndexValue < item.endMonthIndex) {
      return sum + item.amount;
    }
    return sum;
  }, 0);
}

function buildPensionPlanState(birthDate, plans, changes, initialBalance) {
  if (!birthDate || !plans || plans.length === 0) {
    return [];
  }
  const changeMap = new Map();
  (changes || []).forEach((row) => {
    if (!row || !row.planId) {
      return;
    }
    const age = parseNumber(row.age);
    const amount = parseNumber(row.amount);
    if (age === null || amount === null) {
      return;
    }
    const monthIndexValue = monthIndex(addYears(birthDate, age));
    if (!changeMap.has(row.planId)) {
      changeMap.set(row.planId, []);
    }
    changeMap.get(row.planId).push({ monthIndex: monthIndexValue, amount });
  });

  const planStates = plans
    .map((plan) => {
      if (!plan || !plan.id) {
        return null;
      }
      const startAge = parseNumber(plan.startAge);
      if (startAge === null) {
        return null;
      }
      const contributionAmount = parseNumber(plan.amount) || 0;
      const installmentAmount = parseNumber(plan.installmentAmount) || 0;
      const payoutType = plan.payoutType === "lump" ? "lump" : "installment";
      const changesForPlan = (changeMap.get(plan.id) || []).sort(
        (a, b) => a.monthIndex - b.monthIndex
      );
      return {
        id: plan.id,
        name: plan.name || "",
        startMonthIndex: monthIndex(addYears(birthDate, startAge)),
        payoutType,
        contributionAmount,
        installmentAmount,
        changes: changesForPlan,
        balance: 0,
        paid: false,
      };
    })
    .filter(Boolean);

  if (planStates.length === 0) {
    return [];
  }
  const totalBase = planStates.reduce(
    (sum, plan) => sum + Math.max(0, plan.contributionAmount),
    0
  );
  if (Number.isFinite(initialBalance) && initialBalance > 0) {
    if (totalBase > 0) {
      planStates.forEach((plan) => {
        plan.balance +=
          initialBalance * (Math.max(0, plan.contributionAmount) / totalBase);
      });
    } else {
      const perPlan = initialBalance / planStates.length;
      planStates.forEach((plan) => {
        plan.balance += perPlan;
      });
    }
  }
  return planStates;
}

function clonePensionPlanState(planState) {
  if (!planState) {
    return [];
  }
  return planState.map((plan) => ({
    ...plan,
    changes: Array.isArray(plan.changes) ? [...plan.changes] : [],
  }));
}

function getPensionContributionAmount(monthIndexValue, plan) {
  let amount = plan.contributionAmount;
  plan.changes.forEach((change) => {
    if (change.monthIndex <= monthIndexValue) {
      amount = change.amount;
    }
  });
  return amount;
}

function applyPensionPlanFlow(data, monthIndexValue, planState) {
  if (!planState || planState.length === 0) {
    return { contribution: 0, transfer: 0 };
  }
  let contribution = 0;
  let transfer = 0;

  planState.forEach((plan) => {
    // 会計処理：年金拠出フェーズと給付フェーズの区分
    // monthIndexValue < plan.startMonthIndex：拠出期間（開始年齢到達まで毎月拠出）
    // monthIndexValue >= plan.startMonthIndex：給付期間（開始年齢以降から給付開始）
    if (monthIndexValue < plan.startMonthIndex) {
      // 拠出フェーズ：年金資産を積み立てる
      const amount = getPensionContributionAmount(monthIndexValue, plan);
      if (amount > 0) {
        contribution += amount;
        plan.balance += amount;
        data.dc += amount;
        data.cash -= amount;
      }
      return;
    }

    if (plan.payoutType === "lump") {
      // 給付フェーズ（一括給付）：開始年齢に達した月に一度だけ給付
      if (!plan.paid) {
        const transferable = Math.min(plan.balance, data.dc);
        if (transferable > 0) {
          data.dc -= transferable;
          data.cash += transferable;
          transfer += transferable;
          plan.balance -= transferable;
        }
        plan.paid = true;
      }
      return;
    }

    // 給付フェーズ（分割給付）：開始年齢以降、毎年1回に分割給付
    if ((monthIndexValue - plan.startMonthIndex) % 12 === 0) {
      const amount = plan.installmentAmount || 0;
      const transferable = Math.min(plan.balance, amount, data.dc);
      if (transferable > 0) {
        data.dc -= transferable;
        data.cash += transferable;
        transfer += transferable;
        plan.balance -= transferable;
      }
    }
  });

  return { contribution, transfer };
}

function buildBondMaturitySchedule(bondMaturities, usdRate) {
  const schedule = new Map();
  if (!bondMaturities || bondMaturities.length === 0) {
    return schedule;
  }
  bondMaturities.forEach((bond) => {
    if (!bond.maturityDate) {
      return;
    }
    const monthKey = monthIndex(bond.maturityDate);
    const rate = bond.currency === "USD" ? usdRate : 1;
    const faceValue = Number.isFinite(bond.faceValue) ? bond.faceValue : 0;
    const bookValue = Number.isFinite(bond.bookValue)
      ? bond.bookValue
      : faceValue;
    const faceAmount = toYenAmount(faceValue * rate);
    const bookAmount = toYenAmount(bookValue * rate);
    if (!faceAmount && !bookAmount) {
      return;
    }
    const entry = schedule.get(monthKey) || { faceValue: 0, bookValue: 0 };
    entry.faceValue += faceAmount;
    entry.bookValue += bookAmount;
    schedule.set(monthKey, entry);
  });
  return schedule;
}

function applyBondMaturities(data, schedule, monthIndexValue) {
  if (!schedule || !schedule.size) {
    return { faceValue: 0, bookValue: 0, gain: 0 };
  }
  const entry = schedule.get(monthIndexValue);
  if (!entry) {
    return { faceValue: 0, bookValue: 0, gain: 0 };
  }
  const faceAmount = entry.faceValue || 0;
  const bookAmount = entry.bookValue || entry.faceValue || 0;
  if (!faceAmount && !bookAmount) {
    return { faceValue: 0, bookValue: 0, gain: 0 };
  }
  const before = data.bonds;
  const reducible = Math.min(before, bookAmount);
  data.bonds = before - reducible;
  data.cash += faceAmount;
  if (bookAmount > before) {
    console.warn(
      `警告：債券償還時の評価額が残高を超えています（評価額${bookAmount}、残高${before}）。`
    );
  }
  return {
    faceValue: faceAmount,
    bookValue: reducible,
    gain: faceAmount - reducible,
  };
}



function findNegativeCashMonth({
  startCash,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  baseDividendIncome,
  dividendYieldRate,
  retirementAge,
  retirementIncomeEndAge,
  contributionSchedule,
  monthsRemaining,
  startMonthIndex,
}) {
  let cash = startCash;
  const months = Math.max(0, monthsRemaining);
  for (let i = 0; i < months; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let cashFlow = monthlyNetCash + dividendDelta;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash + dividendDelta;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash + dividendDelta;
    }
    cash += cashFlow;
    cash -= getContributionForMonth(monthIndexValue, contributionSchedule);
    if (cash < 0) {
      return monthIndexValue;
    }
  }
  return null;
}

function findNegativeCashMonthDetailed({
  startDate,
  monthsRemaining,
  annualRate,
  categoryRates,
  retirementAge,
  retirementIncomeEndAge,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  baseDividendIncome,
  dividendYieldRate,
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  pensionPlanState,
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const maturitySchedule = buildBondMaturitySchedule(bondMaturities, usdRate);
  const planState = clonePensionPlanState(pensionPlanState);
  const investKeys = [
    "stocks",
    "funds",
    "bonds",
    "insurance",
    "dc",
    "other",
  ];

  for (let i = 0; i < totalMonths; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const monthDate = addMonths(startDate, i);
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let cashFlow = monthlyNetCash + dividendDelta;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash + dividendDelta;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash + dividendDelta;
    }

    investKeys.forEach((key) => {
      const rate = categoryRates?.[key] ?? monthlyRate;
      if (isCompoundingCategory(key)) {
        data[key] += data[key] * rate;
      }
    });

    data.cash += cashFlow;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        const amount = item.amount;
        if (item.category !== "cash") {
          data.cash -= amount;
        }
        data[item.category] += amount;
      }
    });

    applyBondMaturities(data, maturitySchedule, monthIndexValue);
    applyPensionPlanFlow(data, monthIndexValue, planState);

    if (data.cash < 0) {
      return {
        monthIndex: monthIndexValue,
        date: addMonths(startDate, i + 1),
        cash: data.cash,
      };
    }
  }

  return null;
}

function simulateAnnualSeries({
  startDate,
  monthsRemaining,
  annualRate,
  categoryRates,
  retirementAge,
  retirementIncomeEndAge,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  baseDividendIncome,
  dividendYieldRate,
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  pensionPlanState,
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const rows = [];
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const maturitySchedule = buildBondMaturitySchedule(bondMaturities, usdRate);
  const planState = clonePensionPlanState(pensionPlanState);
  const investKeys = [
    "stocks",
    "funds",
    "bonds",
    "insurance",
    "dc",
    "other",
  ];

  for (let i = 0; i < totalMonths; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const monthDate = addMonths(startDate, i);
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let cashFlow = monthlyNetCash + dividendDelta;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash + dividendDelta;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash + dividendDelta;
    }

    investKeys.forEach((key) => {
      const rate = categoryRates?.[key] ?? monthlyRate;
      if (isCompoundingCategory(key)) {
        data[key] += data[key] * rate;
      }
    });

    data.cash += cashFlow;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        const amount = item.amount;
        if (item.category !== "cash") {
          data.cash -= amount;
        }
        data[item.category] += amount;
      }
    });

    applyBondMaturities(data, maturitySchedule, monthIndexValue);
    applyPensionPlanFlow(data, monthIndexValue, planState);

    const isYearEnd = monthDate.getMonth() === 11;
    const isFinal = i === totalMonths - 1;
    if (isYearEnd || isFinal) {
      const date = addMonths(startDate, i + 1);
      const total = sumCategoryTotal(data);
      rows.push({
        date,
        total,
        ...data,
      });
    }
  }

  return rows;
}

function simulateMonthlySeries({
  startDate,
  monthsRemaining,
  annualRate,
  categoryRates,
  retirementAge,
  retirementIncomeEndAge,
  workIncome,
  workExpense,
  retireIncome,
  retireExpense,
  pensionIncome,
  pensionExpense,
  baseDividendIncome,
  dividendYieldRate,
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  pensionPlanState,
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const rows = [];
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const maturitySchedule = buildBondMaturitySchedule(bondMaturities, usdRate);
  const planState = clonePensionPlanState(pensionPlanState);
  const investKeys = [
    "stocks",
    "funds",
    "bonds",
    "insurance",
    "dc",
    "other",
  ];

  for (let i = 0; i < totalMonths; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const monthDate = addMonths(startDate, i);
    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let monthlyIncome = workIncome + dividendDelta;
    let monthlyExpense = workExpense;
    if (monthIndexValue >= retirementAge) {
      monthlyIncome = retireIncome + dividendDelta;
      monthlyExpense = retireExpense;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      monthlyIncome = pensionIncome + dividendDelta;
      monthlyExpense = pensionExpense;
    }

    investKeys.forEach((key) => {
      const rate = categoryRates?.[key] ?? monthlyRate;
      if (isCompoundingCategory(key)) {
        data[key] += data[key] * rate;
      }
    });

    data.cash += monthlyIncome - monthlyExpense;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        const amount = item.amount;
        if (item.category !== "cash") {
          data.cash -= amount;
        }
        data[item.category] += amount;
      }
    });

    applyBondMaturities(data, maturitySchedule, monthIndexValue);
    applyPensionPlanFlow(data, monthIndexValue, planState);

    rows.push({
      date: monthDate,
      ...data,
    });
  }

  return rows;
}

function sumCategoryTotal(data) {
  return (
    data.cash +
    data.stocks +
    data.funds +
    data.bonds +
    data.insurance +
    data.dc +
    (data.usd || 0) +
    data.other
  );
}

function simulateAnnualStatements({
  startDate,
  monthsRemaining,
  annualRate,
  categoryRates,
  retirementAge,
  retirementIncomeEndAge,
  workIncome,
  workExpense,
  retireIncome,
  retireExpense,
  pensionIncome,
  pensionExpense,
  baseDividendIncome,
  dividendYieldRate,
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  pensionPlanState,
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const rows = [];
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const maturitySchedule = buildBondMaturitySchedule(bondMaturities, usdRate);
  const planState = clonePensionPlanState(pensionPlanState);
  const investKeys = [
    "stocks",
    "funds",
    "bonds",
    "insurance",
    "dc",
    "other",
  ];

  // 会計処理：年次決算データ初期化
  // 年度開始時点の残高、年間の収支、運用益を月次で累積して年度決算を作成
  let yearStart = { ...data };  // 期首残高
  let yearIncome = 0;            // 総現金収入
  let yearCashIncome = 0;        // 給与・pension等の現金収入
  let yearInvestmentIncome = 0;  // 配当等の投資収入（未使用）
  let yearExpense = 0;           // 総現金支出
  let yearInvestmentGain = 0;    // 運用益（含み益）
  let yearContribution = 0;      // 投資支出（資産移動）
  let yearBondMaturityFace = 0;  // 債券償還による現金化（額面）
  let yearBondMaturityBook = 0;  // 債券償還による残高減少（評価額）
  let yearPensionTransfer = 0;   // 年金給付による現金化
  let yearContributionByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    dc: 0,
    usd: 0,
    other: 0,
  };
  let yearGainByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    dc: 0,
    usd: 0,
    other: 0,
  };
  let yearWorkMonths = 0;
  let yearRetireMonths = 0;
  let yearPensionMonths = 0;
  let yearMonths = 0;

  for (let i = 0; i < totalMonths; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const monthDate = addMonths(startDate, i);

    const dividendDelta = getDividendDelta(
      baseDividendIncome,
      startMonthIndex,
      monthIndexValue,
      dividendYieldRate
    );
    let monthlyIncome = workIncome + dividendDelta;
    let monthlyExpense = workExpense;
    let phase = "work";
    if (monthIndexValue >= retirementAge) {
      monthlyIncome = retireIncome + dividendDelta;
      monthlyExpense = retireExpense;
      phase = "retire";
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      monthlyIncome = pensionIncome + dividendDelta;
      monthlyExpense = pensionExpense;
      phase = "pension";
    }

    // 会計処理：運用益（含み益）の計算
    // 複利計算対象カテゴリーのみ期中の利息/配当を認識
    // 注：実現益と含み益を区分していない（含み益ベースで計算）
    let monthlyInvestmentGain = 0;
    investKeys.forEach((key) => {
      const rate = categoryRates?.[key] ?? monthlyRate;
      const gain = isCompoundingCategory(key) ? data[key] * rate : 0;
      if (isCompoundingCategory(key)) {
        data[key] += gain;  // 運用益を資産に加算（含み益の認識）
      }
      yearGainByCategory[key] += gain;
      monthlyInvestmentGain += gain;
    });

    data.cash += monthlyIncome - monthlyExpense;
    yearIncome += monthlyIncome;
    yearCashIncome += monthlyIncome;
    yearExpense += monthlyExpense;
    if (phase === "work") {
      yearWorkMonths += 1;
    } else if (phase === "retire") {
      yearRetireMonths += 1;
    } else {
      yearPensionMonths += 1;
    }

    // 会計処理：投資支出（資産配分）
    // 拠出は資産の増加（投資信託や保険への振替）であり、損益計算書上の支出ではない
    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        const amount = item.amount;
        yearContribution += amount;
        if (yearContributionByCategory[item.category] !== undefined) {
          yearContributionByCategory[item.category] += amount;
        }
        if (item.category !== "cash") {
          data.cash -= amount;  // 現金減少
        }
        data[item.category] += amount;  // 対応する資産カテゴリー増加
      }
    });

    const bondMaturityFlow = applyBondMaturities(
      data,
      maturitySchedule,
      monthIndexValue
    );
    yearBondMaturityFace += bondMaturityFlow.faceValue || 0;
    yearBondMaturityBook += bondMaturityFlow.bookValue || 0;
    // 債券の差額は利息（現金）として扱うため運用益には含めない
    const pensionFlow = applyPensionPlanFlow(
      data,
      monthIndexValue,
      planState
    );
    if (pensionFlow.contribution) {
      yearContribution += pensionFlow.contribution;
      yearContributionByCategory.dc += pensionFlow.contribution;
    }
    yearPensionTransfer += pensionFlow.transfer;
    yearInvestmentGain += monthlyInvestmentGain;
    yearMonths += 1;

    const isYearEnd = monthDate.getMonth() === 11;
    const isFinal = i === totalMonths - 1;
    if (isYearEnd || isFinal) {
      const endDate = addMonths(startDate, i + 1);
      const startTotal = sumCategoryTotal(yearStart);
      const endTotal = sumCategoryTotal(data);
      // 会計処理：年度決算の資産増減分析
      // netCash = 現金ベースの収支（収入 - 支出）
      // investmentGain = 運用益（含み益）
      // bondMaturityGain = 債券償還の差額（額面 - 評価額）
      // totalChange = 資産増減の総額（netCash + investmentGain + bondMaturityGain）
      // 期末残高 = 期首残高 + 資産増減総額
      const netCash = yearIncome - yearExpense;
      const bondMaturityGain = yearBondMaturityFace - yearBondMaturityBook;
      const totalChange = netCash + yearInvestmentGain + bondMaturityGain;
      const mismatch =
        Math.round(startTotal + totalChange) !== Math.round(endTotal);
      const yearDividendDelta = getDividendDelta(
        baseDividendIncome,
        startMonthIndex,
        monthIndexValue,
        dividendYieldRate
      );
      rows.push({
        year: monthDate.getFullYear(),
        date: endDate,
        start: { ...yearStart, total: startTotal },
        end: { ...data, total: endTotal },
        income: yearIncome,
        cashIncome: yearCashIncome,
        investmentIncome: yearInvestmentIncome,
        expense: yearExpense,
        netCash,
        investmentGain: yearInvestmentGain,
        totalChange,
        bondMaturity: yearBondMaturityFace,
        bondMaturityBook: yearBondMaturityBook,
        pensionTransfer: yearPensionTransfer,
        months: yearMonths,
        contributions: yearContribution,
        contributionsByCategory: { ...yearContributionByCategory },
        gainsByCategory: { ...yearGainByCategory },
        workMonths: yearWorkMonths,
        retireMonths: yearRetireMonths,
        pensionMonths: yearPensionMonths,
        workIncome: workIncome + yearDividendDelta,
        retireIncome: retireIncome + yearDividendDelta,
        pensionIncome: pensionIncome + yearDividendDelta,
        workExpense,
        retireExpense,
        pensionExpense,
        mismatch,
      });
      yearStart = { ...data };
      yearIncome = 0;
      yearCashIncome = 0;
      yearInvestmentIncome = 0;
      yearExpense = 0;
      yearInvestmentGain = 0;
      yearBondMaturityFace = 0;
      yearBondMaturityBook = 0;
      yearPensionTransfer = 0;
      yearContribution = 0;
      yearContributionByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        dc: 0,
        usd: 0,
        other: 0,
      };
      yearGainByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        dc: 0,
        usd: 0,
        other: 0,
      };
      yearWorkMonths = 0;
      yearRetireMonths = 0;
      yearPensionMonths = 0;
      yearMonths = 0;
    }
  }

  return rows;
}

function getSummaryRowValues(headers, row) {
  const getAmount = (regex) => {
    const idx = mapHeaderIndex(headers, regex);
    if (idx === null) {
      return 0;
    }
    return parseAmount(row[idx] || "") || 0;
  };
  const totalIndex = findSummaryTotalIndex(headers);
  const rawTotal =
    totalIndex === null ? 0 : parseAmount(row[totalIndex] || "") || 0;
  const cash = getAmount(/預金|現金|暗号資産|仮想通貨/);
  const stocks = getAmount(/株式/);
  const funds = getAmount(/投資信託/);
  const bonds = getAmount(/債券/);
  const insurance = getAmount(/保険/);
  const pension = getAmount(/年金/);
  const points = getAmount(/ポイント/);
  const other = getAmount(/その他の資産|その他資産|その他/);
  const cashWithPoints = cash + points;
  const investmentsTotal =
    stocks + funds + bonds + insurance + pension + other;
  const derivedTotal = cashWithPoints + investmentsTotal;
  const total =
    derivedTotal > 0 &&
    cashWithPoints > 0 &&
    (rawTotal <= 0 || rawTotal <= investmentsTotal)
      ? derivedTotal
      : rawTotal;
  return {
    total,
    cash,
    stocks,
    funds,
    bonds,
    insurance,
    pension,
    points,
    other,
  };
}

function getSummaryBreakdown(text) {
  const table = parseTable(text);
  if (!table) {
    return null;
  }

  const dateIndex = mapHeaderIndex(table.headers, /日付/);
  let bestRow = null;
  let bestRank = null;

  table.dataRows.forEach((row, index) => {
    const dateValue =
      dateIndex === null ? null : parseDateValue(row[dateIndex] || "");
    const rank = dateValue ?? index;
    if (bestRank === null || rank > bestRank) {
      bestRank = rank;
      bestRow = row;
    }
  });

  if (!bestRow) {
    return null;
  }

  return getSummaryRowValues(table.headers, bestRow);
}

function buildContributionSchedule(birthDate, options = {}) {
  if (!birthDate) {
    return [];
  }
  const skipDc = options.skipDc === true;
  const insurancePlans = Array.isArray(options.insurancePlans)
    ? options.insurancePlans
    : [];
  const toEndMonth = (input, fallbackAge = 120) => {
    const ageValue = parseNumber(input?.value);
    const age = ageValue ?? fallbackAge;
    return monthIndex(addYears(birthDate, age));
  };
  const insuranceItems = insurancePlans
    .map((plan) => {
      const amount = parseNumber(plan?.amount);
      if (!Number.isFinite(amount) || amount <= 0) {
        return null;
      }
      return {
        category: "insurance",
        amount,
        endMonthIndex: getInsurancePlanEndMonthIndex(plan),
      };
    })
    .filter(Boolean);

  const schedule = [
    {
      category: "stocks",
      amount: parseNumber(contribStocksInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeStocksInput),
    },
    {
      category: "funds",
      amount: parseNumber(contribFundsInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeFundsInput),
    },
    ...(contribBondsInput
      ? [
          {
            category: "bonds",
            amount: parseNumber(contribBondsInput.value) || 0,
            endMonthIndex: toEndMonth(endAgeBondsInput),
          },
        ]
      : []),
    ...(insuranceItems.length
      ? insuranceItems
      : [
          {
      category: "insurance",
      amount: parseNumber(contribInsuranceInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeInsuranceInput),
    },
        ]),
    {
      category: "usd",
      amount: parseNumber(contribUsdInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeUsdInput),
    },
  ];
  if (!skipDc) {
    schedule.push({
      category: "dc",
      amount: parseNumber(contribDcInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeDcInput),
    });
  }
  return schedule;
}

function getInvestmentBalanceTotal() {
  return (
    (parseNumber(balanceStocksInput.value) || 0) +
    (parseNumber(balanceFundsInput.value) || 0) +
    (parseNumber(balanceBondsInput.value) || 0) +
    (parseNumber(balanceInsuranceInput.value) || 0) +
    (parseNumber(balanceUsdInput.value) || 0) +
    (parseNumber(balanceDcInput.value) || 0)
  );
}

function updateCurrentAssetsFromInvestmentBalances() {
  if (!currentAssetsInput) {
    return;
  }
  const investmentTotal = getInvestmentBalanceTotal();
  const summaryBreakdown = getSummaryBreakdownSafe();
  const points = summaryBreakdown?.points || 0;
  const cashInputValue = getCashInputValue();
  const cashAdjustment = getCashAdjustment();
  if (Number.isFinite(cashInputValue)) {
    const cashReclassTotal = getCashReclassTotalYen();
    currentAssetsInput.value = Math.round(
      investmentTotal +
        cashInputValue +
        points -
        cashReclassTotal +
        cashAdjustment
    );
    lastInvestmentBalanceTotal = investmentTotal;
    render();
    return;
  }
  const currentTotal = parseNumber(currentAssetsInput.value);
  const previousInvestmentTotal = lastInvestmentBalanceTotal;
  if (
    Number.isFinite(currentTotal) &&
    Number.isFinite(previousInvestmentTotal)
  ) {
    const cashBaseline = currentTotal - previousInvestmentTotal;
    currentAssetsInput.value = Math.round(investmentTotal + cashBaseline);
    lastInvestmentBalanceTotal = investmentTotal;
    render();
    return;
  }
  currentAssetsInput.value = Math.round(investmentTotal);
  lastInvestmentBalanceTotal = investmentTotal;
  render();
}

function syncCurrentAssetsFromCashInput() {
  if (!currentAssetsInput) {
    return;
  }
  const investmentTotal = getInvestmentBalanceTotal();
  const cashInputValue = getCashInputValue();
  if (!Number.isFinite(cashInputValue)) {
    return;
  }
  const summaryBreakdown = getSummaryBreakdownSafe();
  const points = summaryBreakdown?.points || 0;
  const cashAdjustment = getCashAdjustment();
  const cashReclassTotal = getCashReclassTotalYen();
  currentAssetsInput.value = Math.round(
    investmentTotal +
      cashInputValue +
      points -
      cashReclassTotal +
      cashAdjustment
  );
  lastInvestmentBalanceTotal = investmentTotal;
  render();
}

// 会計処理：初期資産の分類
// マネーフォワードなどからのインポートデータを、各資産カテゴリーに分類
// ユーザーの調整値を加算して初期残高を確定
function buildInitialCategories(summaryBreakdown, currentAssets) {
  const storedBonds = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? storedBonds.usdRate) ?? 0;
  const otherAssetsCashReclass = getOtherAssetsCashReclassTotalYen(
    readOtherAssetsStorage(),
    usdRate
  );
  const readBalanceValue = (input, fallback = 0) => {
    const value = parseNumber(input?.value);
    if (Number.isFinite(value)) {
      return value;
    }
    return Number.isFinite(fallback) ? fallback : 0;
  };
  if (summaryBreakdown) {
    const totalFromInput =
      Number.isFinite(currentAssets) ?
        currentAssets :
        summaryBreakdown.total ||
          summaryBreakdown.cash +
            summaryBreakdown.points +
            summaryBreakdown.stocks +
            summaryBreakdown.funds +
            summaryBreakdown.bonds +
            summaryBreakdown.insurance +
            summaryBreakdown.pension +
            summaryBreakdown.other;
    const pensionTotal = summaryBreakdown.pension || 0;
    const dc = readBalanceValue(balanceDcInput, pensionTotal);
    const usdBalance = readBalanceValue(balanceUsdInput, 0);
    const stocks = readBalanceValue(balanceStocksInput, summaryBreakdown.stocks);
    const funds = readBalanceValue(balanceFundsInput, summaryBreakdown.funds);
    const bondInputTotal = readBalanceValue(balanceBondsInput, summaryBreakdown.bonds);
    const bondReclass = getBondCashReclassTotalYen(storedBonds, usdRate);
    const bonds = bondInputTotal;
    const insurance = readBalanceValue(
      balanceInsuranceInput,
      summaryBreakdown.insurance
    );
    const points = summaryBreakdown.points || 0;
    const other = summaryBreakdown.other || 0;
    const usd = usdBalance;
    const cashFromSummaryBase =
      Number.isFinite(summaryBreakdown.cash) ? summaryBreakdown.cash : null;
    const investmentTotal = stocks + funds + bonds + insurance + dc + other + usd;
    const cashForTotal = cashFromSummaryBase !== null
      ? cashFromSummaryBase + points - (bondReclass + otherAssetsCashReclass)
      : 0;
    const derivedTotal = cashForTotal + investmentTotal;
    const total =
      derivedTotal > 0 && totalFromInput < derivedTotal
        ? derivedTotal
        : totalFromInput;
    const cashInputValue = getCashInputValue();
    const cashAdjustment = getCashAdjustment();
    const cashReclassTotal = bondReclass + otherAssetsCashReclass;
    const baseValue = resolveCashBaseWithoutPoints({
      cashInputValue,
      summaryBreakdown,
      currentAssetsValue: total,
      investmentTotal,
    });
    const cashBaseWithPoints = Number.isFinite(baseValue)
      ? baseValue + points
      : baseValue;
    const cash = Number.isFinite(cashBaseWithPoints)
      ? cashBaseWithPoints - cashReclassTotal + cashAdjustment
      : cashBaseWithPoints;
    if (cash < 0) {
      console.warn(
        `警告：マネーフォワード取込後、投資資産の合計が総資産を超えています。` +
        `調整値を確認してください。現金残高がマイナス（${cash}）です。`
      );
    }
    
    return {
      cash: cash || 0,
      stocks,
      funds,
      bonds,
      insurance,
      dc,
      usd,
      points: 0,
      other,
      total: total || 0,
    };
  }

  const stocks = readBalanceValue(balanceStocksInput, 0);
  const funds = readBalanceValue(balanceFundsInput, 0);
  const bonds = readBalanceValue(balanceBondsInput, 0);
  const insurance = readBalanceValue(balanceInsuranceInput, 0);
  const dc = readBalanceValue(balanceDcInput, 0);
  const usd = readBalanceValue(balanceUsdInput, 0);
  // 会計処理：初期現金残高の計算
  // 現在資産 = 現金 + 各投資資産 の関係から現金を逆算
  // 投資資産の合計が現在資産を超えないようにチェック（超える場合は警告）
  const other = 0;
  const investmentTotal = stocks + funds + bonds + insurance + dc + usd + other;
  const currentAssetsValue = parseNumber(currentAssets) || 0;
  let cash = currentAssetsValue - investmentTotal;
  
  // 注：投資資産が現在資産を超える場合、現金がマイナスになることを警告
  // これはユーザーが入力値を誤った場合に発生
  if (cash < 0) {
    console.warn(
      `警告：投資資産の合計（${investmentTotal}）が現在資産総額（${currentAssetsValue}）を超えています。` +
      `現金残高がマイナス（${cash}）になります。入力値を確認してください。`
    );
  }

  const cashInputValue = getCashInputValue();
  const cashAdjustment = getCashAdjustment();
  const cashReclassTotal = getCashReclassTotalYen();
  if (Number.isFinite(cashInputValue)) {
    cash = cashInputValue - cashReclassTotal;
  } else {
    cash = Math.max(0, cash - cashReclassTotal);
  }
  cash += cashAdjustment;

  return {
    cash,
    stocks,
    funds,
    bonds,
    insurance,
    dc,
    usd,
    points: 0,
    other,
    total: currentAssetsValue,
  };
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getMonthStart(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function formatAgeYears(birthDate, atDate) {
  if (!birthDate || !atDate) {
    return "-";
  }
  const months = fullMonthsBetween(birthDate, atDate);
  return Math.floor(months / 12);
}

function updateRetireExpensePlaceholders() {
  const baseByKey = new Map();
  expenseInputs.forEach((input) => {
    const key = input.dataset.expenseKey;
    if (!key) {
      return;
    }
    const value = parseNumber(input.value);
    baseByKey.set(key, value);
  });
  retireExpenseInputs.forEach((input) => {
    const key = input.dataset.expenseKey;
    if (!key) {
      return;
    }
    const baseValue = baseByKey.get(key);
    const placeholderValue =
      Number.isFinite(baseValue) && baseValue > 0 ? String(Math.round(baseValue)) : "";
    if (input.placeholder !== placeholderValue) {
      input.placeholder = placeholderValue;
    }
    const hint = input
      .closest(".input-with-hint")
      ?.querySelector(".retire-expense-hint");
    if (hint) {
      if (placeholderValue) {
        if (hint.textContent !== placeholderValue) {
          hint.textContent = placeholderValue;
        }
        hint.hidden = false;
      } else {
        hint.textContent = "";
        hint.hidden = true;
      }
    }
  });
}

function getPeriodStartDate(periodEndDate, months) {
  if (!periodEndDate || !Number.isFinite(months)) {
    return null;
  }
  return addMonths(periodEndDate, -months);
}

function formatDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

function getGitHubRepoFromLocation() {
  if (!window.location || !window.location.hostname) {
    return null;
  }
  const host = window.location.hostname;
  if (!host.endsWith("github.io")) {
    return null;
  }
  const owner = host.split(".")[0];
  const pathParts = window.location.pathname.split("/").filter(Boolean);
  const repo = pathParts[0];
  if (!owner || !repo) {
    return null;
  }
  return { owner, repo };
}

function resolveGitHubRepo() {
  return getGitHubRepoFromLocation() || GITHUB_REPO_FALLBACK;
}

function setLastUpdatedText(date) {
  if (!lastUpdated) {
    return;
  }
  lastUpdated.textContent = date ? `更新日時: ${formatDateTime(date)}` : "更新日時: -";
}

function fetchLastPushTimestamp() {
  const repo = resolveGitHubRepo();
  if (!repo) {
    return Promise.resolve(null);
  }
  const url = `https://api.github.com/repos/${repo.owner}/${repo.repo}/commits/main`;
  return fetch(url)
    .then((response) => (response.ok ? response.json() : null))
    .then((data) => {
      const rawDate =
        data?.commit?.committer?.date || data?.commit?.author?.date || null;
      if (!rawDate) {
        return null;
      }
      const parsed = new Date(rawDate);
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    })
    .catch(() => null);
}

function updateLastUpdatedFromPush() {
  const cached = safeParseJson(localStorage.getItem(PUSH_TIMESTAMP_KEY), null);
  if (cached && cached.value) {
    const cachedDate = new Date(cached.value);
    if (!Number.isNaN(cachedDate.getTime())) {
      setLastUpdatedText(cachedDate);
    }
  } else {
    setLastUpdatedText(null);
  }
  fetchLastPushTimestamp().then((date) => {
    if (!date) {
      return;
    }
    setLastUpdatedText(date);
    try {
      localStorage.setItem(
        PUSH_TIMESTAMP_KEY,
        JSON.stringify({ value: date.toISOString() })
      );
    } catch {
      // Ignore storage failures.
    }
  });
}

function toCsvNumber(value) {
  return Math.round(value);
}

function escapeCsvCell(value) {
  const text = String(value).replace(/"/g, '""');
  return `"${text}"`;
}

function buildSignedExpression(terms) {
  if (!terms.length) {
    return "0";
  }
  const [first, ...rest] = terms.map((value) => toCsvNumber(value));
  let expression = `${first}`;
  rest.forEach((value) => {
    if (value < 0) {
      expression += `-${Math.abs(value)}`;
      return;
    }
    expression += `+${value}`;
  });
  return expression;
}

function buildMultiplicationExpression(items) {
  const parts = items
    .filter((item) => item.months > 0 && item.amount !== 0)
    .map((item) => `${toCsvNumber(item.amount)}*${item.months}`);
  if (parts.length === 0) {
    return "0";
  }
  return parts.join("+");
}

function csvCellWithFormula(value, expression) {
  return escapeCsvCell(`${toCsvNumber(value)}\n=${expression}`);
}

async function saveCsvWithPicker(csv, filename) {
  if (!window.showSaveFilePicker) {
    window.alert("このブラウザは保存場所の指定に対応していません。");
    return;
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: filename,
      types: [
        {
          description: "CSVファイル",
          accept: { "text/csv": [".csv"] },
        },
      ],
    });
    const writable = await handle.createWritable();
    await writable.write(new Blob([csv], { type: "text/csv;charset=utf-8;" }));
    await writable.close();
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    window.alert("ファイルの保存に失敗しました。");
  }
}

function downloadCsv(rows, birthDate, options = {}) {
  const todayRow = options.todayRow ?? null;
  const header =
    "日付,年齢,合計（円）,預金・現金・暗号資産（円）,株式(現物)（円）,投資信託（円）,債券（円）,保険（円）,年金（円）";
  const buildLine = (row) =>
    [
      formatDate(row.date),
      formatAgeYears(birthDate, row.date),
      toCsvNumber(row.total),
      toCsvNumber(row.cash),
      toCsvNumber(row.stocks),
      toCsvNumber(row.funds),
      toCsvNumber(row.bonds),
      toCsvNumber(row.insurance),
      toCsvNumber(row.dc || 0),
    ].join(",");
  const lines = [];
  if (todayRow) {
    lines.push(buildLine(todayRow));
  }
  rows.forEach((row) => lines.push(buildLine(row)));
  const csv = [header, ...lines].join("\n");
  void saveCsvWithPicker(csv, "LifeWealth100_annual.csv");
}

function downloadCsvText(csv, filename) {
  void saveCsvWithPicker(csv, filename);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toTsvCell(value) {
  return String(value ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
}

function openSpreadsheetView({ title, headers, rows }) {
  const popup = window.open("", "_blank");
  if (!popup) {
    window.alert("ポップアップがブロックされました。ブラウザ設定を確認してください。");
    return;
  }
  const safeTitle = escapeHtml(title || "CSV");
  const headHtml = headers
    .map((header) => `<th>${escapeHtml(header)}</th>`)
    .join("");
  const bodyHtml = rows
    .map(
      (row) =>
        `<tr>${row
          .map((cell) => `<td>${escapeHtml(cell)}</td>`)
          .join("")}</tr>`
    )
    .join("");
  const tsv = [headers, ...rows]
    .map((row) => row.map(toTsvCell).join("\t"))
    .join("\n");
  const html = `<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${safeTitle}</title>
  <style>
    body { font-family: system-ui, -apple-system, sans-serif; margin: 24px; color: #222; }
    header { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 16px; }
    h1 { font-size: 18px; margin: 0; }
    button { border: none; border-radius: 999px; background: #2f6b5c; color: #fff; padding: 8px 14px; cursor: pointer; }
    table { border-collapse: collapse; width: 100%; font-size: 12px; }
    th, td { border: 1px solid #ddd; padding: 6px 8px; text-align: right; }
    th:first-child, td:first-child { text-align: left; }
    th:nth-child(2), td:nth-child(2) { text-align: left; }
    .copy-panel { margin-top: 16px; display: grid; gap: 8px; }
    textarea { width: 100%; min-height: 140px; padding: 8px; font-size: 12px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
    .note { color: #666; font-size: 12px; margin: 0; }
  </style>
</head>
<body>
  <header>
    <h1>${safeTitle}</h1>
    <button id="copyTsv" type="button">TSVをコピー</button>
  </header>
  <table>
    <thead>
      <tr>${headHtml}</tr>
    </thead>
    <tbody>
      ${bodyHtml}
    </tbody>
  </table>
  <div class="copy-panel">
    <p class="note">スプレッドシートに貼り付ける場合はTSVをコピーしてください。</p>
    <textarea id="tsvArea" readonly></textarea>
  </div>
  <script>
    const tsv = ${JSON.stringify(tsv)};
    const area = document.getElementById("tsvArea");
    area.value = tsv;
    document.getElementById("copyTsv").addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(tsv);
        alert("TSVをコピーしました");
      } catch (error) {
        area.focus();
        area.select();
      }
    });
  </script>
</body>
</html>`;
  popup.document.open();
  popup.document.write(html);
  popup.document.close();
}

function detectDelimiter(lines) {
  if (lines.some((line) => line.includes("\t"))) {
    return "\t";
  }
  if (lines.some((line) => line.includes(","))) {
    return ",";
  }
  if (lines.some((line) => /\s{2,}/.test(line))) {
    return /\s{2,}/;
  }
  return null;
}

function parseDelimitedLine(line, delimiter) {
  if (delimiter instanceof RegExp) {
    return line.split(delimiter).map((cell) => cell.trim());
  }
  const cells = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === "\"") {
      if (inQuotes && line[i + 1] === "\"") {
        current += "\"";
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && ch === delimiter) {
      cells.push(current.trim());
      current = "";
      continue;
    }
    current += ch;
  }
  cells.push(current.trim());
  return cells;
}

function parseTable(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length === 0) {
    return null;
  }

  const delimiter = detectDelimiter(lines);
  if (!delimiter) {
    return null;
  }

  const rows = lines.map((line) => parseDelimitedLine(line, delimiter));
  const headers = rows[0];
  const dataRows = rows.slice(1);
  return { headers, dataRows };
}

function mapHeaderIndex(headers, regex) {
  const index = headers.findIndex((header) => regex.test(header));
  return index === -1 ? null : index;
}

function parseAssetTable(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length === 0) {
    return null;
  }

  const delimiter = detectDelimiter(lines);
  if (!delimiter) {
    return null;
  }

  const rows = lines
    .map((line) => parseDelimitedLine(line, delimiter))
    .filter((row) => row.length > 1);
  if (!rows.length) {
    return null;
  }

  const headerKeywords = [
    /資産区分/,
    /資産種別/,
    /種別/,
    /分類/,
    /カテゴリ/,
    /カテゴリー/,
    /大分類/,
    /中分類/,
    /名称/,
    /銘柄/,
    /商品/,
    /口座/,
    /通貨/,
    /評価額/,
    /残高/,
    /金額/,
    /現在高/,
  ];

  let headerIndex = 0;
  let bestScore = 0;
  rows.forEach((row, idx) => {
    const score = row.reduce((count, cell) => {
      const text = String(cell ?? "");
      return count + (headerKeywords.some((rx) => rx.test(text)) ? 1 : 0);
    }, 0);
    if (score > bestScore) {
      bestScore = score;
      headerIndex = idx;
    }
  });

  const headers = rows[headerIndex];
  const dataRows = rows.slice(headerIndex + 1);
  return { headers, dataRows };
}

function findSummaryTotalIndex(headers) {
  const normalized = headers.map((header) => String(header || ""));
  const find = (regex, excludeRegex) => {
    for (let i = 0; i < normalized.length; i += 1) {
      const header = normalized[i];
      if (regex.test(header) && (!excludeRegex || !excludeRegex.test(header))) {
        return i;
      }
    }
    return null;
  };
  return (
    find(/純資産|資産純額|純資産合計/) ??
    find(/資産合計|総資産|資産総額/, /投資|運用|金融|有価証券/) ??
    find(/合計|総計/, /負債|投資|収支|損益/)
  );
}

function normalizeImportText(value) {
  return String(value ?? "")
    .replace(/[\s　]+/g, "")
    .replace(/[‐‑‒–—−ー－]/g, "-")
    .toLowerCase();
}

function isUsdDepositRow({ type, category, name, rowText }) {
  const text = `${type || ""} ${category || ""} ${name || ""} ${rowText || ""}`;
  const compact = normalizeImportText(text);
  return /米ドル普通/.test(compact);
}

function shouldReclassifyToOther({ type, category, name }) {
  if (isUsdDepositRow({ type, category, name })) {
    return false;
  }
  const text = `${type || ""} ${category || ""} ${name || ""}`;
  const nameText = String(name || "").trim();
  const categoryText = String(category || "").trim();
  if (/暗号資産|仮想通貨|ビットコイン|BTC|ETH|イーサ|モナコイン|Monacoin|MONA|Mona/i.test(text)) {
    return true;
  }
  if (/仕組(?:預金|定期)|仕組み預金/.test(text)) {
    return true;
  }
  if (/(米ドル|USD|ドル).*(普通|定期|預金|口座|積立|建|現金)/i.test(text)) {
    return true;
  }
  if (/^(金|純金|ゴールド|GOLD)$/i.test(nameText)) {
    return true;
  }
  if (/^(金|純金|ゴールド|GOLD)$/i.test(categoryText)) {
    return true;
  }
  if (/金・プラチナ|プラチナ|純金|ゴールド|gold|金(?:地金|積立|現物|ETF|投資)/i.test(text)) {
    return true;
  }
  if (/JA\s*出資|出資金/.test(text)) {
    return true;
  }
  if (/その他資産/.test(text)) {
    return true;
  }
  return false;
}

function findHeaderIndex(headers, includeRegex, excludeRegex) {
  for (let i = 0; i < headers.length; i += 1) {
    const header = String(headers[i] ?? "");
    if (includeRegex.test(header) && (!excludeRegex || !excludeRegex.test(header))) {
      return i;
    }
  }
  return null;
}

function findAssetAmountIndex(headers) {
  return findHeaderIndex(
    headers,
    /(評価額|時価|残高|金額|現在高|口座残高|保有額)/,
    /(損益|増減|評価損|取得|単価|利回り|率)/
  );
}

function isYenHeader(header) {
  return /円|JPY|円換算/.test(String(header || ""));
}

function normalizeCurrencyLabel(value) {
  const text = String(value || "").toUpperCase();
  if (/USD|米ドル|USドル|\$/.test(text)) {
    return "USD";
  }
  if (/JPY|円|日本円/.test(text)) {
    return "JPY";
  }
  return "";
}

function inferCurrencyFromRow({ currencyValue, type, category, name, rowText }) {
  const raw = normalizeCurrencyLabel(currencyValue);
  if (raw) {
    return raw;
  }
  const text = `${type || ""} ${category || ""} ${name || ""} ${rowText || ""}`;
  if (/USD|米ドル|ドル/i.test(text)) {
    return "USD";
  }
  return "JPY";
}

function formatImportAmount(value) {
  if (!Number.isFinite(value)) {
    return "";
  }
  const rounded = Math.round(value * 100) / 100;
  return String(rounded);
}

function parseRateValue(value) {
  const text = String(value ?? "").replace(/[%％]/g, "");
  return parseNumber(text);
}

function parseDateToInput(value) {
  const timestamp = parseDateValue(String(value || ""));
  if (timestamp === null) {
    return "";
  }
  return formatDate(new Date(timestamp));
}

function normalizeRowKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[　\s]+/g, "")
    .replace(/[()（）]/g, "");
}

function makeImportRowKey(row) {
  const nameKey = normalizeRowKey(row?.name);
  const currencyKey = normalizeRowKey(row?.currency || "JPY");
  return `${nameKey}|${currencyKey}`;
}

function mergeImportedRows(existingRows, importedRows, options = {}) {
  const existing = Array.isArray(existingRows) ? existingRows : [];
  const imported = Array.isArray(importedRows) ? importedRows : [];
  const merged = existing.map((row) => ({ ...row }));
  const map = new Map();
  merged.forEach((row, index) => {
    const key = makeImportRowKey(row);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(index);
  });

  const appended = [];
  imported.forEach((row) => {
    const key = makeImportRowKey(row);
    const indices = map.get(key);
    if (indices && indices.length) {
      const idx = indices.shift();
      const target = merged[idx];
      if (row?.faceValue !== undefined && row.faceValue !== "") {
        target.faceValue = row.faceValue;
      }
      if (row?.purchasePrice !== undefined && row.purchasePrice !== "") {
        target.purchasePrice = row.purchasePrice;
      }
      if (row?.rate !== undefined && row.rate !== "") {
        target.rate = row.rate;
      }
      if (!target.name && row?.name) {
        target.name = row.name;
      }
      if (!target.currency && row?.currency) {
        target.currency = row.currency;
      }
      if (
        options.includeMaturity &&
        row?.maturityDate !== undefined &&
        row.maturityDate !== ""
      ) {
        target.maturityDate = row.maturityDate;
      }
      return;
    }
    const newRow = {
      name: row?.name || "",
      cash: row?.cash ?? false,
      currency: row?.currency || "JPY",
      faceValue: row?.faceValue ?? "",
      purchasePrice: row?.purchasePrice ?? "",
      rate: row?.rate ?? "",
    };
    if (options.includeMaturity) {
      newRow.maturityDate = row?.maturityDate ?? "";
    }
    appended.push(newRow);
  });

  return [...merged, ...appended];
}

function getClassLabel(type, category) {
  const typeText = String(type || "").trim();
  if (typeText) {
    return typeText;
  }
  return String(category || "").trim();
}

function isExcludedClass(label) {
  return /投資信託|投信|ファンド|ETF|株式|保険|年金/i.test(label);
}

function classifyImportAsset({ type, category, name }) {
  const classLabel = getClassLabel(type, category);
  const nameText = String(name || "").trim();
  const combined = `${classLabel} ${nameText}`.trim();
  if (!combined || /合計/.test(combined)) {
    return null;
  }
  if (isExcludedClass(classLabel)) {
    return null;
  }
  if (shouldReclassifyToOther({ type: classLabel, category: "", name: nameText })) {
    return "other";
  }
  if (/債券|国債|社債|外債|地方債|公社債|JGB|Treasury|T-?Bill|T-?Note/i.test(classLabel)) {
    return "bond";
  }
  return null;
}

function classifyImportAssetFromRow({ type, category, name, rowText }) {
  if (!rowText || /合計/.test(rowText)) {
    return null;
  }
  const classLabel = getClassLabel(type, category);
  if (isExcludedClass(classLabel)) {
    return null;
  }
  const base = classifyImportAsset({ type, category, name });
  if (base) {
    return base;
  }
  if (/債券|国債|社債|外債|地方債|公社債|JGB|Treasury|T-?Bill|T-?Note/i.test(rowText)) {
    return "bond";
  }
  if (shouldReclassifyToOther({ name: rowText })) {
    return "other";
  }
  return null;
}

function isNumberLikeText(value) {
  const text = String(value ?? "").replace(/\s/g, "");
  if (!text) {
    return false;
  }
  if (/^[¥￥$]?\d[\d,]*(?:\.\d+)?%?$/.test(text)) {
    return true;
  }
  if (/^\(.*\)$/.test(text)) {
    return true;
  }
  return false;
}

function isDateLikeText(value) {
  const text = String(value ?? "").trim();
  if (!text) {
    return false;
  }
  return /^\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}$/.test(text);
}

function resolveNameFromRow(row, headers, indexes) {
  const nameIndex = indexes.nameIndex;
  if (nameIndex !== null && row[nameIndex]) {
    return row[nameIndex];
  }
  const typeIndex = indexes.typeIndex;
  const categoryIndex = indexes.categoryIndex;
  if (categoryIndex !== null && row[categoryIndex]) {
    return row[categoryIndex];
  }
  if (typeIndex !== null && row[typeIndex]) {
    return row[typeIndex];
  }
  let best = "";
  row.forEach((cell, idx) => {
    const value = String(cell ?? "").trim();
    if (!value) {
      return;
    }
    const header = String(headers[idx] ?? "");
    if (/(金額|評価|残高|額面|元本|通貨|利率|利回り|償還|満期|日付|数量|口数|損益|増減|単価|取得|為替|円|JPY|USD)/.test(header)) {
      return;
    }
    if (isNumberLikeText(value) || isDateLikeText(value)) {
      return;
    }
    if (value.length > best.length) {
      best = value;
    }
  });
  return best;
}

function isImportCashReclass(text) {
  return /預金|現金|暗号資産|仮想通貨|仕組預金|(米ドル|USD|外貨).*(普通|定期|現金)/i.test(
    text
  );
}

function parseDecimalAmount(value) {
  const text = String(value ?? "").replace(/\s/g, "");
  let isNegative = false;
  if (/^\(.*\)$/.test(text) || /[▲△]/.test(text)) {
    isNegative = true;
  }
  const match = text.match(/\d[\d,]*(?:\.\d+)?/);
  if (!match) {
    return null;
  }
  const normalized = match[0].replace(/,/g, "");
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  const hasMinus = /-/.test(text);
  return isNegative || hasMinus ? -Math.abs(parsed) : parsed;
}

function getRowAmount(row, index) {
  if (index === null || index === undefined) {
    return null;
  }
  if (!row || row.length <= index) {
    return null;
  }
  return parseDecimalAmount(row[index]);
}

function extractImportedAssetRows(text) {
  const table = parseAssetTable(text);
  if (!table) {
    return { bondRows: [], otherAssetRows: [], otherAssetReclassTotalYen: 0 };
  }

  const headers = table.headers.map((header) => String(header ?? ""));
  const typeIndex = findHeaderIndex(
    headers,
    /資産区分|資産種別|資産タイプ|カテゴリ|カテゴリー|科目|種別|分類/
  );
  const nameIndex = findHeaderIndex(
    headers,
    /名称|銘柄|商品|口座|資産名|ファンド/
  );
  const categoryIndex = findHeaderIndex(
    headers,
    /大分類|中分類|小分類|分類|種別|カテゴリ|カテゴリー|科目/
  );
  const currencyIndex = findHeaderIndex(
    headers,
    /通貨|通貨コード|通貨名|通貨単位|通貨区分/
  );
  const maturityIndex = findHeaderIndex(headers, /償還日|満期|期限|償還期日|満期日/);
  const rateIndex = findHeaderIndex(headers, /利率|利回り|クーポン|金利|利息/);
  const faceValueIndex = findHeaderIndex(headers, /額面|保有額面|元本|償還額|額面金額/);
  const yenAmountIndex = findHeaderIndex(
    headers,
    /(評価額|時価|残高|金額|現在高|口座残高|保有額).*(円|JPY|円換算)/
  );
  const amountIndex = findHeaderIndex(
    headers,
    /(評価額|時価|残高|金額|現在高|口座残高|保有額|数量|口数)/,
    /(損益|増減|評価損|取得|単価|利回り|率)/
  );

  const usdRate =
    parseNumber(bondUsdRateInput?.value ?? readBondStorage().usdRate) ?? 0;
  const bondRows = [];
  const otherAssetRows = [];
  let otherAssetReclassTotalYen = 0;

  table.dataRows.forEach((row) => {
    const type = typeIndex === null ? "" : row[typeIndex] || "";
    const name = nameIndex === null ? "" : row[nameIndex] || "";
    const category = categoryIndex === null ? "" : row[categoryIndex] || "";
    const rowText = row.map((cell) => String(cell ?? "")).join(" ").trim();
    const textValue = `${type} ${category} ${name}`.trim();
    const target = classifyImportAssetFromRow({
      type,
      category,
      name,
      rowText,
    });
    if (!target) {
      return;
    }
    const resolvedName =
      name ||
      resolveNameFromRow(row, headers, { typeIndex, nameIndex, categoryIndex }) ||
      category ||
      type ||
      textValue ||
      rowText;

    let currency = inferCurrencyFromRow({
      currencyValue: currencyIndex === null ? "" : row[currencyIndex],
      type,
      category,
      name,
      rowText,
    });

    const rawAmount = getRowAmount(row, amountIndex);
    const yenAmount = getRowAmount(row, yenAmountIndex);
    const faceRaw = getRowAmount(row, faceValueIndex);
    const amountHeader = amountIndex === null ? "" : headers[amountIndex];
    const faceHeader = faceValueIndex === null ? "" : headers[faceValueIndex];
    const amountIsYen = isYenHeader(amountHeader);
    const faceIsYen = isYenHeader(faceHeader);

    let valuation = null;
    if (currency === "USD") {
      if (rawAmount !== null && !amountIsYen) {
        valuation = rawAmount;
      } else if (yenAmount !== null && usdRate > 0) {
        valuation = yenAmount / usdRate;
      } else if (rawAmount !== null && amountIsYen && usdRate > 0) {
        valuation = rawAmount / usdRate;
      } else {
        currency = "JPY";
        valuation = yenAmount ?? rawAmount;
      }
    } else {
      valuation = yenAmount ?? rawAmount;
    }

    let faceValue = null;
    if (faceRaw !== null) {
      if (currency === "USD" && faceIsYen && usdRate > 0) {
        faceValue = faceRaw / usdRate;
      } else {
        faceValue = faceRaw;
      }
    }
    if (!Number.isFinite(faceValue) && Number.isFinite(valuation)) {
      faceValue = valuation;
    }
    if (!Number.isFinite(valuation) && Number.isFinite(faceValue)) {
      valuation = faceValue;
    }
    if ((!Number.isFinite(valuation) || valuation <= 0) &&
      (!Number.isFinite(faceValue) || faceValue <= 0)) {
      return;
    }
    if ((!Number.isFinite(valuation) || valuation <= 0) &&
      Number.isFinite(faceValue) && faceValue > 0) {
      valuation = faceValue;
    }

    const cash = isImportCashReclass(textValue);
    const valuationValue = Number.isFinite(valuation)
      ? valuation
      : Number.isFinite(faceValue)
        ? faceValue
        : null;
    const faceValueValue = Number.isFinite(faceValue) ? faceValue : valuationValue;
    const faceValueText = formatImportAmount(faceValueValue);
    const purchasePrice = formatImportAmount(valuationValue);
    const rateValue = rateIndex === null ? null : parseRateValue(row[rateIndex]);
    const rateText =
      Number.isFinite(rateValue) ? formatImportAmount(rateValue) : "";
    const maturityText =
      maturityIndex === null ? "" : parseDateToInput(row[maturityIndex] || "");

    const rowYenValue = Number.isFinite(valuation)
      ? toYenAmount(valuation * (currency === "USD" ? usdRate : 1))
      : 0;
    const rowData = {
      name: resolvedName,
      cash,
      currency,
      faceValue: faceValueText,
      purchasePrice,
      rate: rateText,
    };

    if (target === "bond") {
      rowData.maturityDate = maturityText;
      bondRows.push(rowData);
    } else {
      otherAssetRows.push(rowData);
      if (shouldReclassifyToOther({ type, category, name })) {
        otherAssetReclassTotalYen += rowYenValue;
      }
    }
  });

  return { bondRows, otherAssetRows, otherAssetReclassTotalYen };
}

function extractUsdDepositFromAssetList(text) {
  const table = parseAssetTable(text);
  if (!table) {
    return extractUsdDepositFromLooseText(text);
  }

  const headers = table.headers.map((header) => String(header ?? ""));
  const typeIndex = findHeaderIndex(
    headers,
    /資産区分|資産種別|資産タイプ|カテゴリ|カテゴリー|科目|種別|分類/
  );
  const nameIndex = findHeaderIndex(
    headers,
    /名称|銘柄|商品|口座|資産名|ファンド/
  );
  const categoryIndex = findHeaderIndex(
    headers,
    /大分類|中分類|小分類|分類|種別|カテゴリ|カテゴリー|科目/
  );
  const currencyIndex = findHeaderIndex(
    headers,
    /通貨|通貨コード|通貨名|通貨単位|通貨区分/
  );
  const yenAmountIndex = findHeaderIndex(
    headers,
    /(評価額|時価|残高|金額|現在高|口座残高|保有額).*(円|JPY|円換算)/
  );
  const amountIndex = findAssetAmountIndex(headers);
  if (amountIndex === null) {
    return extractUsdDepositFromLooseText(text);
  }

  let matched = null;

  table.dataRows.forEach((row) => {
    const type = typeIndex === null ? "" : row[typeIndex] || "";
    const name = nameIndex === null ? "" : row[nameIndex] || "";
    const category = categoryIndex === null ? "" : row[categoryIndex] || "";
    const rowText = row.map((cell) => String(cell ?? "")).join(" ").trim();
    if (!isUsdDepositRow({ type, category, name, rowText })) {
      return;
    }

    const rawAmount = getRowAmount(row, amountIndex);
    const yenAmount = getRowAmount(row, yenAmountIndex);
    const valuationYen = yenAmount ?? rawAmount;
    if (!Number.isFinite(valuationYen)) {
      return;
    }
    const yenValue = toYenAmount(valuationYen);
    const resolvedName =
      name ||
      resolveNameFromRow(row, headers, { typeIndex, nameIndex, categoryIndex }) ||
      category ||
      type ||
      "米ドル普通";
    const textValue = `${type} ${category} ${name} ${rowText}`.trim();
    matched = {
      name: resolvedName,
      currency: "JPY",
      amount: valuationYen,
      amountYen: yenValue,
      cash: isImportCashReclass(textValue),
    };
  });

  return matched || extractUsdDepositFromLooseText(text);
}

function extractUsdDepositFromLooseText(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  let matched = null;
  lines.forEach((line) => {
    if (!/米ドル普通/.test(line)) {
      return;
    }
    const numbers = line.match(/[+-]?\d[\d,]*(?:\.\d+)?/g);
    if (!numbers || !numbers.length) {
      return;
    }
    const amount = parseDecimalAmount(numbers[numbers.length - 1]);
    if (!Number.isFinite(amount)) {
      return;
    }
    matched = {
      name: "米ドル普通",
      currency: "JPY",
      amount,
      amountYen: toYenAmount(amount),
      cash: isImportCashReclass(line),
    };
  });
  return matched;
}

function upsertUsdDepositBondRow(info) {
  if (!info || !Number.isFinite(info.amount)) {
    return false;
  }
  const stored = readBondStorage();
  const active = Array.isArray(stored.active) ? stored.active : [];
  const targetIndex = active.findIndex((row) =>
    isUsdDepositRow({ name: row?.name })
  );
  const faceValue = formatImportAmount(info.amount);
  const purchasePrice = formatImportAmount(info.amount);
  if (targetIndex >= 0) {
    const target = active[targetIndex];
    target.currency = "JPY";
    if (faceValue !== "") {
      target.faceValue = faceValue;
    }
    if (purchasePrice !== "") {
      target.purchasePrice = purchasePrice;
    }
  } else {
    active.push({
      name: info.name || "米ドル普通",
      cash: info.cash ?? true,
      currency: "JPY",
      faceValue,
      purchasePrice,
      rate: "",
      maturityDate: "",
    });
  }
  stored.active = active;
  writeBondStorage(stored);
  loadBondRows();
  return true;
}

function sumInvestmentsFromList(text) {
  const table = parseAssetTable(text);
  if (!table) {
    return sumInvestmentsFromText(text);
  }

  const typeIndex = mapHeaderIndex(
    table.headers,
    /資産区分|資産種別|資産タイプ|カテゴリ|カテゴリー|科目|種別|分類/
  );
  const nameIndex = mapHeaderIndex(
    table.headers,
    /名称|銘柄|商品|口座|資産名|ファンド/
  );
  const categoryIndex = mapHeaderIndex(
    table.headers,
    /大分類|中分類|小分類|分類|種別|カテゴリ|カテゴリー|科目/
  );
  const amountIndex = findAssetAmountIndex(table.headers);
  if (typeIndex === null || amountIndex === null) {
    return sumInvestmentsFromText(text);
  }

  const totals = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    usd: 0,
    dc: 0,
  };

  table.dataRows.forEach((row) => {
    const type = row[typeIndex] || "";
    const name = nameIndex !== null ? row[nameIndex] || "" : "";
    const category = categoryIndex !== null ? row[categoryIndex] || "" : "";
    const rowText = row.map((cell) => String(cell ?? "")).join(" ").trim();
    const amount = row[amountIndex] ? parseAmount(row[amountIndex]) : null;
    if (amount === null) {
      return;
    }

    if (isUsdDepositRow({ type, category, name, rowText })) {
      totals.usd += amount;
      return;
    }

    if (shouldReclassifyToOther({ type, category, name })) {
      return;
    }

    if (/ニッセイみらいのカタチ/.test(name)) {
      totals.dc += amount;
      return;
    }
    if (/DC|確定拠出|ベネフィット|あおぞら/.test(name)) {
      totals.dc += amount;
      return;
    }
    if (/株式/.test(type)) {
      totals.stocks += amount;
      return;
    }
    if (/投資信託/.test(type)) {
      totals.funds += amount;
      return;
    }
    if (/債券/.test(type)) {
      totals.bonds += amount;
      return;
    }
    if (/保険/.test(type)) {
      totals.insurance += amount;
      return;
    }
    if (/外貨/.test(type) || /米ドル/.test(name)) {
      totals.usd += amount;
      return;
    }
    if (/年金/.test(type)) {
      return;
    }
  });

  const hasAny =
    totals.stocks ||
    totals.funds ||
    totals.bonds ||
    totals.insurance ||
    totals.usd ||
    totals.dc;
  return hasAny ? totals : sumInvestmentsFromText(text);
}

function sumInvestmentsFromText(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  const totals = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    usd: 0,
    dc: 0,
  };

  let currentSection = "";
  let matched = false;

  lines.forEach((line) => {
    if (/株式/.test(line)) {
      currentSection = "stocks";
      return;
    }
    if (/投資信託/.test(line)) {
      currentSection = "funds";
      return;
    }
    if (/債券/.test(line)) {
      currentSection = "bonds";
      return;
    }
    if (/保険/.test(line)) {
      currentSection = "insurance";
      return;
    }
    if (/年金/.test(line)) {
      currentSection = "pension";
      return;
    }

    const amountMatch = line.match(/([+-]?\d[\d,]*)\s*円/);
    if (!amountMatch) {
      return;
    }
    const amount = parseAmount(amountMatch[1]);
    if (amount === null) {
      return;
    }

    if (isUsdDepositRow({ name: line, type: currentSection })) {
      totals.usd += amount;
      matched = true;
      return;
    }

    if (shouldReclassifyToOther({ name: line, type: currentSection })) {
      matched = true;
      return;
    }
    if (/ニッセイみらいのカタチ/.test(line)) {
      totals.dc += amount;
      matched = true;
      return;
    }
    if (/DC|確定拠出|ベネフィット|あおぞら/.test(line)) {
      totals.dc += amount;
      matched = true;
      return;
    }
    if (/米ドル|USD|ドル/.test(line)) {
      totals.usd += amount;
      matched = true;
      return;
    }

    if (currentSection === "stocks") {
      totals.stocks += amount;
      matched = true;
      return;
    }
    if (currentSection === "funds") {
      totals.funds += amount;
      matched = true;
      return;
    }
    if (currentSection === "bonds") {
      totals.bonds += amount;
      matched = true;
      return;
    }
    if (currentSection === "insurance") {
      totals.insurance += amount;
      matched = true;
      return;
    }
    if (currentSection === "pension") {
      return;
    }
  });

  return matched ? totals : null;
}

function parseAmount(value) {
  const text = String(value ?? "").replace(/\s/g, "");
  let isNegative = false;
  if (/^\(.*\)$/.test(text)) {
    isNegative = true;
  }
  if (/[▲△]/.test(text)) {
    isNegative = true;
  }
  const match = text.match(/\d[\d,]*/);
  if (!match) {
    return null;
  }
  const normalized = match[0].replace(/,/g, "");
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return null;
  }
  const hasMinus = /-/.test(text);
  return isNegative || hasMinus ? -Math.abs(parsed) : parsed;
}

function sumInputs(inputs) {
  return inputs.reduce((sum, input) => {
    const value = parseNumber(input.value);
    if (value === null || value < 0) {
      return sum;
    }
    return sum + value;
  }, 0);
}

function mapIncomeInputs(inputs) {
  return inputs.reduce((acc, input) => {
    const key = input.dataset.incomeKey;
    if (!key) {
      return acc;
    }
    const value = parseNumber(input.value);
    acc[key] = value === null || value < 0 ? 0 : value;
    return acc;
  }, {});
}

function getPersistKey(el, index) {
  if (el.id) {
    return el.id;
  }
  const incomeKey = el.dataset.incomeKey;
  if (incomeKey) {
    return `${el.className}:${incomeKey}`;
  }
  const expenseKey = el.dataset.expenseKey;
  if (expenseKey) {
    return `${el.className}:${expenseKey}`;
  }
  return `${el.className || el.tagName}:${index}`;
}

function loadPersistedInputs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return;
    }
    const data = JSON.parse(raw);
    if (!data || typeof data !== "object") {
      return;
    }
    let migrated = false;
    if (Object.prototype.hasOwnProperty.call(data, "balanceNissay")) {
      const nissay = parseNumber(data.balanceNissay);
      const current = parseNumber(data.balanceDc) || 0;
      if (nissay !== null) {
        data.balanceDc = String(Math.round(current + nissay));
      }
      delete data.balanceNissay;
      migrated = true;
    }
    if (Object.prototype.hasOwnProperty.call(data, "contribNissay")) {
      const nissay = parseNumber(data.contribNissay);
      const current = parseNumber(data.contribDc) || 0;
      if (nissay !== null) {
        data.contribDc = String(Math.round(current + nissay));
      }
      delete data.contribNissay;
      migrated = true;
    }
    if (Object.prototype.hasOwnProperty.call(data, "endAgeNissay")) {
      if (
        !Object.prototype.hasOwnProperty.call(data, "endAgeDc") ||
        data.endAgeDc === ""
      ) {
        data.endAgeDc = data.endAgeNissay;
      }
      delete data.endAgeNissay;
      migrated = true;
    }
    if (migrated) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }
    persistInputs.forEach((el, index) => {
      const key = getPersistKey(el, index);
      if (Object.prototype.hasOwnProperty.call(data, key)) {
        el.value = data[key];
        return;
      }
      const legacyKey = `${el.className || el.tagName}:${index}`;
      if (Object.prototype.hasOwnProperty.call(data, legacyKey)) {
        el.value = data[legacyKey];
      }
    });
  } catch {
    // Ignore storage failures.
  }
}

function createBondRow(data = {}) {
  if (!bondTableBody) {
    return;
  }
  const row = document.createElement("tr");

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.value = data.name ?? "";
  nameInput.dataset.key = "name";
  nameInput.classList.add("bond-name");
  nameInput.title = nameInput.value;

  const cashCheck = document.createElement("input");
  cashCheck.type = "checkbox";
  cashCheck.dataset.key = "cash";
  cashCheck.checked =
    data.cash === true || data.cash === "true" || data.cash === "1";

  const currencySelect = document.createElement("select");
  currencySelect.dataset.key = "currency";
  ["JPY", "USD"].forEach((code) => {
    const option = document.createElement("option");
    option.value = code;
    option.textContent = code;
    currencySelect.appendChild(option);
  });
  currencySelect.value = data.currency ?? "JPY";

  const faceValueInput = document.createElement("input");
  faceValueInput.type = "number";
  faceValueInput.min = "0";
  faceValueInput.step = "1000";
  faceValueInput.inputMode = "numeric";
  faceValueInput.value = data.faceValue ?? "";
  faceValueInput.dataset.key = "faceValue";

  const purchasePriceInput = document.createElement("input");
  purchasePriceInput.type = "number";
  purchasePriceInput.min = "0";
  purchasePriceInput.step = "1";
  purchasePriceInput.inputMode = "numeric";
  purchasePriceInput.value = data.purchasePrice ?? "";
  purchasePriceInput.dataset.key = "purchasePrice";

  const maturityDateInput = document.createElement("input");
  maturityDateInput.type = "date";
  maturityDateInput.value = data.maturityDate ?? "";
  maturityDateInput.dataset.key = "maturityDate";

  const rateInput = document.createElement("input");
  rateInput.type = "number";
  rateInput.min = "-100";
  rateInput.max = "100";
  rateInput.step = "0.01";
  rateInput.inputMode = "numeric";
  rateInput.value = data.rate ?? "";
  rateInput.dataset.key = "rate";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistBondRows();
  });

  const cells = [
    nameInput,
    cashCheck,
    currencySelect,
    faceValueInput,
    purchasePriceInput,
    maturityDateInput,
    rateInput,
  ].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });

  const actionCell = document.createElement("td");
  actionCell.className = "bond-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  bondTableBody.appendChild(row);

  [
    nameInput,
    cashCheck,
    currencySelect,
    faceValueInput,
    purchasePriceInput,
    maturityDateInput,
    rateInput,
  ].forEach((input) => {
    const handler = () => {
      if (input === nameInput) {
        nameInput.title = nameInput.value;
      }
      persistBondRows();
    };
    input.addEventListener("input", handler);
    if (input.type === "checkbox") {
      input.addEventListener("change", handler);
    }
    if (input.tagName === "SELECT") {
      input.addEventListener("change", persistBondRows);
    }
  });
  nameInput.addEventListener("focus", () => {
    nameInput.title = nameInput.value;
  });
}

function createMaturedBondRow(data = {}) {
  if (!bondMaturedBody) {
    return;
  }
  const row = document.createElement("tr");
  const cells = [
    data.name ?? "",
    data.currency ?? "",
    data.faceValue ?? "",
    data.purchasePrice ?? "",
    data.maturityDate ?? "",
    data.rate ?? "",
    data.maturedAt ?? "",
  ].map((value) => {
    const td = document.createElement("td");
    td.textContent = value;
    return td;
  });
  cells.forEach((cell) => row.appendChild(cell));
  bondMaturedBody.appendChild(row);
}

function serializeBondRow(row) {
  const data = {};
  row.querySelectorAll("input, select").forEach((input) => {
    if (input.type === "checkbox") {
      data[input.dataset.key] = input.checked;
      return;
    }
    data[input.dataset.key] = input.value;
  });
  return data;
}

function createOtherAssetRow(data = {}) {
  if (!otherAssetTableBody) {
    return;
  }
  const row = document.createElement("tr");

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.value = data.name ?? "";
  nameInput.dataset.key = "name";
  nameInput.classList.add("bond-name");
  nameInput.title = nameInput.value;

  const cashCheck = document.createElement("input");
  cashCheck.type = "checkbox";
  cashCheck.dataset.key = "cash";
  cashCheck.checked =
    data.cash === true || data.cash === "true" || data.cash === "1";

  const currencySelect = document.createElement("select");
  currencySelect.dataset.key = "currency";
  ["JPY", "USD"].forEach((code) => {
    const option = document.createElement("option");
    option.value = code;
    option.textContent = code;
    currencySelect.appendChild(option);
  });
  currencySelect.value = data.currency ?? "JPY";

  const faceValueInput = document.createElement("input");
  faceValueInput.type = "number";
  faceValueInput.min = "0";
  faceValueInput.step = "1000";
  faceValueInput.inputMode = "numeric";
  faceValueInput.value = data.faceValue ?? "";
  faceValueInput.dataset.key = "faceValue";

  const purchasePriceInput = document.createElement("input");
  purchasePriceInput.type = "number";
  purchasePriceInput.min = "0";
  purchasePriceInput.step = "1";
  purchasePriceInput.inputMode = "numeric";
  purchasePriceInput.value = data.purchasePrice ?? "";
  purchasePriceInput.dataset.key = "purchasePrice";

  const rateInput = document.createElement("input");
  rateInput.type = "number";
  rateInput.min = "-100";
  rateInput.max = "100";
  rateInput.step = "0.01";
  rateInput.inputMode = "numeric";
  rateInput.value = data.rate ?? "";
  rateInput.dataset.key = "rate";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistOtherAssetRows();
  });

  const cells = [
    nameInput,
    cashCheck,
    currencySelect,
    faceValueInput,
    purchasePriceInput,
    rateInput,
  ].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });

  const actionCell = document.createElement("td");
  actionCell.className = "bond-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  otherAssetTableBody.appendChild(row);

  [
    nameInput,
    cashCheck,
    currencySelect,
    faceValueInput,
    purchasePriceInput,
    rateInput,
  ].forEach((input) => {
    const handler = () => {
      if (input === nameInput) {
        nameInput.title = nameInput.value;
      }
      persistOtherAssetRows();
    };
    input.addEventListener("input", handler);
    if (input.type === "checkbox") {
      input.addEventListener("change", handler);
    }
    if (input.tagName === "SELECT") {
      input.addEventListener("change", persistOtherAssetRows);
    }
  });
  nameInput.addEventListener("focus", () => {
    nameInput.title = nameInput.value;
  });
}

function serializeOtherAssetRow(row) {
  const data = {};
  row.querySelectorAll("input, select").forEach((input) => {
    if (input.type === "checkbox") {
      data[input.dataset.key] = input.checked;
      return;
    }
    data[input.dataset.key] = input.value;
  });
  return data;
}

function persistOtherAssetRows() {
  if (!otherAssetTableBody) {
    return;
  }
  const rows = Array.from(otherAssetTableBody.querySelectorAll("tr")).map(
    (row) => serializeOtherAssetRow(row)
  );
  writeOtherAssetsStorage(rows);
  updateOtherAssetsTotalFromStorage();
  updateCurrentAssetsFromInvestmentBalances();
  render();
}

function loadOtherAssetRows() {
  if (!otherAssetTableBody) {
    return;
  }
  otherAssetTableBody.innerHTML = "";
  const rows = readOtherAssetsStorage();
  if (!rows.length) {
    createOtherAssetRow();
  } else {
    rows.forEach((row) => createOtherAssetRow(row));
  }
  updateOtherAssetsTotalFromStorage();
}

function createInsurancePlanRow(data = {}) {
  if (!insurancePlanBody) {
    return;
  }
  const row = document.createElement("tr");
  row.dataset.id = data.id || makeRowId();

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.placeholder = "例: 医療保険";
  nameInput.dataset.key = "name";
  nameInput.value = data.name ?? "";

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.min = "0";
  amountInput.step = "1000";
  amountInput.inputMode = "numeric";
  amountInput.dataset.key = "amount";
  amountInput.value = data.amount ?? "";

  const endMonthInput = document.createElement("input");
  endMonthInput.type = "month";
  endMonthInput.dataset.key = "endMonth";
  endMonthInput.value = data.endMonth ?? "";

  const memoInput = document.createElement("input");
  memoInput.type = "text";
  memoInput.placeholder = "自由メモ";
  memoInput.dataset.key = "memo";
  memoInput.value = data.memo ?? "";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistInsurancePlanRows();
  });

  const cells = [nameInput, amountInput, endMonthInput, memoInput].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });
  const actionCell = document.createElement("td");
  actionCell.className = "insurance-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  insurancePlanBody.appendChild(row);

  [nameInput, amountInput, endMonthInput, memoInput].forEach((input) => {
    input.addEventListener("input", persistInsurancePlanRows);
    input.addEventListener("change", persistInsurancePlanRows);
  });
}

function serializeInsurancePlanRow(row) {
  const data = { id: row.dataset.id || makeRowId() };
  row.querySelectorAll("input").forEach((input) => {
    data[input.dataset.key] = input.value;
  });
  return data;
}

function persistInsurancePlanRows() {
  if (!insurancePlanBody) {
    return;
  }
  const rows = Array.from(insurancePlanBody.querySelectorAll("tr")).map((row) =>
    serializeInsurancePlanRow(row)
  );
  writeInsurancePlans(rows);
  render();
  updateInsuranceDetailSummary();
}

function loadInsurancePlanRows() {
  if (!insurancePlanBody) {
    return;
  }
  insurancePlanBody.innerHTML = "";
  const rows = readInsurancePlans();
  if (!rows.length) {
    createInsurancePlanRow();
  } else {
    rows.forEach((row) => createInsurancePlanRow(row));
  }
  updateInsuranceDetailSummary();
}

function updateInsuranceDetailSummary() {
  if (!insuranceCurrentAmount) {
    return;
  }
  const total = getInsurancePlanTotalForDate(readInsurancePlans(), new Date());
  insuranceCurrentAmount.textContent = Number.isFinite(total)
    ? yenFormatter.format(total)
    : "-";
}


function createPensionPlanRow(data = {}) {
  if (!pensionPlanBody) {
    return;
  }
  const row = document.createElement("tr");
  row.dataset.id = data.id || makeRowId();

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.placeholder = "例: 企業年金";
  nameInput.dataset.key = "name";
  nameInput.value = data.name ?? "";

  const startAgeInput = document.createElement("input");
  startAgeInput.type = "number";
  startAgeInput.min = "0";
  startAgeInput.max = "120";
  startAgeInput.step = "1";
  startAgeInput.inputMode = "numeric";
  startAgeInput.dataset.key = "startAge";
  startAgeInput.value = data.startAge ?? "";

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.min = "0";
  amountInput.step = "1000";
  amountInput.inputMode = "numeric";
  amountInput.dataset.key = "amount";
  amountInput.value = data.amount ?? "";

  const payoutSelect = document.createElement("select");
  payoutSelect.dataset.key = "payoutType";
  [
    { value: "lump", label: "一括" },
    { value: "installment", label: "分割" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    payoutSelect.appendChild(option);
  });
  payoutSelect.value = data.payoutType ?? "installment";

  const installmentInput = document.createElement("input");
  installmentInput.type = "number";
  installmentInput.min = "0";
  installmentInput.step = "1000";
  installmentInput.inputMode = "numeric";
  installmentInput.dataset.key = "installmentAmount";
  installmentInput.value = data.installmentAmount ?? "";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistPensionPlanRows();
  });

  const cells = [
    nameInput,
    startAgeInput,
    amountInput,
    payoutSelect,
    installmentInput,
  ].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });
  const actionCell = document.createElement("td");
  actionCell.className = "pension-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  pensionPlanBody.appendChild(row);

  [nameInput, startAgeInput, amountInput, payoutSelect, installmentInput].forEach(
    (input) => {
      input.addEventListener("input", persistPensionPlanRows);
      input.addEventListener("change", persistPensionPlanRows);
    }
  );
}

function serializePensionPlanRow(row) {
  const data = { id: row.dataset.id || makeRowId() };
  row.querySelectorAll("input, select").forEach((input) => {
    data[input.dataset.key] = input.value;
  });
  return data;
}

function persistPensionPlanRows() {
  if (!pensionPlanBody) {
    return;
  }
  const rows = Array.from(pensionPlanBody.querySelectorAll("tr")).map((row) =>
    serializePensionPlanRow(row)
  );
  writePensionPlans(rows);
  refreshPensionChangePlanOptions();
  render();
}

function loadPensionPlanRows() {
  if (!pensionPlanBody) {
    return;
  }
  pensionPlanBody.innerHTML = "";
  const rows = readPensionPlans();
  if (!rows.length) {
    createPensionPlanRow();
  } else {
    rows.forEach((row) => createPensionPlanRow(row));
  }
  refreshPensionChangePlanOptions();
}

function createPensionChangeRow(data = {}) {
  if (!pensionChangeBody) {
    return;
  }
  const row = document.createElement("tr");
  row.dataset.id = data.id || makeRowId();

  const planSelect = document.createElement("select");
  planSelect.dataset.key = "planId";
  planSelect.value = data.planId ?? "";

  const ageInput = document.createElement("input");
  ageInput.type = "number";
  ageInput.min = "0";
  ageInput.max = "120";
  ageInput.step = "1";
  ageInput.inputMode = "numeric";
  ageInput.dataset.key = "age";
  ageInput.value = data.age ?? "";

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.min = "0";
  amountInput.step = "1000";
  amountInput.inputMode = "numeric";
  amountInput.dataset.key = "amount";
  amountInput.value = data.amount ?? "";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistPensionChangeRows();
  });

  const cells = [planSelect, ageInput, amountInput].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });
  const actionCell = document.createElement("td");
  actionCell.className = "pension-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  pensionChangeBody.appendChild(row);

  [planSelect, ageInput, amountInput].forEach((input) => {
    input.addEventListener("input", persistPensionChangeRows);
    input.addEventListener("change", persistPensionChangeRows);
  });

  refreshPensionChangePlanOptions();
}

function serializePensionChangeRow(row) {
  const data = { id: row.dataset.id || makeRowId() };
  row.querySelectorAll("input, select").forEach((input) => {
    data[input.dataset.key] = input.value;
  });
  return data;
}

function persistPensionChangeRows() {
  if (!pensionChangeBody) {
    return;
  }
  const rows = Array.from(pensionChangeBody.querySelectorAll("tr")).map((row) =>
    serializePensionChangeRow(row)
  );
  writePensionChanges(rows);
  render();
}

function loadPensionChangeRows() {
  if (!pensionChangeBody) {
    return;
  }
  pensionChangeBody.innerHTML = "";
  const rows = readPensionChanges();
  if (!rows.length) {
    createPensionChangeRow();
  } else {
    rows.forEach((row) => createPensionChangeRow(row));
  }
}

function refreshPensionChangePlanOptions() {
  if (!pensionChangeBody) {
    return;
  }
  const plans = readPensionPlans();
  const options = plans.map((plan) => ({
    id: plan.id,
    name: plan.name || "名称未入力",
  }));
  pensionChangeBody.querySelectorAll("select[data-key=\"planId\"]").forEach(
    (select) => {
      const current = select.value;
      select.innerHTML = "";
      if (!options.length) {
        const option = document.createElement("option");
        option.value = "";
        option.textContent = "年金がありません";
        select.appendChild(option);
        select.disabled = true;
      } else {
        options.forEach((item) => {
          const option = document.createElement("option");
          option.value = item.id;
          option.textContent = item.name;
          select.appendChild(option);
        });
        select.disabled = false;
        if (options.some((item) => item.id === current)) {
          select.value = current;
        }
      }
    }
  );
}

function updatePensionDetailSummary() {
  if (!pensionCurrentAmount) {
    return;
  }
  const currentAmount = parseNumber(balanceDcInput.value);
  pensionCurrentAmount.textContent = Number.isFinite(currentAmount)
    ? yenFormatter.format(currentAmount)
    : "-";
}

function computePensionContributionForDate(birthDate, plans, changes, atDate) {
  if (!plans || plans.length === 0) {
    return null;
  }
  return plans.reduce((sum, plan) => {
    const amount = parseNumber(plan.amount) || 0;
    return sum + amount;
  }, 0);
}

function syncDcContributionFromPensionDetail(birthDate, plans, changes, atDate) {
  if (!contribDcInput || !plans || plans.length === 0) {
    return false;
  }
  const amount = computePensionContributionForDate(
    birthDate,
    plans,
    changes,
    atDate
  );
  if (!Number.isFinite(amount)) {
    return false;
  }
  const nextValue = String(Math.round(amount));
  if (contribDcInput.value === nextValue) {
    return false;
  }
  contribDcInput.value = nextValue;
  return true;
}

function persistBondRows() {
  if (!bondTableBody) {
    return;
  }
  const rows = Array.from(bondTableBody.querySelectorAll("tr")).map((row) =>
    serializeBondRow(row)
  );
  const stored = readBondStorage();
  stored.active = rows;
  if (bondUsdRateInput) {
    stored.usdRate = bondUsdRateInput.value;
  }
  writeBondStorage(stored);
  updateBondAverageRate();
  updateBondBalanceFromStorage();
  updateCurrentAssetsFromInvestmentBalances();
  render();
}

function isBondMatured(data, today) {
  const maturity = parseDate(data.maturityDate);
  if (!maturity) {
    return false;
  }
  return maturity.getTime() <= today.getTime();
}

function moveMaturedBonds(data, today) {
  const active = [];
  let moved = false;
  data.active.forEach((row) => {
    if (isBondMatured(row, today)) {
      data.matured.push({
        ...row,
        maturedAt: formatDate(today),
      });
      moved = true;
      return;
    }
    active.push(row);
  });
  data.active = active;
  return moved;
}

function loadBondRows() {
  if (!bondTableBody) {
    return;
  }
  bondTableBody.innerHTML = "";
  if (bondMaturedBody) {
    bondMaturedBody.innerHTML = "";
  }
  const stored = readBondStorage();
  const today = new Date();
  const moved = moveMaturedBonds(stored, today);
  if (bondUsdRateInput) {
    bondUsdRateInput.value = stored.usdRate ?? "";
    if (!bondUsdRateListenerBound) {
      bondUsdRateInput.addEventListener("input", () => {
        const next = readBondStorage();
        next.usdRate = bondUsdRateInput.value;
        writeBondStorage(next);
        updateBondAverageRate();
        updateBondBalanceFromStorage();
        updateOtherAssetsTotalFromStorage();
        updateCurrentAssetsFromInvestmentBalances();
        render();
      });
      bondUsdRateListenerBound = true;
    }
  }
  if (!stored.active.length) {
    createBondRow();
  } else {
    stored.active.forEach((row) => createBondRow(row));
  }
  stored.matured.forEach((row) => createMaturedBondRow(row));
  if (moved) {
    writeBondStorage(stored);
  }
  updateBondAverageRate();
  updateBondBalanceFromStorage();
}

function sortBondRowsByMaturity() {
  if (!bondTableBody) {
    return;
  }
  const rows = Array.from(bondTableBody.querySelectorAll("tr")).map(
    (row, index) => {
      const data = serializeBondRow(row);
      return { data, index };
    }
  );
  if (!rows.length) {
    return;
  }
  rows.sort((a, b) => {
    const dateA = parseDate(a.data.maturityDate);
    const dateB = parseDate(b.data.maturityDate);
    if (dateA && dateB) {
      return dateA.getTime() - dateB.getTime();
    }
    if (dateA) {
      return -1;
    }
    if (dateB) {
      return 1;
    }
    return a.index - b.index;
  });
  bondTableBody.innerHTML = "";
  rows.forEach(({ data }) => createBondRow(data));
  persistBondRows();
}

function updateBondAverageRate() {
  if (!bondTableBody || !bondAverageRate) {
    return;
  }
  const rateInputs = Array.from(bondTableBody.querySelectorAll("input")).filter(
    (input) => input.dataset.key === "rate"
  );
  const values = rateInputs
    .map((input) => parseNumber(input.value))
    .filter((value) => value !== null);
  if (!values.length) {
    bondAverageRate.textContent = "平均利率: -";
    if (bondRateAverageDisplay) {
      bondRateAverageDisplay.textContent = "平均利率: -";
    }
    if (rateBondsInput) {
      rateBondsInput.value = "";
    }
    return;
  }
  const avg = values.reduce((sum, value) => sum + value, 0) / values.length;
  bondAverageRate.textContent = `平均利率: ${percentFormatter.format(avg)}%`;
  if (bondRateAverageDisplay) {
    bondRateAverageDisplay.textContent = `平均利率: ${percentFormatter.format(avg)}%`;
  }
  if (rateBondsInput) {
    rateBondsInput.value = String(Math.round(avg * 100) / 100);
  }
}

function getRowValuationYen(row, usdRate) {
  const hasPurchasePrice =
    row?.purchasePrice !== undefined &&
    row?.purchasePrice !== null &&
    String(row.purchasePrice).trim() !== "";
  const purchasePrice = parseNumber(row?.purchasePrice);
  const faceValue = parseNumber(row?.faceValue);
  let baseValue = null;
  if (hasPurchasePrice && Number.isFinite(purchasePrice) && purchasePrice >= 0) {
    baseValue = purchasePrice;
  } else if (Number.isFinite(faceValue) && faceValue >= 0) {
    baseValue = faceValue;
  }
  if (!Number.isFinite(baseValue)) {
    return 0;
  }
  const rate = row?.currency === "USD" ? usdRate : 1;
  return toYenAmount(baseValue * rate);
}

function getBondValuationTotalYen(stored, usdRate) {
  const active = stored?.active || [];
  return active.reduce(
    (sum, row) =>
      isUsdDepositRow({ name: row?.name })
        ? sum
        : sum + getRowValuationYen(row, usdRate),
    0
  );
}

function getOtherAssetsTotalYen(stored, usdRate) {
  const rows = Array.isArray(stored) ? stored : [];
  return rows.reduce(
    (sum, row) => sum + getRowValuationYen(row, usdRate),
    0
  );
}

function getOtherAssetsCashReclassTotalYen(stored, usdRate) {
  const rows = Array.isArray(stored) ? stored : [];
  return rows.reduce((sum, row) => {
    const isCash =
      row?.cash === true || row?.cash === "true" || row?.cash === "1";
    if (!isCash) {
      return sum;
    }
    return sum + getRowValuationYen(row, usdRate);
  }, 0);
}

function getOtherAssetsTotalFromStorage() {
  const stored = readOtherAssetsStorage();
  const usdRate =
    parseNumber(bondUsdRateInput?.value ?? readBondStorage().usdRate) ?? 0;
  return getOtherAssetsTotalYen(stored, usdRate);
}

function updateOtherAssetsTotalFromStorage() {
  const total = getOtherAssetsTotalFromStorage();
  if (otherAssetTotalAmount) {
    otherAssetTotalAmount.textContent = `合計: ${yenFormatter.format(
      Math.round(total)
    )}`;
  }
  if (bondCashReclassTotal) {
    const cashReclassTotal = getCashReclassTotalYen();
    bondCashReclassTotal.textContent = cashReclassTotal > 0
      ? `現金から引く合計: ${yenFormatter.format(Math.round(cashReclassTotal))}`
      : "現金から引く合計: -";
  }
  updateBondBalanceFromStorage();
  return total;
}

function getCashReclassTotalYen() {
  const storedBonds = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? storedBonds.usdRate) ?? 0;
  const bondReclass = getBondCashReclassTotalYen(storedBonds, usdRate);
  const otherReclass = getOtherAssetsCashReclassTotalYen(
    readOtherAssetsStorage(),
    usdRate
  );
  return bondReclass + otherReclass;
}

function getBondCashReclassTotalYen(stored, usdRate) {
  const active = stored?.active || [];
  return active.reduce((sum, row) => {
    if (isUsdDepositRow({ name: row?.name })) {
      return sum;
    }
    const isCash =
      row?.cash === true || row?.cash === "true" || row?.cash === "1";
    if (!isCash) {
      return sum;
    }
    return sum + getRowValuationYen(row, usdRate);
  }, 0);
}

function updateBondBalanceFromStorage() {
  if (!balanceBondsInput && !bondTotalAmount) {
    return;
  }
  const stored = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? stored.usdRate) ?? 0;
  const bondTotal = getBondValuationTotalYen(stored, usdRate);
  const otherTotal = getOtherAssetsTotalYen(readOtherAssetsStorage(), usdRate);
  const baseTotal = bondTotal + otherTotal;
  const adjustment = readAdjustmentValue(adjustBondsInput);
  const combinedTotal = baseTotal + adjustment;
  const nextValue = String(Math.round(combinedTotal));
  if (balanceBondsInput && balanceBondsInput.value !== nextValue) {
    balanceBondsInput.value = nextValue;
  }
  writePrevAdjustmentValue(adjustBondsInput, adjustment);
  if (bondTotalAmount) {
    bondTotalAmount.textContent = `合計: ${yenFormatter.format(
      Math.round(combinedTotal)
    )}`;
  }
  if (bondCashReclassTotal) {
    const cashReclassTotal = getCashReclassTotalYen();
    bondCashReclassTotal.textContent = cashReclassTotal > 0
      ? `現金から引く合計: ${yenFormatter.format(Math.round(cashReclassTotal))}`
      : "現金から引く合計: -";
  }
  updateBondCombinedTotal();
}

function updateBondCombinedTotal() {
  if (!bondCombinedTotal) {
    return;
  }
  const storedBonds = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? storedBonds.usdRate) ?? 0;
  const bondTotal = getBondValuationTotalYen(storedBonds, usdRate);
  const otherTotal = getOtherAssetsTotalYen(readOtherAssetsStorage(), usdRate);
  const adjustment = readAdjustmentValue(adjustBondsInput);
  const combinedTotal = bondTotal + otherTotal + adjustment;
  bondCombinedTotal.textContent = `債券・その他資産 合計: ${yenFormatter.format(
    Math.round(combinedTotal)
  )}`;
}

function persistInputsToStorage() {
  try {
    const data = {};
    persistInputs.forEach((el, index) => {
      const key = getPersistKey(el, index);
      data[key] = el.value;
    });
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch {
    // Ignore storage failures.
  }
}

function safeParseJson(raw, fallback) {
  if (!raw) {
    return fallback;
  }
  try {
    const parsed = JSON.parse(raw);
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function buildStorageSnapshot() {
  try {
    const snapshot = {};
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith(STORAGE_PREFIX)) {
        continue;
      }
      const value = localStorage.getItem(key);
      if (value !== null) {
        snapshot[key] = value;
      }
    }
    return snapshot;
  } catch {
    return null;
  }
}

function restoreStorageSnapshot(snapshot) {
  if (!snapshot || typeof snapshot !== "object") {
    return false;
  }
  try {
    Object.entries(snapshot).forEach(([key, value]) => {
      if (!key || !key.startsWith(STORAGE_PREFIX)) {
        return;
      }
      if (typeof value === "string") {
        localStorage.setItem(key, value);
      }
    });
    return true;
  } catch {
    return false;
  }
}

function getInsurancePlanSnapshot() {
  if (insurancePlanBody) {
    return Array.from(insurancePlanBody.querySelectorAll("tr")).map((row) =>
      serializeInsurancePlanRow(row)
    );
  }
  return readInsurancePlans();
}

function getPensionPlanSnapshot() {
  if (pensionPlanBody) {
    return Array.from(pensionPlanBody.querySelectorAll("tr")).map((row) =>
      serializePensionPlanRow(row)
    );
  }
  return readPensionPlans();
}

function getPensionChangeSnapshot() {
  if (pensionChangeBody) {
    return Array.from(pensionChangeBody.querySelectorAll("tr")).map((row) =>
      serializePensionChangeRow(row)
    );
  }
  return readPensionChanges();
}

function buildSyncPayload() {
  persistInputsToStorage();
  persistBondRows();
  persistOtherAssetRows();
  persistInsurancePlanRows();
  persistPensionPlanRows();
  persistPensionChangeRows();
  const inputs = safeParseJson(localStorage.getItem(STORAGE_KEY), {});
  const bonds = readBondStorage();
  const otherAssets = readOtherAssetsStorage();
  const insurancePlans = getInsurancePlanSnapshot();
  const pensionPlans = getPensionPlanSnapshot();
  const pensionChanges = getPensionChangeSnapshot();
  const storage = buildStorageSnapshot();
  return {
    version: 3,
    exportedAt: new Date().toISOString(),
    storage,
    inputs,
    bonds,
    otherAssets,
    insurancePlans,
    pensionPlans,
    pensionChanges,
  };
}

async function handleExportSyncFolder() {
  if (!window.showDirectoryPicker) {
    window.alert("このブラウザはフォルダー選択に対応していません。");
    return;
  }
  try {
    const dirHandle = await window.showDirectoryPicker();
    const fileHandle = await dirHandle.getFileHandle("LifeWealth100_sync.json", {
      create: true,
    });
    const writable = await fileHandle.createWritable();
    const payload = buildSyncPayload();
    const json = JSON.stringify(payload, null, 2);
    await writable.write(json);
    await writable.close();
    if (syncStatus) {
      syncStatus.textContent = "フォルダーに保存しました";
    }
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    window.alert("フォルダーへの保存に失敗しました。");
  }
}

function importSyncPayload(raw) {
  const trimmed = raw.trim();
  if (!trimmed) {
    window.alert("同期データが空です。");
    return false;
  }
  let data;
  try {
    data = JSON.parse(trimmed);
  } catch {
    window.alert("同期データのJSONが正しくありません。");
    return false;
  }
  if (!isPlainObject(data)) {
    window.alert("同期データの形式が正しくありません。");
    return false;
  }
  const storage = isPlainObject(data.storage) ? data.storage : null;
  const inputs = isPlainObject(data.inputs) ? data.inputs : null;
  const bonds = data.bonds ? normalizeBondStorage(data.bonds) : null;
  const otherAssets = Array.isArray(data.otherAssets) ? data.otherAssets : null;
  const insurancePlans = Array.isArray(data.insurancePlans)
    ? data.insurancePlans
    : null;
  const insuranceScheduleLegacy = Array.isArray(data.insuranceSchedule)
    ? data.insuranceSchedule
    : null;
  const pensionPlans = Array.isArray(data.pensionPlans)
    ? data.pensionPlans
    : null;
  const pensionChanges = Array.isArray(data.pensionChanges)
    ? data.pensionChanges
    : null;
  if (
    !storage &&
    !inputs &&
    !bonds &&
    !otherAssets &&
    !insurancePlans &&
    !insuranceScheduleLegacy &&
    !pensionPlans &&
    !pensionChanges
  ) {
    window.alert("同期データに読み込める内容がありません。");
    return false;
  }
  try {
    if (storage) {
      restoreStorageSnapshot(storage);
    }
    if (inputs) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(inputs));
    }
    if (bonds) {
      writeBondStorage(bonds);
    }
    if (otherAssets) {
      writeOtherAssetsStorage(otherAssets);
    }
    if (insurancePlans) {
      writeInsurancePlans(insurancePlans);
    } else if (insuranceScheduleLegacy) {
      localStorage.setItem(
        INSURANCE_SCHEDULE_LEGACY_KEY,
        JSON.stringify(insuranceScheduleLegacy)
      );
    }
    if (pensionPlans) {
      writePensionPlans(pensionPlans);
    }
    if (pensionChanges) {
      writePensionChanges(pensionChanges);
    }
  } catch {
    window.alert("同期データの保存に失敗しました。");
    return false;
  }
  loadPersistedInputs();
  cashInputManual = readCashManualFlag();
  if (balanceCashInput && balanceCashInput.value === "") {
    cashInputManual = false;
    writeCashManualFlag(false);
  }
  initializeAdjustments({ applyToBalance: false, includeBonds: true });
  writeAdjustmentsAppliedFlag(true);
  loadBondRows();
  loadOtherAssetRows();
  loadInsurancePlanRows();
  loadPensionPlanRows();
  loadPensionChangeRows();
  render();
  if (syncStatus) {
    syncStatus.textContent = "インポート済み";
  }
  return true;
}

function extractSectionTotals(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  const totals = [];

  lines.forEach((line) => {
    const match = line.match(/合計[:：]?\s*([+-]?\d[\d,]*)\s*円?/);
    if (!match) {
      return;
    }
    const amount = parseAmount(match[1]);
    if (amount === null) {
      return;
    }
    totals.push(amount);
  });

  return totals;
}

function sumAssetsFromList(text) {
  const sectionTotals = extractSectionTotals(text);
  if (sectionTotals.length > 0) {
    return sectionTotals.reduce((sum, value) => sum + value, 0);
  }

  const table = parseAssetTable(text) || parseTable(text);
  if (!table) {
    return null;
  }

  const amountIndex = findAssetAmountIndex(table.headers);
  if (amountIndex === null) {
    return null;
  }

  let total = 0;
  let hasValue = false;

  table.dataRows.forEach((row) => {
    if (row.length <= amountIndex) {
      return;
    }
    const amount = parseAmount(row[amountIndex]);
    if (amount === null) {
      return;
    }
    total += amount;
    hasValue = true;
  });

  return hasValue ? total : null;
}

function parseDateValue(value) {
  const match = value.match(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);
  if (!match) {
    return null;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }
  return new Date(year, month - 1, day).getTime();
}

function latestTotalFromSummaryFallback(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  let best = null;

  lines.forEach((line, index) => {
    const dateValue = parseDateValue(line);
    if (dateValue === null) {
      return;
    }
    const afterDate = line.replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/, "").trim();
    const amount = parseAmount(afterDate);
    if (amount === null) {
      return;
    }
    const rank = dateValue ?? index;
    if (!best || rank > best.rank) {
      best = { total: amount, rank };
    }
  });

  return best ? best.total : null;
}

function latestTotalFromSummary(text) {
  const table = parseTable(text);
  if (!table) {
    return latestTotalFromSummaryFallback(text);
  }

  const dateIndex = table.headers.findIndex((header) => /日付/.test(header));
  const totalIndex = findSummaryTotalIndex(table.headers);
  if (totalIndex === null) {
    return null;
  }

  let best = null;

  table.dataRows.forEach((row, index) => {
    const values = getSummaryRowValues(table.headers, row);
    const total = Number.isFinite(values.total) ? values.total : null;
    const dateValue = dateIndex === -1 ? null : parseDateValue(row[dateIndex]);
    const rank = dateValue ?? index;

    if (!best || rank > best.rank) {
      best = { total, rank };
    }
  });

  return best ? best.total : null;
}

function sumInvestmentsFromSummary(text) {
  const table = parseTable(text);
  if (!table) {
    return null;
  }

  const totalIndex = mapHeaderIndex(table.headers, /合計/);
  if (totalIndex === null) {
    return null;
  }

  let bestRow = null;
  let bestRank = null;
  const dateIndex = mapHeaderIndex(table.headers, /日付/);

  table.dataRows.forEach((row, index) => {
    if (!row[totalIndex]) {
      return;
    }
    const dateValue =
      dateIndex === null ? null : parseDateValue(row[dateIndex] || "");
    const rank = dateValue ?? index;
    if (bestRank === null || rank > bestRank) {
      bestRank = rank;
      bestRow = row;
    }
  });

  if (!bestRow) {
    return null;
  }

  const getAmount = (regex) => {
    const idx = mapHeaderIndex(table.headers, regex);
    if (idx === null) {
      return 0;
    }
    const value = bestRow[idx] || "";
    return parseAmount(value) || 0;
  };

  return {
    stocks: getAmount(/株式/),
    funds: getAmount(/投資信託/),
    bonds: getAmount(/債券/),
    insurance: getAmount(/保険/),
    usd: getAmount(/ドル|外貨/),
    dc:
      getAmount(/確定拠出|年金/) + getAmount(/ニッセイみらいのカタチ/),
  };
}

function applyImportedData() {
  const assetText = assetDataInput.value;
  const summaryText = summaryDataInput.value;
  const assetTotal = sumAssetsFromList(assetText);
  const summaryTotal = latestTotalFromSummary(summaryText);
  const listInvestmentTotals = sumInvestmentsFromList(assetText);
  const summaryInvestmentTotals = sumInvestmentsFromSummary(summaryText);
  const importedAssets = extractImportedAssetRows(assetText);
  const usdDeposit = extractUsdDepositFromAssetList(assetText);
  let hasBondImport = importedAssets.bondRows.length > 0;
  const hasOtherImport = importedAssets.otherAssetRows.length > 0;
  const storedBondsSnapshot = readBondStorage();
  const hasExistingBondRows =
    Array.isArray(storedBondsSnapshot.active) &&
    storedBondsSnapshot.active.length > 0;
  let investmentTotals = null;
  if (hasBondImport || hasOtherImport) {
    if (listInvestmentTotals) {
      investmentTotals = listInvestmentTotals;
    } else if (summaryInvestmentTotals) {
      investmentTotals = summaryInvestmentTotals;
    }
  } else if (summaryInvestmentTotals) {
    investmentTotals = summaryInvestmentTotals;
  } else if (listInvestmentTotals) {
    investmentTotals = listInvestmentTotals;
  }
  if (
    !hasBondImport &&
    !hasExistingBondRows &&
    investmentTotals &&
    investmentTotals.bonds > 0
  ) {
    importedAssets.bondRows.push({
      name: "債券合計",
      cash: false,
      currency: "JPY",
      faceValue: formatImportAmount(investmentTotals.bonds),
      purchasePrice: formatImportAmount(investmentTotals.bonds),
      rate: "",
      maturityDate: "",
    });
    hasBondImport = true;
  }

  if (hasBondImport) {
    const stored = readBondStorage();
    stored.active = mergeImportedRows(
      stored.active || [],
      importedAssets.bondRows,
      { includeMaturity: true }
    );
    writeBondStorage(stored);
    loadBondRows();
  }
  if (hasOtherImport) {
    const existingOther = readOtherAssetsStorage();
    const mergedOther = mergeImportedRows(
      existingOther || [],
      importedAssets.otherAssetRows
    );
    writeOtherAssetsStorage(mergedOther);
    loadOtherAssetRows();
  }
  if (usdDeposit) {
    upsertUsdDepositBondRow(usdDeposit);
  }

  if (summaryTotal !== null) {
    currentAssetsInput.value = Math.round(summaryTotal);
    importStatus.textContent = `資産推移の最新合計を反映: ${yenFormatter.format(
      summaryTotal
    )}`;
  } else if (assetTotal !== null) {
    currentAssetsInput.value = Math.round(assetTotal);
    importStatus.textContent = `資産一覧の合計を反映: ${yenFormatter.format(assetTotal)}`;
  } else {
    importStatus.textContent = "取り込みに失敗しました（列名を確認してください）。";
  }

  if (investmentTotals) {
    balanceStocksInput.value = Math.round(investmentTotals.stocks);
    balanceFundsInput.value = Math.round(investmentTotals.funds);
    if (!hasBondImport) {
      balanceBondsInput.value = Math.round(investmentTotals.bonds);
    }
    balanceInsuranceInput.value = Math.round(investmentTotals.insurance);
    if (!hasOtherImport) {
      balanceUsdInput.value = Math.round(investmentTotals.usd);
    }
    if (usdDeposit && Number.isFinite(usdDeposit.amountYen)) {
      balanceUsdInput.value = Math.round(usdDeposit.amountYen);
    }
    balanceDcInput.value = Math.round(investmentTotals.dc);
    initializeAdjustments({ applyToBalance: true, includeBonds: false });
    persistInputsToStorage();
  } else if (usdDeposit && Number.isFinite(usdDeposit.amountYen)) {
    balanceUsdInput.value = Math.round(usdDeposit.amountYen);
    initializeAdjustments({ applyToBalance: true, includeBonds: false });
    persistInputsToStorage();
  }

  importDirty = false;
  render();
}

function markImportDirty() {
  importDirty = true;
  importStatus.textContent = "貼り付けデータは未反映です。";
}

function handleExportCsv() {
  const context = getSimulationContext();
  if (!context) {
    window.alert("入力値を確認してください。");
    return;
  }

  const annualStartDate = getMonthStart(context.today);
  const annualMonthsRemaining = fullMonthsBetween(
    annualStartDate,
    addYears(context.birthDate, 100)
  );
  const rows = simulateAnnualSeries({
    startDate: annualStartDate,
    monthsRemaining: annualMonthsRemaining,
    annualRate: context.annualRate,
    categoryRates: context.categoryRates,
    retirementAge: context.retirementMonthIndex,
    retirementIncomeEndAge: context.retirementIncomeEndMonthIndex,
    monthlyNetCash: context.incomeTotal - context.expenseTotal,
    retirementMonthlyNetCash:
      context.retireBaseIncome + context.ongoingIncome - context.retireExpenseTotal,
    postRetirementMonthlyNetCash:
      context.pensionIncomeTotal +
      context.ongoingIncome -
      context.retireExpenseTotal,
    baseDividendIncome: context.baseDividendIncome,
    dividendYieldRate: context.dividendYieldRate,
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    pensionPlanState: context.pensionPlanState,
  });

  if (!rows.length) {
    window.alert("出力できるデータがありません。");
    return;
  }

  const todaySnapshot = {
    date: context.today,
    total: sumCategoryTotal(context.categories),
    ...context.categories,
  };
  downloadCsv(rows, context.birthDate, { todayRow: todaySnapshot });
}

function handleOpenCsv() {
  const context = getSimulationContext();
  if (!context) {
    window.alert("入力値を確認してください。");
    return;
  }

  const annualStartDate = getMonthStart(context.today);
  const annualMonthsRemaining = fullMonthsBetween(
    annualStartDate,
    addYears(context.birthDate, 100)
  );
  const rows = simulateAnnualSeries({
    startDate: annualStartDate,
    monthsRemaining: annualMonthsRemaining,
    annualRate: context.annualRate,
    categoryRates: context.categoryRates,
    retirementAge: context.retirementMonthIndex,
    retirementIncomeEndAge: context.retirementIncomeEndMonthIndex,
    monthlyNetCash: context.incomeTotal - context.expenseTotal,
    retirementMonthlyNetCash:
      context.retireBaseIncome + context.ongoingIncome - context.retireExpenseTotal,
    postRetirementMonthlyNetCash:
      context.pensionIncomeTotal +
      context.ongoingIncome -
      context.retireExpenseTotal,
    baseDividendIncome: context.baseDividendIncome,
    dividendYieldRate: context.dividendYieldRate,
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    pensionPlanState: context.pensionPlanState,
  });

  if (!rows.length) {
    window.alert("出力できるデータがありません。");
    return;
  }

  const header = [
    "日付",
    "年齢",
    "合計（円）",
    "預金・現金・暗号資産（円）",
    "株式(現物)（円）",
    "投資信託（円）",
    "債券（円）",
    "保険（円）",
    "年金（円）",
  ];
  const displayRows = [];
  const pushRow = (row) => {
    displayRows.push([
      formatDate(row.date),
      formatAgeYears(context.birthDate, row.date),
      toCsvNumber(row.total),
      toCsvNumber(row.cash),
      toCsvNumber(row.stocks),
      toCsvNumber(row.funds),
      toCsvNumber(row.bonds),
      toCsvNumber(row.insurance),
      toCsvNumber(row.dc || 0),
    ]);
  };

  const todaySnapshot = {
    date: context.today,
    total: sumCategoryTotal(context.categories),
    ...context.categories,
  };
  pushRow(todaySnapshot);
  rows.forEach((row) => pushRow(row));
  openSpreadsheetView({
    title: "年次CSV",
    headers: header,
    rows: displayRows,
  });
}

function getSimulationContext() {
  const birthDate = parseDate(birthDateInput.value);
  const currentAssets = parseNumber(currentAssetsInput.value);
  const retirementAgeYears = parseNumber(retirementAgeInput.value);
  const retirementIncomeEndAgeYears =
    parseNumber(retirementIncomeEndAgeInput.value) ?? 100;

  if (
    birthDate === null ||
    currentAssets === null ||
    retirementAgeYears === null ||
    retirementIncomeEndAgeYears === null
  ) {
    return null;
  }

  const today = new Date();
  const hundredthBirthday = addYears(birthDate, 100);
  const monthsRemaining = fullMonthsBetween(today, hundredthBirthday);
  if (monthsRemaining <= 0) {
    return null;
  }

  const annualRate = 0;
  const rateStocks = parseNumber(rateStocksInput?.value);
  const rateFunds = parseNumber(rateFundsInput?.value);
  const rateBonds = parseNumber(rateBondsInput?.value);
  const rateInsurance = parseNumber(rateInsuranceInput?.value);
  const toMonthlyRate = (rate) => Math.pow(1 + rate, 1 / 12) - 1;
  const defaultMonthlyRate = toMonthlyRate(annualRate);
  const categoryRates = {
    stocks:
      rateStocks === null ? defaultMonthlyRate : toMonthlyRate(rateStocks / 100),
    funds:
      rateFunds === null ? defaultMonthlyRate : toMonthlyRate(rateFunds / 100),
    bonds:
      rateBonds === null ? defaultMonthlyRate : toMonthlyRate(rateBonds / 100),
    insurance:
      rateInsurance === null
        ? defaultMonthlyRate
        : toMonthlyRate(rateInsurance / 100),
    dc: 0,
    usd: 0,
    other: 0,
  };
  const storedBonds = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? storedBonds.usdRate) ?? 0;
  const bondMaturities = (storedBonds.active || []).map((row) => ({
    maturityDate: parseDate(row.maturityDate),
    currency: row.currency || "JPY",
    faceValue: parseNumber(row.faceValue) || 0,
    bookValue:
      parseNumber(row.purchasePrice) ?? parseNumber(row.faceValue) ?? 0,
  }));
  const insurancePlans = readInsurancePlans();
  const pensionPlans = readPensionPlans();
  const activePensionPlans = getActivePensionPlans(pensionPlans);
  const hasPensionPlans = activePensionPlans.length > 0;
  const contributionSchedule = buildContributionSchedule(birthDate, {
    skipDc: hasPensionPlans,
    insurancePlans,
  });
  const expenseTotal = sumInputs(expenseInputs);
  const incomeTotal = sumInputs(incomeInputs);
  const retireExpenseTotal = sumInputs(retireExpenseInputs);
  const incomeMap = mapIncomeInputs(incomeInputs);
  const retireIncomeMap = mapIncomeInputs(retireIncomeInputs);
  const retireBaseIncome =
    (retireIncomeMap.salary || 0) +
    (retireIncomeMap.bonus || 0);
  const dividendYieldPercent = parseNumber(dividendYieldInput?.value);
  const dividendYieldRate =
    dividendYieldPercent === null ? 0 : dividendYieldPercent / 100;
  const baseDividendIncome = incomeMap.dividend || 0;
  const pensionIncomeTotal = sumInputs(pensionIncomeInputs);
  const ongoingIncome = (incomeMap.dividend || 0) + (incomeMap.other || 0);
  const summaryBreakdown = importDirty
    ? null
    : getSummaryBreakdown(summaryDataInput.value);
  const initial = buildInitialCategories(summaryBreakdown, currentAssets);
  const categories = {
    cash: initial.cash,
    stocks: initial.stocks,
    funds: initial.funds,
    bonds: initial.bonds,
    insurance: initial.insurance,
    dc: initial.dc,
    usd: initial.usd,
    other: initial.other,
  };
  const pensionPlanState = buildPensionPlanState(
    birthDate,
    activePensionPlans,
    [],
    categories.dc
  );

  return {
    birthDate,
    currentAssets,
    annualRate,
    categoryRates,
    today,
    monthsRemaining,
    retirementAgeYears,
    retirementIncomeEndAgeYears,
    retirementMonthIndex: monthIndex(addYears(birthDate, retirementAgeYears)),
    retirementIncomeEndMonthIndex: monthIndex(
      addYears(birthDate, retirementIncomeEndAgeYears)
    ),
    expenseTotal,
    incomeTotal,
    retireExpenseTotal,
    retireBaseIncome,
    pensionIncomeTotal,
    ongoingIncome,
    baseDividendIncome,
    dividendYieldRate,
    contributionSchedule,
    categories,
    bondMaturities,
    usdRate,
    pensionPlanState,
  };
}

function buildStatementRows({ showAlert }) {
  const context = getSimulationContext();
  if (!context) {
    if (showAlert) {
      window.alert("入力値を確認してください。");
    }
    return null;
  }

  const rows = simulateAnnualStatements({
    startDate: context.today,
    monthsRemaining: context.monthsRemaining,
    annualRate: context.annualRate,
    categoryRates: context.categoryRates,
    retirementAge: context.retirementMonthIndex,
    retirementIncomeEndAge: context.retirementIncomeEndMonthIndex,
    workIncome: context.incomeTotal,
    workExpense: context.expenseTotal,
    retireIncome: context.retireBaseIncome + context.ongoingIncome,
    retireExpense: context.retireExpenseTotal,
    pensionIncome: context.pensionIncomeTotal + context.ongoingIncome,
    pensionExpense: context.retireExpenseTotal,
    baseDividendIncome: context.baseDividendIncome,
    dividendYieldRate: context.dividendYieldRate,
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    pensionPlanState: context.pensionPlanState,
  });

  if (!rows.length) {
    if (showAlert) {
      window.alert("出力できるデータがありません。");
    }
    return null;
  }

  const mismatchRow = rows.find((row) => row.mismatch);
  if (mismatchRow) {
    if (showAlert) {
      window.alert(
        `${mismatchRow.year}年の計算が一致しません。入力値を見直してください。`
      );
    }
    return null;
  }

  return rows;
}

function updateStatementYearOptions(statementRows) {
  if (!statementYearFromSelect || !statementYearToSelect) {
    return;
  }
  statementYearFromSelect.innerHTML = "";
  statementYearToSelect.innerHTML = "";

  if (!statementRows || statementRows.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "選択できません";
    statementYearFromSelect.appendChild(option);
    statementYearFromSelect.disabled = true;
    const optionTo = document.createElement("option");
    optionTo.value = "";
    optionTo.textContent = "選択できません";
    statementYearToSelect.appendChild(optionTo);
    statementYearToSelect.disabled = true;
    if (exportBalanceSheetDecadeButton) {
      exportBalanceSheetDecadeButton.disabled = true;
    }
    if (openBalanceSheetDecadeButton) {
      openBalanceSheetDecadeButton.disabled = true;
    }
    if (exportProfitLossDecadeButton) {
      exportProfitLossDecadeButton.disabled = true;
    }
    if (openProfitLossDecadeButton) {
      openProfitLossDecadeButton.disabled = true;
    }
    return;
  }

  const sortedRows = [...statementRows].sort((a, b) => a.year - b.year);
  sortedRows.forEach((row) => {
    const optionFrom = document.createElement("option");
    optionFrom.value = row.year;
    optionFrom.textContent = `${row.year}年`;
    statementYearFromSelect.appendChild(optionFrom);
    const optionTo = document.createElement("option");
    optionTo.value = row.year;
    optionTo.textContent = optionFrom.textContent;
    statementYearToSelect.appendChild(optionTo);
  });
  const selectedFrom = parseNumber(statementYearFromSelect.value);
  const selectedTo = parseNumber(statementYearToSelect.value);
  const fallbackYear = sortedRows[sortedRows.length - 1].year;
  const matchedFrom = sortedRows.find((row) => row.year === selectedFrom);
  const matchedTo = sortedRows.find((row) => row.year === selectedTo);
  statementYearFromSelect.value = String(
    matchedFrom ? matchedFrom.year : fallbackYear
  );
  statementYearToSelect.value = String(
    matchedTo ? matchedTo.year : statementYearFromSelect.value
  );
  statementYearFromSelect.disabled = false;
  statementYearToSelect.disabled = false;
  if (exportBalanceSheetDecadeButton) {
    exportBalanceSheetDecadeButton.disabled = false;
  }
  if (openBalanceSheetDecadeButton) {
    openBalanceSheetDecadeButton.disabled = false;
  }
  if (exportProfitLossDecadeButton) {
    exportProfitLossDecadeButton.disabled = false;
  }
  if (openProfitLossDecadeButton) {
    openProfitLossDecadeButton.disabled = false;
  }
}

function fillYearSelect(select, years) {
  if (!select) {
    return;
  }
  select.innerHTML = "";
  if (!years.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "選択できません";
    select.appendChild(option);
    select.disabled = true;
    return;
  }
  years.forEach((year) => {
    const option = document.createElement("option");
    option.value = year;
    option.textContent = `${year}年`;
    select.appendChild(option);
  });
  const selected = parseNumber(select.value);
  const fallback = years[years.length - 1];
  select.value = String(years.includes(selected) ? selected : fallback);
  select.disabled = false;
}

function updateAssetDetailYearOptions(statementRows) {
  const years = statementRows
    ? [...new Set(statementRows.map((row) => row.year))].sort((a, b) => a - b)
    : [];
  assetYearSelects.forEach((select) => {
    fillYearSelect(select, years);
  });
  fillYearSelect(assetDetailYearSelect, years);
  if (years.length) {
    const fallback = years[years.length - 1];
    if (!years.includes(assetDetailState.year)) {
      assetDetailState.year = fallback;
    }
  }
  const disabled = years.length === 0;
  assetDetailButtons.forEach((button) => {
    button.disabled = disabled;
    if (disabled) {
      button.title = "入力値を確認してください";
    } else {
      button.removeAttribute("title");
    }
  });
}

function renderAssetDetail() {
  if (!assetDetailTitle || !assetDetailSubtitle || !assetDetailTableBody) {
    return;
  }
  if (assetDetailYearDelta) {
    assetDetailYearDelta.textContent = "-";
  }
  const context = getSimulationContext();
  if (!context) {
    assetDetailTitle.textContent = "資産の月次推移";
    assetDetailSubtitle.textContent = "入力値を確認してください。";
    assetDetailTableBody.innerHTML = "";
    return;
  }
  const assetKey = assetDetailState.key || "cash";
  const label = getAssetLabel(assetKey);
  const selectedYear = parseNumber(assetDetailYearSelect?.value);
  const yearValue = Number.isFinite(selectedYear)
    ? selectedYear
    : assetDetailState.year;
  if (Number.isFinite(yearValue)) {
    assetDetailState.year = yearValue;
  }
  if (assetDetailYearSelect && Number.isFinite(assetDetailState.year)) {
    assetDetailYearSelect.value = String(assetDetailState.year);
  }
  assetDetailTitle.textContent = `${label}の月次推移`;
  assetDetailSubtitle.textContent = `${assetDetailState.year}年の月次推移を表示します。`;

  const startDate = getMonthStart(context.today);
  const monthsRemaining = fullMonthsBetween(
    startDate,
    addYears(context.birthDate, 100)
  );
  const monthlyRows = simulateMonthlySeries({
    startDate,
    monthsRemaining,
    annualRate: context.annualRate,
    categoryRates: context.categoryRates,
    retirementAge: context.retirementMonthIndex,
    retirementIncomeEndAge: context.retirementIncomeEndMonthIndex,
    workIncome: context.incomeTotal,
    workExpense: context.expenseTotal,
    retireIncome: context.retireBaseIncome + context.ongoingIncome,
    retireExpense: context.retireExpenseTotal,
    pensionIncome: context.pensionIncomeTotal + context.ongoingIncome,
    pensionExpense: context.retireExpenseTotal,
    baseDividendIncome: context.baseDividendIncome,
    dividendYieldRate: context.dividendYieldRate,
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    pensionPlanState: context.pensionPlanState,
  });

  const categoryKey = ASSET_CATEGORY_KEYS[assetKey] || assetKey;
  const baseValue = Number.isFinite(context.categories?.[categoryKey])
    ? context.categories[categoryKey]
    : 0;
  let previousValue = Number.isFinite(baseValue) ? baseValue : null;
  const rowsWithDelta = monthlyRows.map((row) => {
    const value = Number.isFinite(row[categoryKey]) ? row[categoryKey] : 0;
    const delta = previousValue === null ? null : value - previousValue;
    previousValue = value;
    return { row, value, delta };
  });
  const rowsForYear = rowsWithDelta.filter(
    (entry) => entry.row.date.getFullYear() === assetDetailState.year
  );
  assetDetailTableBody.innerHTML = "";
  if (!rowsForYear.length) {
    const emptyRow = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 3;
    cell.textContent = "対象年のデータがありません。";
    emptyRow.appendChild(cell);
    assetDetailTableBody.appendChild(emptyRow);
    return;
  }
  let totalDelta = 0;
  rowsForYear.forEach((entry) => {
    const { row, value, delta } = entry;
    if (delta !== null) {
      totalDelta += delta;
    }
    const monthLabel = `${row.date.getFullYear()}-${String(
      row.date.getMonth() + 1
    ).padStart(2, "0")}`;
    const tr = document.createElement("tr");
    const monthCell = document.createElement("td");
    monthCell.textContent = monthLabel;
    const valueCell = document.createElement("td");
    valueCell.textContent = yenFormatter.format(Math.round(value));
    const deltaCell = document.createElement("td");
    deltaCell.textContent =
      delta === null ? "-" : yenFormatter.format(Math.round(delta));
    tr.appendChild(monthCell);
    tr.appendChild(valueCell);
    tr.appendChild(deltaCell);
    assetDetailTableBody.appendChild(tr);
  });
  if (assetDetailYearDelta) {
    assetDetailYearDelta.textContent = yenFormatter.format(
      Math.round(totalDelta)
    );
  }
}

function openAssetDetail(assetKey, year) {
  assetDetailState = {
    key: assetKey,
    year: Number.isFinite(year) ? year : assetDetailState.year,
  };
  if (assetDetailYearSelect && Number.isFinite(assetDetailState.year)) {
    assetDetailYearSelect.value = String(assetDetailState.year);
  }
  renderAssetDetail();
  setActivePage("asset-detail");
}

function buildBalanceSheetCsv(row, birthDate) {
  const header =
    "年度,年齢,対象月数,期首合計（円）,期首預金・現金・暗号資産（円）,期首株式(現物)（円）,期首投資信託（円）,期首債券（円）,期首保険（円）,期首年金（円）,期末合計（円）,期末預金・現金・暗号資産（円）,期末株式(現物)（円）,期末投資信託（円）,期末債券（円）,期末保険（円）,期末年金（円）";
  const line = buildBalanceSheetCsvLine(row, birthDate);
  return [header, line].join("\n");
}

function gainForCategoryInBalance(row, key) {
  return isCompoundingCategory(key) ? row.gainsByCategory[key] : 0;
}

function sumPensionFromData(data) {
  return data.dc || 0;
}

function buildBalanceSheetCsvLine(row, birthDate) {
  const periodStartDate = getPeriodStartDate(row.date, row.months);
  const startTotalExpression = buildSignedExpression([
    row.start.cash,
    row.start.stocks,
    row.start.funds,
    row.start.bonds,
    row.start.insurance,
    row.start.dc,
    row.start.usd,
    row.start.other,
  ]);
  const cashIncomeExpression = buildMultiplicationExpression([
    { amount: row.workIncome, months: row.workMonths },
    { amount: row.retireIncome, months: row.retireMonths },
    { amount: row.pensionIncome, months: row.pensionMonths },
  ]);
  const expenseExpression = buildMultiplicationExpression([
    { amount: row.workExpense, months: row.workMonths },
    { amount: row.retireExpense, months: row.retireMonths },
    { amount: row.pensionExpense, months: row.pensionMonths },
  ]);
  const incomeExpression = `${cashIncomeExpression}`;
  const netExpression = `(${incomeExpression})-(${expenseExpression})`;
  const investmentGainExpression = buildSignedExpression([
    row.gainsByCategory.funds,
    row.gainsByCategory.insurance,
  ]);
  const bondMaturityCashExpression = toCsvNumber(row.bondMaturity || 0);
  const bondMaturityBookExpression = toCsvNumber(row.bondMaturityBook || 0);
  const bondMaturityGainExpression = `(${bondMaturityCashExpression})-(${bondMaturityBookExpression})`;
  const pensionTransferExpression = toCsvNumber(row.pensionTransfer || 0);
  const cashEndExpression = `${toCsvNumber(row.start.cash)}+(${cashIncomeExpression})-(${expenseExpression})-${toCsvNumber(
    row.contributions
  )}+${bondMaturityCashExpression}+${pensionTransferExpression}`;
  const endTotalExpression = `${toCsvNumber(
    row.start.total
  )}+(${netExpression})+(${investmentGainExpression})+(${bondMaturityGainExpression})`;
  return [
    row.year,
    formatAgeYears(birthDate, periodStartDate),
    row.months,
    csvCellWithFormula(row.start.total, startTotalExpression),
    csvCellWithFormula(row.start.cash, `${toCsvNumber(row.start.cash)}`),
    csvCellWithFormula(row.start.stocks, `${toCsvNumber(row.start.stocks)}`),
    csvCellWithFormula(row.start.funds, `${toCsvNumber(row.start.funds)}`),
    csvCellWithFormula(row.start.bonds, `${toCsvNumber(row.start.bonds)}`),
    csvCellWithFormula(row.start.insurance, `${toCsvNumber(row.start.insurance)}`),
    csvCellWithFormula(
      sumPensionFromData(row.start),
      buildSignedExpression([row.start.dc])
    ),
    csvCellWithFormula(row.end.total, endTotalExpression),
    csvCellWithFormula(row.end.cash, cashEndExpression),
    csvCellWithFormula(
      row.end.stocks,
      buildSignedExpression([
        row.start.stocks,
        row.contributionsByCategory.stocks,
        gainForCategoryInBalance(row, "stocks"),
      ])
    ),
    csvCellWithFormula(
      row.end.funds,
      buildSignedExpression([
        row.start.funds,
        row.contributionsByCategory.funds,
        gainForCategoryInBalance(row, "funds"),
      ])
    ),
    csvCellWithFormula(
      row.end.bonds,
      buildSignedExpression([
        row.start.bonds,
        row.contributionsByCategory.bonds,
        -bondMaturityBookExpression,
        gainForCategoryInBalance(row, "bonds"),
      ])
    ),
    csvCellWithFormula(
      row.end.insurance,
      buildSignedExpression([
        row.start.insurance,
        row.contributionsByCategory.insurance,
        gainForCategoryInBalance(row, "insurance"),
      ])
    ),
    csvCellWithFormula(
      sumPensionFromData(row.end),
      buildSignedExpression([
        row.start.dc,
        row.contributionsByCategory.dc,
        -pensionTransferExpression,
        gainForCategoryInBalance(row, "dc"),
      ])
    ),
  ].join(",");
}

function buildProfitLossCsv(row, birthDate) {
  // 会計処理：損益計算書（P/L）の構造
  // 本シミュレーターの「損益計算書」は、実際にはキャッシュフロー分析に近い構造
  // 運用益（含み益）は期中の時価変動を反映しており、実現益と区分されていない
  // 資産増減 = 収支 + 運用益（含み益） + 債券償還差額
  const header =
    "年度,年齢,対象月数,収入（円）,支出（円）,収支（円）,運用益（円）,資産増減（円）,期首合計（円）,期末合計（円）";
  const line = buildProfitLossCsvLine(row, birthDate);
  return [header, line].join("\n");
}

function buildProfitLossCsvLine(row, birthDate) {
  const periodStartDate = getPeriodStartDate(row.date, row.months);
  const startTotalExpression = buildSignedExpression([
    row.start.cash,
    row.start.stocks,
    row.start.funds,
    row.start.bonds,
    row.start.insurance,
    row.start.dc,
    row.start.usd,
    row.start.other,
  ]);
  const cashIncomeExpression = buildMultiplicationExpression([
    { amount: row.workIncome, months: row.workMonths },
    { amount: row.retireIncome, months: row.retireMonths },
    { amount: row.pensionIncome, months: row.pensionMonths },
  ]);
  const expenseExpression = buildMultiplicationExpression([
    { amount: row.workExpense, months: row.workMonths },
    { amount: row.retireExpense, months: row.retireMonths },
    { amount: row.pensionExpense, months: row.pensionMonths },
  ]);
  const incomeExpression = `${cashIncomeExpression}`;
  const netExpression = `(${incomeExpression})-(${expenseExpression})`;
  const investmentGainExpression = buildSignedExpression([
    row.gainsByCategory.stocks,
    row.gainsByCategory.funds,
    row.gainsByCategory.insurance,
    row.gainsByCategory.dc,
    row.gainsByCategory.other,
  ]);
  const bondMaturityCashExpression = toCsvNumber(row.bondMaturity || 0);
  const bondMaturityBookExpression = toCsvNumber(row.bondMaturityBook || 0);
  const bondMaturityGainExpression = `(${bondMaturityCashExpression})-(${bondMaturityBookExpression})`;
  const totalChangeExpression = `(${netExpression})+(${investmentGainExpression})+(${bondMaturityGainExpression})`;
  const endTotalExpression = `${toCsvNumber(
    row.start.total
  )}+(${totalChangeExpression})`;
  return [
    row.year,
    formatAgeYears(birthDate, periodStartDate),
    row.months,
    csvCellWithFormula(row.income, incomeExpression),
    csvCellWithFormula(row.expense, expenseExpression),
    csvCellWithFormula(row.netCash, netExpression),
    csvCellWithFormula(row.investmentGain, investmentGainExpression),
    csvCellWithFormula(row.totalChange, totalChangeExpression),
    csvCellWithFormula(row.start.total, startTotalExpression),
    csvCellWithFormula(row.end.total, endTotalExpression),
  ].join(",");
}

function getSelectedYearRange(statementRows) {
  if (!statementRows || statementRows.length === 0) {
    return null;
  }
  const years = statementRows.map((row) => row.year);
  const maxYear = Math.max(...years);
  const minYear = Math.min(...years);
  const selectedFrom = parseNumber(statementYearFromSelect?.value);
  const selectedTo = parseNumber(statementYearToSelect?.value);
  let startYear = selectedFrom ?? maxYear;
  let endYear = selectedTo ?? startYear;
  if (!Number.isFinite(startYear)) {
    startYear = maxYear;
  }
  if (!Number.isFinite(endYear)) {
    endYear = startYear;
  }
  startYear = Math.min(Math.max(startYear, minYear), maxYear);
  endYear = Math.min(Math.max(endYear, minYear), maxYear);
  if (startYear > endYear) {
    [startYear, endYear] = [endYear, startYear];
  }
  return { startYear, endYear };
}

function getStatementRowsInRange(statementRows, startYear, endYear) {
  if (!statementRows || statementRows.length === 0) {
    return [];
  }
  return statementRows.filter(
    (row) => row.year >= startYear && row.year <= endYear
  );
}

function handleExportBalanceSheetDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const range = getSelectedYearRange(statementRows);
  if (!range) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const rangeRows = getStatementRowsInRange(
    statementRows,
    range.startYear,
    range.endYear
  );
  if (!rangeRows.length) {
    window.alert("対象の期間データがありません。");
    return;
  }
  const header =
    "年度,年齢,対象月数,期首合計（円）,期首預金・現金・暗号資産（円）,期首株式(現物)（円）,期首投資信託（円）,期首債券（円）,期首保険（円）,期首年金（円）,期末合計（円）,期末預金・現金・暗号資産（円）,期末株式(現物)（円）,期末投資信託（円）,期末債券（円）,期末保険（円）,期末年金（円）";
  const birthDate = parseDate(birthDateInput.value);
  const lines = rangeRows.map((row) => buildBalanceSheetCsvLine(row, birthDate));
  const csv = [header, ...lines].join("\n");
  downloadCsvText(
    csv,
    `LifeWealth100_balance_sheet_${rangeRows[0].year}_to_${rangeRows[rangeRows.length - 1].year}.csv`
  );
}

function handleOpenBalanceSheetDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const range = getSelectedYearRange(statementRows);
  if (!range) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const rangeRows = getStatementRowsInRange(
    statementRows,
    range.startYear,
    range.endYear
  );
  if (!rangeRows.length) {
    window.alert("対象の期間データがありません。");
    return;
  }
  const birthDate = parseDate(birthDateInput.value);
  const header = [
    "年度",
    "年齢",
    "対象月数",
    "期首合計（円）",
    "期首預金・現金・暗号資産（円）",
    "期首株式(現物)（円）",
    "期首投資信託（円）",
    "期首債券（円）",
    "期首保険（円）",
    "期首年金（円）",
    "期末合計（円）",
    "期末預金・現金・暗号資産（円）",
    "期末株式(現物)（円）",
    "期末投資信託（円）",
    "期末債券（円）",
    "期末保険（円）",
    "期末年金（円）",
  ];
  const rows = rangeRows.map((row) => {
    const periodStartDate = getPeriodStartDate(row.date, row.months);
    return [
      row.year,
      formatAgeYears(birthDate, periodStartDate),
      toCsvNumber(row.months),
      toCsvNumber(row.start.total),
      toCsvNumber(row.start.cash),
      toCsvNumber(row.start.stocks),
      toCsvNumber(row.start.funds),
      toCsvNumber(row.start.bonds),
      toCsvNumber(row.start.insurance),
      toCsvNumber(row.start.dc),
      toCsvNumber(row.end.total),
      toCsvNumber(row.end.cash),
      toCsvNumber(row.end.stocks),
      toCsvNumber(row.end.funds),
      toCsvNumber(row.end.bonds),
      toCsvNumber(row.end.insurance),
      toCsvNumber(row.end.dc),
    ];
  });
  openSpreadsheetView({
    title: `期間指定貸借対照表 ${rangeRows[0].year}年〜${rangeRows[rangeRows.length - 1].year}年`,
    headers: header,
    rows,
  });
}

function handleExportProfitLossDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const range = getSelectedYearRange(statementRows);
  if (!range) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const rangeRows = getStatementRowsInRange(
    statementRows,
    range.startYear,
    range.endYear
  );
  if (!rangeRows.length) {
    window.alert("対象の期間データがありません。");
    return;
  }
  const header =
    "年度,年齢,対象月数,収入（円）,支出（円）,収支（円）,運用益（円）,資産増減（円）,期首合計（円）,期末合計（円）";
  const birthDate = parseDate(birthDateInput.value);
  const lines = rangeRows.map((row) => buildProfitLossCsvLine(row, birthDate));
  const csv = [header, ...lines].join("\n");
  downloadCsvText(
    csv,
    `LifeWealth100_profit_loss_${rangeRows[0].year}_to_${rangeRows[rangeRows.length - 1].year}.csv`
  );
}

function handleOpenProfitLossDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const range = getSelectedYearRange(statementRows);
  if (!range) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const rangeRows = getStatementRowsInRange(
    statementRows,
    range.startYear,
    range.endYear
  );
  if (!rangeRows.length) {
    window.alert("対象の期間データがありません。");
    return;
  }
  const birthDate = parseDate(birthDateInput.value);
  const header = [
    "年度",
    "年齢",
    "対象月数",
    "収入（円）",
    "支出（円）",
    "収支（円）",
    "運用益（円）",
    "資産増減（円）",
    "期首合計（円）",
    "期末合計（円）",
  ];
  const rows = rangeRows.map((row) => {
    const periodStartDate = getPeriodStartDate(row.date, row.months);
    return [
      row.year,
      formatAgeYears(birthDate, periodStartDate),
      toCsvNumber(row.months),
      toCsvNumber(row.income),
      toCsvNumber(row.expense),
      toCsvNumber(row.netCash),
      toCsvNumber(row.investmentGain),
      toCsvNumber(row.totalChange),
      toCsvNumber(row.start.total),
      toCsvNumber(row.end.total),
    ];
  });
  openSpreadsheetView({
    title: `期間指定損益計算書 ${rangeRows[0].year}年〜${rangeRows[rangeRows.length - 1].year}年`,
    headers: header,
    rows,
  });
}

function render() {
  const birthDate = parseDate(birthDateInput.value);
  let currentAssets = parseNumber(currentAssetsInput.value);
  const annualRatePercent = 0;
  const retirementAgeYears = parseNumber(retirementAgeInput.value);
  const retirementIncomeEndAgeYears =
    parseNumber(retirementIncomeEndAgeInput.value) ?? 100;
  const expenseTotal = sumInputs(expenseInputs);
  const incomeTotal = sumInputs(incomeInputs);
  const retireExpenseTotal = sumInputs(retireExpenseInputs);
  const incomeMap = mapIncomeInputs(incomeInputs);
  const retireIncomeMap = mapIncomeInputs(retireIncomeInputs);
  const retireBaseIncome =
    (retireIncomeMap.salary || 0) +
    (retireIncomeMap.bonus || 0);
  const dividendYieldPercent = parseNumber(dividendYieldInput?.value);
  const dividendYieldRate =
    dividendYieldPercent === null ? 0 : dividendYieldPercent / 100;
  const baseDividendIncome = incomeMap.dividend || 0;
  const pensionIncomeTotal = sumInputs(pensionIncomeInputs);
  const ongoingIncome = (incomeMap.dividend || 0) + (incomeMap.other || 0);
  const retireIncomeTotal = retireBaseIncome;
  const today = new Date();
  updateBondBalanceFromStorage();
  const insurancePlans = readInsurancePlans();
  syncInsuranceContributionFromDetail(insurancePlans, today);
  const pensionPlans = readPensionPlans();
  const pensionChanges = readPensionChanges();
  const activePensionPlans = getActivePensionPlans(pensionPlans);
  const hasPensionPlans = activePensionPlans.length > 0;
  const syncedDcContribution = syncDcContributionFromPensionDetail(
    birthDate,
    activePensionPlans,
    pensionChanges,
    today
  );
  const contributionSchedule = buildContributionSchedule(birthDate, {
    skipDc: hasPensionPlans,
    insurancePlans,
  });
  const investmentBalanceTotal = getInvestmentBalanceTotal();
  if (!Number.isFinite(lastInvestmentBalanceTotal)) {
    lastInvestmentBalanceTotal = investmentBalanceTotal;
  }
  const investmentContributionTotal =
    (parseNumber(contribStocksInput.value) || 0) +
    (parseNumber(contribFundsInput.value) || 0) +
    (parseNumber(contribBondsInput?.value) || 0) +
    (parseNumber(contribInsuranceInput.value) || 0) +
    (parseNumber(contribUsdInput.value) || 0) +
    (parseNumber(contribDcInput.value) || 0);

  const hasBirthDate = birthDate !== null && birthDate <= today;
  const hasRetirement =
    retirementAgeYears !== null &&
    retirementIncomeEndAgeYears !== null &&
    retirementIncomeEndAgeYears >= retirementAgeYears;
  let hasAssets = currentAssets !== null && currentAssets >= 0;

  const annualRate = 0;
  const monthlyNetCash = incomeTotal - expenseTotal;
  const retirementMonthlyNetCash =
    retireIncomeTotal + ongoingIncome - retireExpenseTotal;
  const postRetirementMonthlyNetCash =
    pensionIncomeTotal + ongoingIncome - retireExpenseTotal;

  if (syncedDcContribution) {
    persistInputsToStorage();
  }

  let retirementDate = null;
  let retirementIncomeEndDate = null;
  let monthsRemaining = null;

  if (hasBirthDate && hasRetirement) {
    retirementDate = addYears(birthDate, retirementAgeYears);
    retirementIncomeEndDate = addYears(birthDate, retirementIncomeEndAgeYears);
    const hundredthBirthday = addYears(birthDate, 100);
    monthsRemaining = fullMonthsBetween(today, hundredthBirthday);
  }

  if (investmentTotal) {
    if (Number.isFinite(currentAssets)) {
      const cashReclassTotal = getCashReclassTotalYen();
      if (reclassStatus) {
        reclassStatus.textContent = cashReclassTotal > 0
          ? `現金から引く合計: ${yenFormatter.format(
            Math.round(cashReclassTotal)
          )}`
          : "現金から引く合計: -";
      }
      const summaryBreakdown = getSummaryBreakdownSafe();
      const cashBreakdown = buildCashBreakdown({
        summaryBreakdown,
        currentAssetsValue: currentAssets,
        investmentTotal: investmentBalanceTotal,
      });
      const cashBase = cashBreakdown.baseWithoutPoints;
      const cashNow = cashBreakdown.cashFinal;
      const derivedTotal = Number.isFinite(cashNow)
        ? cashNow + investmentBalanceTotal
        : null;
      if (
        currentAssetsInput &&
        Number.isFinite(derivedTotal) &&
        (summaryBreakdown || cashInputManual) &&
        !currentAssetsInput.matches(":focus")
      ) {
        const roundedTotal = Math.round(derivedTotal);
        currentAssetsInput.value = roundedTotal;
        currentAssets = roundedTotal;
        hasAssets = roundedTotal >= 0;
        lastInvestmentBalanceTotal = investmentBalanceTotal;
      }
      if (
        balanceCashInput &&
        !cashInputManual &&
        Number.isFinite(cashNow) &&
        !balanceCashInput.matches(":focus")
      ) {
        balanceCashInput.value = Math.round(cashNow);
      }
      const cashAfterValue = cashNow - investmentContributionTotal;
      const warningCashNow = cashNow;
      const warningCashAfter = cashAfterValue;
      investmentTotal.textContent = yenFormatter.format(
        Math.round(investmentBalanceTotal)
      );
      cashBalance.textContent = yenFormatter.format(
        Math.round(Number.isFinite(cashNow) ? cashNow : 0)
      );
      if (cashDeduction) {
        cashDeduction.textContent = cashReclassTotal > 0
          ? `現金から引く: ${yenFormatter.format(Math.round(cashReclassTotal))}`
          : "現金から引く: -";
      }
      if (cashBaseDisplay) {
        cashBaseDisplay.textContent = Number.isFinite(cashBase)
          ? `元の現金残高: ${yenFormatter.format(Math.round(cashBase))}`
          : "元の現金残高: -";
      }
      updateCashDetailDisplay(cashBreakdown);
      investmentContribTotal.textContent = yenFormatter.format(
        Math.round(investmentContributionTotal)
      );
      if (investmentAfter) {
        investmentAfter.textContent = yenFormatter.format(
          Math.round(investmentBalanceTotal + investmentContributionTotal)
        );
      }
      if (cashAfter) {
        cashAfter.textContent = yenFormatter.format(
          Math.round(Math.max(0, cashAfterValue))
        );
      }
      if (investmentAlert) {
        const canForecast =
          hasBirthDate && hasRetirement && monthsRemaining !== null;
        let negativeRow = null;
        // 会計チェック：初期現金残高がマイナスの場合
        if (cashNow < 0) {
          investmentAlert.textContent =
            "エラー：現在の現金残高がマイナスです。投資資産の合計が現在資産を超えています。調整値を確認してください。";
          investmentAlert.hidden = false;
        } else if (canForecast) {
          const simulationContext = getSimulationContext();
          const initial = buildInitialCategories(summaryBreakdown, currentAssets);
          const categories = {
            cash: initial.cash,
            stocks: initial.stocks,
            funds: initial.funds,
            bonds: initial.bonds,
            insurance: initial.insurance,
            dc: initial.dc,
            usd: initial.usd,
            other: initial.other,
          };
          negativeRow = findNegativeCashMonthDetailed({
            startDate: today,
            monthsRemaining,
            annualRate,
            categoryRates: simulationContext?.categoryRates,
            retirementAge: monthIndex(retirementDate),
            retirementIncomeEndAge: monthIndex(retirementIncomeEndDate),
            monthlyNetCash,
            retirementMonthlyNetCash,
            postRetirementMonthlyNetCash,
            baseDividendIncome,
            dividendYieldRate,
            contributionSchedule,
            categories,
            bondMaturities: simulationContext?.bondMaturities,
            usdRate: simulationContext?.usdRate,
            pensionPlanState: simulationContext?.pensionPlanState,
          });
        }
        if (warningCashNow < 0) {
          investmentAlert.textContent =
            "エラー：現在の現金残高がマイナスです。投資額を減らすか、現在資産を増やしてください。";
          investmentAlert.hidden = false;
        } else if (negativeRow && birthDate) {
          const ageMonths = fullMonthsBetween(birthDate, negativeRow.date);
          const ageYears = Math.floor(ageMonths / 12);
          const ageRemain = ageMonths % 12;
          investmentAlert.textContent = `警告：将来、現金残高がマイナスになります（${ageYears}歳${ageRemain}か月）。投資額を減らすか収入を増やしてください。`;
          investmentAlert.hidden = false;
        } else if (warningCashAfter < 0) {
          investmentAlert.textContent =
            "警告：積立後の現金残高がマイナスです。積立額を訂正してください。";
          investmentAlert.hidden = false;
        } else {
          investmentAlert.hidden = true;
        }
      }
    } else {
      investmentTotal.textContent = "-";
      cashBalance.textContent = "-";
      investmentContribTotal.textContent = "-";
      if (cashDeduction) {
        cashDeduction.textContent = "現金から引く: -";
      }
      if (cashBaseDisplay) {
        cashBaseDisplay.textContent = "元の現金残高: -";
      }
      updateCashDetailDisplay(null);
    if (reclassStatus) {
      reclassStatus.textContent = "現金から引く合計: -";
    }
    if (investmentAfter) {
      investmentAfter.textContent = "-";
    }
      if (cashAfter) {
        cashAfter.textContent = "-";
      }
      if (investmentAlert) {
        investmentAlert.hidden = true;
      }
    }
  }

  if (!hasBirthDate || !hasRetirement || !hasAssets) {
    updateStatementYearOptions(null);
    resultValue.textContent = "-";
    resultMeta.textContent = "入力値を確認してください。";
    return;
  }

  const ageMonths = fullMonthsBetween(birthDate, today);
  const ageYears = Math.floor(ageMonths / 12);
  const ageRemainMonths = ageMonths % 12;

  const hundredthBirthday = addYears(birthDate, 100);
  monthsRemaining = fullMonthsBetween(today, hundredthBirthday);
  const yearsRemaining = Math.floor(monthsRemaining / 12);
  const remainMonths = monthsRemaining % 12;
  const monthsLabel =
    monthsRemaining === 0
      ? "すでに100歳以上"
      : `100歳まで${yearsRemaining}年${remainMonths}か月`;

  retirementDate = addYears(birthDate, retirementAgeYears);
  retirementIncomeEndDate = addYears(birthDate, retirementIncomeEndAgeYears);

  const context = getSimulationContext();
  const { assets } = context
      ? simulateToAge100Detailed({
        startDate: today,
        monthsRemaining,
        retirementAge: monthIndex(retirementDate),
        retirementIncomeEndAge: monthIndex(retirementIncomeEndDate),
        monthlyNetCash,
        retirementMonthlyNetCash,
        postRetirementMonthlyNetCash,
        baseDividendIncome,
        dividendYieldRate,
        contributionSchedule,
        categories: context.categories,
        categoryRates: context.categoryRates,
        bondMaturities: context.bondMaturities,
        usdRate: context.usdRate,
        pensionPlanState: context.pensionPlanState,
      })
    : simulateToAge100({
        currentAssets,
        annualRate,
        retirementAge: monthIndex(retirementDate),
        retirementIncomeEndAge: monthIndex(retirementIncomeEndDate),
        monthlyNetCash,
        retirementMonthlyNetCash,
        postRetirementMonthlyNetCash,
        baseDividendIncome,
        dividendYieldRate,
        monthsRemaining,
        startMonthIndex: monthIndex(today),
      });

  resultValue.textContent = yenFormatter.format(Math.round(assets));
  resultValue.classList.remove("reveal");
  void resultValue.offsetWidth;
  resultValue.classList.add("reveal");

  const ageLabel =
    ageMonths === 0 ? "0歳0か月" : `${ageYears}歳${ageRemainMonths}か月`;
  resultMeta.textContent = `前提: 誕生日${birthDateInput.value} / 現在年齢${ageLabel} / 定年${retirementAgeYears}歳 / 現在資産${yenFormatter.format(
    currentAssets
  )} / 月収${yenFormatter.format(incomeTotal)} / 月支出${yenFormatter.format(
    expenseTotal
  )} / 定年後月収${yenFormatter.format(
    retireIncomeTotal
  )} / 年金生活月収${yenFormatter.format(
    pensionIncomeTotal
  )} / 定年後月支出${yenFormatter.format(
    retireExpenseTotal
  )} / 給与等終了${retirementIncomeEndAgeYears}歳 / ${monthsLabel}`;

  monthlyExpense.textContent = yenFormatter.format(Math.round(expenseTotal));
  monthlyIncome.textContent = yenFormatter.format(Math.round(incomeTotal));
  monthlyRetireExpense.textContent = yenFormatter.format(
    Math.round(retireExpenseTotal)
  );
  monthlyRetireIncome.textContent = yenFormatter.format(
    Math.round(retireIncomeTotal)
  );
  monthlyPensionIncome.textContent = yenFormatter.format(
    Math.round(pensionIncomeTotal)
  );
  updateRetireExpensePlaceholders();

  // Investment summary is handled before the main validation.

  const statementRows = buildStatementRows({ showAlert: false });
  updateStatementYearOptions(statementRows);
  updateAssetDetailYearOptions(statementRows);
  const isAssetDetailActive = pages.some(
    (page) => page.dataset.page === "asset-detail" && page.classList.contains("is-active")
  );
  if (isAssetDetailActive) {
    renderAssetDetail();
  }
  updateInsuranceDetailSummary();
  updatePensionDetailSummary();
}

manualAdjustmentPairs.forEach(({ balance, adjust }) => {
  if (!balance || !adjust) {
    return;
  }
  adjust.addEventListener("input", () => {
    applyAdjustmentChange(balance, adjust);
  });
});
if (bondAdjustmentPair.balance && bondAdjustmentPair.adjust) {
  bondAdjustmentPair.adjust.addEventListener("input", () => {
    updateBondBalanceFromStorage();
  });
}

[
  birthDateInput,
  currentAssetsInput,
  retirementAgeInput,
  retirementIncomeEndAgeInput,
  ...expenseInputs,
  ...incomeInputs,
  ...retireExpenseInputs,
  ...retireIncomeInputs,
  ...pensionIncomeInputs,
  balanceCashInput,
  adjustCashInput,
  balanceStocksInput,
  balanceFundsInput,
  balanceBondsInput,
  balanceInsuranceInput,
  balanceUsdInput,
  balanceDcInput,
  adjustStocksInput,
  adjustFundsInput,
  adjustBondsInput,
  adjustInsuranceInput,
  adjustUsdInput,
  adjustDcInput,
  rateStocksInput,
  rateFundsInput,
  rateBondsInput,
  rateInsuranceInput,
  contribStocksInput,
  contribFundsInput,
  contribBondsInput,
  contribInsuranceInput,
  contribUsdInput,
  contribDcInput,
  endAgeStocksInput,
  endAgeFundsInput,
  endAgeBondsInput,
  endAgeInsuranceInput,
  endAgeUsdInput,
  endAgeDcInput,
].filter(Boolean).forEach((input) => {
  input.addEventListener("input", render);
  input.addEventListener("input", persistInputsToStorage);
});

[
  balanceStocksInput,
  balanceFundsInput,
  balanceBondsInput,
  balanceInsuranceInput,
  balanceUsdInput,
  balanceDcInput,
  adjustStocksInput,
  adjustFundsInput,
  adjustBondsInput,
  adjustInsuranceInput,
  adjustUsdInput,
  adjustDcInput,
].forEach((input) => {
  input.addEventListener("input", updateCurrentAssetsFromInvestmentBalances);
});

if (balanceCashInput) {
  balanceCashInput.addEventListener("input", () => {
    const nextManual = balanceCashInput.value !== "";
    if (cashInputManual !== nextManual) {
      cashInputManual = nextManual;
      writeCashManualFlag(cashInputManual);
    }
    if (cashInputManual) {
      syncCurrentAssetsFromCashInput();
    }
  });
}
if (adjustCashInput) {
  adjustCashInput.addEventListener("input", syncCurrentAssetsFromCashInput);
}

importButton.addEventListener("click", applyImportedData);
if (exportButton) {
  exportButton.addEventListener("click", handleExportCsv);
}
if (openCsvButton) {
  openCsvButton.addEventListener("click", handleOpenCsv);
}
if (exportBalanceSheetDecadeButton) {
  exportBalanceSheetDecadeButton.addEventListener(
    "click",
    handleExportBalanceSheetDecade
  );
}
if (openBalanceSheetDecadeButton) {
  openBalanceSheetDecadeButton.addEventListener(
    "click",
    handleOpenBalanceSheetDecade
  );
}
if (exportProfitLossDecadeButton) {
  exportProfitLossDecadeButton.addEventListener(
    "click",
    handleExportProfitLossDecade
  );
}
if (openProfitLossDecadeButton) {
  openProfitLossDecadeButton.addEventListener(
    "click",
    handleOpenProfitLossDecade
  );
}
if (exportSyncFolderButton) {
  exportSyncFolderButton.addEventListener("click", () => {
    handleExportSyncFolder();
  });
}
if (importSyncFileButton && syncFileInput) {
  importSyncFileButton.addEventListener("click", () => {
    syncFileInput.value = "";
    syncFileInput.click();
    if (syncStatus) {
      syncStatus.textContent = "ファイルを選択してください";
    }
  });
  syncFileInput.addEventListener("change", () => {
    const file = syncFileInput.files ? syncFileInput.files[0] : null;
    if (!file) {
      if (syncStatus) {
        syncStatus.textContent = "ファイルの選択がキャンセルされました";
      }
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const text = typeof reader.result === "string" ? reader.result : "";
      if (!text) {
        window.alert("ファイルの読み込みに失敗しました。");
        if (syncStatus) {
          syncStatus.textContent = "読み込みに失敗しました";
        }
        return;
      }
      if (syncStatus) {
        syncStatus.textContent = "復元を実行中...";
      }
      importSyncPayload(text);
    };
    reader.onerror = () => {
      window.alert("ファイルの読み込みに失敗しました。");
      if (syncStatus) {
        syncStatus.textContent = "読み込みに失敗しました";
      }
    };
    reader.readAsText(file);
  });
}
if (addBondRowButton) {
  addBondRowButton.addEventListener("click", () => {
    createBondRow();
    persistBondRows();
  });
}
if (sortBondRowsButton) {
  sortBondRowsButton.addEventListener("click", () => {
    sortBondRowsByMaturity();
  });
}
if (addOtherAssetRowButton) {
  addOtherAssetRowButton.addEventListener("click", () => {
    createOtherAssetRow();
    persistOtherAssetRows();
  });
}
if (bondDetailButton) {
  bondDetailButton.addEventListener("click", () => {
    setActivePage("bond-input");
  });
}
if (bondDetailBackButton) {
  bondDetailBackButton.addEventListener("click", () => {
    setActivePage("investment");
  });
}
if (insuranceDetailButton) {
  insuranceDetailButton.addEventListener("click", () => {
    setActivePage("insurance-detail");
  });
}
if (insuranceDetailBackButton) {
  insuranceDetailBackButton.addEventListener("click", () => {
    setActivePage("investment");
  });
}
if (addInsurancePlanRowButton) {
  addInsurancePlanRowButton.addEventListener("click", () => {
    createInsurancePlanRow();
    persistInsurancePlanRows();
  });
}
if (pensionDetailButton) {
  pensionDetailButton.addEventListener("click", () => {
    setActivePage("pension-detail");
  });
}
if (pensionDetailBackButton) {
  pensionDetailBackButton.addEventListener("click", () => {
    setActivePage("investment");
  });
}
if (cashDetailButton) {
  cashDetailButton.addEventListener("click", () => {
    const summaryBreakdown = getSummaryBreakdownSafe();
    const breakdown = buildCashBreakdown({
      summaryBreakdown,
      currentAssetsValue: parseNumber(currentAssetsInput?.value),
      investmentTotal: getInvestmentBalanceTotal(),
    });
    updateCashDetailDisplay(breakdown);
    setActivePage("cash-detail");
  });
}
if (cashDetailBackButton) {
  cashDetailBackButton.addEventListener("click", () => {
    setActivePage("investment");
  });
}
if (addPensionPlanRowButton) {
  addPensionPlanRowButton.addEventListener("click", () => {
    createPensionPlanRow();
    persistPensionPlanRows();
  });
}
if (addPensionChangeRowButton) {
  addPensionChangeRowButton.addEventListener("click", () => {
    createPensionChangeRow();
    persistPensionChangeRows();
  });
}
assetDetailButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const key = button.dataset.assetKey;
    const select = assetYearSelects.find(
      (item) => item.dataset.assetKey === key
    );
    const year = parseNumber(select?.value);
    openAssetDetail(key, year);
  });
});
if (assetDetailYearSelect) {
  assetDetailYearSelect.addEventListener("change", () => {
    const year = parseNumber(assetDetailYearSelect.value);
    if (Number.isFinite(year)) {
      assetDetailState.year = year;
    }
    renderAssetDetail();
  });
}
if (assetDetailBackButton) {
  assetDetailBackButton.addEventListener("click", () => {
    setActivePage("investment");
  });
}
assetDataInput.addEventListener("input", markImportDirty);
summaryDataInput.addEventListener("input", markImportDirty);

persistInputs.forEach((input) => {
  input.addEventListener("input", persistInputsToStorage);
});

cashInputManual = readCashManualFlag();
loadPersistedInputs();
if (balanceCashInput && balanceCashInput.value === "") {
  cashInputManual = false;
  writeCashManualFlag(false);
}
loadBondRows();
loadOtherAssetRows();
if (!readAdjustmentsAppliedFlag()) {
  initializeAdjustments({ applyToBalance: true, includeBonds: false });
  writeAdjustmentsAppliedFlag(true);
  persistInputsToStorage();
} else {
  initializeAdjustments({ applyToBalance: false, includeBonds: true });
}
loadInsurancePlanRows();
loadPensionPlanRows();
loadPensionChangeRows();
render();
updateLastUpdatedFromPush();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker
      .register("sw.js", { updateViaCache: "none" })
      .then((registration) => {
        if (typeof registration.update === "function") {
          registration.update();
        }

        if (registration.waiting) {
          registration.waiting.postMessage({ type: "SKIP_WAITING" });
        }

        registration.addEventListener("updatefound", () => {
          const newWorker = registration.installing;
          if (!newWorker) {
            return;
          }
          newWorker.addEventListener("statechange", () => {
            if (newWorker.state === "installed" && navigator.serviceWorker.controller) {
              newWorker.postMessage({ type: "SKIP_WAITING" });
            }
          });
        });

        let refreshing = false;
        navigator.serviceWorker.addEventListener("controllerchange", () => {
          if (refreshing) {
            return;
          }
          refreshing = true;
          window.location.reload();
        });
      })
      .catch(() => {
        // Ignore service worker registration failures.
      });
  });
}

function setActivePage(pageId) {
  pages.forEach((page) => {
    page.classList.toggle("is-active", page.dataset.page === pageId);
  });
  tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.page === pageId);
  });
}

tabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActivePage(button.dataset.page);
  });
});

setActivePage("core");
