/**
 * LifeWealth 100 Simulator
 * 
 * 会計処理に関する注記：
 * 
 * 資産カテゴリー分類：
 * - cash：現金・預金（現金等価物）
 * - stocks：株式・現物株式（当初原価ベース）
 * - funds：投資信託・ファンド（複利計算により含み益を認識）
 * - bonds：債券・固定利付証券（当初原価ベース、満期時に額面現金化）
 * - insurance：保険商品（複利計算により含み益を認識）
 * - dc：確定拠出年金・企業年金（拠出フェーズと給付フェーズで区分）
 * - points：ポイント・マイル等（雑資産、時価評価）
 * - other：その他資産（外貨等、当初原価ベース）
 * 
 * 運用益の認識：
 * - 複利計算対象（funds、insurance）：期中の利息・配当を含み益として計上
 * - その他：期中の時価変動を反映しない（当初原価ベース）
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
const statementYearSelect = document.getElementById("statementYear");
const exportBalanceSheetButton = document.getElementById("exportBalanceSheet");
const exportProfitLossButton = document.getElementById("exportProfitLoss");
const exportBalanceSheetDecadeButton = document.getElementById(
  "exportBalanceSheetDecade"
);
const exportProfitLossDecadeButton = document.getElementById(
  "exportProfitLossDecade"
);
const bondTableBody = document.getElementById("bondTableBody");
const bondMaturedBody = document.getElementById("bondMaturedBody");
const addBondRowButton = document.getElementById("addBondRow");
const sortBondRowsButton = document.getElementById("sortBondRows");
const bondAverageRate = document.getElementById("bondAverageRate");
const bondUsdRateInput = document.getElementById("bondUsdRate");
const importStatus = document.getElementById("importStatus");
const exportSyncFolderButton = document.getElementById("exportSyncFolder");
const importSyncFileButton = document.getElementById("importSyncFile");
const syncFileInput = document.getElementById("syncFileInput");
const syncStatus = document.getElementById("syncStatus");
const expenseInputs = Array.from(document.querySelectorAll(".expense-input"));
const monthlyExpense = document.getElementById("monthlyExpense");
const incomeInputs = Array.from(document.querySelectorAll(".income-input"));
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
const investmentTotal = document.getElementById("investmentTotal");
const cashBalance = document.getElementById("cashBalance");
const investmentContribTotal = document.getElementById("investmentContribTotal");
const investmentAfter = document.getElementById("investmentAfter");
const cashAfter = document.getElementById("cashAfter");
const investmentAlert = document.getElementById("investmentAlert");
const insuranceDetailButton = document.getElementById("insuranceDetailButton");
const insuranceCurrentAmount = document.getElementById("insuranceCurrentAmount");
const insuranceScheduleBody = document.getElementById("insuranceScheduleBody");
const addInsuranceScheduleRowButton = document.getElementById(
  "addInsuranceScheduleRow"
);
const insuranceFutureBody = document.getElementById("insuranceFutureBody");
const pensionDetailButton = document.getElementById("pensionDetailButton");
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

const STORAGE_KEY = "lifewealth100.inputs.v1";
const BOND_STORAGE_KEY = "lifewealth100.bonds.v1";
const INSURANCE_SCHEDULE_KEY = "lifewealth100.insurance.schedule.v1";
const PENSION_PLANS_KEY = "lifewealth100.pension.plans.v1";
const PENSION_CHANGES_KEY = "lifewealth100.pension.changes.v1";
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
  return key === "funds" || key === "insurance";
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

function readInsuranceSchedule() {
  return safeParseJson(localStorage.getItem(INSURANCE_SCHEDULE_KEY), []);
}

function writeInsuranceSchedule(rows) {
  try {
    localStorage.setItem(INSURANCE_SCHEDULE_KEY, JSON.stringify(rows));
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

function simulateToAge100({
  currentAssets,
  annualRate,
  retirementAge,
  retirementIncomeEndAge,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  monthsRemaining,
  startMonthIndex,
}) {
  const months = Math.max(0, monthsRemaining);
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  let assets = currentAssets;

  for (let i = 0; i < months; i += 1) {
    const startAssets = assets;
    const monthIndex = startMonthIndex + i;
    let cashFlow = monthlyNetCash;
    if (monthIndex >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndex >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
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
  contributionSchedule,
  categories,
  categoryRates,
  bondMaturities,
  usdRate,
  insuranceContributionSchedule,
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
    let cashFlow = monthlyNetCash;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
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
        const amount =
          item.category === "insurance"
            ? getInsuranceContributionAmount(
                monthIndexValue,
                insuranceContributionSchedule,
                item.amount
              )
            : item.amount;
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

function getContributionForMonth(
  monthIndexValue,
  schedule,
  insuranceContributionSchedule
) {
  if (!schedule || schedule.length === 0) {
    return 0;
  }
  return schedule.reduce((sum, item) => {
    if (monthIndexValue < item.endMonthIndex) {
      const amount =
        item.category === "insurance"
          ? getInsuranceContributionAmount(
              monthIndexValue,
              insuranceContributionSchedule,
              item.amount
            )
          : item.amount;
      return sum + amount;
    }
    return sum;
  }, 0);
}

function buildInsuranceContributionSchedule(birthDate, rows) {
  if (!birthDate || !rows || rows.length === 0) {
    return [];
  }
  const entries = rows
    .map((row) => {
      const age = parseNumber(row.age);
      const amount = parseNumber(row.amount);
      if (age === null || amount === null) {
        return null;
      }
      return {
        monthIndex: monthIndex(addYears(birthDate, age)),
        amount,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.monthIndex - b.monthIndex);
  return entries;
}

function getInsuranceContributionAmount(
  monthIndexValue,
  schedule,
  fallbackAmount
) {
  if (!schedule || schedule.length === 0) {
    return fallbackAmount;
  }
  let target = null;
  schedule.forEach((entry) => {
    if (entry.monthIndex <= monthIndexValue) {
      target = entry.amount;
    }
  });
  return target === null ? fallbackAmount : target;
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
    if (!bond.maturityDate || !Number.isFinite(bond.faceValue)) {
      return;
    }
    const monthKey = monthIndex(bond.maturityDate);
    const rate = bond.currency === "USD" ? usdRate : 1;
    const amount = toYenAmount(bond.faceValue * rate);
    if (!amount) {
      return;
    }
    schedule.set(monthKey, (schedule.get(monthKey) || 0) + amount);
  });
  return schedule;
}

function applyBondMaturities(data, schedule, monthIndexValue) {
  if (!schedule || !schedule.size) {
    return 0;
  }
  const amount = schedule.get(monthIndexValue);
  if (!amount) {
    return 0;
  }
  const transferable = Math.min(data.bonds, amount);
  data.bonds -= transferable;
  data.cash += transferable;
  return transferable;
}



function findNegativeCashMonth({
  startCash,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
  retirementAge,
  retirementIncomeEndAge,
  contributionSchedule,
  insuranceContributionSchedule,
  monthsRemaining,
  startMonthIndex,
}) {
  let cash = startCash;
  const months = Math.max(0, monthsRemaining);
  for (let i = 0; i < months; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    let cashFlow = monthlyNetCash;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
    }
    cash += cashFlow;
    cash -= getContributionForMonth(
      monthIndexValue,
      contributionSchedule,
      insuranceContributionSchedule
    );
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
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  insuranceContributionSchedule,
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
    let cashFlow = monthlyNetCash;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
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
        const amount =
          item.category === "insurance"
            ? getInsuranceContributionAmount(
                monthIndexValue,
                insuranceContributionSchedule,
                item.amount
              )
            : item.amount;
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
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  insuranceContributionSchedule,
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
    let cashFlow = monthlyNetCash;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
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
        const amount =
          item.category === "insurance"
            ? getInsuranceContributionAmount(
                monthIndexValue,
                insuranceContributionSchedule,
                item.amount
              )
            : item.amount;
        if (item.category !== "cash") {
          data.cash -= amount;
        }
        data[item.category] += amount;
      }
    });

    applyBondMaturities(data, maturitySchedule, monthIndexValue);
    applyPensionPlanFlow(data, monthIndexValue, planState);

    const isYearEnd = (i + 1) % 12 === 0;
    const isFinal = i === totalMonths - 1;
    if (isYearEnd || isFinal) {
      const date = addMonths(startDate, i + 1);
      const total =
        data.cash +
        data.stocks +
        data.funds +
        data.bonds +
        data.insurance +
        data.dc +
        data.points +
        data.other;
      rows.push({
        date,
        total,
        ...data,
      });
    }
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
    data.points +
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
  contributionSchedule,
  categories,
  bondMaturities,
  usdRate,
  insuranceContributionSchedule,
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
  let yearBondMaturity = 0;      // 債券償還による現金化
  let yearPensionTransfer = 0;   // 年金給付による現金化
  let yearContributionByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    dc: 0,
    other: 0,
  };
  let yearGainByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    dc: 0,
    other: 0,
  };
  let yearWorkMonths = 0;
  let yearRetireMonths = 0;
  let yearPensionMonths = 0;
  let yearMonths = 0;

  for (let i = 0; i < totalMonths; i += 1) {
    const monthIndexValue = startMonthIndex + i;
    const monthDate = addMonths(startDate, i);

    let monthlyIncome = workIncome;
    let monthlyExpense = workExpense;
    let phase = "work";
    if (monthIndexValue >= retirementAge) {
      monthlyIncome = retireIncome;
      monthlyExpense = retireExpense;
      phase = "retire";
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      monthlyIncome = pensionIncome;
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
        const amount =
          item.category === "insurance"
            ? getInsuranceContributionAmount(
                monthIndexValue,
                insuranceContributionSchedule,
                item.amount
              )
            : item.amount;
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

    yearBondMaturity += applyBondMaturities(
      data,
      maturitySchedule,
      monthIndexValue
    );
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
      // totalChange = 資産増減の総額（netCash + investmentGain）
      // 期末残高 = 期首残高 + 資産増減総額
      const netCash = yearIncome - yearExpense;
      const totalChange = netCash + yearInvestmentGain;
      const mismatch =
        Math.round(startTotal + totalChange) !== Math.round(endTotal);
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
        bondMaturity: yearBondMaturity,
        pensionTransfer: yearPensionTransfer,
        months: yearMonths,
        contributions: yearContribution,
        contributionsByCategory: { ...yearContributionByCategory },
        gainsByCategory: { ...yearGainByCategory },
        workMonths: yearWorkMonths,
        retireMonths: yearRetireMonths,
        pensionMonths: yearPensionMonths,
        workIncome,
        retireIncome,
        pensionIncome,
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
      yearBondMaturity = 0;
      yearPensionTransfer = 0;
      yearContribution = 0;
      yearContributionByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        dc: 0,
        other: 0,
      };
      yearGainByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        dc: 0,
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
  const cash = getAmount(/預金|現金|暗号資産/);
  const stocks = getAmount(/株式/);
  const funds = getAmount(/投資信託/);
  const bonds = getAmount(/債券/);
  const insurance = getAmount(/保険/);
  const pension = getAmount(/年金/);
  const points = getAmount(/ポイント/);
  const other = getAmount(/その他/);
  const investmentsTotal =
    stocks + funds + bonds + insurance + pension + points + other;
  const derivedTotal = cash + investmentsTotal;
  const total =
    derivedTotal > 0 &&
    cash > 0 &&
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
  const toEndMonth = (input) =>
    monthIndex(addYears(birthDate, parseNumber(input.value) || 0));

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
    {
      category: "bonds",
      amount: parseNumber(contribBondsInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeBondsInput),
    },
    {
      category: "insurance",
      amount: parseNumber(contribInsuranceInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeInsuranceInput),
    },
    {
      category: "other",
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
    (parseNumber(adjustStocksInput?.value) || 0) +
    (parseNumber(balanceFundsInput.value) || 0) +
    (parseNumber(adjustFundsInput?.value) || 0) +
    (parseNumber(balanceBondsInput.value) || 0) +
    (parseNumber(adjustBondsInput?.value) || 0) +
    (parseNumber(balanceInsuranceInput.value) || 0) +
    (parseNumber(adjustInsuranceInput?.value) || 0) +
    (parseNumber(balanceUsdInput.value) || 0) +
    (parseNumber(adjustUsdInput?.value) || 0) +
    (parseNumber(balanceDcInput.value) || 0) +
    (parseNumber(adjustDcInput?.value) || 0)
  );
}

function getInvestmentAdjustments() {
  return {
    stocks: parseNumber(adjustStocksInput?.value) || 0,
    funds: parseNumber(adjustFundsInput?.value) || 0,
    bonds: parseNumber(adjustBondsInput?.value) || 0,
    insurance: parseNumber(adjustInsuranceInput?.value) || 0,
    usd: parseNumber(adjustUsdInput?.value) || 0,
    dc: parseNumber(adjustDcInput?.value) || 0,
  };
}

function getAdjustedBalance(value, adjustment) {
  return (value || 0) + (adjustment || 0);
}

function updateCurrentAssetsFromInvestmentBalances() {
  if (!currentAssetsInput) {
    return;
  }
  const investmentTotal = getInvestmentBalanceTotal();
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

// 会計処理：初期資産の分類
// マネーフォワードなどからのインポートデータを、各資産カテゴリーに分類
// ユーザーの調整値を加算して初期残高を確定
function buildInitialCategories(summaryBreakdown, currentAssets) {
  const adjustments = getInvestmentAdjustments();
  if (summaryBreakdown) {
    const totalFromInput =
      Number.isFinite(currentAssets) ?
        currentAssets :
        summaryBreakdown.total ||
          summaryBreakdown.cash +
            summaryBreakdown.stocks +
            summaryBreakdown.funds +
            summaryBreakdown.bonds +
            summaryBreakdown.insurance +
            summaryBreakdown.pension +
            summaryBreakdown.points +
            summaryBreakdown.other;
    const pensionTotal = summaryBreakdown.pension || 0;
    const dc = getAdjustedBalance(pensionTotal, adjustments.dc);
    const usdBalance = getAdjustedBalance(
      parseNumber(balanceUsdInput.value),
      adjustments.usd
    );
    const stocks = getAdjustedBalance(summaryBreakdown.stocks, adjustments.stocks);
    const funds = getAdjustedBalance(summaryBreakdown.funds, adjustments.funds);
    const bonds = getAdjustedBalance(summaryBreakdown.bonds, adjustments.bonds);
    const insurance = getAdjustedBalance(
      summaryBreakdown.insurance,
      adjustments.insurance
    );
    const points = summaryBreakdown.points || 0;
    const other = (summaryBreakdown.other || 0) + usdBalance;
    const cashFromSummary =
      Number.isFinite(summaryBreakdown.cash) ? summaryBreakdown.cash : null;
    const investmentTotal = stocks + funds + bonds + insurance + dc + points + other;
    const derivedTotal =
      (cashFromSummary || 0) + investmentTotal;
    const total =
      derivedTotal > 0 && totalFromInput < derivedTotal
        ? derivedTotal
        : totalFromInput;
    const cash =
      cashFromSummary !== null ? cashFromSummary : total - investmentTotal;
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
      points,
      other,
      total: total || 0,
    };
  }

  const stocks = getAdjustedBalance(
    parseNumber(balanceStocksInput.value),
    adjustments.stocks
  );
  const funds = getAdjustedBalance(
    parseNumber(balanceFundsInput.value),
    adjustments.funds
  );
  const bonds = getAdjustedBalance(
    parseNumber(balanceBondsInput.value),
    adjustments.bonds
  );
  const insurance = getAdjustedBalance(
    parseNumber(balanceInsuranceInput.value),
    adjustments.insurance
  );
  const dc = getAdjustedBalance(
    parseNumber(balanceDcInput.value),
    adjustments.dc
  );
  const usd = getAdjustedBalance(
    parseNumber(balanceUsdInput.value),
    adjustments.usd
  );
  // 会計処理：初期現金残高の計算
  // 現在資産 = 現金 + 各投資資産 の関係から現金を逆算
  // 投資資産の合計が現在資産を超えないようにチェック（超える場合は警告）
  const investmentTotal = stocks + funds + bonds + insurance + dc + usd;
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

  return {
    cash: Math.max(0, cash),  // 現金がマイナスになるのを防ぐ（0以上に調整）
    stocks,
    funds,
    bonds,
    insurance,
    dc,
    points: 0,
    other: usd,
    total: currentAssetsValue,
  };
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatAgeYears(birthDate, atDate) {
  if (!birthDate || !atDate) {
    return "-";
  }
  const months = fullMonthsBetween(birthDate, atDate);
  return Math.floor(months / 12);
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

function downloadCsv(rows, birthDate) {
  const header =
    "日付,年齢,合計（円）,預金・現金・暗号資産（円）,株式(現物)（円）,投資信託（円）,債券（円）,保険（円）,年金（円）,ポイント（円）,その他の資産（円）";
  const lines = rows.map((row) =>
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
      toCsvNumber(row.points),
      toCsvNumber(row.other),
    ].join(",")
  );
  const csv = [header, ...lines].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "LifeWealth100_annual.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function downloadCsvText(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
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
    find(/合計/, /負債|投資|収支|損益/)
  );
}

function sumInvestmentsFromList(text) {
  const table = parseTable(text);
  if (!table) {
    return sumInvestmentsFromText(text);
  }

  const typeIndex = mapHeaderIndex(table.headers, /資産区分/);
  const nameIndex = mapHeaderIndex(table.headers, /名称/);
  const amountIndex = mapHeaderIndex(
    table.headers,
    /(金額|評価額|残高).*円/
  );
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
    const amount = row[amountIndex] ? parseAmount(row[amountIndex]) : null;
    if (amount === null) {
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
    if (/代表口座-米ドル普通\s*住信SBIネット銀行/.test(name)) {
      totals.usd += amount;
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
    if (/代表口座-米ドル普通\s*住信SBIネット銀行/.test(line)) {
      totals.usd += amount;
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
    currencySelect,
    faceValueInput,
    purchasePriceInput,
    maturityDateInput,
    rateInput,
  ].forEach((input) => {
    input.addEventListener("input", () => {
      if (input === nameInput) {
        nameInput.title = nameInput.value;
      }
      persistBondRows();
    });
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
    data[input.dataset.key] = input.value;
  });
  return data;
}

function createInsuranceScheduleRow(data = {}) {
  if (!insuranceScheduleBody) {
    return;
  }
  const row = document.createElement("tr");

  const ageInput = document.createElement("input");
  ageInput.type = "number";
  ageInput.min = "0";
  ageInput.max = "120";
  ageInput.step = "1";
  ageInput.inputMode = "numeric";
  ageInput.value = data.age ?? "";
  ageInput.dataset.key = "age";

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.min = "0";
  amountInput.step = "1000";
  amountInput.inputMode = "numeric";
  amountInput.value = data.amount ?? "";
  amountInput.dataset.key = "amount";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.textContent = "削除";
  removeButton.addEventListener("click", () => {
    row.remove();
    persistInsuranceScheduleRows();
  });

  const cells = [ageInput, amountInput].map((input) => {
    const td = document.createElement("td");
    td.appendChild(input);
    return td;
  });
  const actionCell = document.createElement("td");
  actionCell.className = "insurance-action";
  actionCell.appendChild(removeButton);
  cells.push(actionCell);

  cells.forEach((cell) => row.appendChild(cell));
  insuranceScheduleBody.appendChild(row);

  [ageInput, amountInput].forEach((input) => {
    input.addEventListener("input", persistInsuranceScheduleRows);
  });
}

function serializeInsuranceScheduleRow(row) {
  const data = {};
  row.querySelectorAll("input").forEach((input) => {
    data[input.dataset.key] = input.value;
  });
  return data;
}

function persistInsuranceScheduleRows() {
  if (!insuranceScheduleBody) {
    return;
  }
  const rows = Array.from(insuranceScheduleBody.querySelectorAll("tr")).map((row) =>
    serializeInsuranceScheduleRow(row)
  );
  writeInsuranceSchedule(rows);
  updateInsuranceDetailSummary();
}

function loadInsuranceScheduleRows() {
  if (!insuranceScheduleBody) {
    return;
  }
  insuranceScheduleBody.innerHTML = "";
  const rows = readInsuranceSchedule();
  if (!rows.length) {
    createInsuranceScheduleRow();
  } else {
    rows.forEach((row) => createInsuranceScheduleRow(row));
  }
  updateInsuranceDetailSummary();
}

function updateInsuranceDetailSummary() {
  if (!insuranceCurrentAmount || !insuranceFutureBody) {
    return;
  }
  const currentAmount = parseNumber(contribInsuranceInput.value);
  insuranceCurrentAmount.textContent = Number.isFinite(currentAmount)
    ? yenFormatter.format(currentAmount)
    : "-";

  const rawRows = readInsuranceSchedule();
  const map = new Map();
  rawRows.forEach((row) => {
    const age = parseNumber(row.age);
    const amount = parseNumber(row.amount);
    if (age === null || amount === null) {
      return;
    }
    map.set(age, amount);
  });
  const sorted = Array.from(map.entries()).sort((a, b) => a[0] - b[0]);
  insuranceFutureBody.innerHTML = "";
  if (!sorted.length) {
    const emptyRow = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 2;
    cell.textContent = "変更予定がありません。";
    emptyRow.appendChild(cell);
    insuranceFutureBody.appendChild(emptyRow);
    return;
  }
  sorted.forEach(([age, amount]) => {
    const row = document.createElement("tr");
    const ageCell = document.createElement("td");
    ageCell.textContent = `${age}歳`;
    const amountCell = document.createElement("td");
    amountCell.textContent = yenFormatter.format(amount);
    row.appendChild(ageCell);
    row.appendChild(amountCell);
    insuranceFutureBody.appendChild(row);
  });
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
  const changeMap = new Map();
  (changes || []).forEach((row) => {
    if (!row || !row.planId) {
      return;
    }
    const age = parseNumber(row.age);
    const amount = parseNumber(row.amount);
    if (age === null || amount === null || !birthDate) {
      return;
    }
    const monthIndexValue = monthIndex(addYears(birthDate, age));
    if (!changeMap.has(row.planId)) {
      changeMap.set(row.planId, []);
    }
    changeMap.get(row.planId).push({ monthIndex: monthIndexValue, amount });
  });

  if (!birthDate) {
    return plans.reduce(
      (sum, plan) => sum + (parseNumber(plan.amount) || 0),
      0
    );
  }

  const currentMonthIndex = monthIndex(atDate || new Date());
  return plans.reduce((sum, plan) => {
    const baseAmount = parseNumber(plan.amount) || 0;
    let amount = baseAmount;
    const changesForPlan = (changeMap.get(plan.id) || []).sort(
      (a, b) => a.monthIndex - b.monthIndex
    );
    changesForPlan.forEach((change) => {
      if (change.monthIndex <= currentMonthIndex) {
        amount = change.amount;
      }
    });
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
    bondUsdRateInput.addEventListener("input", () => {
      const next = readBondStorage();
      next.usdRate = bondUsdRateInput.value;
      writeBondStorage(next);
    });
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
    return;
  }
  const avg = values.reduce((sum, value) => sum + value, 0) / values.length;
  bondAverageRate.textContent = `平均利率: ${percentFormatter.format(avg)}%`;
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

function buildSyncPayload() {
  persistInputsToStorage();
  persistBondRows();
  persistInsuranceScheduleRows();
  persistPensionPlanRows();
  persistPensionChangeRows();
  const inputs = safeParseJson(localStorage.getItem(STORAGE_KEY), {});
  const bonds = readBondStorage();
  const insuranceSchedule = readInsuranceSchedule();
  const pensionPlans = readPensionPlans();
  const pensionChanges = readPensionChanges();
  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    inputs,
    bonds,
    insuranceSchedule,
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
  const inputs = isPlainObject(data.inputs) ? data.inputs : null;
  const bonds = data.bonds ? normalizeBondStorage(data.bonds) : null;
  const insuranceSchedule = Array.isArray(data.insuranceSchedule)
    ? data.insuranceSchedule
    : null;
  const pensionPlans = Array.isArray(data.pensionPlans)
    ? data.pensionPlans
    : null;
  const pensionChanges = Array.isArray(data.pensionChanges)
    ? data.pensionChanges
    : null;
  if (!inputs && !bonds && !insuranceSchedule && !pensionPlans && !pensionChanges) {
    window.alert("同期データに読み込める内容がありません。");
    return false;
  }
  try {
    if (inputs) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(inputs));
    }
    if (bonds) {
      writeBondStorage(bonds);
    }
    if (insuranceSchedule) {
      writeInsuranceSchedule(insuranceSchedule);
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
  loadBondRows();
  loadInsuranceScheduleRows();
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

  const table = parseTable(text);
  if (!table) {
    return null;
  }

  const amountIndex = table.headers.findIndex((header) =>
    /(金額|評価額|残高).*円/.test(header)
  );
  if (amountIndex === -1) {
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
  const assetTotal = sumAssetsFromList(assetDataInput.value);
  const summaryTotal = latestTotalFromSummary(summaryDataInput.value);
  const investmentTotals =
    sumInvestmentsFromSummary(summaryDataInput.value) ??
    sumInvestmentsFromList(assetDataInput.value);

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
    balanceBondsInput.value = Math.round(investmentTotals.bonds);
    balanceInsuranceInput.value = Math.round(investmentTotals.insurance);
    balanceUsdInput.value = Math.round(investmentTotals.usd);
    balanceDcInput.value = Math.round(investmentTotals.dc);
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

  const rows = simulateAnnualSeries({
    startDate: context.today,
    monthsRemaining: context.monthsRemaining,
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
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    insuranceContributionSchedule: context.insuranceContributionSchedule,
    pensionPlanState: context.pensionPlanState,
  });

  if (!rows.length) {
    window.alert("出力できるデータがありません。");
    return;
  }

  downloadCsv(rows, context.birthDate);
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
    other: 0,
  };
  const storedBonds = readBondStorage();
  const usdRate = parseNumber(bondUsdRateInput?.value ?? storedBonds.usdRate) ?? 0;
  const bondMaturities = (storedBonds.active || []).map((row) => ({
    maturityDate: parseDate(row.maturityDate),
    currency: row.currency || "JPY",
    faceValue: parseNumber(row.faceValue) || 0,
  }));
  const insuranceContributionSchedule = buildInsuranceContributionSchedule(
    birthDate,
    readInsuranceSchedule()
  );
  const pensionPlans = readPensionPlans();
  const hasPensionPlans = pensionPlans.length > 0;
  const contributionSchedule = buildContributionSchedule(birthDate, {
    skipDc: hasPensionPlans,
  });
  const expenseTotal = sumInputs(expenseInputs);
  const incomeTotal = sumInputs(incomeInputs);
  const retireExpenseTotal = sumInputs(retireExpenseInputs);
  const incomeMap = mapIncomeInputs(incomeInputs);
  const retireIncomeMap = mapIncomeInputs(retireIncomeInputs);
  const retireBaseIncome =
    (retireIncomeMap.salary || 0) +
    (retireIncomeMap.bonus || 0);
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
    points: initial.points,
    other: initial.other,
  };
  const pensionPlanState = buildPensionPlanState(
    birthDate,
    pensionPlans,
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
    contributionSchedule,
    categories,
    bondMaturities,
    usdRate,
    insuranceContributionSchedule,
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
    contributionSchedule: context.contributionSchedule,
    categories: context.categories,
    bondMaturities: context.bondMaturities,
    usdRate: context.usdRate,
    insuranceContributionSchedule: context.insuranceContributionSchedule,
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
  if (!statementYearSelect) {
    return;
  }
  statementYearSelect.innerHTML = "";

  if (!statementRows || statementRows.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "選択できません";
    statementYearSelect.appendChild(option);
    statementYearSelect.disabled = true;
    if (exportBalanceSheetButton) {
      exportBalanceSheetButton.disabled = true;
    }
    if (exportProfitLossButton) {
      exportProfitLossButton.disabled = true;
    }
    if (exportBalanceSheetDecadeButton) {
      exportBalanceSheetDecadeButton.disabled = true;
    }
    if (exportProfitLossDecadeButton) {
      exportProfitLossDecadeButton.disabled = true;
    }
    return;
  }

  const selectedYear = parseNumber(statementYearSelect.value);
  const currentYear = new Date().getFullYear();
  const sortedRows = [...statementRows].sort((a, b) => b.year - a.year);
  let currentRowIndex = sortedRows.findIndex((row) => row.year === currentYear);
  if (currentRowIndex === -1) {
    let closestIndex = 0;
    let closestDiff = Infinity;
    sortedRows.forEach((row, index) => {
      const diff = Math.abs(row.year - currentYear);
      if (diff < closestDiff) {
        closestDiff = diff;
        closestIndex = index;
      }
    });
    currentRowIndex = closestIndex;
  }
  if (currentRowIndex > 0) {
    const [currentRow] = sortedRows.splice(currentRowIndex, 1);
    sortedRows.unshift(currentRow);
  }
  sortedRows.forEach((row) => {
    const option = document.createElement("option");
    option.value = row.year;
    option.textContent =
      row.months < 12 ? `${row.year}年（${row.months}か月）` : `${row.year}年`;
    statementYearSelect.appendChild(option);
  });
  const matched = sortedRows.find((row) => row.year === selectedYear);
  const fallbackYear = sortedRows[0].year;
  statementYearSelect.value = String(matched ? matched.year : fallbackYear);
  statementYearSelect.disabled = false;
  if (exportBalanceSheetButton) {
    exportBalanceSheetButton.disabled = false;
  }
  if (exportProfitLossButton) {
    exportProfitLossButton.disabled = false;
  }
  if (exportBalanceSheetDecadeButton) {
    exportBalanceSheetDecadeButton.disabled = false;
  }
  if (exportProfitLossDecadeButton) {
    exportProfitLossDecadeButton.disabled = false;
  }
}

function buildBalanceSheetCsv(row, birthDate) {
  const header =
    "年度,年齢,対象月数,期首合計（円）,期首預金・現金・暗号資産（円）,期首株式(現物)（円）,期首投資信託（円）,期首債券（円）,期首保険（円）,期首年金（円）,期首ポイント（円）,期首その他の資産（円）,期末合計（円）,期末預金・現金・暗号資産（円）,期末株式(現物)（円）,期末投資信託（円）,期末債券（円）,期末保険（円）,期末年金（円）,期末ポイント（円）,期末その他の資産（円）";
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
    row.start.points,
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
  const bondMaturityExpression = toCsvNumber(row.bondMaturity || 0);
  const pensionTransferExpression = toCsvNumber(row.pensionTransfer || 0);
  const cashEndExpression = `${toCsvNumber(row.start.cash)}+(${cashIncomeExpression})-(${expenseExpression})-${toCsvNumber(
    row.contributions
  )}+${bondMaturityExpression}+${pensionTransferExpression}`;
  const endTotalExpression = `${toCsvNumber(
    row.start.total
  )}+(${netExpression})+(${investmentGainExpression})`;
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
    csvCellWithFormula(row.start.points, `${toCsvNumber(row.start.points)}`),
    csvCellWithFormula(row.start.other, `${toCsvNumber(row.start.other)}`),
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
        -bondMaturityExpression,
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
    csvCellWithFormula(
      row.end.points,
      buildSignedExpression([row.start.points])
    ),
    csvCellWithFormula(
      row.end.other,
      buildSignedExpression([
        row.start.other,
        row.contributionsByCategory.other,
        gainForCategoryInBalance(row, "other"),
      ])
    ),
  ].join(",");
}

function buildProfitLossCsv(row, birthDate) {
  // 会計処理：損益計算書（P/L）の構造
  // 本シミュレーターの「損益計算書」は、実際にはキャッシュフロー分析に近い構造
  // 運用益（含み益）は期中の時価変動を反映しており、実現益と区分されていない
  // 資産増減 = 収支 + 運用益（含み益）
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
    row.start.points,
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
    row.gainsByCategory.bonds,
    row.gainsByCategory.insurance,
    row.gainsByCategory.dc,
    row.gainsByCategory.other,
  ]);
  const totalChangeExpression = `(${netExpression})+(${investmentGainExpression})`;
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

function getStatementRow(statementRows) {
  if (!statementRows || statementRows.length === 0) {
    return null;
  }
  const selectedYear = parseNumber(statementYearSelect?.value);
  if (selectedYear === null) {
    return statementRows[statementRows.length - 1];
  }
  return statementRows.find((row) => row.year === selectedYear) ??
    statementRows[statementRows.length - 1];
}

function getDecadeRows(statementRows) {
  if (!statementRows || statementRows.length === 0) {
    return [];
  }
  const selectedYear = parseNumber(statementYearSelect?.value);
  const startYear =
    selectedYear !== null ? selectedYear : statementRows[0].year;
  const endYear = startYear + 10;
  return statementRows.filter(
    (row) => row.year >= startYear && row.year < endYear
  );
}

function handleExportBalanceSheet() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const row = getStatementRow(statementRows);
  if (!row) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const birthDate = parseDate(birthDateInput.value);
  const csv = buildBalanceSheetCsv(row, birthDate);
  downloadCsvText(csv, `LifeWealth100_balance_sheet_${row.year}.csv`);
}

function handleExportProfitLoss() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const row = getStatementRow(statementRows);
  if (!row) {
    window.alert("対象年度のデータがありません。");
    return;
  }
  const birthDate = parseDate(birthDateInput.value);
  const csv = buildProfitLossCsv(row, birthDate);
  downloadCsvText(csv, `LifeWealth100_profit_loss_${row.year}.csv`);
}

function handleExportBalanceSheetDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const decadeRows = getDecadeRows(statementRows);
  if (!decadeRows.length) {
    window.alert("対象の10年データがありません。");
    return;
  }
  const header =
    "年度,年齢,対象月数,期首合計（円）,期首預金・現金・暗号資産（円）,期首株式(現物)（円）,期首投資信託（円）,期首債券（円）,期首保険（円）,期首年金（円）,期首ポイント（円）,期首その他の資産（円）,期末合計（円）,期末預金・現金・暗号資産（円）,期末株式(現物)（円）,期末投資信託（円）,期末債券（円）,期末保険（円）,期末年金（円）,期末ポイント（円）,期末その他の資産（円）";
  const birthDate = parseDate(birthDateInput.value);
  const lines = decadeRows.map((row) => buildBalanceSheetCsvLine(row, birthDate));
  const csv = [header, ...lines].join("\n");
  downloadCsvText(
    csv,
    `LifeWealth100_balance_sheet_${decadeRows[0].year}_to_${decadeRows[decadeRows.length - 1].year}.csv`
  );
}

function handleExportProfitLossDecade() {
  const statementRows = buildStatementRows({ showAlert: true });
  if (!statementRows) {
    return;
  }
  const decadeRows = getDecadeRows(statementRows);
  if (!decadeRows.length) {
    window.alert("対象の10年データがありません。");
    return;
  }
  const header =
    "年度,年齢,対象月数,収入（円）,支出（円）,収支（円）,運用益（円）,資産増減（円）,期首合計（円）,期末合計（円）";
  const birthDate = parseDate(birthDateInput.value);
  const lines = decadeRows.map((row) => buildProfitLossCsvLine(row, birthDate));
  const csv = [header, ...lines].join("\n");
  downloadCsvText(
    csv,
    `LifeWealth100_profit_loss_${decadeRows[0].year}_to_${decadeRows[decadeRows.length - 1].year}.csv`
  );
}

function render() {
  if (lastUpdated) {
    lastUpdated.textContent = `更新日時: ${formatDateTime(new Date())}`;
  }

  const birthDate = parseDate(birthDateInput.value);
  const currentAssets = parseNumber(currentAssetsInput.value);
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
  const pensionIncomeTotal = sumInputs(pensionIncomeInputs);
  const ongoingIncome = (incomeMap.dividend || 0) + (incomeMap.other || 0);
  const retireIncomeTotal = retireBaseIncome;
  const today = new Date();
  const pensionPlans = readPensionPlans();
  const pensionChanges = readPensionChanges();
  const hasPensionPlans = pensionPlans.length > 0;
  const syncedDcContribution = syncDcContributionFromPensionDetail(
    birthDate,
    pensionPlans,
    pensionChanges,
    today
  );
  const contributionSchedule = buildContributionSchedule(birthDate, {
    skipDc: hasPensionPlans,
  });
  const investmentBalanceTotal = getInvestmentBalanceTotal();
  if (!Number.isFinite(lastInvestmentBalanceTotal)) {
    lastInvestmentBalanceTotal = investmentBalanceTotal;
  }
  const investmentContributionTotal =
    (parseNumber(contribStocksInput.value) || 0) +
    (parseNumber(contribFundsInput.value) || 0) +
    (parseNumber(contribBondsInput.value) || 0) +
    (parseNumber(contribInsuranceInput.value) || 0) +
    (parseNumber(contribUsdInput.value) || 0) +
    (parseNumber(contribDcInput.value) || 0);

  const hasBirthDate = birthDate !== null && birthDate <= today;
  const hasRetirement =
    retirementAgeYears !== null &&
    retirementIncomeEndAgeYears !== null &&
    retirementIncomeEndAgeYears >= retirementAgeYears;
  const hasAssets = currentAssets !== null && currentAssets >= 0;

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
      const summaryBreakdown = importDirty
        ? null
        : getSummaryBreakdown(summaryDataInput.value);
      const cashNow =
        summaryBreakdown && Number.isFinite(summaryBreakdown.cash)
          ? summaryBreakdown.cash
          : currentAssets - investmentBalanceTotal;
      const cashAfterValue = cashNow - investmentContributionTotal;
      const warningCashNow = cashNow;
      const warningCashAfter = cashAfterValue;
      investmentTotal.textContent = yenFormatter.format(
        Math.round(investmentBalanceTotal)
      );
      cashBalance.textContent = yenFormatter.format(Math.round(Math.max(0, cashNow)));
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
            points: initial.points,
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
            contributionSchedule,
            categories,
            bondMaturities: simulationContext?.bondMaturities,
            usdRate: simulationContext?.usdRate,
            insuranceContributionSchedule:
              simulationContext?.insuranceContributionSchedule,
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
        contributionSchedule,
        categories: context.categories,
        categoryRates: context.categoryRates,
        bondMaturities: context.bondMaturities,
        usdRate: context.usdRate,
        insuranceContributionSchedule: context.insuranceContributionSchedule,
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

  // Investment summary is handled before the main validation.

  updateStatementYearOptions(buildStatementRows({ showAlert: false }));
  updateInsuranceDetailSummary();
  updatePensionDetailSummary();
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
].forEach((input) => {
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

importButton.addEventListener("click", applyImportedData);
if (exportButton) {
  exportButton.addEventListener("click", handleExportCsv);
}
if (exportBalanceSheetButton) {
  exportBalanceSheetButton.addEventListener("click", handleExportBalanceSheet);
}
if (exportProfitLossButton) {
  exportProfitLossButton.addEventListener("click", handleExportProfitLoss);
}
if (exportBalanceSheetDecadeButton) {
  exportBalanceSheetDecadeButton.addEventListener(
    "click",
    handleExportBalanceSheetDecade
  );
}
if (exportProfitLossDecadeButton) {
  exportProfitLossDecadeButton.addEventListener(
    "click",
    handleExportProfitLossDecade
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
if (insuranceDetailButton) {
  insuranceDetailButton.addEventListener("click", () => {
    setActivePage("insurance-detail");
  });
}
if (addInsuranceScheduleRowButton) {
  addInsuranceScheduleRowButton.addEventListener("click", () => {
    createInsuranceScheduleRow();
    persistInsuranceScheduleRows();
  });
}
if (pensionDetailButton) {
  pensionDetailButton.addEventListener("click", () => {
    setActivePage("pension-detail");
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
assetDataInput.addEventListener("input", markImportDirty);
summaryDataInput.addEventListener("input", markImportDirty);

persistInputs.forEach((input) => {
  input.addEventListener("input", persistInputsToStorage);
});

loadPersistedInputs();
loadBondRows();
loadInsuranceScheduleRows();
loadPensionPlanRows();
loadPensionChangeRows();
render();

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
