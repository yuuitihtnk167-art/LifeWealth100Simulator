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
const addBondRowButton = document.getElementById("addBondRow");
const sortBondRowsButton = document.getElementById("sortBondRows");
const bondAverageRate = document.getElementById("bondAverageRate");
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
const balanceNissayInput = document.getElementById("balanceNissay");
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
const contribNissayInput = document.getElementById("contribNissay");
const endAgeStocksInput = document.getElementById("endAgeStocks");
const endAgeFundsInput = document.getElementById("endAgeFunds");
const endAgeBondsInput = document.getElementById("endAgeBonds");
const endAgeInsuranceInput = document.getElementById("endAgeInsurance");
const endAgeUsdInput = document.getElementById("endAgeUsd");
const endAgeDcInput = document.getElementById("endAgeDc");
const endAgeNissayInput = document.getElementById("endAgeNissay");
const investmentTotal = document.getElementById("investmentTotal");
const cashBalance = document.getElementById("cashBalance");
const investmentContribTotal = document.getElementById("investmentContribTotal");
const investmentAfter = document.getElementById("investmentAfter");
const cashAfter = document.getElementById("cashAfter");
const investmentAlert = document.getElementById("investmentAlert");
const tabButtons = Array.from(document.querySelectorAll(".tab-button"));
const pages = Array.from(document.querySelectorAll(".page"));
const resultValue = document.getElementById("resultValue");
const resultMeta = document.getElementById("resultMeta");
const lastUpdated = document.getElementById("lastUpdated");
let importDirty = false;

const STORAGE_KEY = "lifewealth100.inputs.v1";
const BOND_STORAGE_KEY = "lifewealth100.bonds.v1";
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

function parseNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
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
}) {
  const months = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const investKeys = ["stocks", "funds", "bonds", "insurance", "pension", "other"];

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
      data[key] += data[key] * rate;
    });

    data.cash += cashFlow;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        if (item.category !== "cash") {
          data.cash -= item.amount;
        }
        data[item.category] += item.amount;
      }
    });
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

function findNegativeCashMonth({
  startCash,
  monthlyNetCash,
  retirementMonthlyNetCash,
  postRetirementMonthlyNetCash,
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
    let cashFlow = monthlyNetCash;
    if (monthIndexValue >= retirementAge) {
      cashFlow = retirementMonthlyNetCash;
    }
    if (monthIndexValue >= retirementIncomeEndAge) {
      cashFlow = postRetirementMonthlyNetCash;
    }
    cash += cashFlow;
    cash -= getContributionForMonth(monthIndexValue, contributionSchedule);
    if (cash < 0) {
      return monthIndexValue;
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
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const rows = [];
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const investKeys = ["stocks", "funds", "bonds", "insurance", "pension", "other"];

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
      data[key] += data[key] * rate;
    });

    data.cash += cashFlow;

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        if (item.category !== "cash") {
          data.cash -= item.amount;
        }
        data[item.category] += item.amount;
      }
    });

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
        data.pension +
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
    data.pension +
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
}) {
  const monthlyRate = Math.pow(1 + annualRate, 1 / 12) - 1;
  const rows = [];
  const totalMonths = Math.max(0, monthsRemaining);
  let data = { ...categories };
  const startMonthIndex = monthIndex(startDate);
  const investKeys = ["stocks", "funds", "bonds", "insurance", "pension", "other"];

  let yearStart = { ...data };
  let yearIncome = 0;
  let yearCashIncome = 0;
  let yearInvestmentIncome = 0;
  let yearExpense = 0;
  let yearInvestmentGain = 0;
  let yearContribution = 0;
  let yearContributionByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    pension: 0,
    other: 0,
  };
  let yearGainByCategory = {
    stocks: 0,
    funds: 0,
    bonds: 0,
    insurance: 0,
    pension: 0,
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

    let investmentIncome = 0;
    investKeys.forEach((key) => {
      const rate = categoryRates?.[key] ?? monthlyRate;
      const gain = data[key] * rate;
      data[key] += gain;
      yearGainByCategory[key] += gain;
      investmentIncome += gain;
    });
    yearInvestmentIncome += investmentIncome;

    data.cash += monthlyIncome - monthlyExpense;
    yearIncome += monthlyIncome + investmentIncome;
    yearCashIncome += monthlyIncome;
    yearExpense += monthlyExpense;
    if (phase === "work") {
      yearWorkMonths += 1;
    } else if (phase === "retire") {
      yearRetireMonths += 1;
    } else {
      yearPensionMonths += 1;
    }

    contributionSchedule.forEach((item) => {
      if (monthIndexValue < item.endMonthIndex) {
        yearContribution += item.amount;
        if (yearContributionByCategory[item.category] !== undefined) {
          yearContributionByCategory[item.category] += item.amount;
        }
        if (item.category !== "cash") {
          data.cash -= item.amount;
        }
        data[item.category] += item.amount;
      }
    });

    yearInvestmentGain += investmentIncome;
    yearMonths += 1;

    const isYearEnd = monthDate.getMonth() === 11;
    const isFinal = i === totalMonths - 1;
    if (isYearEnd || isFinal) {
      const endDate = addMonths(startDate, i + 1);
      const startTotal = sumCategoryTotal(yearStart);
      const endTotal = sumCategoryTotal(data);
      const netChange = yearIncome - yearExpense;
      const mismatch =
        Math.round(startTotal + netChange) !== Math.round(endTotal);
      rows.push({
        year: monthDate.getFullYear(),
        date: endDate,
        start: { ...yearStart, total: startTotal },
        end: { ...data, total: endTotal },
        income: yearIncome,
        cashIncome: yearCashIncome,
        investmentIncome: yearInvestmentIncome,
        expense: yearExpense,
        netCash: netChange,
        investmentGain: yearInvestmentGain,
        totalChange: endTotal - startTotal,
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
      yearContribution = 0;
      yearContributionByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        pension: 0,
        other: 0,
      };
      yearGainByCategory = {
        stocks: 0,
        funds: 0,
        bonds: 0,
        insurance: 0,
        pension: 0,
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

  const getAmount = (regex) => {
    const idx = mapHeaderIndex(table.headers, regex);
    if (idx === null) {
      return 0;
    }
    return parseAmount(bestRow[idx] || "") || 0;
  };

  const total = getAmount(/合計/);
  return {
    total,
    cash: getAmount(/預金|現金|暗号資産/),
    stocks: getAmount(/株式/),
    funds: getAmount(/投資信託/),
    bonds: getAmount(/債券/),
    insurance: getAmount(/保険/),
    pension: getAmount(/年金/),
    points: getAmount(/ポイント/),
    other: getAmount(/その他/),
  };
}

function buildContributionSchedule(birthDate) {
  if (!birthDate) {
    return [];
  }
  const toEndMonth = (input) =>
    monthIndex(addYears(birthDate, parseNumber(input.value) || 0));

  return [
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
    {
      category: "pension",
      amount: parseNumber(contribDcInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeDcInput),
    },
    {
      category: "pension",
      amount: parseNumber(contribNissayInput.value) || 0,
      endMonthIndex: toEndMonth(endAgeNissayInput),
    },
  ];
}

function buildInitialCategories(summaryBreakdown, currentAssets) {
  if (summaryBreakdown) {
    const total =
      summaryBreakdown.total ||
      summaryBreakdown.cash +
        summaryBreakdown.stocks +
        summaryBreakdown.funds +
        summaryBreakdown.bonds +
        summaryBreakdown.insurance +
        summaryBreakdown.pension +
        summaryBreakdown.points +
        summaryBreakdown.other;
    return {
      cash: summaryBreakdown.cash || 0,
      stocks: summaryBreakdown.stocks || 0,
      funds: summaryBreakdown.funds || 0,
      bonds: summaryBreakdown.bonds || 0,
      insurance: summaryBreakdown.insurance || 0,
      pension: summaryBreakdown.pension || 0,
      points: summaryBreakdown.points || 0,
      other: summaryBreakdown.other || 0,
      total: total || 0,
    };
  }

  const stocks = parseNumber(balanceStocksInput.value) || 0;
  const funds = parseNumber(balanceFundsInput.value) || 0;
  const bonds = parseNumber(balanceBondsInput.value) || 0;
  const insurance = parseNumber(balanceInsuranceInput.value) || 0;
  const pension =
    (parseNumber(balanceDcInput.value) || 0) +
    (parseNumber(balanceNissayInput.value) || 0);
  const cash =
    (parseNumber(currentAssets) || 0) -
    (stocks + funds + bonds + insurance + pension);

  return {
    cash,
    stocks,
    funds,
    bonds,
    insurance,
    pension,
    points: 0,
    other: 0,
    total: currentAssets || 0,
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
      toCsvNumber(row.pension),
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

  const rows = lines.map((line) => line.split(delimiter).map((cell) => cell.trim()));
  const headers = rows[0];
  const dataRows = rows.slice(1);
  return { headers, dataRows };
}

function mapHeaderIndex(headers, regex) {
  const index = headers.findIndex((header) => regex.test(header));
  return index === -1 ? null : index;
}

function sumInvestmentsFromList(text) {
  const table = parseTable(text);
  if (!table) {
    return sumInvestmentsFromText(text);
  }

  const typeIndex = mapHeaderIndex(table.headers, /資産区分/);
  const nameIndex = mapHeaderIndex(table.headers, /名称/);
  const amountIndex = mapHeaderIndex(table.headers, /金額.*円/);
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
    nissay: 0,
  };

  table.dataRows.forEach((row) => {
    const type = row[typeIndex] || "";
    const name = nameIndex !== null ? row[nameIndex] || "" : "";
    const amount = row[amountIndex] ? parseAmount(row[amountIndex]) : null;
    if (amount === null) {
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
      if (/ニッセイみらいのカタチ/.test(name)) {
        totals.nissay += amount;
        return;
      }
      if (/DC|確定拠出|ベネフィット|あおぞら/.test(name)) {
        totals.dc += amount;
        return;
      }
    }
  });

  const hasAny =
    totals.stocks ||
    totals.funds ||
    totals.bonds ||
    totals.insurance ||
    totals.usd ||
    totals.dc ||
    totals.nissay;
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
    nissay: 0,
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
      if (/ニッセイみらいのカタチ/.test(line)) {
        totals.nissay += amount;
        matched = true;
        return;
      }
      if (/DC|確定拠出|ベネフィット|あおぞら/.test(line)) {
        totals.dc += amount;
        matched = true;
        return;
      }
    }
  });

  return matched ? totals : null;
}

function parseAmount(value) {
  const match = value.match(/-?\d[\d,]*/);
  if (!match) {
    return null;
  }
  const normalized = match[0].replace(/,/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
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

function persistBondRows() {
  if (!bondTableBody) {
    return;
  }
  const rows = Array.from(bondTableBody.querySelectorAll("tr")).map((row) => {
    const data = {};
    row.querySelectorAll("input, select").forEach((input) => {
      data[input.dataset.key] = input.value;
    });
    return data;
  });
  try {
    localStorage.setItem(BOND_STORAGE_KEY, JSON.stringify(rows));
  } catch {
    // Ignore storage failures.
  }
  updateBondAverageRate();
}

function loadBondRows() {
  if (!bondTableBody) {
    return;
  }
  bondTableBody.innerHTML = "";
  try {
    const raw = localStorage.getItem(BOND_STORAGE_KEY);
    if (!raw) {
      createBondRow();
      updateBondAverageRate();
      return;
    }
    const rows = JSON.parse(raw);
    if (!Array.isArray(rows) || rows.length === 0) {
      createBondRow();
      updateBondAverageRate();
      return;
    }
    rows.forEach((row) => createBondRow(row));
    updateBondAverageRate();
  } catch {
    createBondRow();
    updateBondAverageRate();
  }
}

function sortBondRowsByMaturity() {
  if (!bondTableBody) {
    return;
  }
  const rows = Array.from(bondTableBody.querySelectorAll("tr")).map(
    (row, index) => {
      const data = {};
      row.querySelectorAll("input, select").forEach((input) => {
        data[input.dataset.key] = input.value;
      });
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
  const inputs = safeParseJson(localStorage.getItem(STORAGE_KEY), {});
  const bonds = safeParseJson(localStorage.getItem(BOND_STORAGE_KEY), []);
  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    inputs,
    bonds,
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
  const bonds = Array.isArray(data.bonds) ? data.bonds : null;
  if (!inputs && !bonds) {
    window.alert("同期データに読み込める内容がありません。");
    return false;
  }
  try {
    if (inputs) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(inputs));
    }
    if (bonds) {
      localStorage.setItem(BOND_STORAGE_KEY, JSON.stringify(bonds));
    }
  } catch {
    window.alert("同期データの保存に失敗しました。");
    return false;
  }
  loadPersistedInputs();
  loadBondRows();
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

  const amountIndex = table.headers.findIndex((header) => /金額.*円/.test(header));
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
  const totalIndex = table.headers.findIndex((header) => /合計/.test(header));
  if (totalIndex === -1) {
    return null;
  }

  let best = null;

  table.dataRows.forEach((row, index) => {
    const total = row[totalIndex] ? parseAmount(row[totalIndex]) : null;
    if (total === null) {
      return;
    }
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
    dc: getAmount(/確定拠出|年金/),
    nissay: getAmount(/ニッセイみらいのカタチ/),
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
    balanceNissayInput.value = Math.round(investmentTotals.nissay);
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
    contributionSchedule: buildContributionSchedule(context.birthDate),
    categories: context.categories,
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
    pension: defaultMonthlyRate,
    other: defaultMonthlyRate,
  };
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
  const summaryBreakdown = getSummaryBreakdown(summaryDataInput.value);
  const initial = buildInitialCategories(summaryBreakdown, currentAssets);
  const categories = {
    cash: initial.cash,
    stocks: initial.stocks,
    funds: initial.funds,
    bonds: initial.bonds,
    insurance: initial.insurance,
    pension: initial.pension,
    points: initial.points,
    other: initial.other,
  };

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
    categories,
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
    contributionSchedule: buildContributionSchedule(context.birthDate),
    categories: context.categories,
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
  statementRows.forEach((row) => {
    const option = document.createElement("option");
    option.value = row.year;
    option.textContent =
      row.months < 12 ? `${row.year}年（${row.months}か月）` : `${row.year}年`;
    statementYearSelect.appendChild(option);
  });
  const matched = statementRows.find((row) => row.year === selectedYear);
  const fallbackYear = statementRows[statementRows.length - 1].year;
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

function buildBalanceSheetCsvLine(row, birthDate) {
  const periodStartDate = getPeriodStartDate(row.date, row.months);
  const startTotalExpression = buildSignedExpression([
    row.start.cash,
    row.start.stocks,
    row.start.funds,
    row.start.bonds,
    row.start.insurance,
    row.start.pension,
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
  const incomeExpression = `${cashIncomeExpression}+${toCsvNumber(
    row.investmentIncome
  )}`;
  const netExpression = `(${incomeExpression})-(${expenseExpression})`;
  const cashEndExpression = `${toCsvNumber(row.start.cash)}+(${cashIncomeExpression})-(${expenseExpression})-${toCsvNumber(
    row.contributions
  )}`;
  const endTotalExpression = `${toCsvNumber(row.start.total)}+(${netExpression})`;
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
    csvCellWithFormula(row.start.pension, `${toCsvNumber(row.start.pension)}`),
    csvCellWithFormula(row.start.points, `${toCsvNumber(row.start.points)}`),
    csvCellWithFormula(row.start.other, `${toCsvNumber(row.start.other)}`),
    csvCellWithFormula(row.end.total, endTotalExpression),
    csvCellWithFormula(row.end.cash, cashEndExpression),
    csvCellWithFormula(
      row.end.stocks,
      buildSignedExpression([
        row.start.stocks,
        row.contributionsByCategory.stocks,
        row.gainsByCategory.stocks,
      ])
    ),
    csvCellWithFormula(
      row.end.funds,
      buildSignedExpression([
        row.start.funds,
        row.contributionsByCategory.funds,
        row.gainsByCategory.funds,
      ])
    ),
    csvCellWithFormula(
      row.end.bonds,
      buildSignedExpression([
        row.start.bonds,
        row.contributionsByCategory.bonds,
        row.gainsByCategory.bonds,
      ])
    ),
    csvCellWithFormula(
      row.end.insurance,
      buildSignedExpression([
        row.start.insurance,
        row.contributionsByCategory.insurance,
        row.gainsByCategory.insurance,
      ])
    ),
    csvCellWithFormula(
      row.end.pension,
      buildSignedExpression([
        row.start.pension,
        row.contributionsByCategory.pension,
        row.gainsByCategory.pension,
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
        row.gainsByCategory.other,
      ])
    ),
  ].join(",");
}

function buildProfitLossCsv(row, birthDate) {
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
    row.start.pension,
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
  const incomeExpression = `${cashIncomeExpression}+${toCsvNumber(
    row.investmentIncome
  )}`;
  const netExpression = `(${incomeExpression})-(${expenseExpression})`;
  const investmentGainExpression = buildSignedExpression([
    row.gainsByCategory.stocks,
    row.gainsByCategory.funds,
    row.gainsByCategory.bonds,
    row.gainsByCategory.insurance,
    row.gainsByCategory.pension,
    row.gainsByCategory.other,
  ]);
  const endTotalExpression = `${toCsvNumber(row.start.total)}+(${netExpression})`;
  return [
    row.year,
    formatAgeYears(birthDate, periodStartDate),
    row.months,
    csvCellWithFormula(row.income, incomeExpression),
    csvCellWithFormula(row.expense, expenseExpression),
    csvCellWithFormula(row.netCash, netExpression),
    csvCellWithFormula(row.investmentGain, investmentGainExpression),
    csvCellWithFormula(row.totalChange, netExpression),
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
  const contributionSchedule = buildContributionSchedule(birthDate);
  const investmentBalanceTotal =
    (parseNumber(balanceStocksInput.value) || 0) +
    (parseNumber(balanceFundsInput.value) || 0) +
    (parseNumber(balanceBondsInput.value) || 0) +
    (parseNumber(balanceInsuranceInput.value) || 0) +
    (parseNumber(balanceUsdInput.value) || 0) +
    (parseNumber(balanceDcInput.value) || 0) +
    (parseNumber(balanceNissayInput.value) || 0);
  const investmentContributionTotal =
    (parseNumber(contribStocksInput.value) || 0) +
    (parseNumber(contribFundsInput.value) || 0) +
    (parseNumber(contribBondsInput.value) || 0) +
    (parseNumber(contribInsuranceInput.value) || 0) +
    (parseNumber(contribUsdInput.value) || 0) +
    (parseNumber(contribDcInput.value) || 0) +
    (parseNumber(contribNissayInput.value) || 0);

  const today = new Date();
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
      const cashNow = currentAssets - investmentBalanceTotal;
      const cashAfterValue = cashNow - investmentContributionTotal;
      const summaryBreakdown = getSummaryBreakdown(summaryDataInput.value);
      const warningCashNow =
        summaryBreakdown && Number.isFinite(summaryBreakdown.cash)
          ? summaryBreakdown.cash
          : cashNow;
      const warningCashAfter = warningCashNow - investmentContributionTotal;
      investmentTotal.textContent = yenFormatter.format(
        Math.round(investmentBalanceTotal)
      );
      cashBalance.textContent = yenFormatter.format(Math.round(cashNow));
      investmentContribTotal.textContent = yenFormatter.format(
        Math.round(investmentContributionTotal)
      );
      investmentAfter.textContent = yenFormatter.format(
        Math.round(investmentBalanceTotal + investmentContributionTotal)
      );
      cashAfter.textContent = yenFormatter.format(Math.round(cashAfterValue));
      if (investmentAlert) {
        const canForecast =
          hasBirthDate && hasRetirement && monthsRemaining !== null;
        let negativeRow = null;
        if (canForecast) {
          const initial = buildInitialCategories(summaryBreakdown, currentAssets);
          const categories = {
            cash: initial.cash,
            stocks: initial.stocks,
            funds: initial.funds,
            bonds: initial.bonds,
            insurance: initial.insurance,
            pension: initial.pension,
            points: initial.points,
            other: initial.other,
          };
          const rows = simulateAnnualSeries({
            startDate: today,
            monthsRemaining,
            annualRate,
            categoryRates: getSimulationContext()?.categoryRates,
            retirementAge: monthIndex(retirementDate),
            retirementIncomeEndAge: monthIndex(retirementIncomeEndDate),
            monthlyNetCash,
            retirementMonthlyNetCash,
            postRetirementMonthlyNetCash,
            contributionSchedule,
            categories,
          });
          negativeRow = rows.find((row) => row.cash < 0);
        }
        if (warningCashNow < 0) {
          investmentAlert.textContent =
            "現金残高がマイナスです。投資額を減らすか収入を増やしてください。";
          investmentAlert.hidden = false;
        } else if (negativeRow && birthDate) {
          const ageMonths = fullMonthsBetween(birthDate, negativeRow.date);
          const ageYears = Math.floor(ageMonths / 12);
          const ageRemain = ageMonths % 12;
          investmentAlert.textContent = `将来、現金残高がマイナスになります（${ageYears}歳${ageRemain}か月）。投資額を減らすか収入を増やしてください。`;
          investmentAlert.hidden = false;
        } else if (warningCashAfter < 0) {
          investmentAlert.textContent =
            "積立後の現金残高がマイナスです。積立額を訂正してください。";
          investmentAlert.hidden = false;
        } else {
          investmentAlert.hidden = true;
        }
      }
    } else {
      investmentTotal.textContent = "-";
      cashBalance.textContent = "-";
      investmentContribTotal.textContent = "-";
      investmentAfter.textContent = "-";
      cashAfter.textContent = "-";
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
  balanceNissayInput,
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
  contribNissayInput,
  endAgeStocksInput,
  endAgeFundsInput,
  endAgeBondsInput,
  endAgeInsuranceInput,
  endAgeUsdInput,
  endAgeDcInput,
  endAgeNissayInput,
].forEach((input) => {
  input.addEventListener("input", render);
  input.addEventListener("input", persistInputsToStorage);
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
assetDataInput.addEventListener("input", markImportDirty);
summaryDataInput.addEventListener("input", markImportDirty);

persistInputs.forEach((input) => {
  input.addEventListener("input", persistInputsToStorage);
});

loadPersistedInputs();
loadBondRows();
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
