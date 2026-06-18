// ---- Config ----
const SHEET_ID = '1FpuctduIy7dKMPjjnNNRCeoUurotIdObdFsMQd4XvMY';

function getSheet(tabName) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(tabName);
}

// ---- API Entry Points ----
function doGet(e) {
  const action = e.parameter.action;
  let result;

  switch (action) {
    case 'getFoods':
      result = getFoods();
      break;
    case 'getToday':
      result = getToday();
      break;
    case 'getTodayLog':
      result = getTodayLog();
      break;
    case 'getWeightHistory':
      result = getWeightHistory();
      break;
    default:
      result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = body.action;
  let result;

  switch (action) {
    case 'addFood':
      result = addFood(body);
      break;
    case 'logFood':
      result = logFood(body);
      break;
    case 'logLiquid':
      result = logLiquid(body);
      break;
    case 'seedFoods':
      result = seedFoods();
      break;
    case 'logBodyweight':
      result = logBodyweight(body);
      break;
    default:
      result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- getFoods ----
function getFoods() {
  const sheet = getSheet('Foods');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const foods = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    foods.push({
      id: row[0],
      name: row[1],
      servingLabel: row[2],
      protein: row[3],
      calories: row[4],
      carbs: row[5],
      fat: row[6],
      sugar: row[7],
      sodium: row[8],
      dateAdded: row[9]
    });
  }
  return foods;
}

// ---- addFood ----
function addFood(body) {
  const sheet = getSheet('Foods');
  const foods = getFoods();
  let maxNum = 0;
  foods.forEach(f => {
    const n = parseInt(f.id.replace('f', ''));
    if (n > maxNum) maxNum = n;
  });
  const newId = 'f' + String(maxNum + 1).padStart(3, '0');

  const containers = body.containers || 1;
  const divisor = body.isBatch ? containers : 1;

  const protein = Math.round((body.protein || 0) / divisor);
  const calories = Math.round((body.calories || 0) / divisor);
  const carbs = Math.round((body.carbs || 0) / divisor);
  const fat = Math.round((body.fat || 0) / divisor);
  const sugar = Math.round((body.sugar || 0) / divisor);
  const sodium = Math.round((body.sodium || 0) / divisor);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  sheet.appendRow([
    newId,
    body.name,
    body.servingLabel || '1 serving',
    protein, calories, carbs, fat, sugar, sodium,
    today
  ]);

  return {
    success: true,
    food: { id: newId, name: body.name, servingLabel: body.servingLabel || '1 serving',
            protein, calories, carbs, fat, sugar, sodium, dateAdded: today }
  };
}

// ---- logFood ----
function logFood(body) {
  const foods = getFoods();
  const food = foods.find(f => f.id === body.foodId);
  if (!food) return { error: 'Food not found: ' + body.foodId };

  const servings = body.servings || 1;
  const now = new Date();
  const timestamp = now.toISOString();
  const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const protein = food.protein * servings;
  const calories = food.calories * servings;
  const carbs = food.carbs * servings;
  const fat = food.fat * servings;
  const sugar = food.sugar * servings;
  const sodium = food.sodium * servings;

  const sheet = getSheet('Log');
  sheet.appendRow([
    timestamp, date, 'food',
    food.name + (servings > 1 ? ' x' + servings : ''),
    servings, protein, calories, carbs, fat, sugar, sodium
  ]);

  rebuildDailySummary(date);

  return {
    success: true,
    entry: { timestamp, date, type: 'food', item: food.name, servings, protein, calories, carbs, fat, sugar, sodium }
  };
}

// ---- logLiquid ----
function logLiquid(body) {
  const type = body.type; // water, coffee, coors, bourbon
  const amount = body.amount || 1;
  const now = new Date();
  const timestamp = now.toISOString();
  const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let item, protein = 0, calories = 0, carbs = 0, fat = 0, sugar = 0, sodium = 0;

  switch (type) {
    case 'water':
      item = 'Water (' + amount + ' oz)';
      break;
    case 'coffee':
      item = 'Coffee w/ Oat Milk & Syrup';
      protein = 0; calories = 75; carbs = 17; fat = 1; sugar = 15; sodium = 0;
      break;
    case 'coors':
      item = 'Coors Light';
      protein = 1; calories = 102; carbs = 5; fat = 0; sugar = 0; sodium = 14;
      break;
    case 'bourbon':
      item = 'Bourbon';
      protein = 0; calories = 97; carbs = 0; fat = 0; sugar = 0; sodium = 0;
      break;
    default:
      return { error: 'Unknown liquid type: ' + type };
  }

  const sheet = getSheet('Log');
  sheet.appendRow([
    timestamp, date, type, item, amount,
    protein, calories, carbs, fat, sugar, sodium
  ]);

  rebuildDailySummary(date);

  return {
    success: true,
    entry: { timestamp, date, type, item, servings: amount, protein, calories, carbs, fat, sugar, sodium }
  };
}

// ---- getToday ----
function getToday() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sheet = getSheet('Log');
  const data = sheet.getDataRange().getValues();

  let protein = 0, calories = 0, carbs = 0, fat = 0, sugar = 0, sodium = 0;
  let waterOz = 0, coffeeCount = 0, coorsCount = 0, bourbonCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] !== date && Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') !== date) continue;

    const type = row[2];
    protein += row[5] || 0;
    calories += row[6] || 0;
    carbs += row[7] || 0;
    fat += row[8] || 0;
    sugar += row[9] || 0;
    sodium += row[10] || 0;

    if (type === 'water') waterOz += row[4] || 0;
    if (type === 'coffee') coffeeCount += 1;
    if (type === 'coors') coorsCount += 1;
    if (type === 'bourbon') bourbonCount += 1;
  }

  return { date, protein, calories, carbs, fat, sugar, sodium, waterOz, coffeeCount, coorsCount, bourbonCount };
}

// ---- getTodayLog ----
function getTodayLog() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sheet = getSheet('Log');
  const data = sheet.getDataRange().getValues();
  const entries = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = row[1] instanceof Date
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : row[1];
    if (rowDate !== date) continue;

    entries.push({
      timestamp: row[0],
      date: rowDate,
      type: row[2],
      item: row[3],
      servings: row[4],
      protein: row[5] || 0,
      calories: row[6] || 0,
      carbs: row[7] || 0,
      fat: row[8] || 0,
      sugar: row[9] || 0,
      sodium: row[10] || 0
    });
  }

  return entries;
}

// ---- rebuildDailySummary ----
function rebuildDailySummary(date) {
  if (!date) {
    date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const logSheet = getSheet('Log');
  const logData = logSheet.getDataRange().getValues();

  let protein = 0, calories = 0, waterOz = 0, coffeeCount = 0, coorsCount = 0, bourbonCount = 0;

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const rowDate = row[1] instanceof Date
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : row[1];
    if (rowDate !== date) continue;

    protein += row[5] || 0;
    calories += row[6] || 0;
    const type = row[2];
    if (type === 'water') waterOz += row[4] || 0;
    if (type === 'coffee') coffeeCount += 1;
    if (type === 'coors') coorsCount += 1;
    if (type === 'bourbon') bourbonCount += 1;
  }

  // Check if workout exists for this date
  const workoutSheet = getSheet('Workouts');
  const workoutData = workoutSheet.getDataRange().getValues();
  let workoutDone = 'No';
  for (let i = 1; i < workoutData.length; i++) {
    const rowDate = workoutData[i][0] instanceof Date
      ? Utilities.formatDate(workoutData[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : workoutData[i][0];
    if (rowDate === date) { workoutDone = 'Yes'; break; }
  }

  // Find or create the row for this date in DailySummary
  const summarySheet = getSheet('DailySummary');
  const summaryData = summarySheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < summaryData.length; i++) {
    const rowDate = summaryData[i][0] instanceof Date
      ? Utilities.formatDate(summaryData[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : summaryData[i][0];
    if (rowDate === date) { foundRow = i + 1; break; }
  }

  const rowData = [date, protein, calories, waterOz, coffeeCount, coorsCount, bourbonCount, workoutDone, ''];

  if (foundRow > 0) {
    const range = summarySheet.getRange(foundRow, 1, 1, 9);
    // Preserve bodyweight if already entered
    const existing = range.getValues()[0];
    rowData[8] = existing[8] || '';
    range.setValues([rowData]);
  } else {
    summarySheet.appendRow(rowData);
  }
}

// ---- logBodyweight ----
function logBodyweight(body) {
  const weight = parseFloat(body.weight);
  if (!weight || weight < 50 || weight > 500) return { error: 'Invalid weight' };

  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const summarySheet = getSheet('DailySummary');
  const data = summarySheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : data[i][0];
    if (rowDate === date) { foundRow = i + 1; break; }
  }

  if (foundRow > 0) {
    summarySheet.getRange(foundRow, 9).setValue(weight);
  } else {
    summarySheet.appendRow([date, 0, 0, 0, 0, 0, 0, 'No', weight]);
  }

  return { success: true, date: date, weight: weight };
}

// ---- getWeightHistory ----
function getWeightHistory() {
  const summarySheet = getSheet('DailySummary');
  const data = summarySheet.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const weight = row[8];
    if (!weight) continue;
    const date = row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : row[0];
    history.push({ date: date, weight: parseFloat(weight) });
  }

  history.sort((a, b) => a.date.localeCompare(b.date));
  return history;
}

// ---- Seed Foods ----
function seedFoods() {
  const sheet = getSheet('Foods');
  const existing = sheet.getDataRange().getValues();
  if (existing.length > 1) return { success: true, message: 'Foods already seeded', count: existing.length - 1 };

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const seeds = [
    ['f001', 'Beef Egg Roll Bowl',          '1 container', 16, 230, 14, 12,  5,   0, today],
    ['f002', 'Turkey with Tomato',           '1 container', 26, 160,  4,  6,  2,   0, today],
    ['f003', 'Steamed Broccoli',             '1 serving',    3,  60, 11,  1,  2,   0, today],
    ['f004', 'ONE Bar Reeses PB',            '1 bar',       18, 220, 21,  8,  3,   0, today],
    ['f005', '2-Egg Veggie Omelette',        '1 omelette',  12, 220,  8, 14,  4,   0, today],
    ['f006', 'Coffee w/ Oat Milk & Syrup',   '1 cup',        0,  75, 17,  1, 15,   0, today],
  ];

  seeds.forEach(row => sheet.appendRow(row));
  return { success: true, message: 'Seeded ' + seeds.length + ' foods' };
}
