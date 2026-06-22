// ---- Config ----
const SHEET_ID = '1FpuctduIy7dKMPjjnNNRCeoUurotIdObdFsMQd4XvMY';
const TZ = 'America/New_York';

function getSheet(tabName) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(tabName);
}

function getSheetSuffix(user) { return user === 'michele' ? 'Michele' : ''; }
function getSheetName(base, user) { return base + getSheetSuffix(user); }

// ---- API Entry Points ----
function doGet(e) {
  const action = e.parameter.action;
  const user   = (e.parameter.user || 'aj').toLowerCase();
  let result;

  switch (action) {
    case 'getFoods':
      result = getFoods();
      break;
    case 'getToday':
      result = getToday(e.parameter.date, user);
      break;
    case 'getTodayLog':
      result = getTodayLog(e.parameter.date, user);
      break;
    case 'getDaySummary':
      result = getDaySummary(e.parameter.date, user);
      break;
    case 'getWeightHistory':
      result = getWeightHistory(user);
      break;
    case 'getYesterday':
      result = getYesterday(user);
      break;
    case 'getSummaryHistory':
      result = getSummaryHistory(user);
      break;
    case 'getHomeData':
      result = getHomeData(e.parameter.date, user);
      break;
    case 'getNextWorkout':
      result = getNextWorkout();
      break;
    case 'getExerciseLibrary':
      result = getExerciseLibrary();
      break;
    case 'getWorkoutHistory':
      result = getWorkoutHistory(user);
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
  const user   = (body.user || 'aj').toLowerCase();
  let result;

  switch (action) {
    case 'addFood':
      result = addFood(body);
      break;
    case 'logFood':
      result = logFood(body, user);
      break;
    case 'logLiquid':
      result = logLiquid(body, user);
      break;
    case 'seedFoods':
      result = seedFoods();
      break;
    case 'logBodyweight':
      result = logBodyweight(body, user);
      break;
    case 'buildRecipe':
      result = buildRecipe(body);
      break;
    case 'logFoodDirect':
      result = logFoodDirect(body, user);
      break;
    case 'getCoachAdvice':
      result = getCoachAdvice(body, user);
      break;
    case 'deleteLogEntry':
      result = deleteLogEntry(body, user);
      break;
    case 'logWorkoutSession':
      result = logWorkoutSession(body, user);
      break;
    case 'setupWorkoutSheets':
      result = setupWorkoutSheets();
      break;
    case 'syncExercises':
      result = syncExercises();
      break;
    case 'logActivity':
      result = logActivity(body);
      break;
    case 'setupMicheleSheets':
      result = setupMicheleSheets();
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

  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

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
function logFood(body, user) {
  user = (user || 'aj').toLowerCase();
  const foods = getFoods();
  const food = foods.find(f => f.id === body.foodId);
  if (!food) return { error: 'Food not found: ' + body.foodId };

  const servings = body.servings || 1;
  const now = new Date();
  const timestamp = now.toISOString();
  const date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');

  const protein = food.protein * servings;
  const calories = food.calories * servings;
  const carbs = food.carbs * servings;
  const fat = food.fat * servings;
  const sugar = food.sugar * servings;
  const sodium = food.sodium * servings;

  const sheet = getSheet(getSheetName('Log', user));
  sheet.appendRow([
    timestamp, date, 'food',
    food.name + (servings > 1 ? ' x' + servings : ''),
    servings, protein, calories, carbs, fat, sugar, sodium
  ]);

  rebuildDailySummary(date, user);

  return {
    success: true,
    entry: { timestamp, date, type: 'food', item: food.name, servings, protein, calories, carbs, fat, sugar, sodium }
  };
}

// ---- logLiquid ----
function logLiquid(body, user) {
  user = (user || 'aj').toLowerCase();
  const type   = body.type; // water, coffee, coors, bourbon, mimosa, redwine, surfside, titos, martini
  const amount = body.amount || 1;
  const now    = new Date();
  const timestamp = now.toISOString();
  const date   = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');

  let item, protein = 0, calories = 0, carbs = 0, fat = 0, sugar = 0, sodium = 0;

  switch (type) {
    case 'water':
      item = 'Water (' + amount + ' oz)';
      break;
    case 'coffee':
      item = 'Coffee w/ Oat Milk & Syrup';
      if (user === 'michele') {
        calories = 43; carbs = 9; fat = 1; sugar = 7;
      } else {
        calories = 75; carbs = 17; fat = 1; sugar = 15;
      }
      break;
    case 'coors':
      item = 'Coors Light';
      protein = 1; calories = 102; carbs = 5; fat = 0; sugar = 0; sodium = 14;
      break;
    case 'bourbon':
      item = 'Bourbon';
      protein = 0; calories = 97; carbs = 0; fat = 0; sugar = 0; sodium = 0;
      break;
    case 'mimosa':
      item = 'Mimosa';
      protein = 1; calories = 135; carbs = 16; fat = 0; sugar = 10; sodium = 5;
      break;
    case 'redwine':
      item = 'Red Wine (Can)';
      protein = 0; calories = 200; carbs = 6; fat = 0; sugar = 1; sodium = 0;
      break;
    case 'surfside':
      item = 'Surfside Iced Tea';
      protein = 0; calories = 100; carbs = 2; fat = 0; sugar = 1; sodium = 0;
      break;
    case 'titos':
      item = "Tito's & Soda";
      protein = 0; calories = 98; carbs = 0; fat = 0; sugar = 0; sodium = 0;
      break;
    case 'martini':
      item = 'Dirty Martini';
      protein = 0; calories = 180; carbs = 2; fat = 0; sugar = 0; sodium = 200;
      break;
    default:
      return { error: 'Unknown liquid type: ' + type };
  }

  const sheet = getSheet(getSheetName('Log', user));
  sheet.appendRow([
    timestamp, date, type, item, amount,
    protein, calories, carbs, fat, sugar, sodium
  ]);

  rebuildDailySummary(date, user);

  return {
    success: true,
    entry: { timestamp, date, type, item, servings: amount, protein, calories, carbs, fat, sugar, sodium }
  };
}

// ---- getToday ----
function getToday(dateParam, user) {
  user = (user || 'aj').toLowerCase();
  const date  = dateParam || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const sheet = getSheet(getSheetName('Log', user));
  const data  = sheet.getDataRange().getValues();

  let protein = 0, calories = 0, carbs = 0, fat = 0, sugar = 0, sodium = 0;
  let waterOz = 0, coffeeCount = 0;
  let coorsCount = 0, bourbonCount = 0, mimosaCount = 0;
  let redwineCount = 0, surfsideCount = 0, titosCount = 0, martiniCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] !== date && Utilities.formatDate(new Date(row[1]), TZ, 'yyyy-MM-dd') !== date) continue;

    const type = row[2];
    protein  += row[5] || 0;
    calories += row[6] || 0;
    carbs    += row[7] || 0;
    fat      += row[8] || 0;
    sugar    += row[9] || 0;
    sodium   += row[10] || 0;

    if (type === 'water')    waterOz      += row[4] || 0;
    if (type === 'coffee')   coffeeCount  += 1;
    if (type === 'coors')    coorsCount   += 1;
    if (type === 'bourbon')  bourbonCount += 1;
    if (type === 'mimosa')   mimosaCount  += 1;
    if (type === 'redwine')  redwineCount += 1;
    if (type === 'surfside') surfsideCount+= 1;
    if (type === 'titos')    titosCount   += 1;
    if (type === 'martini')  martiniCount += 1;
  }

  return { date, protein, calories, carbs, fat, sugar, sodium, waterOz, coffeeCount,
           coorsCount, bourbonCount, mimosaCount,
           redwineCount, surfsideCount, titosCount, martiniCount };
}

// ---- getTodayLog ----
function getTodayLog(dateParam, user) {
  user = (user || 'aj').toLowerCase();
  const date = dateParam || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const sheet = getSheet(getSheetName('Log', user));
  const data = sheet.getDataRange().getValues();
  const entries = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = row[1] instanceof Date
      ? Utilities.formatDate(row[1], TZ, 'yyyy-MM-dd')
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
function rebuildDailySummary(date, user) {
  user = (user || 'aj').toLowerCase();
  if (!date) date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  const logSheet = getSheet(getSheetName('Log', user));
  const logData  = logSheet.getDataRange().getValues();

  let protein = 0, calories = 0, waterOz = 0, coffeeCount = 0;
  let coorsCount = 0, bourbonCount = 0, mimosaCount = 0;
  let redwineCount = 0, surfsideCount = 0, titosCount = 0, martiniCount = 0;

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const rowDate = row[1] instanceof Date ? Utilities.formatDate(row[1], TZ, 'yyyy-MM-dd') : row[1];
    if (rowDate !== date) continue;
    protein  += row[5] || 0;
    calories += row[6] || 0;
    const type = row[2];
    if (type === 'water')    waterOz      += row[4] || 0;
    if (type === 'coffee')   coffeeCount  += 1;
    if (type === 'coors')    coorsCount   += 1;
    if (type === 'bourbon')  bourbonCount += 1;
    if (type === 'mimosa')   mimosaCount  += 1;
    if (type === 'redwine')  redwineCount += 1;
    if (type === 'surfside') surfsideCount+= 1;
    if (type === 'titos')    titosCount   += 1;
    if (type === 'martini')  martiniCount += 1;
  }

  const workoutSheetName = user === 'michele' ? 'WorkoutsMichele' : 'Workouts';
  const workoutSheet = getSheet(workoutSheetName);
  let workoutDone = 'No';
  if (workoutSheet) {
    const workoutData = workoutSheet.getDataRange().getValues();
    for (let i = 1; i < workoutData.length; i++) {
      const rowDate = workoutData[i][0] instanceof Date
        ? Utilities.formatDate(workoutData[i][0], TZ, 'yyyy-MM-dd') : workoutData[i][0];
      if (rowDate === date) { workoutDone = 'Yes'; break; }
    }
  }

  const summarySheet = getSheet(getSheetName('DailySummary', user));
  const summaryData  = summarySheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < summaryData.length; i++) {
    const rowDate = summaryData[i][0] instanceof Date
      ? Utilities.formatDate(summaryData[i][0], TZ, 'yyyy-MM-dd') : summaryData[i][0];
    if (rowDate === date) { foundRow = i + 1; break; }
  }

  if (user === 'michele') {
    // [date, protein, calories, waterOz, coffeeCount, redwineCount, surfsideCount, titosCount, martiniCount, workoutDone, bodyweight]
    const rowData = [date, protein, calories, waterOz, coffeeCount, redwineCount, surfsideCount, titosCount, martiniCount, workoutDone, ''];
    if (foundRow > 0) {
      const range = summarySheet.getRange(foundRow, 1, 1, 11);
      rowData[10] = range.getValues()[0][10] || '';
      range.setValues([rowData]);
    } else {
      summarySheet.appendRow(rowData);
    }
  } else {
    // [date, protein, calories, waterOz, coffeeCount, coorsCount, bourbonCount, workoutDone, bodyweight, mimosaCount]
    const rowData = [date, protein, calories, waterOz, coffeeCount, coorsCount, bourbonCount, workoutDone, '', mimosaCount];
    if (foundRow > 0) {
      const range = summarySheet.getRange(foundRow, 1, 1, 10);
      rowData[8] = range.getValues()[0][8] || '';
      range.setValues([rowData]);
    } else {
      summarySheet.appendRow(rowData);
    }
  }
}

// ---- getDaySummary ----
function getDaySummary(dateParam, user) {
  user = (user || 'aj').toLowerCase();
  const date   = dateParam || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const totals = getToday(date, user);
  const log    = getTodayLog(date, user);

  const summarySheet = getSheet(getSheetName('DailySummary', user));
  const summaryData  = summarySheet.getDataRange().getValues();
  totals.workoutDone = 'No';
  const wdCol = user === 'michele' ? 9 : 7;
  for (let i = 1; i < summaryData.length; i++) {
    const rowDate = summaryData[i][0] instanceof Date
      ? Utilities.formatDate(summaryData[i][0], TZ, 'yyyy-MM-dd')
      : summaryData[i][0];
    if (rowDate === date) {
      totals.workoutDone = summaryData[i][wdCol] || 'No';
      break;
    }
  }

  return { totals, log };
}

// ---- getYesterday ----
function getYesterday(user) {
  user = (user || 'aj').toLowerCase();
  const now = new Date();
  now.setDate(now.getDate() - 1);
  const date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');

  const logSheet = getSheet(getSheetName('Log', user));
  const logData  = logSheet.getDataRange().getValues();

  let protein = 0, calories = 0, carbs = 0, fat = 0, waterOz = 0, coffeeCount = 0;
  let coorsCount = 0, bourbonCount = 0, mimosaCount = 0;
  let redwineCount = 0, surfsideCount = 0, titosCount = 0, martiniCount = 0;
  let foodItems = [];

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const rowDate = row[1] instanceof Date ? Utilities.formatDate(row[1], TZ, 'yyyy-MM-dd') : row[1];
    if (rowDate !== date) continue;

    const type = row[2];
    protein  += row[5] || 0;
    calories += row[6] || 0;
    carbs    += row[7] || 0;
    fat      += row[8] || 0;

    if (type === 'water')    waterOz      += row[4] || 0;
    if (type === 'coffee')   coffeeCount  += 1;
    if (type === 'coors')    coorsCount   += 1;
    if (type === 'bourbon')  bourbonCount += 1;
    if (type === 'mimosa')   mimosaCount  += 1;
    if (type === 'redwine')  redwineCount += 1;
    if (type === 'surfside') surfsideCount+= 1;
    if (type === 'titos')    titosCount   += 1;
    if (type === 'martini')  martiniCount += 1;
    if (type === 'food')     foodItems.push(row[3]);
  }

  const workoutSheetName = user === 'michele' ? 'WorkoutsMichele' : 'Workouts';
  const workoutSheet = getSheet(workoutSheetName);
  let workedOut = false, workoutSummary = [];
  if (workoutSheet) {
    const workoutData = workoutSheet.getDataRange().getValues();
    for (let i = 1; i < workoutData.length; i++) {
      const rowDate = workoutData[i][0] instanceof Date
        ? Utilities.formatDate(workoutData[i][0], TZ, 'yyyy-MM-dd') : workoutData[i][0];
      if (rowDate === date) {
        workedOut = true;
        workoutSummary.push(user === 'michele' ? workoutData[i][2] : workoutData[i][1] + ': ' + workoutData[i][2]);
      }
    }
  }

  return {
    date, protein, calories, carbs, fat, waterOz, coffeeCount,
    coorsCount, bourbonCount, mimosaCount,
    redwineCount, surfsideCount, titosCount, martiniCount,
    workedOut, workoutSummary, foodItems,
    hasData: protein > 0 || calories > 0 || waterOz > 0
  };
}

// ---- getTodayWorkout ----
function getTodayWorkout() {
  const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const workoutSheet = getSheet('Workouts');
  const data = workoutSheet.getDataRange().getValues();
  let workedOut = false;
  let exercises = [];

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], TZ, 'yyyy-MM-dd')
      : data[i][0];
    if (rowDate === date) {
      workedOut = true;
      exercises.push({
        muscleGroup: data[i][1],
        exercise: data[i][2],
        sets: data[i][3],
        reps: data[i][4],
        weight: data[i][5]
      });
    }
  }

  return { workedOut, exercises };
}

// ---- logBodyweight ----
function logBodyweight(body, user) {
  user = (user || 'aj').toLowerCase();
  const weight = parseFloat(body.weight);
  if (!weight || weight < 50 || weight > 500) return { error: 'Invalid weight' };

  const date         = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const summarySheet = getSheet(getSheetName('DailySummary', user));
  const data         = summarySheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], TZ, 'yyyy-MM-dd') : data[i][0];
    if (rowDate === date) { foundRow = i + 1; break; }
  }

  // Michele bodyweight = col 11 (1-indexed), AJ = col 9
  const bwCol = user === 'michele' ? 11 : 9;
  if (foundRow > 0) {
    summarySheet.getRange(foundRow, bwCol).setValue(weight);
  } else {
    if (user === 'michele') {
      summarySheet.appendRow([date, 0, 0, 0, 0, 0, 0, 0, 0, 'No', weight]);
    } else {
      summarySheet.appendRow([date, 0, 0, 0, 0, 0, 0, 'No', weight, 0]);
    }
  }

  return { success: true, date, weight };
}

// ---- getWeightHistory ----
function getWeightHistory(user) {
  user = (user || 'aj').toLowerCase();
  const summarySheet = getSheet(getSheetName('DailySummary', user));
  const data         = summarySheet.getDataRange().getValues();
  const bwColIdx     = user === 'michele' ? 10 : 8; // 0-indexed
  const history      = [];

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const weight = row[bwColIdx];
    if (!weight) continue;
    const date = row[0] instanceof Date ? Utilities.formatDate(row[0], TZ, 'yyyy-MM-dd') : row[0];
    history.push({ date, weight: parseFloat(weight) });
  }

  history.sort((a, b) => a.date.localeCompare(b.date));
  return history;
}

// ---- Seed / Sync Foods ----
// Run syncFoods() from the Apps Script editor after updating this list to add new items.
// Safe to run anytime — only adds foods that aren't already in the sheet (matched by name).
function syncFoods() {
  const sheet = getSheet('Foods');
  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => r[1].toString().toLowerCase()));

  let maxNum = 0;
  existing.slice(1).forEach(r => {
    const n = parseInt(String(r[0]).replace('f', ''));
    if (!isNaN(n) && n > maxNum) maxNum = n;
  });

  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const catalog = [
    ['Beef Egg Roll Bowl',         '1 container', 16, 230, 14, 12,  5,   0],
    ['Turkey with Tomato',         '1 container', 26, 160,  4,  6,  2,   0],
    ['Steamed Broccoli',           '1 serving',    3,  60, 11,  1,  2,   0],
    ['ONE Bar Reeses PB',          '1 bar',       18, 220, 21,  8,  3,   0],
    ['2-Egg Veggie Omelette',      '1 omelette',  12, 220,  8, 14,  4,   0],
    ['Coffee w/ Oat Milk & Syrup', '1 cup',        0,  75, 17,  1, 15,   0],
    ['Grilled Chicken Breast',     '5 oz',        35, 185,  0,  5,  0,   0],
  ];

  let added = 0;
  catalog.forEach(([name, ...rest]) => {
    if (!existingNames.has(name.toLowerCase())) {
      maxNum++;
      const newId = 'f' + String(maxNum).padStart(3, '0');
      sheet.appendRow([newId, name, ...rest, today]);
      added++;
    }
  });

  return { success: true, message: added > 0 ? 'Added ' + added + ' new food(s)' : 'All foods already present' };
}

function seedFoods() { return syncFoods(); }

// ---- buildRecipe ----
// Handles paste, url, and ordered modes. Cook It mode is client-side only.
function buildRecipe(body) {
  const mode = body.mode;
  const portions = parseInt(body.portions) || 1;
  let text = body.text || '';

  if (mode === 'cook') {
    // Treat home-cooked descriptions the same as paste — AI estimates from description
    if (!text) return { error: 'Describe what you cooked' };
    const result = callClaude(text, 'cook', portions);
    if (result.error) return result;
    if (!result.servingLabel) {
      result.servingLabel = portions > 1 ? '1 of ' + portions + ' containers' : '1 serving';
    }
    if (body.name && body.name.trim() && (!result.name || result.name.length < 3)) {
      result.name = body.name.trim();
    }
    return result;
  }

  if (mode === 'url') {
    const url = body.url;
    if (!url) return { error: 'URL required' };
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      let html = response.getContentText();
      // Strip scripts, styles, and tags; keep readable text
      text = html
        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 8000);
      if (body.name) text = 'Recipe name: ' + body.name + '\n\n' + text;
    } catch (err) {
      return { error: 'Could not fetch URL: ' + err.message };
    }
  }

  if (!text) return { error: 'No text to analyze' };

  const result = callClaude(text, mode, portions);
  if (result.error) return result;

  if (!result.servingLabel) {
    result.servingLabel = (mode !== 'ordered' && portions > 1)
      ? '1 of ' + portions + ' servings'
      : '1 serving';
  }

  if (body.name && body.name.trim() && (!result.name || result.name.length < 3)) {
    result.name = body.name.trim();
  }

  return result;
}

// ---- callClaude ----
function callClaude(text, mode, portions) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'CLAUDE_API_KEY not set in Script Properties. See setup instructions.' };

  let prompt;
  if (mode === 'ordered') {
    prompt = 'I ordered this meal at a restaurant:\n\n"' + text + '"\n\n' +
      'Estimate realistic nutrition macros for this meal based on typical restaurant portions.\n\n' +
      'Return ONLY a valid JSON object with exactly these fields (numbers only, no units in values):\n' +
      '{"name":"meal name","protein":0,"calories":0,"carbs":0,"fat":0,"sugar":0,"sodium":0,' +
      '"confidence":"high|medium|low","notes":"brief explanation"}';
  } else if (mode === 'cook') {
    prompt = 'I cooked this meal at home:\n\n"' + text + '"\n\n' +
      'This was split into ' + portions + ' portion' + (portions !== 1 ? 's' : '') + '. ' +
      'Estimate realistic nutrition macros PER PORTION based on the ingredients described.\n\n' +
      'Return ONLY a valid JSON object with exactly these fields (numbers only, no units in values):\n' +
      '{"name":"meal name","protein":0,"calories":0,"carbs":0,"fat":0,"sugar":0,"sodium":0,' +
      '"confidence":"high|medium|low","notes":"brief explanation of your estimates"}';
  } else {
    prompt = 'Here is a recipe' + (mode === 'url' ? ' extracted from a webpage' : '') + ':\n\n' + text + '\n\n' +
      'This recipe makes ' + portions + ' serving' + (portions !== 1 ? 's' : '') + '. ' +
      'Calculate nutrition macros PER SERVING.\n\n' +
      'Return ONLY a valid JSON object with exactly these fields (numbers only, no units in values):\n' +
      '{"name":"recipe name","protein":0,"calories":0,"carbs":0,"fat":0,"sugar":0,"sodium":0,' +
      '"confidence":"high|medium|low","notes":"brief explanation or caveats"}';
  }

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 600,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    const responseData = JSON.parse(response.getContentText());
    if (responseData.error) return { error: 'Claude API: ' + responseData.error.message };

    const content = responseData.content[0].text;
    const jsonMatch = content.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return { error: 'Could not parse AI response — try again' };

    const parsed = JSON.parse(jsonMatch[0]);
    // Ensure numeric fields
    ['protein','calories','carbs','fat','sugar','sodium'].forEach(k => {
      parsed[k] = Math.round(parseFloat(parsed[k]) || 0);
    });
    return parsed;
  } catch (err) {
    return { error: 'AI analysis failed: ' + err.message };
  }
}

// ---- logFoodDirect ----
// Writes a food entry directly to Log without requiring a Foods library entry.
function logFoodDirect(body, user) {
  user = (user || 'aj').toLowerCase();
  const now = new Date();
  const timestamp = now.toISOString();
  const date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');

  const protein  = parseFloat(body.protein)  || 0;
  const calories = parseFloat(body.calories) || 0;
  const carbs    = parseFloat(body.carbs)    || 0;
  const fat      = parseFloat(body.fat)      || 0;
  const sugar    = parseFloat(body.sugar)    || 0;
  const sodium   = parseFloat(body.sodium)   || 0;
  const name     = body.name || 'Recipe';

  const sheet = getSheet(getSheetName('Log', user));
  sheet.appendRow([timestamp, date, 'food', name, 1, protein, calories, carbs, fat, sugar, sodium]);

  rebuildDailySummary(date, user);

  return {
    success: true,
    entry: { timestamp, date, type: 'food', item: name, servings: 1, protein, calories, carbs, fat, sugar, sodium }
  };
}

// ---- Michele Activity Logging ----
function logActivity(body) {
  const activityType = body.activityType || 'Activity';
  const date         = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const timestamp    = new Date().toISOString();

  const sheet = getSheet('WorkoutsMichele');
  if (!sheet) return { error: 'WorkoutsMichele not found. Run setupMicheleSheets() first.' };
  sheet.appendRow([date, timestamp, activityType, '', '']);

  rebuildDailySummary(date, 'michele');
  return { success: true, activityType };
}

function getMicheleActivityToday(date) {
  const sheet = getSheet('WorkoutsMichele');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => {
    const rowDate = r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]);
    return rowDate === date;
  }).map(r => ({ date: r[0], activityType: r[2] }));
}

function getMicheleRecentActivities() {
  const sheet = getSheet('WorkoutsMichele');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] && r[2]).map(r => ({
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]),
    activityType: r[2]
  })).sort((a, b) => b.date.localeCompare(a.date));
}

// ---- setupMicheleSheets ----
// Run once from Apps Script editor to create Michele's sheet tabs.
function setupMicheleSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName('LogMichele')) {
    const s = ss.insertSheet('LogMichele');
    s.appendRow(['timestamp','date','type','item','servings','protein','calories','carbs','fat','sugar','sodium']);
  }
  if (!ss.getSheetByName('DailySummaryMichele')) {
    const s = ss.insertSheet('DailySummaryMichele');
    s.appendRow(['date','protein','calories','waterOz','coffeeCount','redwineCount','surfsideCount','titosCount','martiniCount','workoutDone','bodyweight']);
  }
  if (!ss.getSheetByName('WorkoutsMichele')) {
    const s = ss.insertSheet('WorkoutsMichele');
    s.appendRow(['date','timestamp','activityType','notes','extra']);
  }

  return { success: true, message: 'Michele\'s sheets created. She\'s ready to go!' };
}

// ---- getSummaryHistory ----
function getSummaryHistory(user) {
  user = (user || 'aj').toLowerCase();
  const sheet = getSheet(getSheetName('DailySummary', user));
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const history = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const date = row[0] instanceof Date ? Utilities.formatDate(row[0], TZ, 'yyyy-MM-dd') : row[0];
    if (user === 'michele') {
      history.push({
        date,
        protein:       row[1]  || 0,
        calories:      row[2]  || 0,
        waterOz:       row[3]  || 0,
        coffeeCount:   row[4]  || 0,
        redwineCount:  row[5]  || 0,
        surfsideCount: row[6]  || 0,
        titosCount:    row[7]  || 0,
        martiniCount:  row[8]  || 0,
        workoutDone:   row[9]  || 'No',
        bodyweight:    row[10] || null
      });
    } else {
      history.push({
        date,
        protein:      row[1] || 0,
        calories:     row[2] || 0,
        waterOz:      row[3] || 0,
        coffeeCount:  row[4] || 0,
        coorsCount:   row[5] || 0,
        bourbonCount: row[6] || 0,
        workoutDone:  row[7] || 'No',
        bodyweight:   row[8] || null,
        mimosaCount:  row[9] || 0
      });
    }
  }

  history.sort((a, b) => a.date.localeCompare(b.date));
  return history;
}

// ---- setupWorkoutSheets ----
// Run once from Apps Script editor to create tabs and seed exercises.
function setupWorkoutSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName('ExerciseLibrary')) {
    const s = ss.insertSheet('ExerciseLibrary');
    s.appendRow(['id', 'name', 'muscleGroup', 'day', 'defaultWeight', 'isPerSide']);
  }

  if (!ss.getSheetByName('WorkoutSessions')) {
    const s = ss.insertSheet('WorkoutSessions');
    s.appendRow(['sessionId', 'date', 'dayName', 'exerciseName', 'setNumber', 'weight', 'reps', 'feel', 'timestamp']);
  }

  return syncExercises();
}

// ---- syncExercises ----
// Seeds the ExerciseLibrary. Safe to re-run — only adds missing exercises.
function syncExercises() {
  const sheet = getSheet('ExerciseLibrary');
  if (!sheet) return { error: 'ExerciseLibrary sheet not found. Run setupWorkoutSheets() first.' };

  const existing = sheet.getDataRange().getValues();
  const existingNames = new Set(existing.slice(1).map(r => String(r[1]).toLowerCase()));

  let maxNum = 0;
  existing.slice(1).forEach(r => {
    const n = parseInt(String(r[0]).replace('ex', ''));
    if (!isNaN(n) && n > maxNum) maxNum = n;
  });

  const catalog = [
    // CHEST DAY
    ['Flat Bench Press',            'Chest',        'CHEST',     45,  true],
    ['Incline Bench Press',         'Chest',        'CHEST',     45,  true],
    ['Decline Bench Press',         'Chest',        'CHEST',     50,  true],
    ['Chest Fly Machine',           'Chest',        'CHEST',     120, false],
    ['V-Bar Cable Pushdown',        'Triceps',      'CHEST',     140, false],
    ['Overhead Cable Extension',    'Triceps',      'CHEST',     50,  true],
    ['Reverse Grip Pushdown',       'Triceps',      'CHEST',     40,  false],
    ['Skull Crushers',              'Triceps',      'CHEST',     55,  false],
    ['Close Grip Bench Press',      'Triceps',      'CHEST',     95,  false],
    ['Tricep Dips',                 'Triceps',      'CHEST',     0,   false],
    ['Pec Deck',                    'Chest',        'CHEST',     110, false],
    ['Cable Fly',                   'Chest',        'CHEST',     30,  true],
    ['Landmine Press',              'Chest',        'CHEST',     25,  false],
    ['Ab Wheel',                    'Core',         'CHEST',     0,   false],
    ['Landmine Twist',              'Core',         'CHEST',     25,  false],
    // PULL DAY
    ['Wide Grip Lat Pulldown',      'Back',         'PULL',      120, false],
    ['Seated Close Grip Row',       'Back',         'PULL',      120, false],
    ['T-Bar Row',                   'Back',         'PULL',      90,  false],
    ['Landmine Row',                'Back',         'PULL',      45,  false],
    ['Pull-ups',                    'Back',         'PULL',      0,   false],
    ['Chin-ups',                    'Back',         'PULL',      0,   false],
    ['Rope Straight Arm Pulldown',  'Back',         'PULL',      80,  false],
    ['Rear Delt Fly Machine',       'Rear Delts',   'PULL',      90,  false],
    ['Back Extension Machine',      'Lower Back',   'PULL',      25,  false],
    ['Alternating Dumbbell Curls',  'Biceps',       'PULL',      35,  true],
    ['Hammer Curls',                'Biceps',       'PULL',      35,  true],
    ['Cross Body Curls',            'Biceps',       'PULL',      25,  true],
    ['Cable Concentration Curls',   'Biceps',       'PULL',      40,  false],
    ['Preacher Curl',               'Biceps',       'PULL',      50,  false],
    ['EZ Bar Curl',                 'Biceps',       'PULL',      60,  false],
    ['Incline DB Curl',             'Biceps',       'PULL',      20,  true],
    ['KB Single Arm Row',           'Back',         'PULL',      40,  true],
    ['Face Pulls',                  'Rear Delts',   'PULL',      50,  false],
    // SHOULDERS DAY
    ['Plate Loaded Shoulder Press', 'Shoulders',    'SHOULDERS', 45,  true],
    ['Arnold Press',                'Shoulders',    'SHOULDERS', 30,  true],
    ['Barbell OHP',                 'Shoulders',    'SHOULDERS', 95,  false],
    ['Push Press',                  'Shoulders',    'SHOULDERS', 95,  false],
    ['Bradford Press',              'Shoulders',    'SHOULDERS', 45,  false],
    ['Front Raises',                'Shoulders',    'SHOULDERS', 15,  false],
    ['Lateral Raises',              'Shoulders',    'SHOULDERS', 15,  false],
    ['Cable Lateral Raises',        'Shoulders',    'SHOULDERS', 15,  false],
    ['Upright Rows',                'Traps',        'SHOULDERS', 60,  false],
    ['Dumbbell Shrugs',             'Traps',        'SHOULDERS', 60,  true],
    ['Bent Over Rear Delt Raises',  'Rear Delts',   'SHOULDERS', 15,  false],
    ["Farmer's Carries",            'Functional',   'SHOULDERS', 65,  true],
    ['KB Clean and Press',          'Shoulders',    'SHOULDERS', 35,  true],
    ['KB Halo',                     'Shoulders',    'SHOULDERS', 25,  false],
    ['KB Windmill',                 'Functional',   'SHOULDERS', 20,  true],
    // LEGS DAY
    ['Leg Press',                   'Quads',        'LEGS',      180, false],
    ['Barbell Squat',               'Quads',        'LEGS',      135, false],
    ['Hack Squat',                  'Quads',        'LEGS',      90,  false],
    ['Goblet Squat',                'Quads',        'LEGS',      50,  false],
    ['Landmine Squat',              'Quads',        'LEGS',      35,  false],
    ['Sumo Deadlift',               'Hamstrings',   'LEGS',      185, false],
    ['KB Sumo Deadlift',            'Hamstrings',   'LEGS',      62,  false],
    ['Romanian Deadlift',           'Hamstrings',   'LEGS',      95,  false],
    ['Landmine RDL',                'Hamstrings',   'LEGS',      45,  false],
    ['Good Mornings',               'Hamstrings',   'LEGS',      45,  false],
    ['Leg Curl Machine',            'Hamstrings',   'LEGS',      80,  false],
    ['Leg Extension Machine',       'Quads',        'LEGS',      80,  false],
    ['Bulgarian Split Squat',       'Quads/Glutes', 'LEGS',      0,   false],
    ['Landmine Reverse Lunge',      'Quads/Glutes', 'LEGS',      35,  false],
    ['Barbell Hip Thrust',          'Glutes',       'LEGS',      135, false],
    ['Hip Abduction Machine',       'Glutes',       'LEGS',      70,  false],
    ['KB Swings',                   'Functional',   'LEGS',      40,  false],
    ['Box Jumps',                   'Functional',   'LEGS',      0,   false],
    ['Standing Calf Raise Machine', 'Calves',       'LEGS',      90,  false],
    ['Seated Calf Raise',           'Calves',       'LEGS',      50,  false],
    ['Pallof Press',                'Core',         'LEGS',      30,  false],
    ['Cable Woodchop',              'Core',         'LEGS',      30,  false],
    ['Hanging Leg Raise',           'Core',         'LEGS',      0,   false],
    // FUNCTIONAL
    ['KB Turkish Get-Up',           'Functional',   'FUNCTIONAL', 20,  true],
    ['KB Thruster',                 'Functional',   'FUNCTIONAL', 35,  false],
    ['KB Snatch',                   'Functional',   'FUNCTIONAL', 25,  true],
    ['KB Figure 8',                 'Functional',   'FUNCTIONAL', 35,  false],
    ['Battle Ropes',                'Cardio',       'FUNCTIONAL', 0,   false],
    ['Medicine Ball Slams',         'Functional',   'FUNCTIONAL', 20,  false],
    ['Renegade Row',                'Back/Core',    'FUNCTIONAL', 25,  false],
    ['Single Leg KB RDL',           'Hamstrings',   'FUNCTIONAL', 30,  true],
    ['Bear Crawl',                  'Functional',   'FUNCTIONAL', 0,   false],
    ['Sandbag Carry',               'Functional',   'FUNCTIONAL', 50,  false],
    ['KB Clean',                    'Functional',   'FUNCTIONAL', 35,  true],
    ['Broad Jump',                  'Functional',   'FUNCTIONAL', 0,   false],
    // CORE
    ['Plank',                       'Core',         'CORE',       0,   false],
    ['Side Plank',                  'Core',         'CORE',       0,   true],
    ['Dead Bug',                    'Core',         'CORE',       0,   false],
    ['Russian Twist',               'Core',         'CORE',       25,  false],
    ['Bicycle Crunches',            'Core',         'CORE',       0,   false],
    ['V-Ups',                       'Core',         'CORE',       0,   false],
    ['Cable Crunch',                'Core',         'CORE',       60,  false],
    ['Weighted Sit-Up',             'Core',         'CORE',       25,  false],
    ['Dragon Flag',                 'Core',         'CORE',       0,   false],
    ['L-Sit',                       'Core',         'CORE',       0,   false],
    ['Landmine 180s',               'Core',         'CORE',       25,  false],
    ['Windshield Wipers',           'Core',         'CORE',       0,   false],
    ['GHD Sit-Up',                  'Core',         'CORE',       0,   false],
    ['Toe Touches',                 'Core',         'CORE',       0,   false],
  ];

  let added = 0;
  catalog.forEach(([name, muscleGroup, day, defaultWeight, isPerSide]) => {
    if (!existingNames.has(name.toLowerCase())) {
      maxNum++;
      sheet.appendRow(['ex' + String(maxNum).padStart(3, '0'), name, muscleGroup, day, defaultWeight, isPerSide]);
      added++;
    }
  });

  return { success: true, message: added > 0 ? 'Added ' + added + ' exercises' : 'All exercises already present' };
}

// ---- getExerciseLibrary ----
function getExerciseLibrary() {
  const sheet = getSheet('ExerciseLibrary');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).filter(r => r[0]).map(row => ({
    id:            row[0],
    name:          row[1],
    muscleGroup:   row[2],
    day:           row[3],
    defaultWeight: parseFloat(row[4]) || 0,
    isPerSide:     !!row[5]
  }));
}

// ---- getLastWeightsForAllExercises ----
// Returns map of exerciseName → most recent weight logged (last row wins).
function getLastWeightsForAllExercises() {
  const sheet = getSheet('WorkoutSessions');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name   = data[i][3];
    const weight = data[i][5];
    if (name && weight !== '' && weight !== null) {
      map[name] = parseFloat(weight) || 0;
    }
  }
  return map;
}

// ---- getMaxWeightsForAllExercises ----
// Returns map of exerciseName → all-time max weight (personal record).
function getMaxWeightsForAllExercises() {
  const sheet = getSheet('WorkoutSessions');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name   = data[i][3];
    const weight = parseFloat(data[i][5]) || 0;
    if (name && weight > 0) {
      if (!map[name] || weight > map[name]) map[name] = weight;
    }
  }
  return map;
}

// ---- getNextWorkout ----
// Returns all exercises (with lastWeight + prWeight) and the next day in rotation.
function getNextWorkout() {
  try {
    const rotation = ['CHEST', 'PULL', 'SHOULDERS', 'LEGS'];

    const sessSheet = getSheet('WorkoutSessions');
    if (!sessSheet) return { error: 'Run setupWorkoutSheets() from Apps Script editor, then redeploy.' };

    const data = sessSheet.getDataRange().getValues();
    let lastDay = null;
    let latestDate = '';

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][1] instanceof Date
        ? Utilities.formatDate(data[i][1], TZ, 'yyyy-MM-dd')
        : String(data[i][1]);
      if (rowDate > latestDate) {
        latestDate = rowDate;
        lastDay = data[i][2];
      }
    }

    const lastIdx = lastDay ? rotation.indexOf(lastDay) : -1;
    const nextDay = rotation[(lastIdx + 1) % rotation.length];

    const exercises = getExerciseLibrary();
    const lastWeights = getLastWeightsForAllExercises();
    const prWeights   = getMaxWeightsForAllExercises();
    exercises.forEach(ex => {
      ex.lastWeight = lastWeights.hasOwnProperty(ex.name) ? lastWeights[ex.name] : null;
      ex.prWeight   = prWeights.hasOwnProperty(ex.name)   ? prWeights[ex.name]   : null;
    });

    return { nextDay, exercises, lastDay };
  } catch (err) {
    return { error: 'getNextWorkout error: ' + err.message };
  }
}

// ---- logWorkoutSession ----
function logWorkoutSession(body, user) {
  user = (user || 'aj').toLowerCase();
  try {
    const sessSheet = getSheet('WorkoutSessions');
    if (!sessSheet) return { error: 'WorkoutSessions sheet not found. Run setupWorkoutSheets() first.' };

    const sessionId = body.sessionId || ('sess_' + Date.now());
    const date      = body.date;
    const dayName   = body.dayName;
    const sets      = body.sets || [];
    const timestamp = new Date().toISOString();

    sets.forEach(set => {
      sessSheet.appendRow([
        sessionId,
        date,
        dayName,
        set.exerciseName,
        set.setNumber,
        parseFloat(set.weight) || 0,
        parseInt(set.reps)     || 0,
        set.feel || '',
        timestamp
      ]);
    });

    // Write to legacy Workouts sheet so DailySummary workoutDone flag updates
    const workoutSheet = getSheet('Workouts');
    if (workoutSheet) {
      const wData = workoutSheet.getDataRange().getValues();
      const existingDates = wData.slice(1).map(r =>
        r[0] instanceof Date
          ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd')
          : String(r[0])
      );
      if (!existingDates.includes(date)) {
        workoutSheet.appendRow([date, dayName, sets.length + ' sets']);
      }
    }

    rebuildDailySummary(date, user);
    return { success: true, sessionId, setsLogged: sets.length };
  } catch (err) {
    return { error: 'logWorkoutSession error: ' + err.message };
  }
}

// ---- getWorkoutHistory ----
function getWorkoutHistory(user) {
  user = (user || 'aj').toLowerCase();
  if (user === 'michele') return [];  // Michele uses activity logging, not WorkoutSessions
  const sheet = getSheet('WorkoutSessions');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const sessions = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sid = row[0];
    if (!sid) continue;
    if (!sessions[sid]) {
      sessions[sid] = {
        sessionId: sid,
        date: row[1] instanceof Date
          ? Utilities.formatDate(row[1], TZ, 'yyyy-MM-dd')
          : String(row[1]),
        dayName: row[2],
        sets: []
      };
    }
    sessions[sid].sets.push({
      exerciseName: row[3],
      setNumber:    row[4],
      weight:       row[5],
      reps:         row[6],
      feel:         row[7]
    });
  }

  return Object.values(sessions).sort((a, b) => b.date.localeCompare(a.date));
}

// ---- getHomeData ----
// Single endpoint that replaces getFoods + getDaySummary + getYesterday on home screen load.
function getHomeData(dateParam, user) {
  user = (user || 'aj').toLowerCase();
  const date       = dateParam || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const daySummary = getDaySummary(date, user);
  const yesterday  = getYesterday(user);
  const foods      = getFoods();
  const recentHistory = getSummaryHistory(user).slice(-30);

  if (user === 'michele') {
    const todayActivities = getMicheleActivityToday(date);
    return { foods, daySummary, yesterday, nextWorkoutDay: null, recentHistory, todayActivities };
  }

  const rotation = ['CHEST', 'PULL', 'SHOULDERS', 'LEGS'];
  let nextDay = 'CHEST';
  try {
    const sessSheet = getSheet('WorkoutSessions');
    if (sessSheet) {
      const sessData = sessSheet.getDataRange().getValues();
      let lastDay = null, latestDate = '';
      for (let i = 1; i < sessData.length; i++) {
        const rowDate = sessData[i][1] instanceof Date
          ? Utilities.formatDate(sessData[i][1], TZ, 'yyyy-MM-dd')
          : String(sessData[i][1]);
        if (rowDate > latestDate) { latestDate = rowDate; lastDay = sessData[i][2]; }
      }
      const lastIdx = lastDay ? rotation.indexOf(lastDay) : -1;
      nextDay = rotation[(lastIdx + 1) % rotation.length];
    }
  } catch (e) { /* ignore */ }

  return { foods, daySummary, yesterday, nextWorkoutDay: nextDay, recentHistory };
}

// ---- deleteLogEntry ----
// Deletes a single Log row by matching timestamp. Used for accidental entries.
function deleteLogEntry(body, user) {
  user = (user || 'aj').toLowerCase();
  const timestamp = body.timestamp;
  if (!timestamp) return { error: 'timestamp required' };

  const sheet = getSheet(getSheetName('Log', user));
  const data  = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    const rowTs = data[i][0] instanceof Date ? data[i][0].toISOString() : String(data[i][0]);
    if (rowTs === timestamp || rowTs.startsWith(timestamp.replace('Z', ''))) {
      sheet.deleteRow(i + 1);
      const date = data[i][1] instanceof Date
        ? Utilities.formatDate(data[i][1], TZ, 'yyyy-MM-dd') : String(data[i][1]);
      rebuildDailySummary(date, user);
      return { success: true };
    }
  }
  return { error: 'Entry not found' };
}

// ---- getCoachAdvice ----
// Pulls recent data from sheets and sends it + user's message to Claude.
function getCoachAdvice(body, user) {
  user = user || (body.user || 'aj').toLowerCase();
  const message = body.message || 'Review my recent data and give me tips.';
  const targets = body.targets || (user === 'michele'
    ? { protein: 110, calories: 1500, water: 64 }
    : { protein: 190, calories: 2300, water: 128 });

  const summaryHistory = getSummaryHistory(user).slice(-14);

  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const past  = summaryHistory.filter(h => h.date <= today);
  const n     = past.length;

  let ctx = 'TODAY\'S DATE: ' + today + '\n\n';
  ctx += 'USER: ' + (user === 'michele' ? 'Michele' : 'AJ') + '\n';
  ctx += 'USER TARGETS: ' + targets.protein + 'g protein | ' +
         targets.calories + ' cal | ' + targets.water + 'oz water\n\n';

  if (n > 0) {
    const avgProt     = Math.round(past.reduce((s, h) => s + h.protein, 0)  / n);
    const avgCal      = Math.round(past.reduce((s, h) => s + h.calories, 0) / n);
    const avgWater    = Math.round(past.reduce((s, h) => s + h.waterOz, 0)  / n);
    const protHits    = past.filter(h => h.protein >= targets.protein).length;
    const workoutDays = past.filter(h => h.workoutDone === 'Yes').length;
    const drinks = user === 'michele'
      ? past.reduce((s, h) => s + (h.redwineCount||0) + (h.surfsideCount||0) + (h.titosCount||0) + (h.martiniCount||0), 0)
      : past.reduce((s, h) => s + (h.coorsCount||0) + (h.bourbonCount||0) + (h.mimosaCount||0), 0);

    ctx += 'LAST ' + n + ' DAYS AVERAGES:\n';
    ctx += '  Protein: ' + avgProt + 'g avg (goal hit ' + protHits + '/' + n + ' days)\n';
    ctx += '  Calories: ' + avgCal + ' avg\n';
    ctx += '  Water: ' + avgWater + 'oz avg\n';
    ctx += '  ' + (user === 'michele' ? 'Active days' : 'Workouts') + ': ' + workoutDays + ' of ' + n + ' days\n';
    ctx += '  Drinks logged: ' + drinks + ' total\n\n';
  }

  ctx += 'DAILY LOG (most recent ' + Math.min(summaryHistory.length, 14) + ' days):\n';
  summaryHistory.slice().reverse().slice(0, 10).forEach(h => {
    ctx += h.date + ': ' + h.protein + 'p / ' + h.calories + 'cal / ' +
           h.waterOz + 'oz water / workout:' + (h.workoutDone || 'No') +
           (h.bodyweight ? ' / ' + h.bodyweight + 'lbs' : '') + '\n';
  });

  if (user === 'michele') {
    const actHistory = getMicheleRecentActivities();
    if (actHistory.length > 0) {
      ctx += '\nRECENT ACTIVITIES:\n';
      actHistory.slice(0, 10).forEach(a => {
        ctx += a.date + ': ' + a.activityType + '\n';
      });
    }
  } else {
    const workoutHistory = getWorkoutHistory().slice(0, 8);
    if (workoutHistory.length > 0) {
      ctx += '\nRECENT WORKOUTS:\n';
      workoutHistory.slice(0, 5).forEach(s => {
        const exCount = new Set(s.sets.map(x => x.exerciseName)).size;
        ctx += s.date + ': ' + s.dayName + ' — ' + exCount + ' exercises / ' + s.sets.length + ' sets\n';
      });
    }
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'CLAUDE_API_KEY not set in Script Properties.' };

  const systemPrompt = user === 'michele'
    ? 'You are a concise, supportive nutrition and wellness coach in "The Grind" app. ' +
      'Michele is 5\'3", 135 lbs, working to lose 5-10 lbs. She does Club Pilates and 2-mile walks. ' +
      'She tracks protein, calories, water, and bodyweight daily. ' +
      'Be encouraging and direct. Highlight wins, gently flag patterns that might slow progress. ' +
      'Keep responses to 2-3 short paragraphs. Use specific numbers from her data. No generic filler.\n\n' +
      'HER DATA:\n' + ctx
    : 'You are a concise, no-BS fitness and nutrition coach embedded in "The Grind" tracking app. ' +
      'The user tracks protein, calories, water, bodyweight, and gym workouts daily. ' +
      'Give practical, personalized advice based on their actual data. ' +
      'Be direct and encouraging. Keep responses to 2-3 short paragraphs max. ' +
      'Use specific numbers from their data when relevant. No generic filler.\n\n' +
      'THEIR DATA:\n' + ctx;

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 600,
        system: systemPrompt,
        messages: [{ role: 'user', content: message }]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(response.getContentText());
    if (data.error) return { error: 'Claude API: ' + data.error.message };
    return { success: true, reply: data.content[0].text };
  } catch (err) {
    return { error: 'Coach error: ' + err.message };
  }
}
