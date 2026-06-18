# Build Spec: Personal Food and Workout Tracker

## What this is
A single-page HTML app, hosted on GitHub Pages, that logs food, water, alcohol, and
workouts to a Google Sheet in Drive via an Apps Script web app backend. The Sheet is the
source of truth and the trending engine. The app is the fast logging front end.

## Architecture
- Front end: one static HTML file (HTML + CSS + vanilla JS), hosted on GitHub Pages.
- Backend: Google Apps Script web app bound to a Google Sheet, deployed as a web app
  ("execute as me", "anyone with the link"). Exposes a small JSON API.
- Data store: the bound Google Sheet, five tabs.
- Bridge: front end calls the Apps Script URL with fetch. Not a live sync, direct API calls.

## Hard requirement: CORS
Apps Script web apps do not return CORS headers on normal doGet/doPost the way a REST API
would, and preflight on POST with application/json will fail. Use this pattern:
- Send POST requests with `Content-Type: text/plain;charset=utf-8` to avoid the CORS
  preflight. Put the JSON in the body as a string, parse it server side with JSON.parse.
- Return JSON from the script with ContentService
  `.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)`.
- Expect one debugging pass to get this solid. This is the known friction point.

## Google Sheet schema, five tabs

### Tab 1: Foods (the library)
Columns: FoodID | Name | ServingLabel | Protein | Calories | Carbs | Fat | Sugar | Sodium | DateAdded
- FoodID: auto, e.g. "f001"
- ServingLabel: human label for one serving, e.g. "1 container"
- Macros are PER ONE SERVING.

Seed rows (from today's prep, macros already calculated):
f001 | Beef Egg Roll Bowl | 1 container | 16 | 230 | 14 | 12 | 5 | 0 | (today)
f002 | Turkey with Tomato | 1 container | 26 | 160 | 4 | 6 | 2 | 0 | (today)
f003 | Steamed Broccoli | 1 serving | 3 | 60 | 11 | 1 | 2 | 0 | (today)
f004 | ONE Bar Reeses PB | 1 bar | 18 | 220 | 21 | 8 | 3 | 0 | (today)
f005 | 2-Egg Veggie Omelette | 1 omelette | 12 | 220 | 8 | 14 | 4 | 0 | (today)

### Tab 2: Log (daily food/liquid entries)
Columns: Timestamp | Date | Type | Item | Servings | Protein | Calories | Carbs | Fat | Sugar | Sodium
- Type: one of food, water, coffee, coors, bourbon
- Servings: multiplier (food). For liquids, Servings holds the count or oz as relevant.
- Macro columns: the computed totals for that entry (per-serving x Servings).

### Tab 3: Workouts (training log)
Columns: Date | MuscleGroup | Exercise | Sets | Reps | Weight | Notes

### Tab 4: DailySummary (computed roll-up, one row per day)
Columns: Date | TotalProtein | TotalCalories | WaterOz | CoffeeCount | CoorsCount |
BourbonCount | WorkoutDone | Bodyweight
- Bodyweight: manual daily entry. This is the metric the real goal lives on.
- WorkoutDone: yes/no derived from whether Workouts has rows for that date.

### Tab 5: Dashboard (trends)
Pivots and charts off DailySummary:
- Week over week and month over month average protein, calories, water, drinks.
- Best and worst days (by protein hit and by calorie target adherence).
- Protein target hit rate (days at or above target / total days).
- Bodyweight trend line.

## Targets (constants in the app, editable)
- Protein target: 180 g/day
- Calorie target: set a sensible cut number for a 6'7", 241 lb, 46 yo male training 3-4x/wk.
  Start at 2,300 and make it editable.
- Water target: 1 gallon (128 oz)

## Apps Script API (function contract)
Implement doGet and doPost dispatching on an `action` field.

- getFoods() -> array of all Foods rows. Called on app load to populate the library.
- addFood({name, servingLabel, protein, calories, carbs, fat, sugar, sodium})
    -> writes one row to Foods, returns the new FoodID. This is the "remember my meal" call.
- logEntry({type, item, foodId, servings}) -> looks up per-serving macros from Foods,
    multiplies by servings, writes to Log, returns the entry. Handles type=food.
- logLiquid({type, amount}) -> writes a water/coffee/coors/bourbon row to Log.
    water amount in oz, others count = 1 per press.
- logWorkout({date, muscleGroup, exercise, sets, reps, weight, notes}) -> writes to Workouts.
- getToday() -> returns today's running totals: protein, calories, water oz, coffee count,
    coors count, bourbon count. Drives the live dashboard on the main screen.
- rebuildDailySummary() -> recomputes DailySummary from Log and Workouts. Run on each
    logEntry/logLiquid, or nightly via a time trigger.

## Plain-language meal entry (the chicken example)
Flow: user types free text like
"4 chicken breasts at 5oz each, asparagus, split into 4 containers"
App does NOT need to auto-calculate macros from the text in v1 (manual macros path).
Instead:
1. User types the name/description and the number of containers it was split into.
2. App prompts for total batch macros OR per-container macros (toggle).
3. If user enters batch totals, app divides by container count to get per-serving.
4. Saves to Foods with ServingLabel "1 container".
5. From then on, it appears as a tappable item. User taps it, picks servings (1, 2, ...),
   app logs per-serving macros x servings.
Multi-serving selection is required: a stepper or quick 1/2/3 buttons on each food.

## Quick-tap liquid buttons (main screen, top priority UX)
Four big buttons, one tap logs immediately, no dropdowns:
- Coffee  (logs 1 coffee)
- Water   (let user pick 12 or 24 oz, two buttons or a toggle, since those are the only sizes)
- Coors   (logs 1 beer)
- Bourbon (logs 1 pour)
Each press writes to Log and updates the live totals instantly. These are the most-used
controls on the page. Make them the fastest thing to reach.

## Main screen layout
1. Top: today's live totals. Protein and Calories big, against target with a progress bar.
   Water, coffee, Coors, bourbon counts as smaller stat chips.
2. The four quick-tap liquid buttons.
3. Food library as a tappable grid/list, each with a serving stepper and a log button.
4. "Add new food" entry (the plain-language + macros flow above).
5. Workout log section (collapsible): quick entry for the day's exercises.
6. Bodyweight quick entry (one number, once a day).

## Build sequence
1. Stand up the Sheet with the five tabs and seed the Foods rows.
2. Write the Apps Script, deploy as web app, grab the URL.
3. Build the HTML front end against that URL. Start with getFoods + the four liquid buttons
   + live totals. Prove the round trip works (this is where CORS gets debugged).
4. Add food logging with multi-serving, then add-new-food.
5. Add workout logging and bodyweight.
6. Build the Dashboard tab formulas/charts last, once real data exists.

## v1 scope discipline
- Manual macros only. No third-party nutrition API in v1.
- Food module is the priority. Workout logging can be minimal in v1 and expanded later.
- Ship the logging loop first. Trending comes after real data accumulates.

## Notes for macro accuracy
Seed macros above are estimates from standard USDA values for the stated ingredients and
quantities. The soft spots are the beef fat (varies with lean ratio) and sauce distribution
across containers. Good enough to track and plan, not lab-precise. User can edit any Foods row.
