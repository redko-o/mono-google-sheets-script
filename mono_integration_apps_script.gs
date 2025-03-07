/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('😺 Mono Menu')
    .addItem('💳 Завантажити нові транзакції', 'uploadAllTransactions')
    .addItem('📅 Завантажити за період...', 'uploadTransactionsForCustomPeriod')
    .addItem('📃 Застосувати правила', 'applyRulesToTransactions')
    .addItem('❗️ Створити/перестворити табличку', 'initialCreate')
    .addToUi();
}

const MONO_TOKEN = getScriptSecret("MONO_TOKEN")
const jsonString = HtmlService.createHtmlOutputFromFile("mcc.html").getContent();
const jsonObject = JSON.parse(jsonString);

let columns = [
  "Джерело", "Баланс","Час транзакції", "Сума транзакції", "Категорія", "Опис",
  "Коментар", "MCC", "Original MCC", "Назва категорії","Кешбек"
]

let columnsWidths = [85, 75, 80, 75, 140, 150, 100, 75, 75, 120, 70]

let floatColumns = ["Баланс", "Сума транзакції", "Кешбек"]
let textColumns = ["Опис", "Коментар"]
let datetimeColumns = ["Час транзакції"]

let categories = [
  "💸 Базові витрати", "💅 Краса і здоровʼя", "💃 Відпустка", "👵🏻 Батьки", "⚽️ Хобі",
  "🎓 Навчання і освіта", "🪖 Перемога", "📦 Оновлення речей", "🎁 Подарунки",
  "🎡 Дозвілля", "💰Інвестиції", "🍋 Дохід", "🔄 Транзакції", "Інше"
]

let sources = ["Mono", "Mono біла", "Готівка"]

// Додаємо масив з обома рахунками
const accounts = [
    { account: 'MONO_BLACK', source: 'Mono' },      // Основний рахунок
    { account: 'MONO_WHITE', source: 'Mono біла' }  // Додатковий рахунок
];

function getRules(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName(sheetName);

  if (!rulesSheet) {
    throw new Error(`Вкладку '${sheetName}' не знайдено!`);
  }

  const rulesData = rulesSheet.getDataRange().getValues();
  const headers = rulesData.shift(); // Remove headers
  const ruleFieldNameIdx = headers.indexOf("Назва поля");
  const ruleValueIdx = headers.indexOf("Значення");
  const ruleThenIdx = headers.indexOf("Тоді");
  const ruleEqualsIdx = headers.indexOf("Дорівнює");

  if ([ruleFieldNameIdx, ruleValueIdx, ruleThenIdx, ruleEqualsIdx].includes(-1)) {
    throw new Error("Заголовки мають бути 'Назва поля', 'Значення', 'Тоді', 'Дорівнює'.");
  }

  return { rules: rulesData, indices: { ruleFieldNameIdx, ruleValueIdx, ruleThenIdx, ruleEqualsIdx } };
}

function applyRulesToTransactionsData(transactions, transactionHeaders, rules, ruleIndices) {
  for (let i = 0; i < transactions.length; i++) {
    let transaction = transactions[i];

    for (let j = 0; j < rules.length; j++) {
      const rule = rules[j];

      const fieldIndex = transactionHeaders.indexOf(rule[ruleIndices.ruleFieldNameIdx]);
      const thenFieldIndex = transactionHeaders.indexOf(rule[ruleIndices.ruleThenIdx]);
      if (fieldIndex === -1 || thenFieldIndex === -1) {
        Logger.log(`Поле ${rule[ruleIndices.ruleFieldNameIdx]} або ${rule[ruleIndices.ruleThenIdx]} не знайдено.`);
        continue; // Пропускаємо, якщо поле не знайдено
      }

      if (transaction[fieldIndex]?.toString().includes(rule[ruleIndices.ruleValueIdx])) {
        transaction[thenFieldIndex] = rule[ruleIndices.ruleEqualsIdx];
        Logger.log(`Застосовано правило: ${rule[ruleIndices.ruleFieldNameIdx]} -> ${rule[ruleIndices.ruleThenIdx]}`);
      }
    }
  }
  return transactions;
}

function applyRulesToTransactions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetByName("Усі транзакції");

    if (!transactionsSheet) {
      throw new Error("Вкладку 'Усі транзакції' не знайдено!");
    }

    const transactionsData = transactionsSheet.getDataRange().getValues();
    const transactionHeaders = transactionsData.shift(); // Remove headers

    const { rules, indices } = getRules("Правила");

    const updatedTransactions = applyRulesToTransactionsData(transactionsData, transactionHeaders, rules, indices);

    // Записуємо оновлені транзакції у таблицю
    transactionsSheet.getRange(2, 1, updatedTransactions.length, updatedTransactions[0].length).setValues(updatedTransactions);
    SpreadsheetApp.getUi().alert("Правила успішно застосовані до транзакцій!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Помилка: ${e.message}`);
  }
}

function initialCreate() {
  // create new sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Усі транзакції";

  let oldSheet = ss.getSheetByName(sheetName);
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }

  let newSheet = ss.insertSheet(sheetName, 0); // Creates a new sheet at the beginning of the spreadsheet

  // Add header row with filters
  let headerRowRange = newSheet.getRange(1, 1, 1, columns.length);
  headerRowRange.setValues([columns]);
  headerRowRange.setFontWeight("bold");
  newSheet.setFrozenRows(1); // Freeze the header row

  let maxRows = newSheet.getMaxRows();
  let lastColumn = newSheet.getLastColumn();
  // Apply filters
  let dataRange = newSheet.getDataRange();
  dataRange.createFilter();

  // apply color schema
  let range = newSheet.getRange(1, 1, maxRows, lastColumn);
  range.applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW);

  for (const [index, width] of columnsWidths.entries()) {
    newSheet.setColumnWidth(index + 1, width);
  }

  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // apply data types
  // drop down for source
  let sourceColumnIndex = columns.indexOf("Джерело") + 1;
  let sourceColumn = newSheet.getRange(2, sourceColumnIndex, maxRows); // start from 2 to ignore header
  let sourceRule = SpreadsheetApp.newDataValidation().requireValueInList(sources).build();
  sourceColumn.setDataValidation(sourceRule);

  // drop down for categories
  let catColumnIndex = columns.indexOf("Категорія") + 1;
  let catColumn = newSheet.getRange(2, catColumnIndex, maxRows); // start from 2 to ignore header
  let catRule = SpreadsheetApp.newDataValidation().requireValueInList(categories).build();
  catColumn.setDataValidation(catRule);

  //Color RED for category "Інше"
  let range1 = newSheet.getRange("E:E");
  let catRuleFormat = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Інше")
    .setBackground("#F5CBCC").setRanges([range1]).build();
  const rules = newSheet.getConditionalFormatRules();
  rules.push(catRuleFormat);
  newSheet.setConditionalFormatRules(rules);

  // set datatypes 
  applyFormating(floatColumns, newSheet, "#,##0.00");
  applyFormating(textColumns, newSheet, "@")
  applyFormating(datetimeColumns, newSheet, "dd.mm.yyyy, HH:mm");
}

function applyFormating(columnsToApply, sheet, format) {
  let ranges = columnsToApply.map(column => {
    let columnIndex = columns.indexOf(column) + 1;
    let columnRange = sheet.getRange(1, columnIndex, sheet.getMaxRows(), 1);
    return columnRange;
  });
  ranges.map(range => { range.setNumberFormat(format); });
}

function getCategoryNameByMCC(mcc) {
  const match = jsonObject.find(item => item.mcc === mcc.toString());
  return match ? match.shortDescription.uk : "Категорія не знайдена";
}

var lastApiRequest;

function uploadAllTransactions() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Усі транзакції");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Отримуємо правила один раз
  const { rules, indices } = getRules("Правила");
  
  let to = Date.now();
  let from = getLatestTransactionTs() + 1001;
  
  // Проходимо по кожному рахунку
  accounts.forEach(({ account, source }) => {
    Logger.log(`Завантажуємо транзакції для рахунку: ${source}`);
    
    let periods = getTimePeriods(from, to);
    
    periods.forEach(([periodFrom, periodTo]) => {
      let transactions = getTransactions(account, periodFrom, periodTo);
      
      // Обробляємо кожну транзакцію
      for (let step = transactions.length - 1; step >= 0; step--) {
        let transaction = transactions[step];
        transaction.source = source; // Встановлюємо правильне джерело
        
        // Формуємо рядок для транзакції
        let entry = headers.map(col => {
          if (col === "Назва категорії") {
            return getCategoryNameByMCC(transaction.mcc);
          } else {
            return transaction.columnMap().get(col);
          }
        });
        
        // Застосовуємо правила до однієї транзакції
        let updatedEntry = applyRulesToTransactionsData(
          [entry], // масив з однієї транзакції
          headers,
          rules,
          indices
        )[0]; // беремо перший (і єдиний) елемент
        
        // Шукаємо позицію для вставки
        let insertPosition = 2;
        let currentData = sheet.getDataRange().getValues();
        currentData.shift(); // Пропускаємо заголовок
        
        let timestampIndex = headers.indexOf("Час транзакції");
        let transactionDate = new Date(transaction.time);
        
        // Знаходимо позицію для збереження сортування
        while (insertPosition - 1 < currentData.length) {
          let currentRowDate = new Date(currentData[insertPosition - 2][timestampIndex]);
          if (transactionDate > currentRowDate) {
            break;
          }
          insertPosition++;
        }
        
        // Вставляємо транзакцію з застосованими правилами
        sheet.insertRowBefore(insertPosition);
        sheet.getRange(insertPosition, 1, 1, updatedEntry.length).setValues([updatedEntry]);
      }
    });
  });
  
  SpreadsheetApp.getUi().alert("Транзакції успішно завантажені!");
}

function getLatestTransactionTs() {
  Logger.info("Отримуємо час останньої завантаженої транзакції")
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Усі транзакції");
  let transactionsTable = sheet.getDataRange().getValues();
  let headers = transactionsTable.shift();

  let timestampIndex = headers.indexOf("Час транзакції");

  var from = 0;
  // Шукаємо останню транзакцію
  for (let step = 0; step < transactionsTable.length; step++) {
    let transactionTsCell = transactionsTable[step][timestampIndex]
    if (!transactionTsCell) { continue }

    let transactionTs = transactionTsCell.valueOf()
    if (transactionTs > from) {
      from = transactionTs;
      Logger.info(`Час останньої транзакції - ${new Date(from).toISOString()}`);
      break;
    }
  }
  // якщо транзакцій Моно ще не було, то беремо дані за останні 30 днів
  if (from == 0) {
    let lastMonth = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000)).getTime();
    Logger.info(`Останньої транзакції не знайдено, завантажуємо транзакції за останні 30 днів ${lastMonth}`);
    from = lastMonth;
  }
  return from;
}

function getTimePeriods(fromRaw, toRaw) {
  // swap if needed
  let [from, to] = fromRaw < toRaw ? [fromRaw, toRaw] : [toRaw, fromRaw];

  Logger.info(`Розбиваємо період (${new Date(from).toISOString()}, ${new Date(to).toISOString()}) на проміжки не більші за 31 добу + 1 годину (2682000 секунд)`);
  // "Максимальний час, за який можливо отримати виписку — 31 доба + 1 година (2682000 секунд)" (c) документація
  const maxPeriodMillis = 2682000 * 1000;
  const oneDayMillis = 24 * 60 * 60 * 1000;

  var chunks = [];
  
  // якщо період менший за максимальний
  if (to - from < maxPeriodMillis) {
    chunks.push([from, to]);
  } else {
    // розбиваємо на проміжки
    let chunkFrom = from;
    while (chunkFrom < to) {
      let chunkTo = Math.min(chunkFrom + maxPeriodMillis - oneDayMillis, to);
      chunks.push([chunkFrom, chunkTo]);
      chunkFrom = chunkTo + 1; // Додаємо 1 мілісекунду, щоб уникнути перекриття
    }
  }

  let prettyChunks = chunks.map(([from, to]) => [new Date(from).toISOString(), new Date(to).toISOString()]);
  Logger.info(`Отримані проміжки ${prettyChunks.join(', ')}`);
  return chunks;
}

function getTransactions(account, from, to) {
  var transactions = [];
  var newFrom = from;
  var transactionsCnt;
  
  Logger.info(`Отримуємо транзакції для рахунку ${account} за період (${new Date(newFrom).toISOString()}, ${new Date(to).toISOString()})`);

  do {
    // Перевіряємо, чи потрібно чекати
    if (lastApiRequest) {
      let currentTime = Date.now();
      let timeSinceLastRequest = currentTime - lastApiRequest;
      let waitTime = Math.max(61000 - timeSinceLastRequest, 0); // Мінімум 0 мс

      if (waitTime > 0) {
        Logger.info(`Чекаємо ${Math.round(waitTime/1000)} секунд перед наступним запитом`);
        Utilities.sleep(waitTime);
      }
    }
    
    // Робимо запит
    let newTransactions = makeRequest(account, newFrom, to);
    lastApiRequest = Date.now();

    if (newTransactions.length == 0) break;
    
    newFrom = newTransactions.at(-1).time;
    transactionsCnt = newTransactions.length;
    transactions.push(newTransactions);
  } while (transactionsCnt == 500);

  return transactions.flat();
}

function makeRequest(account, from, to) {
  // Використовуємо переданий account
  let account_id = getScriptSecret(account);
  let URL_STRING = `https://api.monobank.ua/personal/statement/${account}/${from}/${to}`;
  let options = {
    'method': 'get',
    'headers': { 'X-Token': MONO_TOKEN },
    'muteHttpExceptions': true
  };
  Logger.log(`Робимо запит: ${URL_STRING}`);

  let response = UrlFetchApp.fetch(URL_STRING, options);
  let responseCode = response.getResponseCode();
  let json = response.getContentText();

  if (responseCode == 429) {
    throw new Error('Забагато запитів за короткий проміжок часу. Почекайте 1 хвилину і спробуйте ще раз');
  } else if (responseCode >= 300) {
    throw new Error(`${responseCode}: ${json}`);
  }

  let transactions = JSON.parse(json).map(MonoTransaction.fromJSON);

  return transactions
}

function getScriptSecret(key) {
  let secret = PropertiesService.getScriptProperties().getProperty(key);
  if (!secret) throw Error(`Ключ ${key} не знайдено. Будь ласка, додайте його в "Властивості скрипта"`);
  return secret;
}

class MonoTransaction {
  constructor({
    time,
    description,
    mcc,
    originalMcc,
    amount,
    cashbackAmount,
    balance,
    comment,
    source
  }) {
    // переводимо epoch seconds в timestamp, а копійки в гривні
    this.time = new Date(time * 1000);
    this.amount = amount / -100;
    this.cashbackAmount = cashbackAmount / 100;
    this.description = description;
    this.mcc = mcc;
    this.originalMcc = originalMcc;
    this.comment = comment;
    this.balance = balance / 100;

    this.source = source || 'Mono';
    this.category = 'Інше';
  }

  columnMap(){
     return new Map([
      ["Джерело", this.source],
      ["Баланс", this.balance],
      ["Сума транзакції", this.amount],
      ["Кешбек", this.cashbackAmount],
      ["Опис", this.description],
      ["MCC", this.mcc],
      ["Original MCC", this.originalMcc],
      ["Коментар", this.comment],
      ["Час транзакції", this.time],
      ["Категорія", this.category],
    ])
  }

  static fromJSON(json) {
    return new MonoTransaction({
      time: json.time,
      description: json.description,
      mcc: json.mcc,
      originalMcc: json.originalMcc,
      amount: json.amount,
      cashbackAmount: json.cashbackAmount,
      balance: json.balance,
      comment: json.comment,
      source: json.source
    }
    );
  }
}

function uploadTransactionsForCustomPeriod() {
  var ui = SpreadsheetApp.getUi();
  
  // Запитуємо початкову дату та час
  var fromResponse = ui.prompt(
    'Початкова дата та час',
    'Введіть дату та час початку у форматі DD.MM.YYYY HH:mm',
    ui.ButtonSet.OK_CANCEL);
  
  if (fromResponse.getSelectedButton() != ui.Button.OK) return;
  
  // Запитуємо кінцеву дату та час
  var toResponse = ui.prompt(
    'Кінцева дата та час',
    'Введіть дату та час кінця у форматі DD.MM.YYYY HH:mm',
    ui.ButtonSet.OK_CANCEL);
  
  if (toResponse.getSelectedButton() != ui.Button.OK) return;
  
  // Парсимо дати
  var from = parseDateTime(fromResponse.getResponseText());
  var to = parseDateTime(toResponse.getResponseText());
  
  // Перевіряємо валідність дат
  if (!from || !to) {
    ui.alert('Помилка', 'Неправильний формат дати та часу. Використовуйте формат DD.MM.YYYY HH:mm', ui.ButtonSet.OK);
    return;
  }
  
  // Перевіряємо період
  var daysDiff = Math.ceil((to - from) / (1000 * 60 * 60 * 24));
  if (daysDiff > 90) {
    ui.alert('Помилка', 'Період не може перевищувати 90 днів', ui.ButtonSet.OK);
    return;
  }
  
  if (to < from) {
    ui.alert('Помилка', 'Кінцева дата не може бути раніше початкової', ui.ButtonSet.OK);
    return;
  }
  
  uploadTransactionsForPeriod(from.getTime(), to.getTime());
}

function parseDateTime(dateTimeStr) {
  // Очікуваний формат: "DD.MM.YYYY HH:mm"
  try {
    const [dateStr, timeStr] = dateTimeStr.split(' ');
    const [day, month, year] = dateStr.split('.');
    const [hours, minutes] = timeStr.split(':');
    
    // Місяці в JavaScript починаються з 0
    const date = new Date(year, month - 1, day, hours, minutes);
    
    return isNaN(date.getTime()) ? null : date;
  } catch (e) {
    return null;
  }
}

function uploadTransactionsForPeriod(from, to) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Усі транзакції");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Отримуємо правила один раз
    const { rules, indices } = getRules("Правила");
    
    // Проходимо по кожному рахунку
    accounts.forEach(({ account, source }) => {
      Logger.log(`Завантажуємо транзакції для рахунку: ${source}`);
      
      let periods = getTimePeriods(from, to);
      
      periods.forEach(([periodFrom, periodTo]) => {
        let transactions = getTransactions(account, periodFrom, periodTo);
        
        // Обробляємо кожну транзакцію
        for (let step = transactions.length - 1; step >= 0; step--) {
          let transaction = transactions[step];
          transaction.source = source;
          
          // Формуємо рядок для транзакції
          let entry = headers.map(col => {
            if (col === "Назва категорії") {
              return getCategoryNameByMCC(transaction.mcc);
            } else {
              return transaction.columnMap().get(col);
            }
          });
          
          // Застосовуємо правила до однієї транзакції
          let updatedEntry = applyRulesToTransactionsData(
            [entry],
            headers,
            rules,
            indices
          )[0];
          
          // Шукаємо позицію для вставки
          let insertPosition = 2;
          let currentData = sheet.getDataRange().getValues();
          currentData.shift();
          
          let timestampIndex = headers.indexOf("Час транзакції");
          let transactionDate = new Date(transaction.time);
          
          while (insertPosition - 1 < currentData.length) {
            let currentRowDate = new Date(currentData[insertPosition - 2][timestampIndex]);
            if (transactionDate > currentRowDate) {
              break;
            }
            insertPosition++;
          }
          
          sheet.insertRowBefore(insertPosition);
          sheet.getRange(insertPosition, 1, 1, updatedEntry.length).setValues([updatedEntry]);
        }
      });
    });
    
    SpreadsheetApp.getUi().alert("Транзакції успішно завантажені!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Помилка: ${e.message}`);
  }
}