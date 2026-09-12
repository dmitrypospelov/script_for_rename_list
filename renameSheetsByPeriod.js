function startSettings() {
  var html = HtmlService.createHtmlOutput(getStep1Html())
    .setWidth(400)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Бюджет')
    .addItem('Настроить таблицу', 'startSettings')
    .addToUi();
}

function getStep1Html() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    'body { font-family: Arial, sans-serif; padding: 25px; background: #f8f9fa; margin: 0; }' +
    'h3 { margin-top: 0; color: #333; text-align: center; font-size: 17px; }' +
    'select { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; box-sizing: border-box; background: white; margin-top: 20px; }' +
    'button { margin-top: 20px; width: 100%; padding: 12px; background: #458275; color: #F0EBE2; border: none; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; }' +
    'button:hover { background: #36695e; }' +
    '.error { font-size: 11px; color: #c62828; text-align: center; margin-top: 12px; display: none; line-height: 1.4; }' +
    '</style></head><body>' +
    '<h3>С какого месяца Вы планируете начать вести бюджет?</h3>' +
    '<select id="month">' +
    '<option value="">Выбрать месяц</option>' +
    '<option>Январь</option><option>Февраль</option><option>Март</option>' +
    '<option>Апрель</option><option>Май</option><option>Июнь</option>' +
    '<option>Июль</option><option>Август</option><option>Сентябрь</option>' +
    '<option>Октябрь</option><option>Ноябрь</option><option>Декабрь</option>' +
    '</select>' +
    '<div class="error" id="errorMonth">Пожалуйста, выберите месяц из списка</div>' +
    '<button onclick="save()">Далее</button>' +
    '<script>' +
    'function save() {' +
    '  var month = document.getElementById("month").value;' +
    '  var error = document.getElementById("errorMonth");' +
    '  if (!month) {' +
    '    error.style.display = "block";' +
    '    return;' +
    '  }' +
    '  error.style.display = "none";' +
    '  google.script.run.saveMonth(month);' +
    '}' +
    '</script>' +
    '</body></html>';
}

function getStep2Html() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; background: #f8f9fa; margin: 0; }' +
    'h3 { margin-top: 0; color: #333; text-align: center; font-size: 15px; }' +
    '.hint { font-size: 11px; color: #888; text-align: center; margin-bottom: 15px; line-height: 1.4; }' +
    '.buttons { display: flex; flex-direction: column; gap: 8px; }' +
    'button { padding: 10px; border: none; border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; }' +
    '.btn-yes { background: #458275; color: #F0EBE2; }' +
    '.btn-yes:hover { background: #36695e; }' +
    '.btn-no { background: #e0e0e0; color: #333; }' +
    '.btn-no:hover { background: #d0d0d0; }' +
    '.sum-input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 6px; font-size: 13px; box-sizing: border-box; margin-bottom: 8px; }' +
    '.save-btn { display: block; margin: 0 auto; }' +
    '.error { font-size: 11px; color: #c62828; text-align: center; margin-bottom: 12px; display: none; line-height: 1.4; }' +
    '</style></head><body>' +
    '<h3>Хотите ли указать Начальную сумму?</h3>' +
    '<div class="hint">Начальная сумма — сумма всех текущих наличных и безналичных денег</div>' +
    '<div id="sumBlock" style="display:none;">' +
    '<input type="text" id="sum" class="sum-input" placeholder="Введите сумму">' +
    '<div class="error" id="errorNumber">' +
    '  Пожалуйста, укажите число, например, 2000<br>' +
    '  Все расчеты в таблице ведутся в рублях автоматически' +
    '</div>' +
    '<div class="error" id="errorSpace">' +
    '  Уберите, пожалуйста, пробел' +
    '</div>' +
    '<button class="btn-yes save-btn" onclick="saveSum()">Сохранить</button>' +
    '</div>' +
    '<div class="buttons" id="buttonsBlock">' +
    '<button class="btn-yes" onclick="showInput()">Да</button>' +
    '<button class="btn-no" onclick="skip()">Нет, спасибо, укажу позже</button>' +
    '</div>' +
    '<script>' +
    'function showInput() {' +
    '  document.getElementById("buttonsBlock").style.display = "none";' +
    '  document.getElementById("sumBlock").style.display = "block";' +
    '  document.getElementById("sum").focus();' +
    '}' +
    'function saveSum() {' +
    '  var input = document.getElementById("sum");' +
    '  var val = input.value;' +
    '  var errNum = document.getElementById("errorNumber");' +
    '  var errSpc = document.getElementById("errorSpace");' +
    '  errNum.style.display = "none";' +
    '  errSpc.style.display = "none";' +
    '  if (val.indexOf(" ") !== -1) {' +
    '    errSpc.style.display = "block";' +
    '    return;' +
    '  }' +
    '  var trimmed = val.trim();' +
    '  if (trimmed === "") {' +
    '    errNum.style.display = "block";' +
    '    return;' +
    '  }' +
    '  var num = Number(trimmed);' +
    '  if (isNaN(num) || num < 0) {' +
    '    errNum.style.display = "block";' +
    '    return;' +
    '  }' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { google.script.host.close(); })' +
    '    .withFailureHandler(function(err) { alert("Ошибка: " + err.message); })' +
    '    .saveStartSum(num);' +
    '}' +
    'function skip() {' +
    '  google.script.run.skipSum();' +
    '}' +
    '</script>' +
    '</body></html>';
}

function getWaitingHtml() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    'body { font-family: Arial; display:flex; align-items:center; justify-content:center; height:100vh; margin:0; background:#f8f9fa; overflow:hidden; }' +
    '.container { text-align:center; padding:30px; }' +
    '.icon { font-size:48px; margin-bottom:15px; display:inline-block; filter: hue-rotate(120deg) saturate(1.2); }' +
    '.message { font-size:16px; color:#333; font-weight:bold; margin-bottom:8px; }' +
    '.submessage { font-size:14px; color:#666; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="icon">❤️</div>' +
    '<div class="message">Подождите, пожалуйста,</div>' +
    '<div class="submessage">готовим таблицу для Вас!</div>' +
    '</div>' +
    '</body></html>';
}

function getDoneHtml() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    'body { font-family: Arial; display:flex; align-items:center; justify-content:center; height:100vh; margin:0; background:#f8f9fa; overflow:hidden; }' +
    '.container { text-align:center; padding:30px; }' +
    '.icon { font-size:48px; margin-bottom:15px; display:inline-block; filter: hue-rotate(120deg) saturate(1.2); }' +
    '.message { font-size:16px; color:#333; font-weight:normal; line-height:1.5; margin-bottom:25px; }' +
    'button { background: #458275; color: #F0EBE2; border: none; padding: 12px 30px; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; }' +
    'button:hover { background: #36695e; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="icon">❤️</div>' +
    '<div class="message">Спасибо большое за ожидание, Ваш бюджет готов!<br>Осталось заполнить категории Расходов и Доходов.</div>' +
    '<button onclick="google.script.host.close()">Спасибо, заполню!</button>' +
    '</div>' +
    '</body></html>';
}

function saveMonth(month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName("Настройка");
  if (!settings) return "Лист 'Настройка' не найден!";
  
  settings.getRange("B16").setValue(month);
  
  var html = HtmlService.createHtmlOutput(getStep2Html())
    .setWidth(380)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
  
  return "Месяц сохранён";
}

function skipSum() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    'body { font-family: Arial; display:flex; align-items:center; justify-content:center; height:100vh; margin:0; background:#f8f9fa; overflow:hidden; }' +
    '.container { text-align:center; padding:30px; }' +
    '.icon { font-size:48px; margin-bottom:15px; display:inline-block; filter: hue-rotate(120deg) saturate(1.2); }' +
    '.message { font-size:15px; color:#333; line-height:1.5; margin-bottom:25px; }' +
    'button { background: #458275; color: #F0EBE2; border: none; padding: 12px 30px; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; }' +
    'button:hover { background: #36695e; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="icon">❤️</div>' +
    '<div class="message">Вы всегда можете указать Начальную сумму позже в ячейке под текстом «Начальная сумма»</div>' +
    '<button onclick="google.script.host.close(); google.script.run.startRename();">Спасибо!</button>' +
    '</div>' +
    '</body></html>'
  ).setWidth(420).setHeight(280);
  
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function startRename() {
  renameSheetsByPeriod();
}

function saveStartSum(sum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName("Настройка");
  if (!settings) return "Лист 'Настройка' не найден!";
  
  var cleanSum = Number(String(sum).replace(/\s/g, '').replace(',', '.'));
  if (isNaN(cleanSum) || cleanSum < 0) cleanSum = 0;
  
  settings.getRange("K14").setValue(cleanSum);
  
  renameSheetsByPeriod();
  
  return "Сумма сохранена";
}

function renameSheetsByPeriod() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var html = HtmlService.createHtmlOutput(getWaitingHtml())
    .setWidth(400)
    .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
  
  Utilities.sleep(300);
  
  try {
    var settings = ss.getSheetByName("Настройка");
    if (!settings) {
      SpreadsheetApp.getUi().alert("Лист 'Настройка' не найден!");
      return;
    }
    
    var startMonthName = String(settings.getRange("B16").getValue()).trim();
    
    var rawYear = settings.getRange("C16").getValue();
    var startYear;
    if (rawYear instanceof Date) {
      startYear = rawYear.getFullYear();
    } else {
      startYear = parseInt(String(rawYear).trim());
    }
    
    if (isNaN(startYear)) {
      SpreadsheetApp.getUi().alert("Не удалось определить год в C16! Значение: " + rawYear);
      return;
    }
    
    var monthMap = {
      "Январь": 0, "Февраль": 1, "Март": 2, "Апрель": 3,
      "Май": 4, "Июнь": 5, "Июль": 6, "Август": 7,
      "Сентябрь": 8, "Октябрь": 9, "Ноябрь": 10, "Декабрь": 11
    };
    
    var shortNames = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн",
                      "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"];
    
    var fullNames = ["Январь", "Февраль", "Март", "Апрель",
                     "Май", "Июнь", "Июль", "Август",
                     "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
    
    var startMonthIndex = monthMap[startMonthName];
    if (startMonthIndex === undefined) {
      SpreadsheetApp.getUi().alert("Не найдено название месяца: '" + startMonthName + "'\nПроверьте B16!");
      return;
    }
    
    var serviceSheets = ["FAQ", "Настройка", "Год", "Данные для диаграмм"];
    var allSheets = ss.getSheets();
    var monthSheets = [];
    
    for (var i = 0; i < allSheets.length; i++) {
      var name = allSheets[i].getName();
      if (serviceSheets.indexOf(name) === -1) {
        monthSheets.push(allSheets[i]);
        if (monthSheets.length === 12) break;
      }
    }
    
    if (monthSheets.length !== 12) {
      SpreadsheetApp.getUi().alert("Найдено " + monthSheets.length + " листов, нужно 12.");
      return;
    }
    
    var newNames = [];
    var newFullNames = [];
    var newYears = [];
    
    var currentMonth = startMonthIndex;
    var currentYear = startYear;
    
    for (var i = 0; i < 12; i++) {
      newNames.push(shortNames[currentMonth] + " " + String(currentYear).slice(-2));
      newFullNames.push(fullNames[currentMonth]);
      newYears.push(currentYear);
      
      currentMonth++;
      if (currentMonth > 11) { currentMonth = 0; currentYear++; }
    }
    
    var tmpPrefix = "_tmp_" + new Date().getTime() + "_";
    var renamedCount = 0;
    
    for (var i = 0; i < 12; i++) {
      try { monthSheets[i].setName(tmpPrefix + i); } catch (e) {}
    }
    
    for (var i = 0; i < 12; i++) {
      try {
        monthSheets[i].setName(newNames[i]);
        monthSheets[i].getRange("B3").setValue(newFullNames[i]);
        monthSheets[i].getRange("D3").setValue(newYears[i]);
        renamedCount++;
      } catch (e) {}
    }
    
    var yearSheet = ss.getSheetByName("Год");
    if (yearSheet) {
      try {
        yearSheet.getRange("D22:O22").setNumberFormat("@");
        yearSheet.getRange("D22:O22").setValues([newNames]);
        
        yearSheet.getRange("D50:O50").setNumberFormat("@");
        yearSheet.getRange("D50:O50").setValues([newNames]);
      } catch (e) {
        Logger.log("Ошибка записи на листе 'Год': " + e.message);
      }
    }
    
    var htmlDone = HtmlService.createHtmlOutput(getDoneHtml())
      .setWidth(450)
      .setHeight(270);
    
    SpreadsheetApp.getUi().showModalDialog(htmlDone, ' ');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("Ошибка: " + e.message);
  }
}
