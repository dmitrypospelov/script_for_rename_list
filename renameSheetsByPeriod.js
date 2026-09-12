function renameSheetsByPeriod() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // === 1. Показываем окно "Подождите..." БЕЗ кнопки ===
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<style>' +
    'body {' +
    '  font-family: Arial, sans-serif;' +
    '  display: flex;' +
    '  align-items: center;' +
    '  justify-content: center;' +
    '  height: 100vh;' +
    '  margin: 0;' +
    '  background: #f8f9fa;' +
    '  overflow: hidden;' +
    '}' +
    '.container {' +
    '  text-align: center;' +
    '  padding: 30px;' +
    '}' +
    '.icon {' +
    '  font-size: 48px;' +
    '  margin-bottom: 15px;' +
    '  display: inline-block;' +
    '  filter: hue-rotate(120deg) saturate(1.2);' +
    '}' +
    '.message {' +
    '  font-size: 16px;' +
    '  color: #333;' +
    '  font-weight: bold;' +
    '  margin-bottom: 8px;' +
    '}' +
    '.submessage {' +
    '  font-size: 14px;' +
    '  color: #666;' +
    '}' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="container">' +
    '<div class="icon">❤️</div>' +
    '<div class="message">Подождите, пожалуйста,</div>' +
    '<div class="submessage">переименовываем Ваши листы</div>' +
    '</div>' +
    '</body>' +
    '</html>'
  ).setWidth(400).setHeight(280);
  
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
  
  // === 2. Небольшая пауза ===
  Utilities.sleep(300);
  
  // === 3. Основная логика ===
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
    
    // === 4. Показываем окно "Готово!" с кнопкой "Спасибо!" ===
    var htmlDone = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<style>' +
      'body {' +
      '  font-family: Arial, sans-serif;' +
      '  display: flex;' +
      '  align-items: center;' +
      '  justify-content: center;' +
      '  height: 100vh;' +
      '  margin: 0;' +
      '  background: #f8f9fa;' +
      '  overflow: hidden;' +
      '}' +
      '.container {' +
      '  text-align: center;' +
      '  padding: 30px;' +
      '}' +
      '.icon {' +
      '  font-size: 48px;' +
      '  margin-bottom: 15px;' +
      '  display: inline-block;' +
      '  filter: hue-rotate(120deg) saturate(1.2);' +
      '}' +
      '.message {' +
      '  font-size: 16px;' +
      '  color: #333;' +
      '  font-weight: bold;' +
      '  margin-bottom: 25px;' +
      '}' +
      'button {' +
      '  background: #458275;' +
      '  color: #F0EBE2;' +
      '  border: none;' +
      '  padding: 12px 30px;' +
      '  border-radius: 6px;' +
      '  font-size: 15px;' +
      '  font-weight: bold;' +
      '  cursor: pointer;' +
      '  transition: background 0.2s;' +
      '}' +
      'button:hover {' +
      '  background: #36695e;' +
      '}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="container">' +
      '<div class="icon">❤️</div>' +
      '<div class="message">Ваши листы переименованы</div>' +
      '<button onclick="google.script.host.close()">Спасибо!</button>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(400).setHeight(260);
    
    SpreadsheetApp.getUi().showModalDialog(htmlDone, ' ');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("Ошибка: " + e.message);
  }
}
