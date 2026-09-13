function startSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(400)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function mainEdit(e) {
  Logger.log("=== mainEdit вызван ===");
  
  if (!e || !e.range) {
    Logger.log("Нет e — выходим");
    return;
  }
  
  var sheet = e.range.getSheet();
  Logger.log("Лист: " + sheet.getName());
  Logger.log("Ячейка: " + e.range.getA1Notation());
  Logger.log("Значение: [" + e.value + "]");
  Logger.log("Тип: " + typeof e.value);
  
  if (sheet.getName() !== "Настройка") {
    Logger.log("Не Настройка — выходим");
    return;
  }
  if (e.range.getA1Notation() !== "C9") {
    Logger.log("Не C9 — выходим");
    return;
  }
  if (e.value !== "TRUE" && e.value !== true) {
    Logger.log("Не TRUE — выходим");
    return;
  }
  
  Logger.log("Проверки пройдены — снимаем галочку");
  e.range.setValue(false);
  SpreadsheetApp.flush();
  
  Logger.log("Запускаем renameSheetsOnly...");
  var result = renameSheetsOnly();
  Logger.log("Результат renameSheetsOnly: [" + result + "]");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (result === "") {
    ss.toast('Готово! Ваши листы переименованы', '✅ Успех', 5);
  } else {
    ss.toast(result, '⚠️ Ошибка', 10);
  }
  
  Logger.log("=== mainEdit завершён ===");
}

function saveMonth(month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName("Настройка");
  if (!settings) return "Лист 'Настройка' не найден!";
  
  var oldMonth = String(settings.getRange("B16").getValue()).trim();
  settings.getRange("B16").setValue(month);
  SpreadsheetApp.flush();
  
  if (oldMonth !== month) {
    return renameSheetsOnly();
  }
  
  return "";
}

function saveStartSum(sum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName("Настройка");
  if (!settings) return "Лист 'Настройка' не найден!";
  
  var cleanSum = Number(String(sum).replace(/\s/g, '').replace(',', '.'));
  if (isNaN(cleanSum) || cleanSum < 0) cleanSum = 0;
  
  settings.getRange("K14").setValue(cleanSum);
  SpreadsheetApp.flush();
  
  return "Сумма сохранена";
}

function renameSheetsOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  
  try {
    var settings = ss.getSheetByName("Настройка");
    if (!settings) {
      Logger.log("Нет листа 'Настройка'");
      return "Нет листа 'Настройка'";
    }
    
    var startMonthName = String(settings.getRange("B16").getValue()).trim();
    log.push("Месяц: [" + startMonthName + "]");
    
    var rawYear = settings.getRange("C16").getValue();
    var startYear;
    if (rawYear instanceof Date) {
      startYear = rawYear.getFullYear();
    } else {
      startYear = parseInt(String(rawYear).trim());
    }
    log.push("Год: " + startYear);
    
    if (isNaN(startYear)) {
      Logger.log(log.join(" | ") + " | Год не число");
      return "Год в C16 — не число";
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
      Logger.log(log.join(" | ") + " | Месяц не найден");
      return "Месяц не найден: " + startMonthName;
    }
    
    var serviceSheets = ["FAQ", "Настройка", "Год", "Данные для диаграмм"];
    var allSheets = ss.getSheets();
    var monthSheets = [];
    
    for (var i = 0; i < allSheets.length; i++) {
      var name = allSheets[i].getName();
      if (serviceSheets.indexOf(name) === -1) {
        monthSheets.push(allSheets[i]);
        log.push("Месячный лист: " + name);
      }
    }
    
    log.push("Найдено листов: " + monthSheets.length);
    
    if (monthSheets.length !== 12) {
      Logger.log(log.join(" | "));
      return "Найдено " + monthSheets.length + " листов, нужно 12";
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
    
    log.push("Новые имена: " + newNames.join(", "));
    
    var tmpPrefix = "_tmp_" + new Date().getTime() + "_";
    
    for (var i = 0; i < 12; i++) {
      try { monthSheets[i].setName(tmpPrefix + i); } catch (err) {
        log.push("Ошибка tmp " + i + ": " + err.message);
      }
    }
    
    for (var i = 0; i < 12; i++) {
      try {
        monthSheets[i].setName(newNames[i]);
        monthSheets[i].getRange("B3").setValue(newFullNames[i]);
        monthSheets[i].getRange("D3").setValue(newYears[i]);
      } catch (err) {
        log.push("Ошибка rename " + i + ": " + err.message);
      }
    }
    
    var yearSheet = ss.getSheetByName("Год");
    if (yearSheet) {
      try {
        yearSheet.getRange("D22:O22").setNumberFormat("@");
        yearSheet.getRange("D22:O22").setValues([newNames]);
        
        yearSheet.getRange("D50:O50").setNumberFormat("@");
        yearSheet.getRange("D50:O50").setValues([newNames]);
      } catch (err) {
        log.push("Ошибка Год: " + err.message);
      }
    }
    
    Logger.log("УСПЕХ: " + log.join(" | "));
    return "";
    
  } catch (err) {
    Logger.log("КРИТИЧЕСКАЯ ОШИБКА: " + err.message);
    return "Ошибка: " + err.message;
  }
}
