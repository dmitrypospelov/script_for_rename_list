function renameSheetsByPeriod() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName("Настройка");
  
  if (!settings) {
    SpreadsheetApp.getUi().alert("Лист 'Настройка' не найден!");
    return;
  }
  
  // === 1. Читаем параметры ===
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
  
  // === 2. Словари ===
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
  
  // === 3. Проверяем месяц ===
  var startMonthIndex = monthMap[startMonthName];
  if (startMonthIndex === undefined) {
    SpreadsheetApp.getUi().alert("Не найдено название месяца: '" + startMonthName + "'\nПроверьте B16!");
    return;
  }
  
  // === 4. Определяем листы для переименования ===
  // Служебные листы, которые НЕ трогаем
  var serviceSheets = ["FAQ", "Настройка", "Год", "Данные для диаграмм"];
  
  // Берём ВСЕ листы слева-направо, кроме служебных
  var allSheets = ss.getSheets();
  var monthSheets = [];
  
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (serviceSheets.indexOf(name) === -1) {
      monthSheets.push(allSheets[i]);
    }
  }
  
  // Проверяем, что нашли ровно 12 листов
  if (monthSheets.length !== 12) {
    SpreadsheetApp.getUi().alert(
      "Найдено " + monthSheets.length + " листов для переименования, а нужно 12.\n\n" +
      "Сейчас найдены: " + monthSheets.map(function(s) { return s.getName(); }).join(", ") + "\n\n" +
      "Проверьте, что:\n" +
      "• Все 12 месячных листов не входят в служебные\n" +
      "• В serviceSheets указаны все служебные листы:\n" +
      serviceSheets.join(", ")
    );
    return;
  }
  
  // === 5. Переименовываем ===
  var currentMonth = startMonthIndex;
  var currentYear = startYear;
  var renamedCount = 0;
  
  for (var i = 0; i < 12; i++) {
    var sheet = monthSheets[i];
    
    var yearShort = String(currentYear).slice(-2);
    var newName = shortNames[currentMonth] + " " + yearShort;
    
    var existing = ss.getSheetByName(newName);
    if (existing && existing.getName() !== sheet.getName()) {
      Logger.log("Имя '" + newName + "' занято, пропускаем.");
    } else {
      sheet.setName(newName);
      sheet.getRange("B3").setValue(fullNames[currentMonth]);
      sheet.getRange("D3").setValue(currentYear);
      renamedCount++;
    }
    
    currentMonth++;
    if (currentMonth > 11) {
      currentMonth = 0;
      currentYear++;
    }
  }
  
  SpreadsheetApp.getUi().alert("Готово! Переименовано листов: " + renamedCount);
}