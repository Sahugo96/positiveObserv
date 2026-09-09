console.log("Background script запущен");

// Глобальная переменная для хранения данных
let allData = [];
const STORAGE_KEY = "rouletteAllData";

console.log("XLSX-генерация выполняется в content.js (там подключена библиотека).");

// Функция для загрузки существующих данных из локального хранилища
async function loadExistingFile() {
  return new Promise((resolve) => {
    chrome.storage.local.get([STORAGE_KEY], function (result) {
      const raw = result[STORAGE_KEY];
      if (raw) {
        try {
          allData = JSON.parse(raw);
        } catch (e) {
          console.error("Ошибка при парсинге данных:", e);
          allData = [];
        }
      } else {
        allData = [];
      }
      resolve();
    });
  });
}

// Функция для сохранения данных в локальное хранилище
function persistData() {
  chrome.storage.local.set({ [STORAGE_KEY]: JSON.stringify(allData) }, function () {
    if (chrome.runtime.lastError) {
      console.error("Ошибка при сохранении данных в хранилище:", chrome.runtime.lastError.message);
    } else {
      console.log("Данные сохранены в локальное хранилище. Записей:", allData.length);
    }
  });
}

// Обработчик сообщений
chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
  console.log("Получено сообщение:", message.action);

  if (message.action === "appendToExcel") {
    loadExistingFile().then(() => {
      allData.push(message.newEntry);
      console.log("Добавлена новая запись, всего записей:", allData.length);
      persistData();
      sendResponse({
        status: "success",
        dataLength: allData.length,
      });
    });
    return true; // Важно для асинхронного ответа
  } else if (message.action === "getExcelData") {
    loadExistingFile().then(() => {
      sendResponse({
        status: "success",
        dataLength: allData.length,
        storedData: JSON.stringify(allData),
      });
    });
    return true; // Важно для асинхронного ответа
  } else if (message.action === "clearAllData") {
    allData = [];
    chrome.storage.local.remove(STORAGE_KEY, function () {
      if (chrome.runtime.lastError) {
        console.error("Ошибка очистки storage:", chrome.runtime.lastError.message);
        sendResponse({ status: "error" });
      } else {
        console.log("Все данные в storage очищены.");
        sendResponse({ status: "success" });
      }
    });
    return true; // Важно для асинхронного ответа
  }
});

// Инициализация при запуске
console.log("Инициализация background script");
loadExistingFile();
