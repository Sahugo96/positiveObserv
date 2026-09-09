"use strict";

const STORAGE_KEY = "rouletteAllData";

let storedEntries = [];

// Loads and validates the saved roulette entries from Chrome extension storage.
function readFromStorage() {
  return new Promise((resolve) => {
    chrome.storage.local.get([STORAGE_KEY], (result) => {
      const rawEntries = result[STORAGE_KEY];

      if (!rawEntries) {
        storedEntries = [];
        resolve();
        return;
      }

      try {
        const parsedEntries = JSON.parse(rawEntries);
        storedEntries = Array.isArray(parsedEntries) ? parsedEntries : [];
      } catch (error) {
        console.error("Ошибка при чтении сохранённых данных:", error);
        storedEntries = [];
      }

      resolve();
    });
  });
}

// Serializes the in-memory entries and writes them to Chrome extension storage.
function writeToStorage() {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(
      { [STORAGE_KEY]: JSON.stringify(storedEntries) },
      () => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
          return;
        }

        console.log(
          "Данные сохранены в локальное хранилище. Записей:",
          storedEntries.length,
        );
        resolve();
      },
    );
  });
}

// Removes the persisted entry collection after the user clears it.
function removeFromStorage() {
  return new Promise((resolve, reject) => {
    chrome.storage.local.remove(STORAGE_KEY, () => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
        return;
      }

      resolve();
    });
  });
}

// Adds one entry to storage and returns the updated entry count to the sender.
async function appendEntry(message) {
  await readFromStorage();
  storedEntries.push(message.newEntry);
  await writeToStorage();

  return {
    status: "success",
    dataLength: storedEntries.length,
  };
}

// Retrieves all saved entries in the format expected by the content script.
async function getEntries() {
  await readFromStorage();

  return {
    status: "success",
    dataLength: storedEntries.length,
    storedData: JSON.stringify(storedEntries),
  };
}

// Clears both the in-memory cache and the durable storage value.
async function clearEntries() {
  storedEntries = [];
  await removeFromStorage();
  console.log("Все данные в storage очищены.");

  return { status: "success" };
}

const messageHandlers = {
  appendToExcel: appendEntry,
  getExcelData: getEntries,
  clearAllData: clearEntries,
};

// Routes supported content-script messages and keeps the response channel open.
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  const handler = Object.hasOwn(messageHandlers, message.action)
    ? messageHandlers[message.action]
    : null;

  if (!handler) {
    return false;
  }

  handler(message)
    .then(sendResponse)
    .catch((error) => {
      console.error(`Ошибка при выполнении ${message.action}:`, error);
      sendResponse({ status: "error" });
    });

  return true;
});

// Populate the in-memory cache whenever the service worker starts.
readFromStorage().then(() => {
  console.log("Roulette Observer background script запущен.");
});
