// Счетчик несохраненных записей
let unsavedEntriesCount = 0;
const DOWNLOAD_THRESHOLD_KEY = "rouletteObserverThreshold";

// Порог скачивания хранится в localStorage страницы и перечитывается при
// каждом добавлении записи. По умолчанию — 10.
function getDownloadThreshold() {
  try {
    const raw = localStorage.getItem(DOWNLOAD_THRESHOLD_KEY);
    const v = raw ? parseInt(raw, 10) : 10;
    return v > 0 ? v : 10;
  } catch (e) {
    return 10;
  }
}

function setDownloadThreshold(n) {
  const v = typeof n === "number" && n > 0 ? Math.floor(n) : 10;
  try {
    localStorage.setItem(DOWNLOAD_THRESHOLD_KEY, String(v));
  } catch (e) {
    console.error("Не удалось записать порог в localStorage:", e);
  }
  return v;
}
// ----- Fairplay data из первой (самой свежей) записи истории -----
function readFairplayData() {
  const out = { id: "", color: "", hash: "", result: "", salt: "" };
  try {
    const entry = document.querySelector(
      ".fairplay_history .roulette_history_entry:first-child"
    );
    if (!entry) return out;

    const idNode = entry.querySelector(".roulette_history_id");
    const colorNode = idNode ? idNode.querySelector("b") : null;
    out.color = colorNode ? colorNode.className : "";
    out.id = idNode ? (idNode.textContent.match(/\d+/) || [""])[0] : "";

    const info = entry.querySelector(".round_info");
    if (!info) return out;

    const read = (sel) => {
      const n = info.querySelector(sel);
      return n ? n.textContent.trim() : "";
    };
    out.hash = read(".roulette_history_value:first-child b");
    out.result =
      read(".roulette_history_value:first-child + div > .half:first-child b") ||
      read(".half:first-child .roulette_history_value b");
    out.salt =
      read(".roulette_history_value:first-child + div > .half + .half b") ||
      read(".half + .half .roulette_history_value b");
  } catch (e) {
    console.error("readFairplayData error:", e);
  }
  return out;
}

// Ждёт, пока в первой записи истории появятся Result и Salt.
// Клик по записи уже сделан; данные приходят после окончания раунда.
function waitForFairplayResult(timeoutMs) {
  return new Promise((resolve) => {
    const start = Date.now();
    const tick = () => {
      const data = readFairplayData();
      const ready = data.result && data.salt;
      if (ready || Date.now() - start > timeoutMs) {
        resolve(data);
      } else {
        setTimeout(tick, 200);
      }
    };
    tick();
  });
}

// ----- Локальное хранилище данных + генерация Excel (XLSX доступен здесь) -----
let allData = [];

function loadStoredData(callback) {
  try {
    if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage({ action: "getExcelData" }, function (response) {
        if (chrome.runtime.lastError || !response || !response.storedData) {
          // background недоступен/пуст — пробуем резервную копию из localStorage
          allData = loadLocalBackup();
        } else {
          try {
            allData = JSON.parse(response.storedData) || [];
          } catch (e) {
            allData = [];
          }
        }
        console.log("Загружено сохранённых записей:", allData.length);
        unsavedEntriesCount = allData.length % getDownloadThreshold();
        if (callback) callback();
      });
    } else {
      allData = loadLocalBackup();
      console.log("Загружено из localStorage (background недоступен):", allData.length);
      unsavedEntriesCount = allData.length % getDownloadThreshold();
      if (callback) callback();
    }
  } catch (e) {
    allData = loadLocalBackup();
    unsavedEntriesCount = allData.length % getDownloadThreshold();
    if (callback) callback();
  }
}

// Чтение резервной копии из localStorage (если chrome.runtime/storage недоступны)
function loadLocalBackup() {
  try {
    const raw = localStorage.getItem(LOCAL_BACKUP_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function buildExcelBase64() {
  if (typeof XLSX === "undefined") {
    console.error("Библиотека XLSX не загружена в content script!");
    return null;
  }
  try {
    const ws = XLSX.utils.json_to_sheet(allData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Данные");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
    return btoa(wbout);
  } catch (e) {
    console.error("Ошибка при создании Excel:", e);
    return null;
  }
}

function appendEntry(newEntry) {
  allData.push(newEntry);
  unsavedEntriesCount++;

  // Каждый раз заново читаем порог из localStorage: если пользователь поменял
  // его (например с 10 на 40) — для следующего скачивания возьмётся 40.
  const threshold = getDownloadThreshold();
  const shouldDownload = unsavedEntriesCount >= threshold;

  // Сохраняем в фоне в chrome.storage (через background script).
  // Если chrome.runtime недоступен (контекст расширения «протух», скрипт
  // запущен в основной области и т.п.) — не падаем, а делаем резервную
  // копию в localStorage страницы.
  try {
    if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage(
        { action: "appendToExcel", newEntry: newEntry },
        function (response) {
          if (chrome.runtime.lastError) {
            console.error("Ошибка сохранения в storage:", chrome.runtime.lastError.message);
            fallbackSaveLocal();
          } else if (response) {
            console.log("Сохранено в storage. Всего:", response.dataLength);
            persistLocalBackup();
          }
        }
      );
    } else {
      console.warn("chrome.runtime недоступен — использую резервное хранение localStorage.");
      fallbackSaveLocal();
    }
  } catch (e) {
    console.error("Ошибка отправки в background:", e);
    fallbackSaveLocal();
  }

  // Скачиваем Excel (строим локально, т.к. XLSX есть только в content script)
  if (shouldDownload) {
    const b64 = buildExcelBase64();
    if (b64) {
      downloadExcelFileFromBase64(b64);
      unsavedEntriesCount = 0;
      persistLocalBackup();
      console.log(`Файл Excel скачан. Счётчик сброшен. Следующий порог: ${getDownloadThreshold()}`);
    }
  } else {
    console.log(
      `Данные сохранены. Несохранённых: ${unsavedEntriesCount}/${threshold}. Для скачивания: downloadExcelFile()`
    );
  }
}

// --- Резервное хранение в localStorage (если background недоступен) ---
const LOCAL_BACKUP_KEY = "rouletteObserverLocal";

function persistLocalBackup() {
  try {
    localStorage.setItem(LOCAL_BACKUP_KEY, JSON.stringify(allData));
  } catch (e) {
    console.error("Не удалось сохранить резервную копию в localStorage:", e);
  }
}

function fallbackSaveLocal() {
  persistLocalBackup();
}

function startObserver() {
  //   console.log("Content script запущен");
  const timerElement = document.querySelector(".roulette_timer");
  const targetText = "14.90";

  if (!timerElement) {
    console.log("Элемент .roulette_timer не найден на странице.");
    return;
  }

  //   console.log("Элемент .roulette_timer найден, начинаем наблюдение");

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      if (
        mutation.type === "childList" &&
        timerElement.textContent.trim() === targetText
      ) {
        // console.log("Обнаружено совпадение текста:", targetText);

        const targetElementClass = document.querySelector(
          ".fairplay_history .roulette_history_entry:first-child .roulette_history_id b",
        );
        const targetElementClick = document.querySelector(
          ".fairplay_history .roulette_history_entry:first-child .roulette_history_id",
        );

        let className = targetElementClass
          ? targetElementClass.className
          : "Не найдено";

        // console.log("Получен класс элемента:", className);

        observer.disconnect();
        // console.log("Наблюдатель остановлен");

        // Сначала выполняем клик, чтобы загрузить дополнительные данные
        if (targetElementClick) {
          targetElementClick.click();
          //   console.log("Клик выполнен на элемент:", targetElementClick);

          // Добавляем задержку перед получением textContent, чтобы дать время для загрузки данных
          setTimeout(async () => {
            // Ждём появления Result/Salt в первой записи истории
            // (клик сделан; данные приходят после окончания раунда)
            const fairplay = await waitForFairplayResult(20000);

            const newEntry = {
              Дата: new Date().toLocaleString(),
              Цвет: fairplay.color || className,
              Число: fairplay.result || "Не найдено",
              "Раунд ID": fairplay.id,
              Hash: fairplay.hash,
              Salt: fairplay.salt,
            };

            // Сохраняем запись, при достижении порога — скачиваем Excel
            appendEntry(newEntry);

            // console.log("Перезапускаем наблюдатель через 18 секунд");
            setTimeout(startObserver, 18000);
          }, 1000); // Задержка в 1 секунду для загрузки данных после клика
        } else {
          console.log("Элемент для клика не найден");

          // Если элемент для клика не найден, все равно продолжаем с пустым textContent
          const fairplay = readFairplayData();
          const newEntry = {
            Дата: new Date().toLocaleString(),
            Цвет: fairplay.color || className,
            Число: fairplay.result || "Не найдено (элемент для клика не найден)",
            "Раунд ID": fairplay.id,
            Hash: fairplay.hash,
            Salt: fairplay.salt,
          };

          // Сохраняем запись, при достижении порога — скачиваем Excel
          appendEntry(newEntry);

          //   console.log("Перезапускаем наблюдатель через 18 секунд");
          setTimeout(startObserver, 18000);
        }
      }
    });
  });

  observer.observe(timerElement, { childList: true, subtree: true });

  //   console.log(
  //     'Сниппет запущен. Ожидание текста "14.90" в элементе .roulette_timer...'
  //   );
}

// Функция для скачивания Excel файла по команде из консоли
function downloadExcelFile() {
  const b64 = buildExcelBase64();
  if (b64) {
    downloadExcelFileFromBase64(b64);
    unsavedEntriesCount = 0;
    console.log("Файл Excel скачан. Счётчик сброшен.");
  } else {
    console.log("Не удалось построить Excel.");
  }
}

// Функция для скачивания Excel файла из base64 данных
function downloadExcelFileFromBase64(base64data) {
  try {
    // console.log("Начинаем скачивание Excel файла из base64 данных");

    // Преобразуем base64 в бинарные данные
    const binary = atob(base64data);
    const array = [];
    for (let i = 0; i < binary.length; i++) {
      array.push(binary.charCodeAt(i));
    }

    // Создаем Blob из бинарных данных
    const blob = new Blob([new Uint8Array(array)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    // Создаем URL для Blob
    const url = URL.createObjectURL(blob);

    // Создаем элемент <a> для скачивания
    const a = document.createElement("a");
    a.href = url;
    a.download = "roulette_data.xlsx";
    a.style.display = "none";

    // Добавляем элемент в DOM
    document.body.appendChild(a);

    // Имитируем клик
    a.click();

    // Удаляем элемент из DOM и освобождаем URL
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      console.log("Файл Excel успешно скачан");
    }, 100);
  } catch (e) {
    console.error("Ошибка при скачивании Excel файла:", e);
  }
}

// Функция для получения текущего количества несохраненных записей
function getUnsavedCount() {
  console.log(
    `Текущее количество несохраненных записей: ${unsavedEntriesCount}/${getDownloadThreshold()}`,
  );
  return unsavedEntriesCount;
}

// Функция для сброса счетчика несохраненных записей
function resetUnsavedCount() {
  unsavedEntriesCount = 0;
  console.log("Счетчик несохраненных записей сброшен");
  return unsavedEntriesCount;
}

// Функция для изменения порога автоматического скачивания.
// Значение сохраняется в localStorage страницы и при каждом добавлении
// записи порог перечитывается заново.
function setAutoDownloadThreshold(newThreshold) {
  const oldThreshold = getDownloadThreshold();
  const v = setDownloadThreshold(newThreshold);
  if (typeof newThreshold === "number" && newThreshold > 0) {
    console.log(
      `Порог автоматического скачивания изменен с ${oldThreshold} на ${v} (сохранён в localStorage)`,
    );
  } else {
    console.error("Ошибка: новый порог должен быть положительным числом");
  }
  return v;
}

// Функция для получения текущего порога скачивания из localStorage
function getAutoDownloadThreshold() {
  return getDownloadThreshold();
}

// Делаем функции доступными глобально
window.downloadExcelFile = downloadExcelFile;
window.getUnsavedCount = getUnsavedCount;
window.resetUnsavedCount = resetUnsavedCount;
window.setAutoDownloadThreshold = setAutoDownloadThreshold;
window.getAutoDownloadThreshold = getAutoDownloadThreshold;

// Загружаем сохранённые данные, затем запускаем основной наблюдатель
loadStoredData(function () {
  startObserver();
});

// Выводим инструкцию в консоль
// console.log(`Расширение запущено. Порог скачивания хранится в localStorage (rouletteObserverThreshold).`);
// console.log("Доступные команды в консоли:");
console.log("- downloadExcelFile() - скачать файл Excel с текущими данными");
console.log(
  "- getUnsavedCount() - получить текущее количество несохраненных записей",
);
console.log("- resetUnsavedCount() - сбросить счетчик несохраненных записей");
console.log(
  "- setAutoDownloadThreshold(число) - изменить порог автоматического скачивания",
);


let balance = parseFloat(document.querySelector('.balance_cont .balance').childNodes[0].textContent.trim());
let balanceGem = parseFloat(document.querySelector('.bonus_cont .bonus').childNodes[0].textContent.trim());


// Функция для суммирования значений в колонке при наличии активной ссылки
function sumHistoryColumn() {
  const activeLink = document.querySelector('a[data-filter="active"].active');
  if (!activeLink) return;

  const table = document.querySelector("table.history");
  if (!table) return;

  // Добавляем крестики в четвёртую колонку каждой строки
  addDeleteButtonsToRows(table);

  const rows = table.querySelectorAll("tbody tr td:nth-child(3)");
  let sum = 0;
  let sumGem = 0;
  rows.forEach((row) => {
    // Проверяем, не помечена ли строка на удаление
    const parentRow = row.parentElement;
    if (parentRow && parentRow.classList.contains('deleted-row')) return;
    
    const count = row.firstChild.data;
    if (count) {
      const num = parseFloat(count.trim());
      if (!isNaN(num)) {
        if (row.querySelector(" i").className == 'gems') {
          sumGem += num;
        } else if (row.querySelector(" i").className == 'rr') {
          sum += num;
        }
      }
    }
  });
  const rowsWon = table.querySelectorAll("tbody tr td:nth-child(5)");
  let sumWon = 0;
  let sumWonGem = 0;
  rowsWon.forEach((row) => {
    const parentRow = row.parentElement;
    if (parentRow && parentRow.classList.contains('deleted-row')) return;
    
    const won = row.firstChild.data;
    if (won) {
      const numWon = parseFloat(won.trim());
      if (!isNaN(numWon)) {
        if (row.querySelector(" i").className == 'gems') {
          sumWonGem += numWon;
        } else if (row.querySelector(" i").className == 'rr') {
          sumWon += numWon;
        }
      }
    }
  });

  const headerCell = table.querySelector("thead tr th:nth-child(3)");
  const headerWon = table.querySelector("thead tr th:nth-child(5)");

  if (headerCell) {
    headerCell.innerHTML  = sum.toString() + "<i class='rr'></i></br>" + sumGem.toString() + "<i class='gems'></i>";
    headerCell.style.transform = "translateY(-10px)";
    headerCell.style.lineHeight = "16px";
    headerWon.innerHTML  = sumWon.toString() + "<i class='rr'></i></br>" + sumWonGem.toString() + "<i class='gems'></i>";
    headerWon.style.transform = "translateY(-10px)";
    headerWon.style.lineHeight = "16px";

    let totalSum = balance + sum;
    let totalSumGem = balanceGem + sumGem;
    let totalWon = balance + sumWon;
    let totalWonGem = balanceGem + sumWonGem;


    document.querySelector('.balance_cont .balance').childNodes[0].textContent = balance + ' (' + totalSum + ' / '+ totalWon + ')';

    document.querySelector('.bonus_cont .bonus').childNodes[0].textContent =  balanceGem + ' (' + totalSumGem + ' / '+ totalWonGem + ')';
  }
}

// Функция для добавления кнопок удаления в четвёртую колонку
function addDeleteButtonsToRows(table) {
  const rows = table.querySelectorAll("tbody tr");
  
  rows.forEach((row) => {
    // Проверяем, есть ли уже кнопка удаления в этой строке
    const secondCell = row.querySelector("td:nth-child(2)");
    if (!secondCell) return;
    if (row.classList.contains('express_icon')) return;
    
    // Если кнопка уже есть, пропускаем
    if (secondCell.querySelector('.delete-row-btn')) return;
    
    // Создаём кнопку-крестик
    const deleteBtn = document.createElement('button');
    deleteBtn.innerHTML = '✕';
    deleteBtn.className = 'delete-row-btn';
    deleteBtn.style.cssText = `
      background: none;
      border: none;
      color: red;
      font-size: 13px;
      font-weight: bold;
      cursor: pointer;
      padding: 0;
      margin: 0;
      transition: transform 0.2s;
    `;
    
    // Добавляем эффект при наведении
    deleteBtn.onmouseenter = () => {
      deleteBtn.style.transform = 'scale(1.2)';
    };
    deleteBtn.onmouseleave = () => {
      deleteBtn.style.transform = 'scale(1)';
    };
    
    // Обработчик удаления строки
    deleteBtn.onclick = (e) => {
      e.stopPropagation();
      
      if (row.classList.contains('deleted-row')) {
        row.classList.remove('deleted-row');
        row.style.opacity = '1';
        row.style.textDecoration = 'unset';
        
      } else {
        row.classList.add('deleted-row');
        row.style.opacity = '0.5';
        row.style.textDecoration = 'line-through';
      }
      
      
      // Пересчитываем сумму после удаления
      setTimeout(() => {
        sumHistoryColumn();
      }, 100);
    };
    
    // Очищаем ячейку и добавляем кнопку
    // secondCell.innerHTML = '';
    secondCell.appendChild(deleteBtn);
    // secondCell.style.textAlign = 'center';
    // secondCell.style.verticalAlign = 'middle';
  });
}

// Функция для восстановления всех удалённых строк (опционально)
// function restoreAllRows() {
//   const table = document.querySelector("table.history");
//   if (!table) return;
  
//   const deletedRows = table.querySelectorAll("tbody tr.deleted-row");
//   deletedRows.forEach((row) => {
//     row.classList.remove('deleted-row');
//     row.style.opacity = '';
//     row.style.textDecoration = '';
//   });
  
//   // Пересчитываем сумму
//   setTimeout(() => {
//     sumHistoryColumn();
//   }, 100);
// }

// Добавляем кнопку для восстановления всех строк (опционально)
// function addRestoreButton() {
//   const table = document.querySelector("table.history");
//   if (!table) return;
  
//   // Проверяем, есть ли уже кнопка восстановления
//   if (document.querySelector('.restore-all-btn')) return;
  
//   const restoreBtn = document.createElement('button');
//   restoreBtn.innerHTML = '🔄 Восстановить все строки';
//   restoreBtn.className = 'restore-all-btn';
//   restoreBtn.style.cssText = `
//     margin: 10px;
//     padding: 5px 10px;
//     background: #4CAF50;
//     color: white;
//     border: none;
//     border-radius: 3px;
//     cursor: pointer;
//     font-size: 14px;
//   `;
  
//   restoreBtn.onclick = restoreAllRows;
  
//   // Вставляем кнопку перед таблицей
//   table.parentNode.insertBefore(restoreBtn, table);
// }

if (document.querySelector('a[data-filter="active"]')) {
  document
    .querySelector('a[data-filter="active"]')
    .addEventListener("click", function () {
      setTimeout(() => {
        sumHistoryColumn();
        // addRestoreButton(); // Добавляем кнопку восстановления
      }, 1200);
    });
}