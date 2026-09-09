// Keeps roulette capture and Excel export isolated from other content features.
(function initializeRouletteObserver() {
  "use strict";

  const DEFAULT_DOWNLOAD_THRESHOLD = 10;
  const DOWNLOAD_THRESHOLD_KEY = "rouletteObserverThreshold";
  const LOCAL_BACKUP_KEY = "rouletteObserverLocal";
  const DOWNLOAD_FILE_NAME = "roulette_data.xlsx";
  const TIMER_TRIGGER_TEXT = "14.90";
  const FAIRPLAY_TIMEOUT_MS = 20_000;
  const ROUND_DETAILS_DELAY_MS = 1_000;
  const OBSERVER_RESTART_DELAY_MS = 18_000;

  const selectors = {
    timer: ".roulette_timer",
    latestHistoryEntry:
      ".fairplay_history .roulette_history_entry:first-child",
    historyId: ".roulette_history_id",
    historyHash: ".roulette_history_value:first-child b",
    historyResult:
      ".roulette_history_value:first-child + div > .half:first-child b",
    historyResultFallback: ".half:first-child .roulette_history_value b",
    historySalt:
      ".roulette_history_value:first-child + div > .half + .half b",
    historySaltFallback: ".half + .half .roulette_history_value b",
  };

  let entries = [];
  let unsavedEntriesCount = 0;

  // Reads the configured auto-download threshold, falling back to the default.
  function getDownloadThreshold() {
    try {
      const rawThreshold = localStorage.getItem(DOWNLOAD_THRESHOLD_KEY);
      const threshold = rawThreshold
        ? Number.parseInt(rawThreshold, 10)
        : DEFAULT_DOWNLOAD_THRESHOLD;

      return threshold > 0 ? threshold : DEFAULT_DOWNLOAD_THRESHOLD;
    } catch (_error) {
      return DEFAULT_DOWNLOAD_THRESHOLD;
    }
  }

  // Validates and stores a new auto-download threshold for this site.
  function setDownloadThreshold(value) {
    const threshold =
      typeof value === "number" && value > 0
        ? Math.floor(value)
        : DEFAULT_DOWNLOAD_THRESHOLD;

    try {
      localStorage.setItem(DOWNLOAD_THRESHOLD_KEY, String(threshold));
    } catch (error) {
      console.error("Не удалось записать порог в localStorage:", error);
    }

    return threshold;
  }

  // Returns trimmed text for an element found within a parent, or an empty string.
  function readText(parent, selector) {
    return parent.querySelector(selector)?.textContent.trim() ?? "";
  }

  // Extracts the latest round's visible Fairplay fields from the history panel.
  function readFairplayData() {
    const fairplay = { id: "", color: "", hash: "", result: "", salt: "" };

    try {
      const entry = document.querySelector(selectors.latestHistoryEntry);
      if (!entry) {
        return fairplay;
      }

      const idElement = entry.querySelector(selectors.historyId);
      fairplay.color = idElement?.querySelector("b")?.className ?? "";
      fairplay.id = idElement?.textContent.match(/\d+/)?.[0] ?? "";

      const roundInfo = entry.querySelector(".round_info");
      if (!roundInfo) {
        return fairplay;
      }

      fairplay.hash = readText(roundInfo, selectors.historyHash);
      fairplay.result =
        readText(roundInfo, selectors.historyResult) ||
        readText(roundInfo, selectors.historyResultFallback);
      fairplay.salt =
        readText(roundInfo, selectors.historySalt) ||
        readText(roundInfo, selectors.historySaltFallback);
    } catch (error) {
      console.error("Не удалось прочитать данные Fairplay:", error);
    }

    return fairplay;
  }

  // Polls the latest history entry until its result and salt are available.
  function waitForFairplayResult(timeoutMs) {
    return new Promise((resolve) => {
      const startedAt = Date.now();

      // Resolves polling when Fairplay details arrive or the timeout is reached.
      function checkResult() {
        const fairplay = readFairplayData();
        const isReady = fairplay.result && fairplay.salt;
        const hasTimedOut = Date.now() - startedAt > timeoutMs;

        if (isReady || hasTimedOut) {
          resolve(fairplay);
          return;
        }

        setTimeout(checkResult, 200);
      }

      checkResult();
    });
  }

  // Reads the page-local backup used when extension messaging is unavailable.
  function loadLocalBackup() {
    try {
      const rawEntries = localStorage.getItem(LOCAL_BACKUP_KEY);
      if (!rawEntries) {
        return [];
      }

      const parsedEntries = JSON.parse(rawEntries);
      return Array.isArray(parsedEntries) ? parsedEntries : [];
    } catch (_error) {
      return [];
    }
  }

  // Stores the current entries as a page-local recovery copy.
  function persistLocalBackup() {
    try {
      localStorage.setItem(LOCAL_BACKUP_KEY, JSON.stringify(entries));
    } catch (error) {
      console.error("Не удалось сохранить резервную копию в localStorage:", error);
    }
  }

  // Checks whether this content script can communicate with the service worker.
  function canMessageBackground() {
    return Boolean(
      typeof chrome !== "undefined" &&
        chrome.runtime?.sendMessage &&
        chrome.runtime.id,
    );
  }

  // Updates the download counter after entries have been loaded, then starts work.
  function finishLoading(callback) {
    unsavedEntriesCount = entries.length % getDownloadThreshold();
    console.log("Загружено сохранённых записей:", entries.length);
    callback();
  }

  // Loads entries from extension storage, with the local backup as a fallback.
  function loadStoredData(callback) {
    if (!canMessageBackground()) {
      entries = loadLocalBackup();
      finishLoading(callback);
      return;
    }

    try {
      chrome.runtime.sendMessage({ action: "getExcelData" }, (response) => {
        if (chrome.runtime.lastError || !response?.storedData) {
          entries = loadLocalBackup();
          finishLoading(callback);
          return;
        }

        try {
          const parsedEntries = JSON.parse(response.storedData);
          entries = Array.isArray(parsedEntries) ? parsedEntries : [];
        } catch (_error) {
          entries = [];
        }

        finishLoading(callback);
      });
    } catch (_error) {
      entries = loadLocalBackup();
      finishLoading(callback);
    }
  }

  // Builds the current entry list as a Base64-encoded XLSX workbook.
  function buildExcelBase64() {
    if (typeof XLSX === "undefined") {
      console.error("Библиотека XLSX не загружена в content script.");
      return null;
    }

    try {
      const worksheet = XLSX.utils.json_to_sheet(entries);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Данные");

      const binaryWorkbook = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "binary",
      });
      return btoa(binaryWorkbook);
    } catch (error) {
      console.error("Ошибка при создании Excel:", error);
      return null;
    }
  }

  // Converts a Base64 workbook into a browser download.
  function downloadExcelFileFromBase64(base64Data) {
    try {
      const binaryData = atob(base64Data);
      const bytes = Uint8Array.from(binaryData, (character) =>
        character.charCodeAt(0),
      );
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const downloadUrl = URL.createObjectURL(blob);
      const downloadLink = document.createElement("a");

      downloadLink.href = downloadUrl;
      downloadLink.download = DOWNLOAD_FILE_NAME;
      downloadLink.hidden = true;
      document.body.appendChild(downloadLink);
      downloadLink.click();

      setTimeout(() => {
        downloadLink.remove();
        URL.revokeObjectURL(downloadUrl);
        console.log("Файл Excel успешно скачан.");
      }, 100);
    } catch (error) {
      console.error("Ошибка при скачивании Excel файла:", error);
    }
  }

  // Sends a new entry to durable storage and falls back to page-local storage on failure.
  function saveEntryToBackground(newEntry) {
    if (!canMessageBackground()) {
      console.warn(
        "chrome.runtime недоступен — используется резервное хранение localStorage.",
      );
      persistLocalBackup();
      return;
    }

    try {
      chrome.runtime.sendMessage(
        { action: "appendToExcel", newEntry },
        (response) => {
          if (chrome.runtime.lastError || response?.status !== "success") {
            console.error(
              "Ошибка сохранения в storage:",
              chrome.runtime.lastError?.message ?? "неизвестная ошибка",
            );
            persistLocalBackup();
            return;
          }

          console.log("Сохранено в storage. Всего:", response.dataLength);
          persistLocalBackup();
        },
      );
    } catch (error) {
      console.error("Ошибка отправки в background:", error);
      persistLocalBackup();
    }
  }

  // Adds an entry, persists it, and downloads an Excel file at the configured threshold.
  function appendEntry(newEntry) {
    entries.push(newEntry);
    unsavedEntriesCount += 1;

    const threshold = getDownloadThreshold();
    const shouldDownload = unsavedEntriesCount >= threshold;

    saveEntryToBackground(newEntry);

    if (!shouldDownload) {
      console.log(
        `Данные сохранены. Несохранённых: ${unsavedEntriesCount}/${threshold}. ` +
          "Для скачивания: downloadExcelFile()",
      );
      return;
    }

    const workbook = buildExcelBase64();
    if (!workbook) {
      return;
    }

    downloadExcelFileFromBase64(workbook);
    unsavedEntriesCount = 0;
    persistLocalBackup();
    console.log(
      `Файл Excel скачан. Следующий порог: ${getDownloadThreshold()}`,
    );
  }

  // Formats scraped Fairplay data into the row shape used by the Excel export.
  function createRoundEntry(fairplay, fallbackColor, missingResultMessage) {
    return {
      Дата: new Date().toLocaleString(),
      Цвет: fairplay.color || fallbackColor,
      Число: fairplay.result || missingResultMessage,
      "Раунд ID": fairplay.id,
      Hash: fairplay.hash,
      Salt: fairplay.salt,
    };
  }

  // Pauses asynchronous round processing for a fixed amount of time.
  function delay(timeoutMs) {
    return new Promise((resolve) => setTimeout(resolve, timeoutMs));
  }

  // Opens the latest round, waits for its Fairplay details, and records the result.
  async function captureLatestRound(historyIdElement, fallbackColor) {
    let fairplay;
    let missingResultMessage = "Не найдено";

    if (historyIdElement) {
      historyIdElement.click();
      await delay(ROUND_DETAILS_DELAY_MS);
      fairplay = await waitForFairplayResult(FAIRPLAY_TIMEOUT_MS);
    } else {
      console.log("Элемент истории для клика не найден.");
      fairplay = readFairplayData();
      missingResultMessage = "Не найдено (элемент для клика не найден)";
    }

    appendEntry(
      createRoundEntry(fairplay, fallbackColor, missingResultMessage),
    );
    setTimeout(startObserver, OBSERVER_RESTART_DELAY_MS);
  }

  // Watches the roulette timer and starts a single capture when it reaches the trigger.
  function startObserver() {
    const timerElement = document.querySelector(selectors.timer);
    if (!timerElement) {
      console.log(`Элемент ${selectors.timer} не найден на странице.`);
      return;
    }

    const observer = new MutationObserver((mutations) => {
      const timerReachedTrigger = mutations.some(
        (mutation) =>
          mutation.type === "childList" &&
          timerElement.textContent.trim() === TIMER_TRIGGER_TEXT,
      );

      if (!timerReachedTrigger) {
        return;
      }

      observer.disconnect();

      const latestEntry = document.querySelector(selectors.latestHistoryEntry);
      const historyIdElement = latestEntry?.querySelector(selectors.historyId);
      const fallbackColor = historyIdElement?.querySelector("b")?.className ?? "";

      void captureLatestRound(historyIdElement, fallbackColor);
    });

    observer.observe(timerElement, { childList: true, subtree: true });
  }

  // Exports all currently loaded entries and resets the pending-download counter.
  function downloadExcelFile() {
    const workbook = buildExcelBase64();
    if (!workbook) {
      console.log("Не удалось построить Excel.");
      return;
    }

    downloadExcelFileFromBase64(workbook);
    unsavedEntriesCount = 0;
    console.log("Файл Excel скачан. Счётчик сброшен.");
  }

  // Reports the number of entries added since the most recent download.
  function getUnsavedCount() {
    console.log(
      `Текущее количество несохраненных записей: ${unsavedEntriesCount}/${getDownloadThreshold()}`,
    );
    return unsavedEntriesCount;
  }

  // Clears the pending-download counter without deleting stored entries.
  function resetUnsavedCount() {
    unsavedEntriesCount = 0;
    console.log("Счетчик несохраненных записей сброшен.");
    return unsavedEntriesCount;
  }

  // Updates the auto-download threshold exposed to users through the console.
  function setAutoDownloadThreshold(newThreshold) {
    const previousThreshold = getDownloadThreshold();
    const threshold = setDownloadThreshold(newThreshold);

    if (typeof newThreshold !== "number" || newThreshold <= 0) {
      console.error("Ошибка: новый порог должен быть положительным числом.");
      return threshold;
    }

    console.log(
      `Порог автоматического скачивания изменен с ${previousThreshold} на ${threshold}.`,
    );
    return threshold;
  }

  window.downloadExcelFile = downloadExcelFile;
  window.getUnsavedCount = getUnsavedCount;
  window.resetUnsavedCount = resetUnsavedCount;
  window.setAutoDownloadThreshold = setAutoDownloadThreshold;
  window.getAutoDownloadThreshold = getDownloadThreshold;

  // Load persisted data before the observer starts so exports include prior rounds.
  loadStoredData(startObserver);

  console.log("Доступные команды Roulette Observer:");
  console.log("- downloadExcelFile() — скачать Excel с текущими данными");
  console.log("- getUnsavedCount() — показать число несохранённых записей");
  console.log("- resetUnsavedCount() — сбросить счётчик");
  console.log("- setAutoDownloadThreshold(число) — изменить порог скачивания");
})();
