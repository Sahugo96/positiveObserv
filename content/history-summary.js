// Adds history-table totals and row-exclusion controls without sharing global state.
(function initializeHistorySummary() {
  "use strict";

  const DELETED_ROW_CLASS = "deleted-row";
  const DELETE_BUTTON_CLASS = "delete-row-btn";

  const selectors = {
    activeFilter: 'a[data-filter="active"]',
    selectedActiveFilter: 'a[data-filter="active"].active',
    balance: ".balance_cont .balance",
    bonusBalance: ".bonus_cont .bonus",
    historyTable: "table.history",
  };

  // Parses the balance text displayed by the site, using zero when it is absent.
  function readElementNumber(element) {
    const rawValue = element?.childNodes[0]?.textContent.trim();
    const value = Number.parseFloat(rawValue);
    return Number.isNaN(value) ? 0 : value;
  }

  const balanceElement = document.querySelector(selectors.balance);
  const bonusBalanceElement = document.querySelector(selectors.bonusBalance);
  const initialBalance = readElementNumber(balanceElement);
  const initialBonusBalance = readElementNumber(bonusBalanceElement);

  // Totals regular and gem values in one table column, skipping excluded rows.
  function sumCurrencyColumn(table, columnNumber) {
    const totals = { regular: 0, gems: 0 };
    const cells = table.querySelectorAll(
      `tbody tr td:nth-child(${columnNumber})`,
    );

    cells.forEach((cell) => {
      if (cell.parentElement?.classList.contains(DELETED_ROW_CLASS)) {
        return;
      }

      const value = Number.parseFloat(cell.firstChild?.textContent.trim() ?? "");
      const currencyIcon = cell.querySelector("i");
      if (Number.isNaN(value) || !currencyIcon) {
        return;
      }

      if (currencyIcon.classList.contains("gems")) {
        totals.gems += value;
      } else if (currencyIcon.classList.contains("rr")) {
        totals.regular += value;
      }
    });

    return totals;
  }

  // Displays regular and gem totals in a history-table header cell.
  function renderColumnTotals(header, totals) {
    if (!header) {
      return;
    }

    header.innerHTML =
      `${totals.regular}<i class="rr"></i><br>` +
      `${totals.gems}<i class="gems"></i>`;
    header.style.transform = "translateY(-10px)";
    header.style.lineHeight = "16px";
  }

  // Updates a balance label with its original value plus placed and won totals.
  function renderBalance(element, initialValue, placed, won) {
    if (!element?.childNodes[0]) {
      return;
    }

    element.childNodes[0].textContent =
      `${initialValue} (${initialValue + placed} / ${initialValue + won})`;
  }

  // Marks or restores a row so it is respectively excluded from or included in totals.
  function toggleDeletedRow(row) {
    const isDeleted = row.classList.toggle(DELETED_ROW_CLASS);
    row.style.opacity = isDeleted ? "0.5" : "1";
    row.style.textDecoration = isDeleted ? "line-through" : "unset";
  }

  // Creates the control that lets a user exclude one history row from calculations.
  function createDeleteButton(row) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "✕";
    button.className = DELETE_BUTTON_CLASS;
    button.title = "Исключить строку из расчёта";
    button.setAttribute("aria-label", "Исключить строку из расчёта");
    button.style.cssText = `
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

    button.addEventListener("mouseenter", () => {
      button.style.transform = "scale(1.2)";
    });
    button.addEventListener("mouseleave", () => {
      button.style.transform = "scale(1)";
    });
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleDeletedRow(row);
      setTimeout(sumHistoryColumns, 100);
    });

    return button;
  }

  // Adds an exclusion control to each eligible history row exactly once.
  function addDeleteButtons(table) {
    table.querySelectorAll("tbody tr").forEach((row) => {
      const buttonCell = row.querySelector("td:nth-child(2)");
      const shouldSkip =
        !buttonCell ||
        row.classList.contains("express_icon") ||
        buttonCell.querySelector(`.${DELETE_BUTTON_CLASS}`);

      if (!shouldSkip) {
        buttonCell.appendChild(createDeleteButton(row));
      }
    });
  }

  // Recalculates table headers and balances for the currently active history filter.
  function sumHistoryColumns() {
    if (!document.querySelector(selectors.selectedActiveFilter)) {
      return;
    }

    const table = document.querySelector(selectors.historyTable);
    if (!table) {
      return;
    }

    addDeleteButtons(table);

    const placedTotals = sumCurrencyColumn(table, 3);
    const wonTotals = sumCurrencyColumn(table, 5);

    renderColumnTotals(
      table.querySelector("thead tr th:nth-child(3)"),
      placedTotals,
    );
    renderColumnTotals(
      table.querySelector("thead tr th:nth-child(5)"),
      wonTotals,
    );
    renderBalance(
      balanceElement,
      initialBalance,
      placedTotals.regular,
      wonTotals.regular,
    );
    renderBalance(
      bonusBalanceElement,
      initialBonusBalance,
      placedTotals.gems,
      wonTotals.gems,
    );
  }

  // Wait for the site to refresh the active-history table before calculating totals.
  const activeFilter = document.querySelector(selectors.activeFilter);
  activeFilter?.addEventListener("click", () => {
    setTimeout(sumHistoryColumns, 1_200);
  });
})();
