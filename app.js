import { searchWorkbook } from "./searcher.js";

let initialized = false;
let activeSearchSheetId = null;
let searchInProgress = false;
let pendingSearch = false;

const setStatus = (message) => {
  const element = document.getElementById("status");
  if (element) element.textContent = message;
};

async function runSearch() {
  if (searchInProgress) {
    pendingSearch = true;
    return;
  }

  searchInProgress = true;
  try {
    const result = await Excel.run((context) => searchWorkbook(context.workbook));
    setStatus(result.message);
  } catch (error) {
    setStatus(error?.message || "Ocurrió un error al buscar.");
  } finally {
    searchInProgress = false;
    if (pendingSearch) {
      pendingSearch = false;
      void runSearch();
    }
  }
}

async function registerG2Listener() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("buscador");
    sheet.load(["id", "name"]);
    await context.sync();
    if (sheet.isNullObject) {
      setStatus("No encontré la hoja buscador.");
      return;
    }

    if (activeSearchSheetId === sheet.id) return;
    sheet.onChanged.add(async (event) => {
      const changedAddress = String(event.address || "").toUpperCase().replace(/\$/gu, "");
      const finalAddress = changedAddress.includes("!")
        ? changedAddress.slice(changedAddress.lastIndexOf("!") + 1)
        : changedAddress;
      if (finalAddress === "G2") await runSearch();
    });
    activeSearchSheetId = sheet.id;
    await context.sync();
    setStatus("Listo. El buscador vigila buscador!G2.");
  });
}

async function enableAutoSearch() {
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    await registerG2Listener();
    setStatus("Búsqueda automática activada para este libro. Puedes cerrar el panel.");
  } catch (error) {
    setStatus(error?.message || "No se pudo activar el inicio automático en este Excel.");
  }
}

function completeCommand(event) {
  if (event && typeof event.completed === "function") event.completed();
}

window.searchNow = (event) => {
  void runSearch().finally(() => completeCommand(event));
};

window.enableAutoSearch = (event) => {
  void enableAutoSearch().finally(() => completeCommand(event));
};

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel || initialized) return;
  initialized = true;

  document.getElementById("searchButton")?.addEventListener("click", () => void runSearch());
  document.getElementById("autoButton")?.addEventListener("click", () => void enableAutoSearch());

  try {
    const behavior = await Office.addin.getStartupBehavior();
    if (behavior === Office.StartupBehavior.load) await registerG2Listener();
    else setStatus("Pulsa ACTIVAR AUTOMÁTICO una vez en este libro; después podrás cerrar el panel.");
  } catch {
    setStatus("Abre este panel y activa la búsqueda automática para este libro.");
  }
});
