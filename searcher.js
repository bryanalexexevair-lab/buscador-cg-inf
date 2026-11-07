/* global Excel */

// Search logic for the Buscador CG INF Excel add-in.
// The monthly worksheet must have one Spanish month name and a four-digit year.
const MONTH_NAMES = [
  "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];

function isMonthlyWorksheetName(name) {
  const normalized = String(name).trim().toLocaleUpperCase("es");
  return MONTH_NAMES.some((month) =>
    new RegExp(`^${month}\\s+\\d{4}$`, "u").test(normalized)
  );
}

function normalize(value) {
  return String(value ?? "").toLocaleLowerCase("es").replace(/\s+/gu, "");
}

/** Runs one search using the workbook open in Excel. */
export async function searchWorkbook(workbook) {
  const sheets = workbook.worksheets;
  sheets.load("items/name");
  await workbook.context.sync();

  const monthlySheets = sheets.items.filter((sheet) => isMonthlyWorksheetName(sheet.name));
  if (monthlySheets.length !== 1) {
    throw new Error(
      monthlySheets.length === 0
        ? "No encontré una hoja con nombre MES AÑO (por ejemplo, NOVIEMBRE 2026)."
        : "Encontré más de una hoja mensual; debe haber solo una."
    );
  }

  const dataSheet = monthlySheets[0];
  const searchSheet = sheets.getItemOrNullObject("buscador");
  searchSheet.load("name");
  await workbook.context.sync();

  if (searchSheet.isNullObject) {
    throw new Error("No encontré la hoja buscador.");
  }

  const queryRange = searchSheet.getRange("G2");
  queryRange.load("values");
  await workbook.context.sync();
  const query = normalize(queryRange.values[0][0]);
  if (!query) {
    searchSheet.getRange("D6:O20000").clear(Excel.ClearApplyTo.contents);
    return { count: 0, message: "Escribe un dato en G2 para buscar." };
  }

  const usedRange = dataSheet.getUsedRange(true);
  usedRange.load(["rowIndex", "rowCount"]);
  await workbook.context.sync();
  const lastExcelRow = usedRange.rowIndex + usedRange.rowCount;
  if (lastExcelRow < 2) {
    searchSheet.getRange("D6:O20000").clear(Excel.ClearApplyTo.contents);
    searchSheet.getRange("D6").values = [["Sin resultados"]];
    await workbook.context.sync();
    return { count: 0, message: "Sin resultados." };
  }

  const dataRange = dataSheet.getRange(`A2:M${lastExcelRow}`);
  dataRange.load("values");

  const comments = dataSheet.comments;
  comments.load("items");
  await workbook.context.sync();

  const loadedComments = [];
  for (const comment of comments.items) {
    const location = comment.getLocation();
    location.load("rowIndex");
    comment.load("content");
    comment.replies.load("items");
    loadedComments.push({ comment, location });
  }
  await workbook.context.sync();

  for (const { comment } of loadedComments) {
    for (const reply of comment.replies.items) reply.load("content");
  }
  await workbook.context.sync();

  const commentRows = new Map();
  for (const { comment, location } of loadedComments) {
    const rowIndex = location.rowIndex;
    const messages = [comment.content];
    for (const reply of comment.replies.items) {
      messages.push(reply.content);
    }
    const existing = commentRows.get(rowIndex) ?? [];
    existing.push(...messages.filter((message) => Boolean(message)));
    commentRows.set(rowIndex, existing);
  }
  const output = [];
  for (let i = 0; i < dataRange.values.length; i++) {
    const row = dataRange.values[i];
    const excelRowIndex = i + 1; // Data starts at Excel row 2; API indexes are zero-based.
    const threadMessages = commentRows.get(excelRowIndex) ?? [];
    const cellMatch = row.some((cell) => normalize(cell).includes(query));
    const commentMatch = threadMessages.some((message) => normalize(message).includes(query));

    if (cellMatch || commentMatch) {
      const resultRow = row.slice(0, 12); // A:L -> D:O, as in the existing script.
      if (threadMessages.length > 0) resultRow[11] = threadMessages.join("\n\n");
      output.push(resultRow);
    }
  }

  const outputRange = searchSheet.getRange("D6:O20000");
  outputRange.clear(Excel.ClearApplyTo.contents);
  if (output.length === 0) {
    searchSheet.getRange("D6").values = [["Sin resultados"]];
    await workbook.context.sync();
    return { count: 0, message: "Sin resultados." };
  }

  if (output.length > 19995) {
    throw new Error("La búsqueda encontró más resultados que el espacio disponible desde D6 hasta O20000.");
  }

  const results = searchSheet.getRangeByIndexes(5, 3, output.length, 12);
  results.values = output;
  searchSheet.getRangeByIndexes(5, 3, output.length, 1).numberFormat = "m/d/yyyy";
  searchSheet.getRangeByIndexes(5, 11, output.length, 2).numberFormat = "hh:mm:ss";
  results.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  results.format.verticalAlignment = Excel.VerticalAlignment.center;
  results.format.wrapText = true;
  await workbook.context.sync();

  return { count: output.length, message: `${output.length} resultado(s).` };
}
