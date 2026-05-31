/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearInterval, console, setInterval, Office, document, Excel */

Office.onReady((_) => {
  buildTaskpane();
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Initialisierung des Protokollwerts beim Laden der Seite
    const protocol = getFromLocalStorage("protocol") || "http";
    document.getElementById("httpsCheckbox").checked = protocol === "https";

    // Ereignishandler für Kontrollkästchen hinzufügen
    document.getElementById("httpsCheckbox").addEventListener("change", function () {
      const isHttps = this.checked ? "https" : "http";
      setFromLocalStorage("protocol", isHttps);
    });
  }
});

// TODO: Find way to share code between this and functions.js
const mandantenSelectionDefaultValue = "Auswahl Mandanten";
const jahrSelectionDefaultValue = "Auswahl Jahr";
const uStrukturSelectionDefaultValue = "Auswahl UStruktur";

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;
  const storageKey = myPartitionKey ? myPartitionKey + key : key;

  if (value === null || value === undefined) {
    // Remove the item from localStorage if value is null or undefined
    localStorage.removeItem(storageKey);
  } else {
    // Store the value as a string
    localStorage.setItem(storageKey, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;
  const storageKey = myPartitionKey ? myPartitionKey + key : key;

  const value = localStorage.getItem(storageKey);
  return value === "null" ? null : value;
}

async function buildTaskpane() {
  console.log("buildTaskpane");

  // Retrieve stored settings using getFromLocalStorage
  var gespeicherterServer = getFromLocalStorage("server");
  var gespeicherterPort = getFromLocalStorage("port");

  console.log("Taskpane: Gespeicherter Server: " + gespeicherterServer);
  console.log("Taskpane: Gespeicherter Port: " + gespeicherterPort);

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  if (gespeicherterPort != null && gespeicherterServer != null) {
    console.log("Taskpane: Found server + port ");
    document.getElementById("subheading").textContent = "Sie sind bereits eingeloggt.";
    document.getElementById("saveServerParams").textContent = "Ausloggen";
    document.getElementById("saveServerParams").onclick = removeLoginData;
    document.getElementById("loginContainer").style.visibility = "hidden";
  } else {
    console.log("Taskpane: Could not find server/port");
    document.getElementById("subheading").textContent =
      "Bitte hier URL und Port des API-Dienstes eingeben";
    document.getElementById("loginContainer").style.visibility = "visible";
    document.getElementById("saveServerParams").textContent = "Einloggen";
    document.getElementById("saveServerParams").onclick = saveServerParams;
  }
}

async function removeLoginData() {
  try {
    // Remove settings using localStorage
    setInLocalStorage("server", null);
    setInLocalStorage("port", null);
    console.log("Deleted stored settings");
  } catch (error) {
    console.error("Fehler: " + error.message);
  } finally {
    buildTaskpane();
  }
}

async function saveServerParams() {
  try {
    // Get input values from the task pane
    const serverInput = document.getElementById("serverInput").value;
    const portInput = document.getElementById("portInput").value;

    // Save settings using localStorage
    setInLocalStorage("server", serverInput);
    setInLocalStorage("port", portInput);

    console.log("Taskpane: Die Einstellungen wurden erfolgreich gespeichert");

    // Now, retrieve and log the saved settings
    var gespeicherterServer = getFromLocalStorage("server");
    var gespeicherterPort = getFromLocalStorage("port");

    console.log("Taskpane: Gespeicherter Server: " + gespeicherterServer);
    console.log("Taskpane: Gespeicherter Port: " + gespeicherterPort);
  } catch (error) {
    console.error("Fehler: " + error.message);
  } finally {
    // Refresh your task pane UI
    buildTaskpane();
  }
}

export async function callURL(url) {
  const response = await fetch(url);

  // Check if the response status code is an error
  if (!response.ok) {
    let responseBody = "";
    try {
      responseBody = await response.text();
    } catch (error) {
      responseBody = "Failed to read response body";
    }

    // Log detailed error info including URL, status code, status text, and part of the response body
    console.error(
      `HTTP error! URL: ${url}, Status: ${response.status}, ${response.statusText}, Response Body: ${responseBody.slice(
        0,
        200
      )}`
    );
    // Slice the response body to include only the first 200 characters in the error message
    throw new Error(`HTTP error! Status: ${response.status}, ${response.statusText}`);
  }
  // If response is OK, parse it as JSON
  const data = await response.json();
  return data;
}

async function mandantenAuswahl(event) {
  console.log("mandantenAuswahl");
  try {
    await Excel.run(async (context) => {
      const currentCell = context.workbook.getActiveCell();

      const auswahlMandanten = currentCell.getOffsetRange(1, 1);
      const auswahlJahre = currentCell.getOffsetRange(2, 1);
      const auswahlUStruktur = currentCell.getOffsetRange(3, 1);
      const auswahlDatenart = currentCell.getOffsetRange(4, 1);

      const cellAddress = currentCell.getOffsetRange(1, 1);
      const cellAddress2 = currentCell.getOffsetRange(2, 1);
      cellAddress.load("address");
      cellAddress2.load("address");
      await context.sync();
      console.log(`Loaded cell addresses: ${cellAddress.address}, ${cellAddress2.address}`);

      //dropdown Auswahl_Mandanten
      try {
        console.log("Clearing previous data validation for Mandanten");
        auswahlMandanten.dataValidation.clear();
        const udfMandanten = currentCell.getOffsetRange(0, 6);
        udfMandanten.formulas = [["=BA2.STAMM_MANDANTEN()"]];
        await context.sync();
        console.log("UDF Mandanten formulas set");

        udfMandanten.load("values");
        await context.sync();

        auswahlMandanten.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: udfMandanten.getResizedRange(70), // TODO: Remove first entry (header)
          },
        };

        const auswahlMandantenBinding = context.workbook.bindings.add(
          auswahlMandanten,
          "Range",
          "resetUStrukturUndJahre"
        );
        auswahlMandantenBinding.onDataChanged.add(resetUStrukturUndJahre.bind(null, auswahlJahre, auswahlUStruktur));
        console.log("Mandanten binding added");
        await context.sync();
      } catch (error) {
        console.error("Error while setting Mandanten dropdown: ", error);
      }

      // Dropdown Auswahl_Jahre
      try {
        console.log("Clearing previous data validation for Jahre");
        auswahlJahre.dataValidation.clear();
        const udfGeschaeftsjahre = currentCell.getOffsetRange(0, 4);
        udfGeschaeftsjahre.formulas = [["=BA2.STAMM_GESCHAEFTSJAHRE(" + cellAddress.address + ")"]];
        await context.sync();
        console.log("UDF Geschäftsjahre formulas set");

        udfGeschaeftsjahre.load("values");
        await context.sync();
        udfGeschaeftsjahre.format.autofitColumns();

        auswahlJahre.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: udfGeschaeftsjahre.getResizedRange(40), // TODO: Remove first entry (header)
          },
        };

        const auswahlJahreBinding = context.workbook.bindings.add(auswahlJahre, "Range", "resetUStruktur");
        auswahlJahreBinding.onDataChanged.add(resetUStruktur.bind(null, auswahlUStruktur));
        console.log("Jahre binding added");
        await context.sync();
      } catch (error) {
        console.error("Error while setting Jahre dropdown: ", error);
      }

      // Dropdown Auswahl_UStruktur
      try {
        console.log("Clearing previous data validation for UStruktur");
        auswahlUStruktur.dataValidation.clear();
        const udfUStruktur = currentCell.getOffsetRange(0, 8);
        udfUStruktur.formulas = [[`=BA2.STAMM_USTRUKTUR(${cellAddress.address},${cellAddress2.address})`]];
        await context.sync();
        console.log("UDF UStruktur formulas set");

        udfUStruktur.load("values");
        await context.sync();
        udfUStruktur.format.autofitColumns();

        auswahlUStruktur.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: udfUStruktur.getResizedRange(70), // TODO: Remove first entry (header)
          },
        };
        console.log("UStruktur dropdown set");
      } catch (error) {
        console.error("Error while setting UStruktur dropdown: ", error);
      }

      // Dropdown Auswahl_Datenart
      try {
        console.log("Clearing previous data validation for Datenart");
        auswahlDatenart.dataValidation.clear();
        auswahlDatenart.format.autofitColumns();
        await context.sync();

        auswahlDatenart.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "Plan, Ist, Vorschau",
          },
        };
        console.log("Datenart dropdown set");
      } catch (error) {
        console.error("Error while setting Datenart dropdown: ", error);
      }

      const leerzelle = currentCell.getOffsetRange(0, 0);
      const daten = currentCell.getOffsetRange(0, 1);
      const mandanten = currentCell.getOffsetRange(1, 0);
      const jahre = currentCell.getOffsetRange(2, 0);
      const ustruktur = currentCell.getOffsetRange(3, 0);
      const datenart = currentCell.getOffsetRange(4, 0);

      //layout
      console.log("Setting layout colors");
      leerzelle.format.fill.color = "#515151";
      daten.format.fill.color = "#FFCC00";
      mandanten.format.fill.color = "#A0A0A0";
      auswahlMandanten.format.fill.color = "#D7D7D7";
      jahre.format.fill.color = "#A0A0A0";
      auswahlJahre.format.fill.color = "#D7D7D7";
      ustruktur.format.fill.color = "#A0A0A0";
      auswahlUStruktur.format.fill.color = "#D7D7D7";
      datenart.format.fill.color = "#A0A0A0";
      auswahlDatenart.format.fill.color = "#D7D7D7";

      //Namen von Liste
      console.log("Setting default values");
      auswahlMandanten.values = [[mandantenSelectionDefaultValue]];
      auswahlJahre.values = [[jahrSelectionDefaultValue]];
      auswahlUStruktur.values = [[uStrukturSelectionDefaultValue]];
      auswahlDatenart.values = [["Auswahl Datenart"]];
      mandanten.values = [["Mandanten"]];
      // mandanten.values.format.font.color = "#FFFFFF";
      jahre.values = [["Jahre"]];
      daten.values = [["Daten:"]];
      ustruktur.values = [["UStruktur"]];
      datenart.values = [["Datenart"]];

      console.log("sechste");
      await context.sync();
      console.log("FINISHED mandantenAuswahl()");
      event.completed();
    });
  } catch (error) {
    console.log("Fehler in Mandantenauswahl.");
    console.error(error);
    console.error(error.stack);
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}
Office.actions.associate("mandantenAuswahl", mandantenAuswahl);

/**
 * Abfrage der GuV-Daten für einen bestimmten Mandanten, ein Bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart.
 * @customfunction DATEN_GUV
 * @param {string} Mandanten_ID_Name - Mandanten ID.
 * @param {string} Jahr_ID - Jahr.
 * @param {string} UStruktur_ID_Name - U-Struktur ID.
 * @param {string} Datenart_ID - Datenart-ID.
 * @returns {string[][]} GuV des ausgewählten Mandantens.
 */
async function Daten_GuV(Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, Datenart_ID) {
  try {
    const Mandanten_ID = getMandantenIdOrUStrukturIdFromName(Mandanten_ID_Name);
    const UStruktur_ID = getUStrukturIDFromName(UStruktur_ID_Name);
    let selectedDatenartID = Datenart_ID;

    // Проверяем, является ли Datenart_ID строкой и соответствует ли она "plan", "ist" или "vorschau"
    if (typeof Datenart_ID === "string") {
      switch (Datenart_ID.toLowerCase()) {
        case "plan":
          selectedDatenartID = 2;
          break;
        case "ist":
          selectedDatenartID = 3;
          break;
        case "vorschau":
          selectedDatenartID = 4;
          break;
        default:
          selectedDatenartID = Datenart_ID;
      }
    }
    const protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
      Mandanten_ID +
      "/" +
      Jahr_ID +
      "/" +
      UStruktur_ID +
      "/" +
      selectedDatenartID;
    console.log(url);
    const myData = await callURL(url);

    const ergebnis = [
      [
        "Reihenfolge",
        "Name",
        "Monat1",
        "Monat2",
        "Monat3",
        "Monat4",
        "Monat5",
        "Monat6",
        "Monat7",
        "Monat8",
        "Monat9",
        "Monat10",
        "Monat11",
        "Monat12",
        "Jahreswert",
      ],
    ];

    myData.forEach((item) => {
      ergebnis.push([
        item.Reihenfolge,
        item.Position_Name,
        item.Monat1,
        item.Monat2,
        item.Monat3,
        item.Monat4,
        item.Monat5,
        item.Monat6,
        item.Monat7,
        item.Monat8,
        item.Monat9,
        item.Monat10,
        item.Monat11,
        item.Monat12,
        item.Jahreswert,
      ]);
    });

    return ergebnis;
  } catch (error) {
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}
CustomFunctions.associate("DATEN_GUV", Daten_GuV);
/*
import { Daten_GuV } from "../functions/functions.js";
async function importDaten_GuV() {
  const result = await Daten_GuV(Param, Jahr_ID, UStruktur_ID_Name, Datenart_ID);
  console.log(result);
} */

async function resetUStruktur(ustrukturRange) {
  await Excel.run(async (context) => {
    ustrukturRange.values = [[uStrukturSelectionDefaultValue]];
    await context.sync();
  });
}
Office.actions.associate("resetUStruktur", resetUStruktur);

async function resetUStrukturUndJahre(jahreRange, ustrukturRange) {
  await Excel.run(async (context) => {
    jahreRange.values = [[jahrSelectionDefaultValue]];
    ustrukturRange.values = [[uStrukturSelectionDefaultValue]];
    await context.sync();
  });
}
Office.actions.associate("resetUStrukturUndJahre", resetUStrukturUndJahre);

async function layout(event) {
  //console.log("layout!");

  try {
    let server = getFromLocalStorage("server");
    let port = getFromLocalStorage("port");

    // Prüfen, ob der Server und der Port verfügbar sind
    if (!server || !port) {
      const errorMessage = "Server oder Port ist nicht angegeben. Bitte überprüfen Sie die API Einstellung";
      console.error(errorMessage);
      // Ein Pop-up-Fenster mit einer Fehlermeldung anzeigen
      showError(errorMessage);
      openFirstTaskpane();

      // Abonnieren, um Einstellungen zu ändern
      Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, async function () {
        server = getFromLocalStorage("server");
        port = getFromLocalStorage("port");
      });
      //await processRangeInContext();
    }

    await processRangeInContext();
  } catch (error) {
    console.error(error);
    showError("Es ist ein Fehler aufgetreten: " + error.message);
  }

  async function processRangeInContext() {
    await Excel.run(async (context) => {
      // Ermittlung der aktiven Zelle
      const activeCell = context.workbook.getActiveCell();
      activeCell.load("address");
      await context.sync();

      // Abrufen des ausgewählten Bereichs
      const range = context.workbook.getSelectedRange();
      range.load("formulas");
      range.load("values");
      await context.sync();

      // Deaktivieren der Bildschirmaktualisierung und der manuellen Berechnung
      context.application.suspendScreenUpdatingUntilNextSync();
      context.application.calculate(Excel.CalculationType.manual);

      // Synchronisierung des Kontexts
      await context.sync();

      // Aufruf der Funktion processRange nur, wenn Server und Port angegeben sind
      try {
        processRange(range);
      } catch (error) {
        showError("Fehler bei der Verarbeitung des Bereichs: " + error.message);
        openFirstTaskpane();
      }

      // Bildschirmaktualisierung und automatische Berechnung aktivieren
      context.application.calculate(Excel.CalculationType.automatic);
      context.application.screenUpdating = true;
      event.completed();
    });
  }
}

Office.actions.associate("layout", layout);

function openFirstTaskpane() {
  Office.addin.showAsTaskpane(); //  Tuskpane anzeigen
}

async function processRange(range) {
  try {
    //console.log("processRangeworks");
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheet = workbook.worksheets.getActiveWorksheet();

      await context.sync();

      const formula = range.formulas[0][0];

      if (formula != null) {
        const formulaParts = formula.split(",");

        if (formulaParts.length !== 4) {
          return;
        }
        let firstFormulaPartMatch = formulaParts[0].match(/\(([^)]+)/);
        if (!firstFormulaPartMatch) {
          return; // TODO SOME ERROR
        }
        let Mandanten_ID = await processFormulaPart(
          firstFormulaPartMatch[1],
          workbook,
          worksheet,
          context,
          parseMandantenIDOrUStruktur
        );
        let Jahr_ID = await processFormulaPart(formulaParts[1], workbook, worksheet, context, (input) => input);
        let UStruktur_ID_Name = await processFormulaPart(
          formulaParts[2],
          workbook,
          worksheet,
          context,
          parseMandantenIDOrUStruktur
        );
        let Datenart_ID = await processFormulaPart(
          formulaParts[3].substring(0, formulaParts[3].length - 1),
          workbook,
          worksheet,
          context,
          (input) => input
        );
        if (formula.startsWith("=BA2.DATEN_GUV")) {
          formatFormulaA(range, Mandanten_ID, Jahr_ID, UStruktur_ID_Name, Datenart_ID);
        } else if (formula.startsWith("=BA2.DATEN_BILANZ")) {
          formatFormulaB(range, Mandanten_ID, Jahr_ID, UStruktur_ID_Name, Datenart_ID);
        } else {
          return;
        }
        await context.sync();
      }
    });
  } catch (error) {
    console.error(error);
  }
}

async function processFormulaPart(formelParameter, workbook, worksheet, context, parser) {
  try {
    let rawCellValue; // can either be raw number or number - someMandantenName
    const exclamationIndex = formelParameter.indexOf("!");

    const cellReferenceRegex = /^[A-Z]+[0-9]+$/;
    const parameterIsCellReference = cellReferenceRegex.test(formelParameter);

    if (exclamationIndex !== -1) {
      try {
        let parts = formelParameter.split("!");
        let coordinates = processFormulaPart(parts[1]);

        let sheet = workbook.worksheets.getItem(parts[0]);
        let cell = sheet.getCell(coordinates.row, coordinates.col);
        cell.load("address, values");
        await context.sync();

        rawCellValue = cell.values[0][0];
      } catch (innerError) {
        console.error("Error processing cell reference with sheet name.");
        console.error(innerError);
        console.error(innerError.stack);
        throw innerError; // Optional: Rethrow the error to be caught by the outer catch block
      }
    } else if (parameterIsCellReference) {
      try {
        let coordinates = convertCellToCoordinates(formelParameter);
        const cell = worksheet.getCell(coordinates.row, coordinates.col);
        cell.load("values");
        await context.sync();
        rawCellValue = cell.values[0][0]; // value = [[]]
      } catch (innerError) {
        console.error("Error processing direct cell reference.");
        console.error(innerError);
        console.error(innerError.stack);
        throw innerError; // Optional: Rethrow the error to be caught by the outer catch block
      }
    } else {
      // Direct input, e.g. 2020 for year
      return formelParameter;
    }

    return parser(rawCellValue);
  } catch (error) {
    console.error("Error in processFormulaPart function.");
    console.error(error);
    console.error(error.stack);
    throw error; // Rethrow the error if you want it to be handled further up the call stack
  }
}

// Parser
async function parseMandantenIDOrUStruktur(input) {
  let isNumber = /^[0-9]+$/.test(input);
  if (!isNumber) {
    return getMandantenIdOrUStrukturIdFromName(input);
  } else {
    return input;
  }
}

async function formatFormulaA(range, Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, Datenart_ID) {
  await Excel.run(async (context) => {
    try {
      const selectedcell = context.workbook.getActiveCell();
      const Mandanten_ID = getMandantenIdOrUStrukturIdFromName(Mandanten_ID_Name);
      const UStruktur_ID = getUStrukturIDFromName(UStruktur_ID_Name);
      let selectedDatenartID = Datenart_ID;

      if (typeof Datenart_ID === "string") {
        switch (Datenart_ID.toLowerCase()) {
          case "plan":
            selectedDatenartID = 2;
            break;
          case "ist":
            selectedDatenartID = 3;
            break;
          case "vorschau":
            selectedDatenartID = 4;
            break;
          default:
            selectedDatenartID = parseInt(Datenart_ID);
        }
      }

      const cellValue = range.values[0][0];
      const protocol = getFromLocalStorage("protocol") || "http";
      const url =
        `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
        Mandanten_ID_Name +
        "/" +
        Jahr_ID +
        "/" +
        UStruktur_ID_Name +
        "/" +
        selectedDatenartID;
      console.log(url);
      const myData = await callURL(url);

      myData.forEach((data, index) => {
        //console.log("zeile:%d", index);

        var currentCell = selectedcell;
        currentCell = currentCell.getOffsetRange(index + 1, 0);
        currentCell = currentCell.getResizedRange(0, 14);
        currentCell.numberFormat = "#,##0.00;-#,##0.00"; //#.##0,00;-#.##0,00 normal
        currentCell.format.fill.color = convertDelphiColorToRGB(data.ZL_Zeilenfarbe);

        currentCell.fontSize = data.ZL_Schrift_Groesse;
        currentCell.fontName = data.ZL_Schrift_Name;

        if (data.ZL_Schrift_Style & (1 << 0)) {
          currentCell.fontBold = true;
        } else {
          currentCell.fontBold = false;
        }

        if (data.ZL_Schrift_Style & (1 << 1)) {
          currentCell.fontItalic = true;
        } else {
          currentCell.fontItalic = false;
        }

        if (data.ZL_Schrift_Style & (1 << 2)) {
          currentCell.fontUnderline = true;
        } else {
          currentCell.fontUnderline = false;
        }

        if (data.ZL_Schrift_Style & (1 << 3)) {
          currentCell.fontStrikethrough = true;
        } else {
          currentCell.fontStrikethrough = false;
        }
        currentCell.format.font.color = convertDelphiColorToRGB(data.ZL_Textfarbe);
      });
    } catch (error) {
      //console.log("FehlerinSchleife");
      console.error(error);
    }
    //console.log("SchleifeEndet");
  });
}
async function formatFormulaB(range, Mandanten_ID, Jahr_ID, UStruktur_ID, Datenart_ID) {
  //try {
  console.log("formatFormulaB function called");
  await Excel.run(async (context) => {
    try {
      const selectedcell = context.workbook.getActiveCell();

      Mandanten_ID = await getMandantenIdOrUStrukturIdFromName(Mandanten_ID);

      const cellValue = range.values[0][0];
      const protocol = getFromLocalStorage("protocol") || "http";
      const url =
        `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
        Mandanten_ID +
        "/" +
        Jahr_ID +
        "/" +
        UStruktur_ID +
        "/" +
        selectedDatenartID;
      console.log(url);
      const myData = await callURL(url);

      myData.forEach((data, index) => {
        console.log("zeile:%d", index);

        var currentCell = selectedcell;
        currentCell = currentCell.getOffsetRange(index + 1, 0);
        currentCell = currentCell.getResizedRange(0, 14);
        //const currentCell = range.getOffsetRange(index + 1, 0).getResizedRange(1, 15);
        //console.log("Fehler1");
        currentCell.numberFormat = "#,##";
        //console.log("Fehler2");
        currentCell.format.fill.color = convertDelphiColorToRGB(data.ZL_Zeilenfarbe); //convertDelphiColorToExcelColor(data.ZL_Zeilenfarbe); //format.fill.color = "Orange";
        //console.log("Fehler3");
        // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
        console.log("farbe:%d", currentCell.format.fill.color);
        console.log("farbe:%s", convertDelphiColorToRGB(data.ZL_Zeilenfarbe));
        currentCell.fontSize = data.ZL_Schrift_Groesse;
        currentCell.fontName = data.ZL_Schrift_Name;

        if (data.ZL_Schrift_Style & (1 << 0)) {
          currentCell.fontBold = true;
        } else {
          currentCell.fontBold = false;
        }

        if (data.ZL_Schrift_Style & (1 << 1)) {
          currentCell.fontItalic = true;
        } else {
          currentCell.fontItalic = false;
        }

        if (data.ZL_Schrift_Style & (1 << 2)) {
          currentCell.fontUnderline = true;
        } else {
          currentCell.fontUnderline = false;
        }

        if (data.ZL_Schrift_Style & (1 << 3)) {
          currentCell.fontStrikethrough = true;
        } else {
          currentCell.fontStrikethrough = false;
        }
        currentCell.fontColor = convertDelphiColorToExcelColor(data.ZL_Textfarbe); //ausprobieren TextFarbe
      });
    } catch (error) {
      console.log("FehlerinSchleife");
      console.error(error);
    }
    console.log("SchleifeEndet");
  });
}
async function formatFormulaC(range, Mandanten_ID, Jahr_ID, UStruktur_ID, Datenart_ID) {
  try {
    console.log("itworksB");
    Mandanten_ID = await getMandantenIdOrUStrukturIdFromName(Mandanten_ID);

    const cellValue = range.values[0][0];

    const protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
      Mandanten_ID +
      "/" +
      Jahr_ID +
      "/" +
      UStruktur_ID +
      "/" +
      selectedDatenartID;
    console.log(url);
    const myData = await callURL(url);

    myData.forEach((data, index) => {
      const currentCell = range.getOffsetRange(index + 1, 0).getResizedRange(1, 8);

      currentCell.numberFormat = "#,##";
      currentCell.setColor = convertDelphiColorToExcelColor(data.ZL_Zeilenfarbe); //format.fill.color
      currentCell.fontSize = data.ZL_Schrift_Groesse;
      currentCell.fontName = data.ZL_Schrift_Name;

      if (data.ZL_Schrift_Style & (1 << 0)) {
        currentCell.fontBold = true;
      } else {
        currentCell.fontBold = false;
      }

      if (data.ZL_Schrift_Style & (1 << 1)) {
        currentCell.fontItalic = true;
      } else {
        currentCell.fontItalic = false;
      }

      if (data.ZL_Schrift_Style & (1 << 2)) {
        currentCell.fontUnderline = true;
      } else {
        currentCell.fontUnderline = false;
      }

      if (data.ZL_Schrift_Style & (1 << 3)) {
        currentCell.fontStrikethrough = true;
      } else {
        currentCell.fontStrikethrough = false;
      }

      currentCell.fontColor = convertDelphiColorToExcelColor(data.ZL_Textfarbe);
    });
  } catch (error) {
    console.error(error);
  }
}
function formatNumberToGerman(value) {
  return value
    .toFixed(2) // Rundet die Zahl auf 2 Dezimalstellen
    .replace(".", ",") // Ersetzt den Dezimalpunkt durch ein Komma
    .replace(/\B(?=(\d{3})+(?!\d))/g, "."); // Fügt Punkte als Tausendertrennzeichen ein
}

/*async function umwandeln(event) {
  try {
    console.log("Kommentare aus dem ausgewählten Bereich extrahieren...");

    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "rowCount", "columnCount", "address", "text"]);
      await context.sync();

      const comments = context.workbook.comments;
      comments.load("items");
      await context.sync();

      if (comments.items.length === 0) {
        console.log("Keine Kommentare gefunden!");
        return;
      }

      // Iteration durch alle Zellen im ausgewählten Bereich
      for (let row = 0; row < selectedRange.rowCount; row++) {
        for (let col = 0; col < selectedRange.columnCount; col++) {
          const cell = selectedRange.getCell(row, col);
          cell.load("address");
          //await context.sync();

          const cellAddress = cell.address;
          console.log(`Überprüfung der Kommentare in der Zelle ${cellAddress}...`);

          const comment = comments.items.find((item) => item.cellAddress === cellAddress);

          if (comment) {
            console.log(`Kommentar in der Zelle ${cellAddress}: "${comment.content}"`);
          } else {
            console.log(`Kein Kommentar in der Zelle ${cellAddress} gefunden.`);
          }
        }
      }

      event.completed();
    });
  } catch (error) {
    console.error("Fehler bei der Ausführung der Funktion umwandeln:", error);
    event.completed();
  }
}
Office.actions.associate("umwandeln", umwandeln);*/

function getMandantenIdOrUStrukturIdFromName(Mandanten_Name) {
  let Mandanten_ID = -1;

  if (Mandanten_Name && typeof Mandanten_Name === "string") {
    const mandantenIdNurZahlen = Mandanten_Name.match(/\d+/);

    if (mandantenIdNurZahlen) {
      Mandanten_ID = parseInt(mandantenIdNurZahlen[0]);
    }
  }
  return Mandanten_ID;
}
function convertDelphiColorToRGB(delphiColor) {
  let r = (delphiColor >> 16) & 0xff; // Rot in BGR (Delphi) entspricht Blau in RGB
  let g = (delphiColor >> 8) & 0xff; // Grün bleibt gleich
  let b = delphiColor & 0xff; // Blau in BGR (Delphi) entspricht Rot in RGB
  return "#" + ((1 << 24) + (b << 16) + (g << 8) + r).toString(16).slice(1).toUpperCase();
}

function getUStrukturIDFromName(UStruktur_Name) {
  let UStruktur_ID = 0;

  if (UStruktur_Name && typeof UStruktur_Name === "string") {
    const ustrukturIdNurZahlen = UStruktur_Name.match(/\d+/);

    if (ustrukturIdNurZahlen) {
      UStruktur_ID = parseInt(ustrukturIdNurZahlen[0]);
    }
  }
  return UStruktur_ID;
}

function convertDelphiColorToExcelColor(delphiColor) {
  let r = (delphiColor >> 16) & 0xff;
  let g = (delphiColor >> 8) & 0xff;
  let b = delphiColor & 0xff;
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}
function extractCellReference(input) {
  // Prüfen, ob ein '!' im String vorhanden ist.
  let exclamationIndex = input.indexOf("!");
  if (exclamationIndex !== -1) {
    // Wenn ja, extrahiere alles nach dem '!' als Zellenbezug.
    return input.substring(exclamationIndex + 1);
  } else {
    // Wenn nicht, gehe davon aus, dass der gesamte String der Zellenbezug ist.
    return input;
  }
}

function columnToNumber(column) {
  let number = 0;
  let multiplier = 1;

  for (let i = column.length - 1; i >= 0; i--) {
    number += (column.charCodeAt(i) - 65 + 1) * multiplier;
    multiplier *= 26;
  }

  return number;
}
function convertCellToCoordinates(cellName) {
  const columnLetter = cellName.match(/[A-Z]+/)[0];
  const rowNumber = cellName.match(/[0-9]+/)[0];
  // Berechnung des Spaltenindex
  let columnIndex = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    columnIndex = columnIndex * 26 + (columnLetter.charCodeAt(i) - "A".charCodeAt(0)) + 1;
  }
  columnIndex -= 1; // Spaltenindex bei 0 beginnen lassen
  const rowIndex = parseInt(rowNumber, 10) - 1; // Zeilenindex bei 0 beginnen lassen
  return { row: rowIndex, col: columnIndex };
}

function logStuff() {
  console.log("Hallo");
}
Office.actions.associate("logStuff", logStuff);
