/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* global clearInterval, console, setInterval, Office, Mandanten_ID, */

// Angenommen, server und port sind Ihre Variablen

export async function callURL(url) {
  const response = await fetch(url);

  if (!response.ok) {
    let responseBody = "";
    try {
      responseBody = await response.text(); // Using text() as it's more general than json()
    } catch (error) {
      responseBody = "Failed to read response body";
    }

    console.error(
      `HTTP error! URL: ${url}, Status: ${response.status}, ${response.statusText}, Response Body: ${responseBody.slice(
        0,
        200
      )}`
    );

    throw new Error(`HTTP error! Status: ${response.status}, ${response.statusText}`);
  }

  const data = await response.json();
  return data;
}

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

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const protocol = getFromLocalStorage("protocol") || "http";
    document.getElementById("httpsCheckbox").checked = protocol === "https";

    document.getElementById("httpsCheckbox").addEventListener("change", function () {
      const isHttps = this.checked ? "https" : "http";
      setFromLocalStorage("protocol", isHttps);
      Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(asyncResult.error.message);
        } else {
          console.log(`Protocol saved: ${isHttps}`);
        }
      });
    });
  }
});

/**
 * Abfrage der Mandanten-ID und -Bezeichnung in einer Tabelle
 * @customfunction Stamm_Mandanten
 * @returns {string[][]} A dynamic array with multiple results.
 */
export async function Stamm_Mandanten() {
  try {
    let server = getFromLocalStorage("server");
    let port = getFromLocalStorage("port");
    let protocol = getFromLocalStorage("protocol") || "http";

    if (server != null && port != null) {
      const url = `${protocol}://${server}:${port}/api/clients/`;
      const myData = await callURL(url);
      console.log(url);

      const ergebnis = [["ID - Name"]];

      myData.forEach((item) => {
        ergebnis.push([`${item.Mandant_ID} - ${item.Name}`]);
      });
      //console.log("Mandanten: ", ergebnis);
      return ergebnis;
    }
  } catch (error) {
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}

/**
 * Abfrage der Jahreszahlen für einen bestimmten Mandanten
 * @customfunction Stamm_Geschaeftsjahre
 * @param {string} Mandanten_ID  - Mandanten-ID.
 * @returns {string[][]} A dynamic array with multiple results.
 */
export async function Stamm_Geschaeftsjahre(Mandanten_ID) {
  const ergebnis = [["Jahre"]];
  if (
    Mandanten_ID === undefined ||
    Mandanten_ID === -1 ||
    Mandanten_ID === "" ||
    Mandanten_ID === "Auswahl Mandanten"
  ) {
    //console.log("MANDANTEN_ID_NAME is undefined, empty or -1");
    return ergebnis;
  }
  try {
    let server = getFromLocalStorage("server");
    let port = getFromLocalStorage("port");
    let protocol = getFromLocalStorage("protocol") || "http";

    if (server != null && port != null) {
      const Mandanten = getMandantenIDFromName(Mandanten_ID);
      const url = `${protocol}://${server}:${port}/api/data/clients/`;
      console.log(url);
      const myData = await callURL(url);

      Object.keys(myData).forEach((key) => {
        const item = myData[key];
        if (item.Mandant_ID === Mandanten) {
          for (let i = item.Erstes_GJ_Co; i <= item.Letztes_GJ; i++) {
            ergebnis.push([i]);
          }
        }
      });
      //console.log("Jahre: ", ergebnis);
      //console.log("finished Geshaeftsjahre()");

      return ergebnis;
    }
  } catch (error) {
    console.error("Fehler:", "Fehler beim Abrufen der Daten.", error.message);
    return [["Fehler:", "Fehler beim Abrufen der Daten."]];
  }
}

/**
 * Abfrage der Jahreszahlen für einen bestimmten Mandanten
 * @customfunction Stamm_Ustruktur
 * @param {string} Mandanten_ID  - Mandanten-ID.
 * @param {string} Jahr - Jahr_ID.
 * @returns {string[][]} A dynamic array with multiple results.
 */
export async function Stamm_Ustruktur(Mandanten_ID, Jahr) {
  const ergebnis = [["ID - Name"]];
  if (
    Mandanten_ID === undefined ||
    Mandanten_ID === -1 ||
    Mandanten_ID === "" ||
    Mandanten_ID === "Auswahl Mandanten"
  ) {
    //console.log("MANDANTEN_ID_NAME is undefined, empty or -1");
    return ergebnis;
  } else if (Jahr === undefined || Jahr === -1 || Jahr === "" || Jahr === "Auswahl Jahr") {
    //console.log("JAHR_ID is undefined, empty or -1");
    return ergebnis;
  }
  try {
    let server = getFromLocalStorage("server");
    let port = getFromLocalStorage("port");
    let protocol = getFromLocalStorage("protocol") || "http";

    if (server != null && port != null) {
      //console.log("Mandanten_ID = " + Mandanten_ID);
      //console.log("Jahr = " + Jahr);

      const Mandanten = getMandantenIDFromName(Mandanten_ID);

      //console.log("Mandanten_ID = " + Mandanten_ID);
      const url =
        `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/clients/` +
        Mandanten +
        "/ustruktur";
      console.log(url);
      const myData = await callURL(url);
      //console.log("U-Struktur: JSON-Daten:", myData);

      Object.keys(myData).forEach((key) => {
        const item = myData[key];
        // console.log("suche Jahr");
        // console.log(item.Jahr_ID + " <> " + Jahr);
        if (item.Jahr_ID == Jahr) {
          // console.log("Jahr gefunden");
          ergebnis.push([`${item.U_Struktur_ID} - ${item.U_Name}`]);
        }
      });
      //console.log("acc endet");

      return ergebnis;
    }
  } catch (error) {
    console.error("Fehler:", "Fehler beim Abrufen der Daten.", error.message);
    return [["Fehler:", "Fehler beim Abrufen der Daten."]];
  }
}

function getMandantenIDFromName(Mandanten_Name) {
  let Mandanten_ID = -1;

  if (Mandanten_Name && typeof Mandanten_Name === "string") {
    const mandantenIdNurZahlen = Mandanten_Name.match(/\d+/);

    if (mandantenIdNurZahlen) {
      Mandanten_ID = parseInt(mandantenIdNurZahlen[0]);
    }
  }
  return Mandanten_ID;
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

/**
 * Abfrage des Aufbaus der GuV für einen bestimmten Mandanten
 * @customfunction Aufbau_GuV
 * @param {string} Mandanten_ID_Name - Mandanten_ID_Name.
 * @returns {string[][]} A dynamic array with multiple results.
 */
async function Aufbau_GuV(Mandanten_ID_Name) {
  try {
    let server = getFromLocalStorage("server");
    let port = getFromLocalStorage("port");
    let protocol = getFromLocalStorage("protocol") || "http";

    if (server != null && port != null) {
      const Mandanten_ID = getMandantenIDFromName(Mandanten_ID_Name);
      const url =
        `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/clients/` +
        Mandanten_ID +
        "/guv";
      console.log(url);
      const myData = await callURL(url);

      const ergebnis = [["ID", "Reihenfolge", "Name"]];

      myData.forEach((item) => {
        ergebnis.push([item.Pos_GuV_ID, item.Reihenfolge, item.Position_Name]);
      });

      return ergebnis;
    }
  } catch (error) {
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}

/**
 * Abfrage des Aufbaus der Aktiva für einen bestimmten Mandanten
 * @customfunction Aufbau_Aktiva
 * @param {string} Mandanten_ID_Name - Mandanten_ID_Name.
 * @returns {string[][]} A dynamic array with multiple results.
 */
async function Aufbau_Aktiva(Mandanten_ID_Name) {
  try {
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/clients/` +
      Mandanten_ID +
      "/aktiva";
    console.log(url);
    const myData = await callURL(url);

    const ergebnis = [["ID", "Reihenfolge", "Name"]];

    myData.forEach((item) => {
      ergebnis.push([item.Pos_Bilanz_ID, item.Reihenfolge, item.Position_Name]);
    });

    return ergebnis;
  } catch (error) {
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}

/**
 * Abfrage des Aufbaus der Passiva für einen bestimmten Mandanten
 * @customfunction Aufbau_Passiva
 * @param {string} Mandanten_ID_Name - Mandanten_ID_Name.
 * @returns {string[][]} A dynamic array with multiple results.
 */
async function Aufbau_Passiva(Mandanten_ID_Name) {
  try {
    const Mandanten_ID = getMandantenIDFromName(Mandanten_ID_Name);
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/clients/` +
      Mandanten_ID +
      "/passiva";
    console.log(url);
    const myData = await callURL(url);

    const ergebnis = [["ID", "Reihenfolge", "Name"]];

    myData.forEach((item) => {
      ergebnis.push([item.Pos_Bilanz_ID, item.Reihenfolge, item.Position_Name]);
    });

    return ergebnis;
  } catch (error) {
    const ergebnis = [
      ["Fehler:", "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."],
    ];
    return ergebnis;
  }
}
/**
 * Abfrage der GuV-Daten für einen bestimmten Mandanten, ein Bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart.
 * @customfunction DATEN_GUV
 * @param {string} Mandanten_ID - Mandanten_ID .
 * @param {string} Jahr Jahr.
 * @param {string} UStruktur_ID - UStruktur_ID.
 * @param {string} Datenart - Datenart.
 * @returns {string[][]} GuV des ausgewählten Mandantens.
 */
export async function Daten_GuV(Mandanten_ID, Jahr, UStruktur_ID, Datenart) {
  try {
    const Mandanten = getMandantenIDFromName(Mandanten_ID);
    const UStruktur = getUStrukturIDFromName(UStruktur_ID);
    let selectedDatenartID = Datenart;

    if (typeof Datenart === "string") {
      switch (Datenart.toLowerCase()) {
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
          selectedDatenartID = parseInt(Datenart);
      }
    }
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
      Mandanten +
      "/" +
      Jahr +
      "/" +
      UStruktur +
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

/**
 * Abfrage der GuV-Daten für einen bestimmten Mandanten, ein Bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart.
 * Param1- Mandanten_ID; Param2- Jahr_ID; Param3- UStruktur_ID; Param4- Datenart_ID;
 * @customfunction DATEN_GUV_WERT
 * @param {string} Mandanten_ID_Name - Имя или ID клиента.
 * @param {string} Jahr_ID - ID года.
 * @param {string} UStruktur_ID_Name - Имя или ID структуры компании.
 * @param {string} Datenart_ID - ID типа данных.
 * @param {number} row - Номер строки значения.
 * @param {number} column - Номер столбца значения.
 * @returns {string|number} Конкретное значение выбранного клиента.
 */
async function Daten_GuV_Wert(Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, Datenart_ID, row, column) {
  try {
    const Mandanten_ID = getMandantenIDFromName(Mandanten_ID_Name);
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
    let protocol = getFromLocalStorage("protocol") || "http";
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

    const adjustedRow = row - 1;
    const adjustedColumn = column - 1;

    if (adjustedRow < 0 || adjustedRow >= myData.length || adjustedColumn < 0 || adjustedColumn >= 14) {
      throw new Error(
        "Parameter nicht erkannt. Bitte geben Sie einen gültigen Servernamen, Port und Mandanten-ID ein."
      );
    }

    const selectedItem = myData[adjustedRow];
    const columns = [
      "Reihenfolge",
      "Position_Name",
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
    ];

    const value = selectedItem[columns[adjustedColumn]];
    return value;
  } catch (error) {
    return "Ошибка: " + error.message;
  }
}
CustomFunctions.associate("DATEN_GUV_WERT", Daten_GuV_Wert);
/**
 * Abfrage der GuV-Daten für einen bestimmten Mandanten, ein Bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart.
 * @customfunction DATEN_GUV2
 * @param {string} Mandant - Mandanten_ID.
 * @param {string} Jahr Jahr_ID.
 * @param {string} UStruktur - UStruktur_ID_Name.
 * @param {string} Datenart - Datenart_ID.
 * @returns {string[][]} GuV des ausgewählten Mandantens.
 */
async function Daten_GuV2(Mandanten, Jahr, UStruktur, Datenart) {
  try {
    const Mandanten_ID = getMandantenIDFromName(Mandanten);
    const UStruktur_ID = getUStrukturIDFromName(UStruktur);
    let selectedDatenartID = Datenart;

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
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/income/` +
      Mandanten_ID +
      "/" +
      Jahr +
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
CustomFunctions.associate("DATEN_GUV2", Daten_GuV2);

/**
 * Abfrage der Passiva-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart
 * @customfunction Daten_Bilanz
 * @param {string} Mandanten_ID_Name - Name des Mandanten.
 * @param {string} Jahr_ID - Jahr_ID.
 * @param {string} UStruktur_ID_Name - UStruktur_ID.
 * @param {string} Datenart_ID - Datenart-ID.
 * @returns {string[][]} A dynamic array with multiple results.
 */
async function Daten_Bilanz(Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, Datenart_ID) {
  try {
    const Mandanten_ID = getMandantenIDFromName(Mandanten_ID_Name);
    const UStruktur_ID = getUStrukturIDFromName(UStruktur_ID_Name);
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/balance/` +
      Mandanten_ID +
      "/" +
      Jahr_ID +
      "/" +
      UStruktur_ID +
      "/" +
      Datenart_ID;
    console.log(url);
    const myData = await callURL(url);

    const ergebnis = [
      ["ID", "Reihenfolge", "Position_Name", "Anfangsbestand", "Monat_Z", "Monat_A", "Monat_S", "Endbestand"],
    ];

    myData.forEach((item) => {
      ergebnis.push([
        item.Pos_Bilanz_ID,
        item.Reihenfolge,
        item.Position_Name,
        item.Anfangsbestend,
        item.Monat_Z +
          item.Monat1_Z +
          item.Monat2_Z +
          item.Monat3_Z +
          item.Monat4_Z +
          item.Monat5_Z +
          item.Monat6_Z +
          item.Monat7_Z +
          item.Monat8_Z +
          item.Monat9_Z +
          item.Monat10_Z +
          item.Monat11_Z +
          item.Monat12_Z,
        item.Monat_A +
          item.Monat1_A +
          item.Monat2_A +
          item.Monat3_A +
          item.Monat4_A +
          item.Monat5_A +
          item.Monat6_A +
          item.Monat7_A +
          item.Monat8_A +
          item.Monat9_A +
          item.Monat10_A +
          item.Monat11_A +
          item.Monat12_A,
        item.Monat_S +
          item.Monat1_S +
          item.Monat2_S +
          item.Monat3_S +
          item.Monat4_S +
          item.Monat5_S +
          item.Monat6_S +
          item.Monat7_S +
          item.Monat8_S +
          item.Monat9_S +
          item.Monat10_S +
          item.Monat11_S +
          item.Monat12_S,
        item.Endbestand,
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

/**
 * Abfrage der Passiva-Daten für einen bestimmten Mandanten, ein bestimmtes Jahr, eine bestimmte Unternehmensstruktur und eine bestimmte Datenart
 * @customfunction Daten_Bilanz2
 * @param {string} Mandanten_ID_Name - Mandanten_ID_Name.
 * @param {string} Jahr_ID - Jahr_ID.
 * @param {string} UStruktur_ID_Name - UStruktur_ID.
 * @param {string} Datenart_ID - Datenart-ID.
 * @returns {string[][]} A dynamic array with multiple results.
 */
async function Daten_Bilanz2(Mandanten_ID_Name, Jahr_ID, UStruktur_ID_Name, Datenart_ID) {
  try {
    const Mandanten_ID = getMandantenIDFromName(Mandanten_ID_Name);
    const UStruktur_ID = getUStrukturIDFromName(UStruktur_ID_Name);
    let protocol = getFromLocalStorage("protocol") || "http";
    const url =
      `${protocol}://${getFromLocalStorage("server")}:${getFromLocalStorage("port")}/api/data/balance/` +
      Mandanten_ID +
      "/" +
      Jahr_ID +
      "/" +
      UStruktur_ID +
      "/" +
      Datenart_ID;
    console.log(url);
    const myData = await callURL(url);

    const ergebnis = [
      [
        "ID",
        "Reihenfolge",
        "Position_Name",
        "Anfangsbestand",
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
        "Endbestand",
      ],
    ];

    myData.forEach((item) => {
      ergebnis.push([
        item.Pos_Bilanz_ID,
        item.Reihenfolge,
        item.Position_Name,
        item.Anfangsbestend,
        item.Monat1_S,
        item.Monat2_S,
        item.Monat3_S,
        item.Monat4_S,
        item.Monat5_S,
        item.Monat6_S,
        item.Monat7_S,
        item.Monat8_S,
        item.Monat9_S,
        item.Monat10_S,
        item.Monat11_S,
        item.Monat12_S,
        item.Endbestand,
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
const RegistrySubKey = "BA2API";
const RegistryValueServer = "BA2_Server";
const RegistryValuePort = "BA2_Port";

var server = "";

var port = "";

/**
 * Writes a message to console.log().
 * @customfunction DataGuV
 * @param {string} Mandant String to write.
 * @returns String to write.
 */

export function DataGuV(Mandant) {
  console.log(Mandant);

  return Mandant;
}

// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/excel/16-custom-functions/streaming-function.yaml
