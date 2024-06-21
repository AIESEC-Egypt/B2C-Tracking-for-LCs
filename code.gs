const sheet =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Dashoard");
const signups_sheet = SpreadsheetApp.openById(
  "1AE6mPcqDDpHXI7nkNphedqAVDPQyLd3RqszxSUDk2jU"
).getSheetByName("Rank");

function signups() {
  let rows = [];

  for (let i = 3; i <= 21; i++) {
    rows.push(signups_sheet.getRange(i, 2, 1, 5).getValues().flat());
  }
  rows.sort(sortByTotal);

  sheet.getRange(4, 2, 19, 5).setValues(rows);
}

function sortByTotal(a, b) {
  if (a[4] === b[4]) {
    return 0;
  } else {
    return a[4] > b[4] ? -1 : 1;
  }
}

function applicants() {
  const lcCodes = {
    O6U: 2820,
    AASTa: 1788,
    AASTc: 1322,
    ASU: 1789,
    ALEX: 899,
    AUC: 1489,
    "Beni Suef": 2126,
    CU: 1064,
    Damieta: 109,
    GUC: 257,
    Helwan: 2124,
    KSU: 2524,
    "Luxor&Aswan": 2114,
    Mansoura: 171,
    Menofia: 1727,
    MIU: 2125,
    MSA: 2817,
    MUST: 2818,
    Suez: 15,
    Tanta: 1725,
    Zagazig: 1114,
    "6O(Closed)": 152,
    "MC Egypt": 2387,
    "AIESEC in Egypt": 1609,
  };
  let rows = [];
  for (let i = 3; i <= 21; i++) {
    var url =
      "https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=" +
      "&start_date=" +
      "01/02/2023" +
      "&end_date=" +
      "31/01/2024" +
      "&performance_v3%5Boffice_id%5D=" +
      lcCodes[sheet.getRange(i, 8, 1, 1).getValue()];
    var response = UrlFetchApp.fetch(url, { method: "GET" }).getContentText();
    var data = JSON.parse(response);
    let ogv = data[`o_applied_7`].applicants.value;
    let ogta = data[`o_applied_8`].applicants.value;
    let ogte = data[`o_applied_9`].applicants.value;
    let total = ogv + ogta + ogte;
    rows.push([sheet.getRange(i, 8, 1, 1).getValue(), ogv, ogta, ogte, total]);
  }
  Logger.log(rows);
  rows.sort(sortByTotal);
  sheet.getRange(3, 8, 19, 5).setValues(rows);
}
