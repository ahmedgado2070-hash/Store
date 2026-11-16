const PRODUCTS_SHEET = "Products";
const ORDERS_SHEET = "Orders";

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const products = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    products.push({
      id: String(data[i][0]),
      name: String(data[i][1]),
      category: String(data[i][2]),
      image: String(data[i][3]),
      oldPrice: Number(data[i][4]) || 0,
      newPrice: Number(data[i][5]) || 0,
      shipping: Number(data[i][6]) || 0
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(products))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORDERS_SHEET);
    const payload = JSON.parse(e.postData.contents);

    const lastRow = sheet.getLastRow();
    let nextSerial = 1;

    if (lastRow > 1) {
      const lastId = String(sheet.getRange(lastRow, 1).getValue());
      const parts = lastId.split("-");
      const lastNum = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(lastNum)) {
        nextSerial = lastNum + 1;
      } else {
        nextSerial = lastRow;
      }
    }

    const productIds = payload.productIds || "000";
    const orderId = `ORD-${productIds}-${nextSerial}`;

    const now = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([
      orderId,
      now,
      payload.customer,
      "'" + payload.phone,
      payload.address,
      payload.productName,
      payload.qty,
      payload.total
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", orderId: orderId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}