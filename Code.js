// 1. دوال التوجيه الأساسية للصفحات
function doGet(e) {
  var page = e.parameter.page;
  if (page == 'courier') {
    return HtmlService.createHtmlOutputFromFile('Courier').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page == 'order') {
    return HtmlService.createHtmlOutputFromFile('Order').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page == 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page == 'user') {
    return HtmlService.createHtmlOutputFromFile('User').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page == 'tracking') {
    return HtmlService.createHtmlOutputFromFile('Tracking').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page == 'signup') {
    return HtmlService.createHtmlOutputFromFile('Signup').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('Index').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function getAppUrl() { return ScriptApp.getService().getUrl(); }

function getSystemFolder(folderType) {
  var mainFolderName = "Dropex_System";
  var mainFolders = DriveApp.getFoldersByName(mainFolderName);
  var mainFolder = mainFolders.hasNext() ? mainFolders.next() : DriveApp.createFolder(mainFolderName);
  var subFolderName = "";
  if (folderType === "PDF") subFolderName = "Waybills_PDF";
  else if (folderType === "Delivered") subFolderName = "Delivered_Orders";
  else if (folderType === "Returned") subFolderName = "Returned_Orders";
  var subFolders = mainFolder.getFoldersByName(subFolderName);
  return subFolders.hasNext() ? subFolders.next() : mainFolder.createFolder(subFolderName);
}

function courierLogin(email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Couriers");
  var data = sheet.getDataRange().getValues();
  var userEmail = String(email).trim().toLowerCase();
  var userPass = String(password).trim();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][1]).trim().toLowerCase();
    var rowPass = String(data[i][2]).trim();
    if (rowEmail === userEmail && rowPass === userPass) { return { success: true, courierName: data[i][0] }; }
  }
  return { success: false, error: "بيانات الدخول غير صحيحة" };
}

function getAreasAndPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Areas");
  var data = sheet.getDataRange().getValues();
  var areas = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] != "") { areas.push({ name: data[i][0], price: data[i][1] }); }
  }
  return areas;
}

function createNewOrder(formObj) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    var newTrackId = generateTrackingNumber();
    var dateAdded = new Date();
    var dateString = Utilities.formatDate(dateAdded, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var orderPin = Math.floor(100000 + Math.random() * 900000).toString();

    var productPrice = parseFloat(formObj.productPrice) || 0;
    var systemShippingFee = parseFloat(formObj.systemShippingCost) || 0;
    var merchantShippingPrice = parseFloat(formObj.deliveryCost) || systemShippingFee;
    
    var receiverDeliveryShare = 0;
    if (formObj.deliveryPaidBy === "على المستلم") { receiverDeliveryShare = merchantShippingPrice; }
    else if (formObj.deliveryPaidBy === "تقسيم") { receiverDeliveryShare = parseFloat(formObj.receiverShare) || 0; }

    var totalToCollectFromReceiver = productPrice + receiverDeliveryShare;
    var barcodeUrl = "https://quickchart.io/barcode?type=code128&text=" + newTrackId + "&height=60&includeText=true";
    var barcodeBlob = UrlFetchApp.fetch(barcodeUrl).getBlob();
    var base64Barcode = Utilities.base64Encode(barcodeBlob.getBytes());
    var barcodeImgSrc = "data:image/png;base64," + base64Barcode;

    // تجهيز تسميات البيانات الجديدة للبوليصة
    var inspectTxt = formObj.openForInspection ? "نعم (يسمح بالفتح)" : "لا (ممنوع الفتح)";
    var payMethodTxt = "نقداً (COD)";
    if(formObj.shippingPaymentMethod === "Wallet") payMethodTxt = "محفظة إلكترونية";
    else if(formObj.shippingPaymentMethod === "InstaPay") payMethodTxt = "إنستا باي";

    var extraPhonesHtml = "";
    if (formObj.receiverPhone2) extraPhonesHtml += `<div>هاتف إضافي: ${formObj.receiverPhone2}</div>`;
    if (formObj.receiverPhone3) extraPhonesHtml += `<div>هاتف إضافي: ${formObj.receiverPhone3}</div>`;

    var htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: 'Tahoma', 'Arial', sans-serif; color: #333; line-height: 1.4; padding: 10px; }
        .header { text-align: center; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; margin-bottom: 15px; }
        .header h1 { margin: 0; color: #2c3e50; font-size: 24px; font-weight: bold; }
        .barcode { margin-top: 10px; width: 220px; height: 50px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        .info-table td { width: 50%; padding: 10px; vertical-align: top; border: 1px solid #bdc3c7; background-color: #fafafa; }
        .info-table h3 { margin: 0 0 5px 0; color: #2980b9; font-size: 14px; border-bottom: 1px solid #eee; }
        .instructions-table td { padding: 8px; border: 1px solid #bdc3c7; background-color: #fff9f0; font-size: 13px; }
        .financial-table th { background-color: #2c3e50; color: white; padding: 10px; font-size: 16px; }
        .financial-table td { padding: 10px; border: 1px solid #bdc3c7; text-align: center; font-size: 18px; }
        .total-row { background-color: #ecf0f1; font-weight: bold; }
        .amount { color: #e74c3c; font-size: 22px; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>Dropex</h1>
        <p style="font-size:12px; margin:0;">رقم الشحنة: ${newTrackId} | ${dateString}</p>
        <img src="${barcodeImgSrc}" class="barcode" alt="Barcode">
      </div>
      
      <table class="info-table">
        <tr>
          <td>
            <h3>إلى (المستلم)</h3>
            <p><strong>الاسم:</strong> ${formObj.receiverName}</p>
            <p><strong>الهاتف:</strong> <span dir="ltr">${formObj.receiverPhone}</span></p>
            <p><strong>العنوان:</strong> ${formObj.receiverAddress} - ${formObj.receiverArea}</p>
          </td>
          <td>
            <h3>من (المرسل)</h3>
            <p><strong>الاسم:</strong> ${formObj.senderName}</p>
            <p><strong>الهاتف:</strong> <span dir="ltr">${formObj.senderPhone}</span></p>
          </td>
        </tr>
      </table>

      <table class="instructions-table">
        <tr>
          <td width="50%"><strong>مسموح بالفتح للمعاينة:</strong> <span class="badge ${formObj.openForInspection ? 'badge-info' : ''}">${inspectTxt}</span></td>
          <td width="50%"><strong>طريقة دفع الشحن:</strong> ${payMethodTxt}</td>
        </tr>
        <tr>
          <td colspan="2"><strong>الشحن مدفوع بواسطة:</strong> ${formObj.deliveryPaidBy} 
          ${formObj.deliveryPaidBy === 'تقسيم' ? '(' + formObj.divisionDetail + ')' : ''}</td>
        </tr>
      </table>

      <table class="financial-table">
        <tr><th colspan="2">المبالغ المطلوب تحصيلها من المستلم</th></tr>
        <tr><td>سعر المنتج</td><td>${productPrice} ج.م</td></tr>
        <tr><td>نصيب المستلم من الشحن</td><td>${receiverDeliveryShare} ج.م</td></tr>
        <tr class="total-row"><td>الإجمالي المطلوب تحصيله</td><td class="amount">${totalToCollectFromReceiver} ج.م</td></tr>
      </table>
    </body>
    </html>
    `;

    var htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, "temp.html");
    var pdfBlob = htmlBlob.getAs(MimeType.PDF);
    pdfBlob.setName("Waybill_" + newTrackId + ".pdf");
    var pdfFolder = getSystemFolder("PDF");
    var pdfFile = pdfFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var waybillUrl = pdfFile.getUrl();
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // الاستقبال من الواجهة
    var openForInspection = formObj.openForInspection ? "نعم" : "لا";
    var shippingPaymentMethod = formObj.shippingPaymentMethod || "COD";

    // التعديل: زيادة المصفوفة لاستيعاب الهواتف الإضافية و "سعر البوليصة" في AJ
    var rowData = new Array(36).fill("");
    rowData[0] = newTrackId; rowData[1] = String(formObj.senderEmail).trim().toLowerCase();
    rowData[2] = String(formObj.senderName).trim(); rowData[3] = formObj.receiverName;
    rowData[4] = "تم الإنشاء"; rowData[5] = dateAdded;
    rowData[6] = formObj.productPrice; rowData[7] = systemShippingFee; // العمود H لسعر دروبكس الثابت
    rowData[8] = formObj.deliveryPaidBy; rowData[9] = 0;
    rowData[10] = waybillUrl;
    rowData[11] = formatPhoneForSheet(formObj.senderPhone); rowData[12] = formObj.senderAddress;
    rowData[13] = formObj.senderArea; rowData[14] = formObj.receiverEmail;
    rowData[15] = formatPhoneForSheet(formObj.receiverPhone); rowData[16] = formObj.receiverAddress;
    rowData[17] = formObj.receiverArea; rowData[26] = orderPin;
    rowData[27] = totalToCollectFromReceiver;
    rowData[28] = ""; // حالة التصفية
    
    // البيانات التقنية والهواتف والأسعار الإضافية
    rowData[31] = openForInspection; // AF
    rowData[32] = shippingPaymentMethod; // AG
    rowData[33] = formatPhoneForSheet(formObj.receiverPhone2); // AH
    rowData[34] = formatPhoneForSheet(formObj.receiverPhone3); // AI
    rowData[35] = merchantShippingPrice; // AJ (السعر الذي وضعه التاجر وسيظهر للعميل)

    sheet.appendRow(rowData);
    return { success: true, trackingId: newTrackId, pin: orderPin, totalToCollect: totalToCollectFromReceiver, receiverDeliveryShare: receiverDeliveryShare, pdfUrl: waybillUrl, pdfBase64: pdfBase64 };
  } catch (e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

/**
 * Normalizes Egyptian phone numbers to international format (+20...)
 * to ensure Google Sheets treats them as text and for compatibility with WhatsApp/Calling.
 */
function formatPhoneForSheet(phone) {
  if (phone === null || phone === undefined) return "";
  var phoneStr = String(phone).replace(/[\s\-\(\)]/g, ""); // Clean formatting characters
  if (phoneStr === "") return "";

  // 1. If it already starts with +20, just return it
  if (phoneStr.startsWith("+20")) return phoneStr;
  
  // 2. If it starts with 20 (but no +), add the +
  if (phoneStr.startsWith("20") && phoneStr.length >= 12) return "+" + phoneStr;

  // 3. If it starts with 0 (e.g., 012...), add +2
  if (phoneStr.startsWith("0")) return "+2" + phoneStr;

  // 4. If it starts with 1 (and is likely an Egyptian mobile number without 0), add +20
  if (phoneStr.startsWith("1") && (phoneStr.length === 10)) return "+20" + phoneStr;

  // 5. If it's something else but looks like it needs the country code
  if (phoneStr.length === 11 && phoneStr.startsWith("01")) return "+2" + phoneStr;

  return phoneStr;
}

function generateTrackingNumber() {
  var date = new Date();
  var randomNum = Math.floor(100 + Math.random() * 900);
  var timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
  return "TRK-" + timestamp + "-" + randomNum;
}

function getCourierOrders(email, password) {
  var courierSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Couriers");
  var courierData = courierSheet.getDataRange().getValues();
  var courierName = "";
  var isValid = false;
  var userEmail = String(email).trim().toLowerCase();
  var userPass = String(password).trim();

  for (var i = 1; i < courierData.length; i++) {
    var rowEmail = String(courierData[i][1]).trim().toLowerCase();
    var rowPass = String(courierData[i][2]).trim();
    if (rowEmail === userEmail && rowPass === userPass) { courierName = courierData[i][0]; isValid = true; break; }
  }
  if (!isValid) return { error: "الإيميل أو الرقم السري غير صحيح" };

  var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  var orderData = orderSheet.getDataRange().getValues();
  var pendingOrders = [];
  for (var j = 1; j < orderData.length; j++) {
    var status = orderData[j][4];
    var assignedCourier = orderData[j][18];
    if (assignedCourier == courierName && status != "تم التوصيل" && status != "مرتجع" && status != "ملغي") {
      var fullAddress = orderData[j][16] + " (" + orderData[j][17] + ")";
      var totalToCollect = parseFloat(orderData[j][27]) || 0;
      var inspection = orderData[j][31] || "لا";
      var payMethod = orderData[j][32] || "COD";
      
      pendingOrders.push({ 
        row: j + 1, 
        id: orderData[j][0], 
        sender: orderData[j][2], 
        senderPhone: orderData[j][11], 
        receiver: orderData[j][3], 
        receiverPhone: orderData[j][15], 
        address: fullAddress, 
        amount: totalToCollect,
        inspection: inspection,
        payMethod: payMethod
      });
    }
  }
  return { error: null, courierName: courierName, orders: pendingOrders };
}

function processCourierUpdate(rowNumber, actionType, imageData, filename, reason, location) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    var podUrl = "";
    if (imageData && filename) {
      var folderType = (actionType === 'delivered') ? "Delivered" : "Returned";
      var folder = getSystemFolder(folderType);
      var contentType = imageData.substring(5, imageData.indexOf(';'));
      var bytes = Utilities.base64Decode(imageData.split(',')[1]);
      var blob = Utilities.newBlob(bytes, contentType, filename);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      podUrl = file.getUrl();
    }
    var mapUrl = "";
    if (location && location.lat) { mapUrl = "https://www.google.com/maps/search/?api=1&query=" + location.lat + "," + location.lng; }
    var newStatus = (actionType === 'delivered') ? "تم التوصيل" : "مرتجع";
    sheet.getRange(rowNumber, 5).setValue(newStatus);

    // تسجيل التواريخ المخصصة حسب المرحلة ( AE للنتيجة النهائية)
    if (newStatus === "تم التوصيل" || newStatus === "مرتجع" || newStatus === "ملغي") {
      sheet.getRange(rowNumber, 31).setValue(new Date()); // AE (31)
    }

    if (podUrl != "") sheet.getRange(rowNumber, 20).setValue(podUrl);    // T (20)
    if (mapUrl != "") sheet.getRange(rowNumber, 21).setValue(mapUrl);    // U (21)
    if (actionType === 'returned' && reason != "") sheet.getRange(rowNumber, 22).setValue(reason); // V (22)
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getDashboardStats(password) {
  try {
    var adminPassword = "admin123";
    if (password !== adminPassword) return { error: "كلمة المرور غير صحيحة" };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();

    var courierSheet = ss.getSheetByName("Couriers");
    var courierData = courierSheet ? courierSheet.getDataRange().getValues() : [];
    var couriersList = [];
    for (var c = 1; c < courierData.length; c++) { if (courierData[c][0] != "") couriersList.push(courierData[c][0]); }

    var stats = { totalOrders: 0, deliveredOrders: 0, pendingOrders: 0, outForDelivery: 0, totalCollectedAmount: 0, todayCollectedAmount: 0, totalNetProfit: 0, todayNetProfit: 0, recentOrders: [] };
    var todayStr = new Date().toDateString();

    // نمر على البيانات بالعكس لنأخذ أحدث الطلبات أولاً للجدول
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] == "") continue;

      var status = data[i][4];
      stats.totalOrders++;

      var netProfit = parseFloat(data[i][25]) || 0;
      var amountToCollect = parseFloat(data[i][27]) || parseFloat(data[i][6]) || 0;
      var productPrice = parseFloat(data[i][6]) || 0;
      var deliveryCost = parseFloat(data[i][7]) || 0;
      var paidBy = String(data[i][8] || "على المستلم").trim();
      var pickupPrice = parseFloat(data[i][9]) || 0;

      var updateDateStr = "";
      // تحديد تاريخ "اليوم" بناءً على الحالة
      if (status === "تم التوصيل" || status === "مرتجع" || status === "ملغي") {
        if (data[i][30] instanceof Date) updateDateStr = data[i][30].toDateString(); // AE
      } else if (status === "خرج للتوصيل" || status === "خرج للتسليم") {
        if (data[i][29] instanceof Date) updateDateStr = data[i][29].toDateString(); // AD
      } else if (status === "في المخزن") {
        if (data[i][22] instanceof Date) updateDateStr = data[i][22].toDateString(); // W
      }

      if (updateDateStr === "" && data[i][5] instanceof Date) {
        updateDateStr = data[i][5].toDateString(); // F (للحالات الجديدة)
      }

      if (status == "تم التوصيل") {
        stats.deliveredOrders++;
        stats.totalCollectedAmount += amountToCollect;
        stats.totalNetProfit += netProfit;
        if (updateDateStr === todayStr) { stats.todayCollectedAmount += amountToCollect; stats.todayNetProfit += netProfit; }
      } else if (status == "مرتجع") {
        stats.totalNetProfit += netProfit;
        if (updateDateStr === todayStr) { stats.todayNetProfit += netProfit; }
      } else if (status == "خرج للتسليم" || status == "خرج للتوصيل") {
        stats.outForDelivery++;
      } else if (status == "قيد الانتظار" || status == "تم الإنشاء" || status == "في المخزن") {
        stats.pendingOrders++;
      }

      // إضافة الطلب للقائمة (نكتفي بأحدث 500 طلب لتحسين الأداء)
      if (stats.recentOrders.length < 500) {
        var merchantNet = 0;
        if (status === "تم التوصيل") {
          merchantNet = (paidBy === "على المرسل") ? (productPrice - deliveryCost - pickupPrice) : (productPrice - pickupPrice);
        } else if (status === "مرتجع") {
          merchantNet = (paidBy === "على المرسل") ? (0 - deliveryCost - pickupPrice) : (0 - pickupPrice);
        } else {
          merchantNet = (paidBy === "على المرسل") ? (productPrice - deliveryCost - pickupPrice) : (productPrice - pickupPrice);
        }

        stats.recentOrders.push({
          rowIndex: i + 1,
          id: data[i][0],
          pin: data[i][26] || "",
          sender: data[i][2] || "",
          receiver: data[i][3] || "",
          address: (data[i][16] || "") + " - " + (data[i][17] || ""),
          status: status,
          productPrice: productPrice,
          deliveryCost: deliveryCost,
          paidBy: paidBy,
          pickupPrice: pickupPrice,
          merchantNet: merchantNet,
          courier: data[i][18] || "",
          gas: data[i][23] || 0,
          maintenance: data[i][24] || 0,
          netProfit: netProfit,
          senderPhone: data[i][11] || "",
          senderAddress: data[i][12] || "",
          senderArea: data[i][13] || "",
          receiverPhone: data[i][15] || "",
          receiverArea: data[i][17] || "",
          waybillUrl: data[i][10] || "",
          inspection: data[i][31] || "لا",
          payMethod: data[i][32] || "COD",
          isSettled: (String(data[i][28]).trim() === "تمت التصفية" || data[i][28] === true)
        });
      }
    }
    return { error: null, stats: stats, couriersList: couriersList };
  } catch (e) {
    return { error: "حدث خطأ في الخادم: " + e.toString() };
  }
}


function updateOrderFromAdmin(dataObj) {
  try {
    var rowIndex = dataObj.rowIndex;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    var oldData = sheet.getRange(rowIndex, 1, 1, 33).getValues()[0]; // جلب السطر الحالي بالكامل
    
    var oldStatus = oldData[4]; // العمود E
    var newStatus = dataObj.status;
    
    // 1. تحديث البيانات الأساسية
    sheet.getRange(rowIndex, 3).setValue(dataObj.senderName); // C
    sheet.getRange(rowIndex, 4).setValue(dataObj.receiverName); // D
    sheet.getRange(rowIndex, 5).setValue(newStatus); // E
    sheet.getRange(rowIndex, 7).setValue(parseFloat(dataObj.productPrice) || 0); // G
    sheet.getRange(rowIndex, 8).setValue(parseFloat(dataObj.deliveryCost) || 0); // H
    sheet.getRange(rowIndex, 9).setValue(dataObj.paidBy); // I
    sheet.getRange(rowIndex, 10).setValue(parseFloat(dataObj.pickupPrice) || 0); // J
    sheet.getRange(rowIndex, 12).setValue(dataObj.senderPhone); // L
    sheet.getRange(rowIndex, 13).setValue(dataObj.senderAddress); // M
    sheet.getRange(rowIndex, 14).setValue(dataObj.senderArea); // N
    sheet.getRange(rowIndex, 16).setValue(formatPhoneForSheet(dataObj.receiverPhone)); // P
    sheet.getRange(rowIndex, 17).setValue(dataObj.receiverAddress); // Q
    sheet.getRange(rowIndex, 18).setValue(dataObj.receiverArea); // R
    sheet.getRange(rowIndex, 19).setValue(dataObj.courierName); // S
    sheet.getRange(rowIndex, 24).setValue(parseFloat(dataObj.gas) || 0); // X
    sheet.getRange(rowIndex, 25).setValue(parseFloat(dataObj.maintenance) || 0); // Y
    sheet.getRange(rowIndex, 26).setValue(parseFloat(dataObj.netProfit) || 0); // Z
    
    sheet.getRange(rowIndex, 32).setValue(dataObj.inspection); // AF
    sheet.getRange(rowIndex, 33).setValue(dataObj.payMethod); // AG
    sheet.getRange(rowIndex, 36).setValue(parseFloat(dataObj.merchantDeliveryCost) || 0); // AJ

    // 2. منطق التواريخ الذكي: تحديث التاريخ فقط إذا تغيرت الحالة
    if (newStatus !== oldStatus) {
      if (newStatus === "في المخزن") sheet.getRange(rowIndex, 23).setValue(new Date()); // W
      else if (newStatus === "خرج للتوصيل" || newStatus === "خرج للتسليم") sheet.getRange(rowIndex, 30).setValue(new Date()); // AD
      else if (newStatus === "تم التوصيل" || newStatus === "مرتجع" || newStatus === "ملغي") sheet.getRange(rowIndex, 31).setValue(new Date()); // AE
    }

    // 3. منطق البوليصة الذكي: هل نحتاج لتوليد بوليصة جديدة؟
    var waybillFieldsChanged = (
      dataObj.receiverName !== oldData[3] ||
      dataObj.receiverPhone !== oldData[15] ||
      dataObj.receiverAddress !== oldData[16] ||
      dataObj.receiverArea !== oldData[17] ||
      dataObj.productPrice != oldData[6] ||
      dataObj.deliveryCost != oldData[7] ||
      dataObj.paidBy !== oldData[8] ||
      dataObj.inspection !== oldData[31] ||
      dataObj.payMethod !== oldData[32] ||
      dataObj.merchantDeliveryCost != oldData[35]
    );

    if (waybillFieldsChanged) {
      var newWaybillUrl = regenerateWaybill(dataObj, oldData[0], oldData[5]); // rowIndex, id, originalDate
      sheet.getRange(rowIndex, 11).setValue(newWaybillUrl); // K
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * دالة مستقلة لإعادة إنشاء البوليصة PDF عند التعديل
 */
function regenerateWaybill(formObj, trackId, dateAdded) {
  var dateStr = (dateAdded instanceof Date) ? Utilities.formatDate(dateAdded, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(dateAdded);
  
  var inspectTxt = (formObj.inspection === "نعم") ? "نعم (يسمح بالفتح)" : "لا (ممنوع الفتح)";
  var payMethodTxt = "نقداً (COD)";
  if(formObj.payMethod === "WALLET_INSTAPAY") payMethodTxt = "محفظة / إنستا باي";

  var productPrice = parseFloat(formObj.productPrice) || 0;
  var deliveryCost = parseFloat(formObj.deliveryCost) || 0;
  var receiverDeliveryShare = 0;
  if (formObj.paidBy === "على المستلم") receiverDeliveryShare = deliveryCost;
  var totalToCollect = productPrice + receiverDeliveryShare;

  var barcodeUrl = "https://quickchart.io/barcode?type=code128&text=" + trackId + "&height=60&includeText=true";
  var barcodeBlob = UrlFetchApp.fetch(barcodeUrl).getBlob();
  var barcodeImgSrc = "data:image/png;base64," + Utilities.base64Encode(barcodeBlob.getBytes());

  var html = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head><meta charset="UTF-8"><style>
      body { font-family: 'Tahoma', sans-serif; padding: 10px; font-size: 13px; }
      .header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 15px; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
      td, th { border: 1px solid #ccc; padding: 8px; text-align: right; }
      .total { background: #f9f9f9; font-weight: bold; font-size: 16px; color: #d35400; }
    </style></head>
    <body>
      <div class="header">
        <h2>Dropex Shipping</h2>
        <p>رقم التتبع: ${trackId} | ${dateStr}</p>
        <img src="${barcodeImgSrc}" style="width:200px;">
      </div>
      <table>
        <tr>
          <td><strong>المستلم:</strong> ${formObj.receiverName}<br><strong>الهاتف:</strong> ${formObj.receiverPhone}<br><strong>العنوان:</strong> ${formObj.receiverAddress} - ${formObj.receiverArea}</td>
          <td><strong>المرسل:</strong> ${formObj.senderName}<br><strong>الهاتف:</strong> ${formObj.senderPhone}</td>
        </tr>
      </table>
      <table>
        <tr><td><strong>المعاينة:</strong> ${inspectTxt}</td><td><strong>الدفع:</strong> ${payMethodTxt}</td></tr>
      </table>
      <table>
        <tr><th>البيان</th><th>المبلغ</th></tr>
        <tr><td>سعر المنتج</td><td>${productPrice} ج.م</td></tr>
        <tr><td>تكلفة الشحن</td><td>${receiverDeliveryShare} ج.م</td></tr>
        <tr class="total"><td>الإجمالي المطلوب</td><td>${totalToCollect} ج.م</td></tr>
      </table>
    </body></html>`;

  var blob = Utilities.newBlob(html, MimeType.HTML, "temp.html").getAs(MimeType.PDF);
  blob.setName("Waybill_Updated_" + trackId + ".pdf");
  var folder = getSystemFolder("PDF");
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


function getOrderStatus(trackingId, pinCode) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();
    var searchId = String(trackingId).trim().toUpperCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === searchId) {
        var rowData = data[i];
        var status = rowData[4];
        var publicData = {
          id: rowData[0],
          status: status,
          dateAdded: rowData[5],
          lastUpdate: rowData[30] || rowData[29] || rowData[22] || rowData[5],
          timestampCreated: rowData[5],
          timestampWarehouse: rowData[22],
          timestampShipping: rowData[29],
          timestampFinal: rowData[30],
          inspection: rowData[31] || "لا"
        };
        if (!pinCode) return JSON.stringify({ error: null, isPublicOnly: true, data: publicData });
        if (String(pinCode).trim() !== String(rowData[26]).trim()) return JSON.stringify({ error: "الرقم السري غير صحيح." });
        var courierName = rowData[18];
        var courierPhone = "غير متوفر";
        if (courierName && courierName !== "") {
          var courierSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Couriers");
          var courierData = courierSheet.getDataRange().getValues();
          for (var c = 1; c < courierData.length; c++) { if (courierData[c][0] == courierName) { courierPhone = courierData[c][3] || "غير متوفر"; break; } }
        }
        var privateData = {
          sender: rowData[2],
          amount: rowData[27] || rowData[6],
          address: rowData[16] + " - " + rowData[17],
          courier: courierName || "لم يتم التحديد",
          courierPhone: courierPhone,
          payMethod: rowData[32] || "COD"
        };
        return JSON.stringify({ error: null, isPublicOnly: false, data: publicData, privateData: privateData });
      }
    }
    return JSON.stringify({ error: "لم يتم العثور على الشحنة." });
  } catch (e) { return JSON.stringify({ error: "حدث خطأ: " + e.toString() }); }
}

function searchOrders(searchTerm, statusFilter) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  var data = sheet.getDataRange().getValues();
  var results = [];
  searchTerm = searchTerm ? String(searchTerm).trim().toLowerCase() : "";
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == "") continue;
    var id = String(data[i][0]).toLowerCase(); var sender = String(data[i][2]).toLowerCase(); var receiver = String(data[i][3]).toLowerCase(); var phone = String(data[i][15]).toLowerCase(); var status = data[i][4];
    var matchesSearch = (searchTerm === "" || id.includes(searchTerm) || sender.includes(searchTerm) || receiver.includes(searchTerm) || phone.includes(searchTerm));
    var matchesStatus = (!statusFilter || statusFilter === "" || status === statusFilter);
    if (matchesSearch && matchesStatus) { results.push({ rowIndex: i + 1, id: data[i][0], sender: data[i][2], receiver: data[i][3], phone: data[i][15], status: status, total: data[i][27] || data[i][6], date: data[i][5] }); }
  }
  return results;
}

function userLogin(email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return { success: false, error: "شيت Users غير موجود." };
  var data = sheet.getDataRange().getValues();
  var userEmail = String(email).trim().toLowerCase();
  var userPass = String(password).trim();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][1]).trim().toLowerCase();
    var rowPass = String(data[i][2]).trim();
    if (rowEmail === userEmail && rowPass === userPass) {
      return {
        success: true,
        userData: {
          name: String(data[i][0]).trim(),
          email: rowEmail,
          phone: String(data[i][3]).trim(),
          address: data[i][4],
          area: data[i][5],
          address2: data[i][7] || "",
          phone2: data[i][8] || "",
          phone3: data[i][9] || ""
        }
      };
    }
  }
  return { success: false, error: "الإيميل أو الرقم السري غير صحيح" };
}

function getUserDashboardStats(email, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  var data = sheet.getDataRange().getValues();

  var stats = { totalOrders: 0, deliveredOrders: 0, pendingOrders: 0, returnedOrders: 0, currentOwed: 0, totalHistoricalAmount: 0, ordersList: [] };

  var searchEmail = String(email).trim().toLowerCase();
  var searchName = String(name).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == "") continue;
    var rowEmail = String(data[i][1]).trim().toLowerCase();
    var rowName = String(data[i][2]).trim().toLowerCase();

    var isMatch = false;
    if (searchEmail !== "" && rowEmail === searchEmail) isMatch = true;
    if (searchName !== "" && rowName === searchName && searchEmail === "") isMatch = true;

    if (isMatch) {
      stats.totalOrders++;
      var status = data[i][4];

      var productPrice = parseFloat(data[i][6]) || 0;
      var deliveryCost = parseFloat(data[i][7]) || 0;
      var paidBy = String(data[i][8]).trim();
      var pickupPrice = parseFloat(data[i][9]) || 0;

      var merchantNet = 0;

      if (status === "تم التوصيل") {
        if (paidBy === "على المرسل") {
          merchantNet = productPrice - deliveryCost - pickupPrice;
        } else {
          merchantNet = productPrice - pickupPrice;
        }
      }
      else if (status === "مرتجع") {
        if (paidBy === "على المرسل") {
          merchantNet = 0 - deliveryCost - pickupPrice;
        } else {
          merchantNet = 0 - pickupPrice;
        }
      }

      var isSettled = (String(data[i][28]).trim() === "تمت التصفية" || data[i][28] === true);

      // الحالة المالية: تم التوصيل أو مرتجع أو (ملغي ببيك أب)
      if (status === "تم التوصيل") {
        var extraShippingProfit = Math.max(0, (parseFloat(data[i][35]) || 0) - (parseFloat(data[i][7]) || 0));
        var finalNet = (paidBy === "على المرسل") ? (productPrice - deliveryCost - pickupPrice) : (productPrice + extraShippingProfit - pickupPrice);
        
        stats.deliveredOrders++;
        stats.totalHistoricalAmount += productPrice;
        if (!isSettled) {
          stats.currentOwed += finalNet;
        }
      }
      else if (status === "مرتجع") {
        stats.returnedOrders++;
        if (!isSettled) {
          stats.currentOwed += merchantNet;
        }
      }
      else if (status === "ملغي") {
        // الطلب الملغي لا يحسب إلا إذا كان فيه بيك أب ولم يتم تصفيته
        if (pickupPrice > 0 && !isSettled) {
          stats.currentOwed += (0 - pickupPrice); // يخصم سعر البيك أب من التاجر
        }
      }
      else if (status === "قيد الانتظار" || status === "تم الإنشاء" || status === "خرج للتسليم" || status === "خرج للتوصيل" || status === "في المخزن") {
        stats.pendingOrders++;
      }

      var safeDateStr = (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(data[i][5]);
      var deliveryDateStr = "";
      if ((status === "تم التوصيل" || status === "مرتجع") && data[i][22]) {
        deliveryDateStr = (data[i][22] instanceof Date) ? Utilities.formatDate(data[i][22], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(data[i][22]);
      }

      stats.ordersList.unshift({
        id: data[i][0],
        pin: data[i][26],
        receiver: data[i][3],
        address: data[i][16] + " - " + data[i][17],
        status: status,
        orderDate: safeDateStr,
        deliveryDate: deliveryDateStr,
        productPrice: productPrice,
        paidBy: paidBy,
        pickupPrice: pickupPrice,
        merchantNet: merchantNet,
        podImage: data[i][19] || "",
        location: data[i][20] || "",
        waybillUrl: data[i][10]
      });
    }
  }
  return { success: true, stats: stats };
}


function processGridOrders(ordersArray, userData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  var areasData = getAreasAndPrices();
  var addedCount = 0;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    for (var i = 0; i < ordersArray.length; i++) {
      var recName = String(ordersArray[i].name).trim();
      var recPhone = String(ordersArray[i].phone).trim();
      var recEmail = String(ordersArray[i].recEmail).trim();
      var recAddress = String(ordersArray[i].address).trim();
      var recArea = String(ordersArray[i].area).trim();
      var productPrice = parseFloat(ordersArray[i].price) || 0;
      var deliveryPaidBy = String(ordersArray[i].paidBy).trim() || "على المستلم";

      if (recName === "" || recPhone === "") continue;

      var deliveryCost = 0;
      for (var a = 0; a < areasData.length; a++) {
        if (String(areasData[a].name).trim() === recArea) { deliveryCost = parseFloat(areasData[a].price) || 0; break; }
      }

      var receiverDeliveryShare = 0;
      var totalToCollect = productPrice;
      if (deliveryPaidBy === "على المستلم") {
        receiverDeliveryShare = deliveryCost;
        totalToCollect += deliveryCost;
      }

      var newTrackId = generateTrackingNumber();
      var orderPin = Math.floor(100000 + Math.random() * 900000).toString();
      var dateAdded = new Date();
      var dateString = Utilities.formatDate(dateAdded, Session.getScriptTimeZone(), "dd/MM/yyyy");

      // التعديل: إنشاء بوليصة PDF للطلب المجمع
      var barcodeUrl = "https://quickchart.io/barcode?type=code128&text=" + newTrackId + "&height=60&includeText=true";
      var barcodeBlob = UrlFetchApp.fetch(barcodeUrl).getBlob();
      var base64Barcode = Utilities.base64Encode(barcodeBlob.getBytes());
      var barcodeImgSrc = "data:image/png;base64," + base64Barcode;

      var htmlContent = `
      <!DOCTYPE html>
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: 'Tahoma', sans-serif; color: #333; line-height: 1.6; padding: 20px; }
          .header { text-align: center; border-bottom: 2px solid #2c3e50; padding-bottom: 15px; margin-bottom: 20px; }
          .header h1 { margin: 0; color: #2c3e50; font-size: 28px; }
          .barcode { margin-top: 15px; width: 250px; height: 60px; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
          .info-table td { width: 50%; padding: 15px; vertical-align: top; border: 1px solid #bdc3c7; background-color: #fafafa; }
          .financial-table th { background-color: #ecf0f1; padding: 10px; text-align: right; border: 1px solid #bdc3c7; }
          .financial-table td { padding: 10px; border: 1px solid #bdc3c7; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Dropex</h1>
          <p>تاريخ الإنشاء: ${dateString}</p>
          <img src="${barcodeImgSrc}" class="barcode" alt="Barcode">
        </div>
        <table class="info-table">
          <tr>
            <td>
              <h3>بيانات المستلم (إلى)</h3>
              <p>الاسم: ${recName}</p>
              <p>الهاتف: <span dir="ltr">${recPhone}</span></p>
              <p>العنوان: ${recAddress} - ${recArea}</p>
            </td>
            <td>
              <h3>بيانات المرسل (من)</h3>
              <p>الاسم: ${userData.name}</p>
            </td>
          </tr>
        </table>
        <table class="financial-table">
          <tr><th colspan="2">التفاصيل المالية</th></tr>
          <tr><td>سعر المنتج</td><td>${productPrice} ج.م</td></tr>
          <tr><td>رسوم التوصيل</td><td>${receiverDeliveryShare} ج.م</td></tr>
          <tr><td>الإجمالي</td><td><strong>${totalToCollect} ج.م</strong></td></tr>
        </table>
      </body>
      </html>
      `;

      var htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, "temp.html");
      var pdfBlob = htmlBlob.getAs(MimeType.PDF);
      pdfBlob.setName("Waybill_" + newTrackId + ".pdf");
      var pdfFolder = getSystemFolder("PDF");
      var pdfFile = pdfFolder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var waybillUrl = pdfFile.getUrl();

      // التعديل: 36 عمود لحفظ البيانات التقنية والهواتف والأسعار المزدوجة
      var rowData = new Array(36).fill("");
      rowData[0] = newTrackId; rowData[1] = String(userData.email).trim().toLowerCase();
      rowData[2] = String(userData.name).trim(); rowData[3] = recName;
      rowData[4] = "تم الإنشاء"; rowData[5] = dateAdded;
      rowData[6] = productPrice; rowData[7] = deliveryCost; // السعر الأصلي (دروبكس)
      rowData[8] = deliveryPaidBy; rowData[10] = waybillUrl;
      rowData[11] = formatPhoneForSheet(userData.phone);
      rowData[12] = userData.address; rowData[13] = userData.area;
      rowData[14] = recEmail; rowData[15] = formatPhoneForSheet(recPhone);
      rowData[16] = recAddress; rowData[17] = recArea;
      rowData[26] = orderPin; rowData[27] = totalToCollect;
      rowData[28] = ""; // حالة التصفية
      
      rowData[31] = ordersArray[i].inspection || "لا"; // AF
      rowData[32] = ordersArray[i].payMethod || "COD"; // AG
      rowData[33] = ""; // AH (هاتف إضافي 1)
      rowData[34] = ""; // AI (هاتف إضافي 2)
      rowData[35] = deliveryCost; // AJ (سعر البوليصة - افتراضياً نفس سعر دروبكس في الرفع المجمع)

      sheet.appendRow(rowData);
      addedCount++;
    }
    return { success: true, count: addedCount };
  } catch (e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

// ==========================================
// دالة تصفية حسابات التجار (للوحة الإدارة)
// ==========================================
function settleMerchantAccountByName(merchantName, settleType) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();
    var searchName = String(merchantName).trim().toLowerCase();
    var summary = {
      delivered: { count: 0, amount: 0 },
      returned: { count: 0, amount: 0 },
      canceledWithPickup: { count: 0, amount: 0 },
      totalAmount: 0,
      totalCount: 0
    };

    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][2]).trim().toLowerCase();
      var status = data[i][4];
      var isSettled = (String(data[i][28]).trim() === "تمت التصفية" || data[i][28] === true);
      
      // تحديد ما إذا كان الطلب قابلاً للتصفية أصلاً
      var canSettle = (status === "تم التوصيل" || status === "مرتجع" || (status === "ملغي" && parseFloat(data[i][9]) > 0));
      
      // التصفية بناءً على النوع المختار (الكل، واصل، مرتجع، ملغي)
      var matchesType = false;
      if (settleType === 'ALL') matchesType = true;
      else if (settleType === 'DELIVERED' && status === "تم التوصيل") matchesType = true;
      else if (settleType === 'RETURNED' && status === "مرتجع") matchesType = true;
      else if (settleType === 'CANCELED' && status === "ملغي") matchesType = true;

      if (rowName === searchName && canSettle && !isSettled && matchesType) {
        var productPrice = parseFloat(data[i][6]) || 0;
        var deliveryCost = parseFloat(data[i][7]) || 0;
        var paidBy = String(data[i][8]).trim();
        var pickupPrice = parseFloat(data[i][9]) || 0;
        var merchantNet = 0;

        if (status === "تم التوصيل") {
          var systemFee = parseFloat(data[i][7]) || 0; // H
          var merchantFee = parseFloat(data[i][35]) || systemFee; // AJ
          var extraProfit = (paidBy === "على المستلم") ? Math.max(0, merchantFee - systemFee) : 0;
          
          merchantNet = (paidBy === "على المرسل") ? (productPrice - systemFee - pickupPrice) : (productPrice + extraProfit - pickupPrice);
          summary.delivered.count++;
          summary.delivered.amount += merchantNet;
        }
        else if (status === "مرتجع") {
          var systemFee = parseFloat(data[i][7]) || 0;
          merchantNet = (paidBy === "على المرسل") ? (0 - systemFee - pickupPrice) : (0 - pickupPrice);
          summary.returned.count++;
          summary.returned.amount += merchantNet;
        }
        else if (status === "ملغي") {
          merchantNet = 0 - pickupPrice;
          summary.canceledWithPickup.count++;
          summary.canceledWithPickup.amount += merchantNet;
        }

        sheet.getRange(i + 1, 29).setValue("تمت التصفية");
        summary.totalAmount += merchantNet;
        summary.totalCount++;
      }
    }
    return { success: true, summary: summary };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// دالة إضافة تاجر (عميل) جديد من لوحة الإدارة
// ==========================================
function createNewUser(userData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");

    // التأكد من وجود شيت Users
    if (!sheet) return { success: false, error: "شيت Users غير موجود. يرجى إنشاؤه أولاً وتسميته 'Users'." };

    var data = sheet.getDataRange().getValues();
    var newEmail = String(userData.email).trim().toLowerCase();

    // التحقق من عدم وجود الإيميل مسبقاً
    for (var i = 1; i < data.length; i++) {
      var existingEmail = String(data[i][1]).trim().toLowerCase();
      if (existingEmail === newEmail) {
        return { success: false, error: "هذا البريد الإلكتروني مسجل لتاجر آخر بالفعل." };
      }
    }

    // بناء صف البيانات بنفس الترتيب الذي تقرأه دالة userLogin
    // [0]:الاسم, [1]:الإيميل, [2]:الباسورد, [3]:الهاتف, [4]:العنوان, [5]:المحافظة, [6]:نوع الخدمة, [7]:عنوان إضافي (H), [8]:رقم إضافي 1 (I), [9]:رقم إضافي 2 (J)
    var rowData = [
      String(userData.name).trim(),
      newEmail,
      String(userData.password).trim(),
      formatPhoneForSheet(userData.phone),
      String(userData.address).trim(),
      String(userData.area).trim(),
      userData.serviceType ? String(userData.serviceType).trim() : "الشحن فقط",
      String(userData.address2 || "").trim(),
      formatPhoneForSheet(userData.phone2 || ""),
      formatPhoneForSheet(userData.phone3 || ""),
      String(userData.businessType || "").trim()
    ];

    sheet.appendRow(rowData);
    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// التحقق من البريد وإرسال الـ OTP
// ==========================================
function sendVerificationEmail(userData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) return { success: false, error: "شيت Users غير موجود." };

  var data = sheet.getDataRange().getValues();
  var newEmail = String(userData.email).trim().toLowerCase();

  // التحقق من أن الحساب غير مسجل بالفعل
  for (var i = 1; i < data.length; i++) {
    var existingEmail = String(data[i][1]).trim().toLowerCase();
    if (existingEmail === newEmail) {
      return { success: false, error: "هذا البريد الإلكتروني مسجل لتاجر آخر بالفعل." };
    }
  }

  // توليد رمز OTP
  var otp = Math.floor(100000 + Math.random() * 900000).toString();

  // حفظ ה־OTP في الكاش لمدة 10 دقائق (600 ثانية)
  var cache = CacheService.getScriptCache();
  cache.put("OTP_" + newEmail, otp, 600);

  // إرسال الإيميل
  var subject = "رمز التحقق من حسابك - Dropex";
  var htmlBody = `
    <div dir="rtl" style="font-family: Arial, sans-serif; padding: 20px; color: #191c1e;">
      <h2 style="color: #0a1e4d;">مرحباً ${userData.name}،</h2>
      <p>شكراً لاختيارك شركة Dropex. لإكمال تسجيل حسابك، يرجى استخدام رمز التحقق التالي:</p>
      <div style="background-color: #f7f9fb; padding: 15px; border-radius: 8px; text-align: center; margin: 20px 0;">
        <span style="font-size: 24px; font-weight: bold; letter-spacing: 5px; color: #ff6b35;">${otp}</span>
      </div>
      <p>هذا الرمز صالح لمدة 10 دقائق فقط. يرجى عدم مشاركة هذا الرمز مع أي شخص.</p>
      <br>
      <p>مع تحياتنا،<br>فريق Dropex</p>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: newEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    return { success: true };
  } catch (e) {
    return { success: false, error: "حدث خطأ أثناء إرسال البريد: " + e.toString() };
  }
}

// ==========================================
// التحقق من ה-OTP وإنشاء الحساب
// ==========================================
function verifyAndCreateUser(userData, otpCode) {
  var newEmail = String(userData.email).trim().toLowerCase();
  var cache = CacheService.getScriptCache();
  var savedOtp = cache.get("OTP_" + newEmail);

  if (!savedOtp) {
    return { success: false, error: "انتهت صلاحية الرمز، يرجى إعادة المحاولة." };
  }

  if (String(savedOtp).trim() !== String(otpCode).trim()) {
    return { success: false, error: "رمز التحقق غير صحيح، حاول مرة أخرى." };
  }

  // الرمز صحيح، نقوم بإنشاء الحساب
  cache.remove("OTP_" + newEmail); // تنظيف الكاش
  return createNewUser(userData);
}

// ==========================================
// طلب توظيف المندوبين
// ==========================================
function submitEmploymentApplication(jobData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Employment");
    if (!sheet) {
      sheet = spreadsheet.insertSheet("Employment");
      sheet.appendRow(["التاريخ", "الاسم", "رقم الهاتف", "السن", "المحافظة", "المدينة", "نوع المركبة", "الخبرة السابقة"]);
    }

    var rowData = [
      new Date(),
      String(jobData.name).trim(),
      formatPhoneForSheet(jobData.phone),
      String(jobData.age).trim(),
      String(jobData.governorate).trim(),
      String(jobData.city).trim(),
      String(jobData.vehicle).trim(),
      String(jobData.experience).trim()
    ];

    sheet.appendRow(rowData);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// end of file