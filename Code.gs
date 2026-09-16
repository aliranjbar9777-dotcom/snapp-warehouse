/**
 * ============================================================
 *  داشبورد انبار اسنپ — بک‌اند Google Apps Script
 *  Google Sheets به عنوان تنها دیتابیس استفاده می‌شود.
 * ============================================================
 *  نحوه نصب:
 *  1) یک Google Sheet جدید بسازید.
 *  2) از منوی Extensions > Apps Script این فایل را جایگزین Code.gs کنید.
 *  3) تابع setupSheets را یک‌بار اجرا کنید (Run) تا شیت‌ها، هدرها و
 *     کاربر مدیر پیش‌فرض ساخته شوند.
 *  4) روی Deploy > New deployment بزنید، نوع را Web app انتخاب کنید،
 *     Execute as: Me ، Who has access: Anyone را انتخاب و Deploy کنید.
 *  5) آدرس (URL) دیپلوی را در صفحه‌ی داشبورد (index.html) وارد کنید.
 *
 *  کاربر مدیر پیش‌فرض:
 *     نام کاربری: admin
 *     رمز عبور:   admin123
 *  (حتماً بلافاصله بعد از اولین ورود، رمز را عوض کنید یا کاربر
 *   جدیدی با نقش «مدیر» بسازید و کاربر admin را حذف/غیرفعال کنید.)
 * ============================================================
 */

const SHEET_TRANSACTIONS = 'Transactions';
const SHEET_INVENTORY    = 'Inventory';
const SHEET_LISTS        = 'Lists';
const SHEET_USERS        = 'Users';

// رمز عبور پیش‌فرض admin123 به‌صورت هش SHA-256 (کلاینت هم با همین الگوریتم هش می‌کند)
const DEFAULT_ADMIN_HASH = '240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9';

// ---------- ورودی وب‌اپ ----------

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getData';
    if (action === 'getData') {
      return jsonResponse({ ok: true, data: getAllData() });
    }
    return jsonResponse({ ok: false, error: 'اکشن نامعتبر است' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // ورود: برای اینکه بعد از لاگین مجبور به یک درخواست جداگانه‌ی دیگر برای
    // خواندن داده‌ها نباشیم (که کندی محسوسی ایجاد می‌کرد)، همین یک درخواست
    // هم نتیجه‌ی ورود و هم کل داده‌های داشبورد را با هم برمی‌گرداند.
    if (action === 'login') {
      const result = login(body.payload);
      return jsonResponse({ ok: true, result: result, data: getAllData() });
    }

    // برای عملیات‌های نوشتنی، دیگر کل شیت‌ها را دوباره نمی‌خوانیم (که کند بود)؛
    // هر تابع فقط نتیجه‌ی لازم برای به‌روزرسانی محلی در مرورگر را برمی‌گرداند.
    let result;
    if (action === 'addTransaction') {
      result = addTransaction(body.payload);
    } else if (action === 'addListItem') {
      result = addListItem(body.payload);
    } else if (action === 'deleteListItem') {
      result = deleteListItem(body.payload);
    } else if (action === 'addUser') {
      result = addUser(body.payload);
    } else if (action === 'deleteUser') {
      result = deleteUser(body.payload);
    } else if (action === 'changePassword') {
      result = changePassword(body.payload);
    } else if (action === 'deleteTransaction') {
      result = deleteTransaction(body.payload);
    } else if (action === 'editInventory') {
      result = editInventory(body.payload);
    } else {
      throw new Error('اکشن نامعتبر است');
    }

    return jsonResponse({ ok: true, result: result });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- ساخت اولیه‌ی شیت‌ها ----------

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let tx = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!tx) tx = ss.insertSheet(SHEET_TRANSACTIONS);
  if (tx.getLastRow() === 0) {
    tx.appendRow(['شناسه', 'تاریخ و ساعت', 'شهر', 'دسته', 'نوع تجهیز', 'رنگ',
                  'نام شخص/راننده', 'ثبت‌کننده (کارشناس)', 'نوع عملیات', 'تعداد', 'موجودی پس از ثبت']);
    tx.setFrozenRows(1);
  }

  let inv = ss.getSheetByName(SHEET_INVENTORY);
  if (!inv) inv = ss.insertSheet(SHEET_INVENTORY);
  if (inv.getLastRow() === 0) {
    inv.appendRow(['شهر', 'دسته', 'نوع تجهیز', 'رنگ', 'موجودی فعلی']);
    inv.setFrozenRows(1);
  }

  let lists = ss.getSheetByName(SHEET_LISTS);
  if (!lists) lists = ss.insertSheet(SHEET_LISTS);
  if (lists.getLastRow() === 0) {
    lists.appendRow(['شهرها', 'انواع باکس', 'رنگ‌ها', 'سایر تجهیزات']);
    lists.setFrozenRows(1);

    const cities   = ['تهران', 'مشهد', 'اصفهان', 'شیراز', 'تبریز', 'کرج', 'اهواز', 'قم'];
    const boxTypes = ['باکس بزرگ', 'باکس متوسط', 'باکس کوچک'];
    const colors   = ['زرد', 'مشکی', 'سفید', 'قرمز'];
    const others   = ['کیف پول‌گردان', 'ترازو', 'چاپگر برچسب'];

    const maxLen = Math.max(cities.length, boxTypes.length, colors.length, others.length);
    for (let i = 0; i < maxLen; i++) {
      lists.appendRow([cities[i] || '', boxTypes[i] || '', colors[i] || '', others[i] || '']);
    }
  }

  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) users = ss.insertSheet(SHEET_USERS);
  if (users.getLastRow() === 0) {
    users.appendRow(['نام کاربری', 'رمز عبور (هش‌شده)', 'نام و نام خانوادگی', 'نقش', 'فعال', 'تاریخ ایجاد']);
    users.setFrozenRows(1);
    users.appendRow(['admin', DEFAULT_ADMIN_HASH, 'مدیر سیستم', 'مدیر', true,
                      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tehran', 'yyyy/MM/dd HH:mm')]);
  }
}

// ---------- خواندن همه‌ی داده‌ها ----------

function getAllData() {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // تراکنش‌ها
  const tx = ss.getSheetByName(SHEET_TRANSACTIONS);
  const txValues = tx.getDataRange().getValues();
  txValues.shift();
  const transactions = txValues
    .filter(row => row[0])
    .map(row => ({
      id: row[0], datetime: row[1], city: row[2], category: row[3],
      itemType: row[4], color: row[5], person: row[6], registeredBy: row[7],
      operation: row[8], quantity: Number(row[9]), balanceAfter: Number(row[10])
    }))
    .reverse();

  // موجودی
  const inv = ss.getSheetByName(SHEET_INVENTORY);
  const invValues = inv.getDataRange().getValues();
  invValues.shift();
  const inventory = invValues
    .filter(row => row[0])
    .map(row => ({
      city: row[0], category: row[1], itemType: row[2], color: row[3], stock: Number(row[4])
    }));

  // لیست‌های کشویی
  const lists = ss.getSheetByName(SHEET_LISTS);
  const listValues = lists.getDataRange().getValues();
  listValues.shift();
  const cities = [], boxTypes = [], colors = [], others = [];
  listValues.forEach(row => {
    if (row[0]) cities.push(row[0]);
    if (row[1]) boxTypes.push(row[1]);
    if (row[2]) colors.push(row[2]);
    if (row[3]) others.push(row[3]);
  });

  // کاربران (بدون رمز عبور، فقط برای پنل مدیریت)
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const userValues = usersSheet.getDataRange().getValues();
  userValues.shift();
  const users = userValues
    .filter(row => row[0])
    .map(row => ({
      username: row[0], fullName: row[2], role: row[3], active: row[4] === true || row[4] === 'true'
    }));

  return { transactions, inventory, lists: { cities, boxTypes, colors, others }, users };
}

// ---------- ورود کاربر ----------

function login(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = ss.getSheetByName(SHEET_USERS);

  const username = String(payload.username || '').trim();
  const passwordHash = String(payload.passwordHash || '').trim();
  if (!username || !passwordHash) throw new Error('نام کاربری و رمز عبور الزامی است');

  const data = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      const active = data[i][4] === true || data[i][4] === 'true';
      if (!active) throw new Error('این حساب کاربری غیرفعال شده است');
      if (data[i][1] !== passwordHash) throw new Error('نام کاربری یا رمز عبور اشتباه است');
      return { username: data[i][0], fullName: data[i][2], role: data[i][3] };
    }
  }
  throw new Error('نام کاربری یا رمز عبور اشتباه است');
}

// ---------- ثبت تراکنش (ورود/خروج) ----------

function addTransaction(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tx = ss.getSheetByName(SHEET_TRANSACTIONS);
  const inv = ss.getSheetByName(SHEET_INVENTORY);

  const city = String(payload.city || '').trim();
  const category = String(payload.category || 'باکس').trim(); // 'باکس' یا 'سایر تجهیزات'
  const itemType = String(payload.itemType || '').trim();
  const color = String(payload.color || '').trim();
  const person = String(payload.person || '').trim();
  const registeredBy = String(payload.registeredBy || '').trim();
  const operation = String(payload.operation || '').trim(); // 'IN' یا 'OUT'
  const quantity = Number(payload.quantity);

  if (!city || !itemType || !person) throw new Error('لطفاً همه‌ی فیلدهای الزامی را پر کنید');
  if (!quantity || quantity <= 0) throw new Error('تعداد باید عددی مثبت باشد');
  if (operation !== 'IN' && operation !== 'OUT') throw new Error('نوع عملیات نامعتبر است');

  // پیدا کردن یا ساختن ردیف موجودی مربوطه
  const invData = inv.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < invData.length; i++) {
    if (invData[i][0] === city && invData[i][1] === category &&
        invData[i][2] === itemType && invData[i][3] === color) {
      rowIndex = i + 1; // شماره ردیف در شیت (۱-پایه)
      break;
    }
  }

  let currentStock = 0;
  if (rowIndex === -1) {
    inv.appendRow([city, category, itemType, color, 0]);
    rowIndex = inv.getLastRow();
  } else {
    currentStock = Number(invData[rowIndex - 1][4]) || 0;
  }

  let newStock;
  if (operation === 'IN') {
    newStock = currentStock + quantity;
  } else {
    newStock = currentStock - quantity;
    if (newStock < 0) {
      throw new Error('موجودی کافی نیست. موجودی فعلی: ' + currentStock);
    }
  }

  inv.getRange(rowIndex, 5).setValue(newStock);

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Tehran', 'yyyy/MM/dd HH:mm');
  const id = 'TX-' + now.getTime();

  tx.appendRow([id, dateStr, city, category, itemType, color, person, registeredBy, operation, quantity, newStock]);

  return { id, newStock, dateStr };
}

// ---------- افزودن آیتم جدید به لیست‌های کشویی ----------

function addListItem(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lists = ss.getSheetByName(SHEET_LISTS);

  const type = payload.type;   // city | boxType | color | other
  const value = String(payload.value || '').trim();
  if (!value) throw new Error('مقدار نمی‌تواند خالی باشد');

  const colMap = { city: 1, boxType: 2, color: 3, other: 4 };
  const col = colMap[type];
  if (!col) throw new Error('نوع لیست نامعتبر است');

  const lastRow = Math.max(lists.getLastRow(), 1);
  const colValues = lists.getRange(1, col, lastRow).getValues().flat();

  if (colValues.includes(value)) return { message: 'این مورد قبلاً موجود است' };

  let targetRow = -1;
  for (let i = 1; i < colValues.length; i++) {
    if (!colValues[i]) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) targetRow = lastRow + 1;

  lists.getRange(targetRow, col).setValue(value);
  return { message: 'با موفقیت افزوده شد' };
}

// ---------- حذف یک آیتم از لیست‌های کشویی ----------

function deleteListItem(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lists = ss.getSheetByName(SHEET_LISTS);

  const type = payload.type;
  const value = String(payload.value || '').trim();
  const colMap = { city: 1, boxType: 2, color: 3, other: 4 };
  const col = colMap[type];
  if (!col) throw new Error('نوع لیست نامعتبر است');

  const lastRow = Math.max(lists.getLastRow(), 1);
  const colValues = lists.getRange(1, col, lastRow).getValues().flat();

  for (let i = 1; i < colValues.length; i++) {
    if (colValues[i] === value) {
      lists.getRange(i + 1, col).setValue('');
      return { message: 'حذف شد' };
    }
  }
  return { message: 'یافت نشد' };
}

// ---------- مدیریت کاربران (فقط برای نقش «مدیر» در فرانت‌اند کنترل می‌شود) ----------

function addUser(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = ss.getSheetByName(SHEET_USERS);

  const username = String(payload.username || '').trim();
  const passwordHash = String(payload.passwordHash || '').trim();
  const fullName = String(payload.fullName || '').trim();
  const role = String(payload.role || 'کارشناس').trim();

  if (!username || !passwordHash || !fullName) throw new Error('همه‌ی فیلدهای کاربر الزامی است');

  const data = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) throw new Error('این نام کاربری قبلاً ثبت شده است');
  }

  users.appendRow([username, passwordHash, fullName, role, true,
                    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tehran', 'yyyy/MM/dd HH:mm')]);
  return { message: 'کاربر جدید ساخته شد' };
}

function deleteUser(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = ss.getSheetByName(SHEET_USERS);

  const username = String(payload.username || '').trim();
  const data = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (data.length === 2) throw new Error('حداقل یک کاربر باید در سامانه باقی بماند');
      users.deleteRow(i + 1);
      return { message: 'کاربر حذف شد' };
    }
  }
  throw new Error('کاربر یافت نشد');
}

function changePassword(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = ss.getSheetByName(SHEET_USERS);

  const username = String(payload.username || '').trim();
  const oldPasswordHash = String(payload.oldPasswordHash || '').trim();
  const newPasswordHash = String(payload.newPasswordHash || '').trim();
  if (!newPasswordHash) throw new Error('رمز عبور جدید نمی‌تواند خالی باشد');

  const data = users.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (data[i][1] !== oldPasswordHash) throw new Error('رمز عبور فعلی اشتباه است');
      users.getRange(i + 1, 2).setValue(newPasswordHash);
      return { message: 'رمز عبور تغییر کرد' };
    }
  }
  throw new Error('کاربر یافت نشد');
}

// ---------- حذف یک تراکنش (فقط مدیر — کنترل نقش در فرانت‌اند انجام می‌شود) ----------
// موجودی فعلی آن آیتم بر اساس اثر همان تراکنش برگردانده می‌شود؛ خودِ تراکنش از
// تاریخچه پاک می‌شود. برای رکوردهای «اصلاح موجودی» (ADJUST) فقط ردیف حذف می‌شود
// چون خودشان یک تغییر مستقیم مقدار بودند، نه یک ورود/خروج واقعی.
function deleteTransaction(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tx = ss.getSheetByName(SHEET_TRANSACTIONS);
  const inv = ss.getSheetByName(SHEET_INVENTORY);

  const id = String(payload.id || '').trim();
  if (!id) throw new Error('شناسه‌ی تراکنش نامعتبر است');

  const data = tx.getDataRange().getValues();
  let rowIndex = -1, row = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { rowIndex = i + 1; row = data[i]; break; }
  }
  if (rowIndex === -1) throw new Error('تراکنش یافت نشد (شاید قبلاً حذف شده)');

  const city = row[2], category = row[3], itemType = row[4], color = row[5];
  const operation = row[8], quantity = Number(row[9]);

  if (operation === 'IN' || operation === 'OUT') {
    const invData = inv.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][0] === city && invData[i][1] === category && invData[i][2] === itemType && invData[i][3] === color) {
        let stock = Number(invData[i][4]) || 0;
        stock = (operation === 'IN') ? (stock - quantity) : (stock + quantity);
        if (stock < 0) stock = 0;
        inv.getRange(i + 1, 5).setValue(stock);
        break;
      }
    }
  }

  tx.deleteRow(rowIndex);
  return { message: 'تراکنش حذف و موجودی اصلاح شد' };
}

// ---------- ویرایش مستقیم موجودی یک آیتم (فقط مدیر) ----------
// برای مواردی مثل شمارش فیزیکی انبار که با عدد سیستم فرق دارد. یک ردیف
// «اصلاح موجودی» هم در تاریخچه ثبت می‌شود تا همه‌چیز قابل ردیابی بماند.
function editInventory(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(SHEET_INVENTORY);
  const tx = ss.getSheetByName(SHEET_TRANSACTIONS);

  const city = String(payload.city || '').trim();
  const category = String(payload.category || '').trim();
  const itemType = String(payload.itemType || '').trim();
  const color = String(payload.color || '').trim();
  const newStock = Number(payload.newStock);
  const adminUsername = String(payload.adminUsername || '').trim();

  if (!city || !itemType) throw new Error('اطلاعات آیتم ناقص است');
  if (isNaN(newStock) || newStock < 0) throw new Error('مقدار موجودی جدید نامعتبر است');

  const data = inv.getDataRange().getValues();
  let rowIndex = -1, oldStock = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === city && data[i][1] === category && data[i][2] === itemType && data[i][3] === color) {
      rowIndex = i + 1; oldStock = Number(data[i][4]) || 0; break;
    }
  }
  if (rowIndex === -1) {
    inv.appendRow([city, category, itemType, color, newStock]);
  } else {
    inv.getRange(rowIndex, 5).setValue(newStock);
  }

  const delta = newStock - oldStock;
  if (delta !== 0) {
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Tehran', 'yyyy/MM/dd HH:mm');
    const id = 'TX-' + now.getTime();
    tx.appendRow([id, dateStr, city, category, itemType, color, 'اصلاح موجودی توسط مدیر', adminUsername, 'ADJUST', Math.abs(delta), newStock]);
  }

  return { newStock: newStock };
}
