// SAUSLYS ADMIN — one script for the RESPONSES sheet.
// Menu: SAUSLYS → 1 Import LISTS v2 · 2 Sync forms from LISTS · 3 Rebuild LINES · 4 Set up COUNT tab · Health check.
// Runs as sauslysgroup@gmail.com (owner of the forms, editor of the sheet). Safe to run repeatedly; nothing is deleted.
var SS_ID = '1molkQTc36y9kR-QYwW4FoCfSUqF8Yf1anMDSoCQDZi0';
var LISTS_URL = 'https://raw.githubusercontent.com/saadtabani09/sauslys-ops/forms-lists-v2/forms/lists_v2.json';
var FORMS = { P1:'1vNxPHqv6qa9trHHTW0QsB57g0BgG7vj9NNwAJb-1QCI', G2:'1QIJRzb37e4RvPecF5NLPlrDW8k3gTJ4dFvv1rXoQYu4', I1:'14pja6PgIb3HgyeNAAsGH1R5txLP2o7WFkodnOZeY9H8', B2:'1tyfcSg7xfPsoDV8CmYF7Q4ECQc6C8caNcJAPGqNIEHY', S1:'18_B_TPC9e0IGtIUe4S-LW5R1m4X8EATI4DReuLXugFQ', S1P:'1gxiKKNMuqx7P9gTms4ceuuhwlTq05kW6Ot5OwKupq74', R1:'1JSEhL2Kui1B8m7M8PnL1phk_jtwaCEEOL5By8KfDdjg', D1:'1QXnEzinNOsFrVUAf9Ro5TkX9T--lM7R6HWm9R9j590g', G1:'1iQnhGs5PC_uqvOPBM3SrVCjtejr8lJMteHzixLjo6Uk' };
var TAB = { P1:'P1_PR', G2:'G2_GRN', I1:'I1_ISSUE', B2:'B2_BATCH', S1:'S1_COUNT_WASTE', S1P:'S1P_PRODUCT', R1:'R1_REQUEST', D1:'D1_DISPATCH', G1:'G1_RECEIVE' };
// Bangla words are written as \u escapes so the file survives every editor and paste.
var BN = { item:'উপকরণ', product:'পণ্য', itemWord:'আইটেম', qty:'পরিমাণ', received:'পেয়েছেন', other:'অন্য আইটেম (তালিকায় নেই)' };
var ITEM_WORD = { P1:BN.item, G2:BN.item, I1:BN.item, S1:BN.itemWord + ' বা ' + BN.product, B2:BN.product, S1P:BN.product, R1:BN.product, D1:BN.product, G1:BN.product + ' (কম/বেশি হলে)' };
var QTY_WORD = { G1: BN.received + ' / Received qty' };
var BUDGET = 2330; // total dropdown choices per form known to work (G2 v1 = 2,327); 8 x 406 = 3,248 failed

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SAUSLYS')
    .addItem('1. Import LISTS v2 from GitHub', 'importListsV2')
    .addItem('2. Sync forms from LISTS', 'syncForms')
    .addItem('3. Rebuild LINES formula', 'rebuildLines')
    .addItem('4. Set up COUNT tab', 'setupCountTab')
    .addItem('Health check', 'healthCheck').addToUi();
}
function ss_() { return SpreadsheetApp.openById(SS_ID); }
function fetchLists_() { return JSON.parse(UrlFetchApp.fetch(LISTS_URL, { muteHttpExceptions: false }).getContentText()); }
function col_(sh, values, c, header) { // write one column (1-based c) with header, clearing old content below it
  var n = Math.max(values.length + 1, sh.getLastRow());
  sh.getRange(1, c, n, 1).clearContent();
  sh.getRange(1, c, 1, 1).setValue(header);
  if (values.length) sh.getRange(2, c, values.length, 1).setValues(values.map(function (v) { return [v]; }));
}
function tab_(ss, name, rows, cols) { // write an object array to a tab (create if missing)
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  var data = [cols].concat(rows.map(function (r) { return cols.map(function (c) { var v = r[c]; if (v === true) return 'yes'; if (v === false || v === undefined || v === null) return ''; return Array.isArray(v) ? v.join(', ') : v; }); }));
  sh.getRange(1, 1, data.length, cols.length).setValues(data); sh.setFrozenRows(1);
  return sh;
}
// 1. Import: LISTS tab columns A-L keep the v1 layout (STOCK reads A, forms read H-L); M-O are new; full masters go to LISTS_* tabs.
function importListsV2() {
  var L = fetchLists_(); var ss = ss_(); var sh = ss.getSheetByName('LISTS');
  var items = L.ITEM.map(function (e) { return e.name; }), prods = L.PRODUCT.map(function (e) { return e.name; }), inter = L.INTERMEDIATE.map(function (e) { return e.name; });
  var staffAll = L.STAFF.filter(function (s) { return String(s.status).indexOf('Resign') < 0; }).map(function (s) { return s.name; });
  col_(sh, items, 1, 'ITEM (unit)'); col_(sh, prods, 2, 'PRODUCT (unit)'); col_(sh, inter, 3, 'INTERMEDIATE'); col_(sh, L.SUPPLIER_FORMS, 4, 'SUPPLIER');
  col_(sh, staffAll, 5, 'STAFF'); col_(sh, L.STATIC.OUTLET, 6, 'OUTLET'); col_(sh, L.STATIC.DEPT, 7, 'DEPT'); col_(sh, L.ITEM_FORMS, 8, 'ITEM for forms (stock only)');
  col_(sh, L.SUPPLIER_FORMS, 9, 'SUPPLIER for forms (cleaned)'); col_(sh, L.PRODUCT_FORMS, 10, 'PRODUCT for outlets'); col_(sh, L.B2_LIST, 11, 'PRODUCT or INTERMEDIATE for B2'); col_(sh, inter, 12, 'INTERMEDIATE for B2 used');
  col_(sh, L.I1_LIST, 13, 'I1 production ingredients'); col_(sh, L.OUTLET_LIST, 14, 'OUTLET list (products + supplies)'); col_(sh, L.OUTLET_SUPPLY, 15, 'OUTLET SUPPLY items');
  tab_(ss, 'LISTS_ITEM', L.ITEM, ['name','unit','bangla','type','category','subcategory','active_2026','outlet_supply','in_forms','source','confidence','note']);
  tab_(ss, 'LISTS_PRODUCT', L.PRODUCT, ['name','bangla','dept','category','status','sell_unit','pcs_per_pack','pack_g','pos_label','sku','outlet_kitchen','in_forms','source','confidence','note']);
  tab_(ss, 'LISTS_SUPPLIER', L.SUPPLIER, ['name','source','confidence','note']);
  tab_(ss, 'LISTS_STAFF', L.STAFF, ['name','department','location','role','dept_head_of','status','source']);
  tab_(ss, 'ALIAS', L.ALIAS, ['as_recorded','canonical','list','source','method']);
  tab_(ss, 'POS_SKU', L.POS_SKU, ['sku','pos_name','canonical','dept','unit','price','governed_id']);
  tab_(ss, 'DECISION_LOG', L.DECISION_LOG, ['area','name','from','to','why','confidence','source']);
  tab_(ss, 'FORM_STAFF', Object.keys(L.FORM_STAFF).map(function (k) { return { field: k, roles: L.FORM_STAFF[k] }; }), ['field','roles']);
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
  var msg = 'LISTS v2 imported: items ' + items.length + ', products ' + prods.length + ', suppliers ' + L.SUPPLIER_FORMS.length + ', staff ' + staffAll.length;
  Logger.log(msg); return msg;
}
function staffFor_(L, field) {
  var roles = L.FORM_STAFF[field] || ['*']; var all = roles.indexOf('*') >= 0;
  return L.STAFF.filter(function (s) { return String(s.status).indexOf('Resign') < 0 && (all || roles.indexOf(s.role) >= 0); }).map(function (s) { return s.name; });
}
function listFor_(L, code) { return L[L.FORM[code].items]; }
function setChoices_(item, values, out, label) {
  var li = item.asListItem(); var cur = li.getChoices().map(function (c) { return c.getValue(); });
  if (cur.join('') === values.join('')) return;
  li.setChoiceValues(values); out.push(label + ': ' + values.length + ' choices');
}
// Question-title patterns (match by the English part so the Bangla spelling never matters)
var RE = { line: /^(\d+)\.\s/, supplier: /Supplier/i, dept: /\/\s*Dept/i, outlet: /Outlet/i, where: /Where/i, inter: /intermediate/i, issued: /Issued by/i, receivedBy: /Received by/i, who: /Who|Made by|Sent by|Received by|Requested by/i, notStaff: /What|Reason|Purpose|All good|Invoice|Note/i, other: /Other item/i };
// 2. Sync every form: staff dropdowns by role, dept/outlet/supplier lists, every item line from the right list, add missing item lines + an Other-item text.
function syncForms() {
  var L = fetchLists_(); var out = [];
  Object.keys(FORMS).forEach(function (code) {
    var form = FormApp.openById(FORMS[code]); var cfg = L.FORM[code]; var items = listFor_(L, code); var wantLines = cfg.lines;
    var suppliers = code === 'G2' ? (L.SUPPLIER_ACTIVE || L.SUPPLIER_FORMS) : L.SUPPLIER_FORMS;
    var lineItems = []; var otherTotal = 0;
    form.getItems().forEach(function (it) {
      if (it.getType() !== FormApp.ItemType.LIST) return;
      var t = it.getTitle(); var m = t.match(RE.line);
      if (m) { lineItems.push({ n: +m[1], it: it }); return; }
      var vals = null, label = '';
      if (RE.supplier.test(t)) { vals = suppliers; label = 'supplier'; }
      else if (RE.dept.test(t)) { vals = L.STATIC.DEPT; label = 'dept'; }
      else if (RE.where.test(t)) { vals = L.STATIC.S1_WHERE; label = 'where'; }
      else if (RE.outlet.test(t)) { vals = L.STATIC.OUTLET; label = 'outlet'; }
      else if (RE.inter.test(t)) { vals = L.INTERMEDIATE.map(function (e) { return e.name; }); label = 'intermediate'; }
      else if (!RE.notStaff.test(t)) {
        var f = null;
        if (code === 'I1' && RE.issued.test(t)) f = cfg.issued_by; else if (code === 'I1' && RE.receivedBy.test(t)) f = cfg.received_by; else if (RE.who.test(t)) f = cfg.staff;
        if (f) { vals = staffFor_(L, f); label = f; }
      }
      if (vals) { setChoices_(it, vals, out, code + ' ' + label); otherTotal += vals.length; }
      else otherTotal += it.asListItem().getChoices().length;
    });
    lineItems.sort(function (a, b) { return a.n - b.n; });
    var budget = BUDGET - otherTotal; var used = 0; var seenLine = 0;
    lineItems.forEach(function (li) {
      seenLine = Math.max(seenLine, li.n); var cur = li.it.asListItem().getChoices().length;
      if (used + items.length <= budget) { setChoices_(li.it, items, out, code + ' line ' + li.n); used += items.length; }
      else { used += cur; out.push(code + ' line ' + li.n + ': kept old choices (' + cur + ') - cap'); }
    });
    for (var n = seenLine + 1; n <= wantLines; n++) {
      if (used + items.length > budget) { out.push(code + ': stopped adding at line ' + (n - 1) + ' (cap ' + BUDGET + ')'); break; }
      form.addListItem().setTitle(n + '. ' + ITEM_WORD[code] + ' / ' + BN.itemWord + ' ' + n).setChoiceValues(items).setRequired(false);
      form.addTextItem().setTitle(n + '. ' + (QTY_WORD[code] || (BN.qty + ' / Qty')) + ' ' + n).setRequired(false);
      used += items.length; out.push(code + ': added line ' + n);
    }
    var hasOther = form.getItems().some(function (it) { return RE.other.test(it.getTitle()); });
    if (L.FORM.other_item_text && !hasOther) { form.addParagraphTextItem().setTitle(BN.other + ' / Other item not in list: name + qty'); out.push(code + ': added Other-item text'); }
    out.push(code + ': total choices ' + (otherTotal + used) + ' / ' + BUDGET);
  });
  var msg = out.join('\n') || 'nothing to change'; Logger.log(msg); return msg;
}
// 3. Rebuild LINES!A2 from the linked tabs' header rows: one block per item line, 11 columns (col K = Reason). STOCK/VARIANCE/PR_CHECK keep using A-J.
function colL_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function q_(t) { return "'" + t + "'"; }
function rebuildLines() {
  var ss = ss_(); var blocks = [];
  function block(tab, code, whoC, locC, extraC, itemC, qtyC, rateC, reasonC, dateExpr) {
    var T = q_(tab) + '!'; var A = T + 'A2:A';
    var rate = rateC ? T + rateC + '2:' + rateC : 'IF(LEN(' + A + '),"","")'; var reason = reasonC ? T + reasonC + '2:' + reasonC : 'IF(LEN(' + A + '),"","")';
    return '{' + A + ',IF(LEN(' + A + '),"' + code + '",""),ROW(' + A + '),' + T + whoC + '2:' + whoC + ',' + T + locC + '2:' + locC + ',' + T + extraC + '2:' + extraC + ',' + T + itemC + '2:' + itemC + ',' + T + qtyC + '2:' + qtyC + ',' + rate + ',' + (dateExpr || 'INT(' + A + ')') + ',' + reason + '}';
  }
  Object.keys(TAB).forEach(function (code) {
    var sh = ss.getSheetByName(TAB[code]); if (!sh) return; var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var find = function (re) { for (var i = 0; i < hdr.length; i++) if (re.test(hdr[i])) return colL_(i + 1); return null; };
    var who = find(code === 'I1' ? /Issued by/i : code === 'G2' ? /Received by/i : code === 'B2' ? /Made by/i : code === 'D1' ? /Sent by/i : code === 'P1' ? /Who|Requested by/i : /\/ Who$/i) || 'B';
    var loc = find(code === 'P1' ? /Purpose/i : code === 'G2' ? /Supplier/i : (code === 'I1' || code === 'B2') ? /Dept/i : (code === 'S1' || code === 'S1P') ? /Where/i : /Outlet/i) || 'C';
    var extra = find(code === 'G2' ? /Invoice/i : code === 'I1' ? /Received by/i : (code === 'S1' || code === 'S1P') ? /What/i : code === 'G1' ? /All good/i : code === 'P1' ? /Purpose/i : (code === 'B2') ? /Dept/i : /Outlet/i) || loc;
    var reason = (code === 'S1' || code === 'S1P') ? find(/Reason/i) : null; var dateExpr = code === 'R1' ? 'INT(' + q_(TAB[code]) + '!A2:A+0.75)' : null;
    for (var i = 0; i < hdr.length; i++) {
      var m = hdr[i].match(/^(\d+)\.\s[\s\S]*?(\d+)\s*$/); if (!m || /Qty|Received qty/i.test(hdr[i])) continue; var n = m[1];
      var qi = -1; for (var j = 0; j < hdr.length; j++) { if (new RegExp('^' + n + '\\.\\s').test(hdr[j]) && /Qty|Received qty/i.test(hdr[j])) { qi = j; break; } }
      if (qi < 0) continue;
      var rateC = null; if (code === 'G2') { var rc = 15 + (+n); if (hdr[rc - 1] && /Rate/i.test(hdr[rc - 1])) rateC = colL_(rc); }
      blocks.push(block(TAB[code], code, who, loc, extra, colL_(i + 1), colL_(qi + 1), rateC, reason, dateExpr));
    }
    if (code === 'B2') { var ui = find(/Used intermediate/i), uq = find(/Used qty/i); if (ui && uq) blocks.push(block(TAB[code], 'B2-USED', who, loc, extra, ui, uq, null, null, null)); }
  });
  if (ss.getSheetByName('COUNT')) blocks.push("{'COUNT'!A2:A,IF(LEN('COUNT'!A2:A),\"S1\",\"\"),ROW('COUNT'!A2:A),'COUNT'!D2:D,IF(LEN('COUNT'!A2:A),\"Factory Store\",\"\"),IF(LEN('COUNT'!A2:A),\"COUNT (stock now)\",\"\"),'COUNT'!B2:B,'COUNT'!C2:C,IF(LEN('COUNT'!A2:A),\"\",\"\"),INT('COUNT'!A2:A),'COUNT'!E2:E}");
  var f = '=IFERROR(ARRAYFORMULA(QUERY({' + blocks.join(';') + '},"select * where Col7 <> \'\' order by Col1 desc",0)),"no entries yet")';
  var lines = ss.getSheetByName('LINES'); lines.getRange('K1').setValue('Reason'); lines.getRange('A2').setFormula(f); SpreadsheetApp.flush(); Utilities.sleep(3000);
  var v = lines.getRange('A2').getDisplayValue(); var msg = 'LINES rebuilt: ' + blocks.length + ' blocks, A2 shows: ' + v; Logger.log(msg); return msg;
}
// 4. COUNT tab: the weekly full store count is typed straight into the sheet (400 items do not fit a form). Columns: Date | Item | Qty | Counted by | Note.
function setupCountTab() {
  var ss = ss_(); var sh = ss.getSheetByName('COUNT'); if (sh) return 'COUNT tab exists';
  sh = ss.insertSheet('COUNT'); sh.getRange('A1:E1').setValues([['Date (type today)', 'Item (pick from list)', 'Qty (number)', 'Counted by', 'Note']]).setFontWeight('bold'); sh.setFrozenRows(1);
  var lists = ss.getSheetByName('LISTS');
  var itemRule = SpreadsheetApp.newDataValidation().requireValueInRange(lists.getRange('A2:A700'), true).setAllowInvalid(false).build();
  var staffRule = SpreadsheetApp.newDataValidation().requireValueInRange(lists.getRange('E2:E200'), true).setAllowInvalid(true).build();
  sh.getRange('B2:B2000').setDataValidation(itemRule); sh.getRange('D2:D2000').setDataValidation(staffRule); sh.getRange('A2:A2000').setNumberFormat('yyyy-mm-dd'); sh.getRange('C2:C2000').setNumberFormat('0.###');
  return 'COUNT tab created; Rebuild LINES makes counts feed STOCK';
}
function healthCheck() {
  var ss = ss_(); var out = [];
  Object.keys(TAB).forEach(function (code) { var sh = ss.getSheetByName(TAB[code]); out.push(TAB[code] + ': ' + (sh ? (sh.getFormUrl() ? 'linked' : 'NOT LINKED') + ', cols ' + sh.getLastColumn() + ', rows ' + Math.max(0, sh.getLastRow() - 1) : 'MISSING')); });
  ['LINES', 'VARIANCE', 'STOCK', 'PR_CHECK'].forEach(function (n) { var s = ss.getSheetByName(n); out.push(n + '!A2: ' + (s ? s.getRange('A2').getDisplayValue() : 'MISSING')); });
  out.push('LISTS first tab: ' + (ss.getSheets()[0].getName() === 'LISTS')); var msg = out.join('\n'); Logger.log(msg); return msg;
}
// One-shot for the first run (GPT / owner): everything in order, one log.
function runAll() { var a = importListsV2(); var b = syncForms(); var c = setupCountTab(); var d = rebuildLines(); var e = healthCheck(); var msg = [a, b, c, d, e].join('\n----\n'); Logger.log(msg); return msg; }
