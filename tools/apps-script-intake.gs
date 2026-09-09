/**
 * SUCCESS CORE J LLC — бизнесийн оношилгооны хариу хүлээн авагч
 * scorej.biz/shalgalt/ хуудаснаас POST хүлээж авч:
 *   1) энэ Google Sheet-д мөр болгон нэмнэ
 *   2) дотоод мэдэгдлийг hello@scorej.biz руу илгээнэ
 *   3) БИЗНЕС ЭРХЛЭГЧ рүү нь дүгнэлт + үнийн саналыг илгээнэ
 *
 * Байршуулах: Deploy → New deployment → Web app
 *   Execute as:      Me
 *   Who has access:  Anyone
 * Кодыг өөрчилсний дараа Deploy → Manage deployments → ✏️ → Version: New version.
 */

var SHEET_NAME = 'Хариултууд';
var NOTIFY_TO  = 'hello@scorej.biz';
var COMPANY    = 'SUCCESS CORE J LLC';
var PHONE_1    = '+976 8889-7485';
var PHONE_2    = '+976 7220-8459';
var SITE       = 'www.scorej.biz';

var HEADERS = [
  'Огноо', 'Нэр', 'Утас', 'И-мэйл', 'Байгууллага', 'Албан тушаал',
  'Вэб', 'Facebook', 'Instagram',
  'Санал болгосон', 'Дараагийн алхам',
  'Цаг/сард', 'Хэмнэх цаг', 'Алдагдсан захиалга',
  'Бүх хариулт', 'Хуудас', 'Хариу илгээсэн эсэх'
];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return reply({ ok: false, error: 'empty body' });

    var d = JSON.parse(e.postData.contents);
    var c = d.contact || {};

    if (!String(c.name || '').trim() || !String(c.phone || '').trim()) {
      return reply({ ok: false, error: 'name and phone required' });
    }

    var answersText = (d.answers || []).map(function (a) {
      return '• ' + a.q + '\n  → ' + (a.a || '—');
    }).join('\n');

    // 1) үйлчлүүлэгч рүү дүгнэлт илгээх — хамгийн чухал нь тул эхэлж оролдоно
    var sent = '';
    if (String(c.email || '').indexOf('@') > 0) {
      try {
        sendReport_(c, d);
        sent = 'Тийм';
      } catch (mailErr) {
        sent = 'Үгүй — ' + String(mailErr);
      }
    } else {
      sent = 'И-мэйл өгөөгүй';
    }

    // 2) хүснэгтэд бичих
    sheet_().appendRow([
      new Date(),
      c.name || '', c.phone || '', c.email || '', c.company || '', c.role || '',
      c.website || '', c.facebook || '', c.instagram || '',
      d.recommended || '', d.second || '',
      d.monthlyHrs || '', d.saveHrs || '', d.lostOrders || '',
      answersText,
      d.page || '',
      sent
    ]);

    // 3) дотоод мэдэгдэл
    notify_(c, d, answersText, sent);
    return reply({ ok: true, reportSent: sent });

  } catch (err) {
    try {
      MailApp.sendEmail(NOTIFY_TO, 'scorej.biz оношилгоо — алдаа', String(err) + '\n\n' +
        (e && e.postData ? e.postData.contents : '(no body)'));
    } catch (ignored) {}
    return reply({ ok: false, error: String(err) });
  }
}

/** Хөтчөөс хаягийг шалгахад ажиллаж байгааг харуулна */
function doGet() {
  return reply({ ok: true, service: 'scorej diagnostic intake' });
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(15, 420);
  } else if (sh.getLastColumn() < HEADERS.length) {
    // багана нэмэгдсэн бол гарчгийг нөхөж бичнэ
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sh;
}

/* ============================================================
   ҮЙЛЧЛҮҮЛЭГЧ РҮҮ ИЛГЭЭХ ДҮГНЭЛТ
   Зураг, вэб фонт ашиглаагүй — бүх и-мэйл клиентэд ижил харагдана.
   ============================================================ */
function sendReport_(c, d) {
  var pkg = d.pkg || {};
  var name = String(c.name || '').split(' ')[0] || c.name || '';

  var subject = 'Танай бизнесийн оношилгооны дүгнэлт — ' +
                (pkg.name ? pkg.name + ' ' + (pkg.price || '') : COMPANY);

  MailApp.sendEmail({
    to: c.email,
    subject: subject,
    body: reportText_(c, d, name),
    htmlBody: reportHtml_(c, d, name),
    replyTo: NOTIFY_TO,
    name: COMPANY
  });
}

function reportHtml_(c, d, name) {
  var pkg = d.pkg || {}, sec = d.secondPkg;
  var GREEN = '#12805a', DEEP = '#0b5d41', INK = '#16211d', MUTED = '#5d716a', LINE = '#e2eae6';
  var F = 'font-family:Arial,Helvetica,sans-serif;';

  function h2(t) {
    return '<tr><td style="padding:26px 0 10px;"><div style="' + F +
      'font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' + GREEN +
      ';font-weight:bold;">' + esc_(t) + '</div>' +
      '<div style="height:2px;width:34px;background:' + GREEN + ';margin-top:6px;"></div></td></tr>';
  }
  function bullets(arr, mark, color) {
    if (!arr || !arr.length) return '';
    return '<tr><td style="padding:4px 0 0;"><table cellpadding="0" cellspacing="0" border="0" width="100%">' +
      arr.map(function (x) {
        return '<tr>' +
          '<td width="18" valign="top" style="' + F + 'font-size:14px;line-height:1.6;color:' +
            color + ';padding:3px 8px 3px 0;">' + mark + '</td>' +
          '<td style="' + F + 'font-size:14px;line-height:1.6;color:' + INK +
            ';padding:3px 0;">' + esc_(x) + '</td></tr>';
      }).join('') + '</table></td></tr>';
  }

  var figs = [];
  if (d.monthlyHrs) figs.push([d.monthlyHrs, 'цаг / сард мессежид']);
  if (d.saveHrs)    figs.push([d.saveHrs, 'цаг / сард хэмнэнэ']);
  if (d.lostOrders) figs.push([d.lostOrders, 'захиалга / сард алдаж байна']);

  var figRow = '';
  if (figs.length) {
    figRow = '<tr><td style="padding:14px 0 0;"><table cellpadding="0" cellspacing="0" border="0" width="100%">' +
      '<tr>' + figs.map(function (f) {
        return '<td width="' + Math.floor(100 / figs.length) + '%" valign="top" ' +
          'style="border:1px solid ' + LINE + ';padding:14px 12px;">' +
          '<div style="' + F + 'font-size:26px;font-weight:bold;color:' + DEEP + ';line-height:1;">' +
            esc_(f[0]) + '</div>' +
          '<div style="' + F + 'font-size:11px;letter-spacing:1px;text-transform:uppercase;color:' +
            MUTED + ';padding-top:6px;">' + esc_(f[1]) + '</div></td>';
      }).join('<td width="10"></td>') + '</tr></table></td></tr>';
  }

  var items = (pkg.items || []).map(function (x) {
    return '<tr><td width="16" valign="top" style="' + F + 'font-size:14px;color:' + GREEN +
      ';padding:3px 8px 3px 0;">•</td><td style="' + F + 'font-size:14px;line-height:1.55;color:' +
      INK + ';padding:3px 0;">' + esc_(x) + '</td></tr>';
  }).join('');

  var offer =
    '<tr><td style="padding:16px 0 0;">' +
      '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border:2px solid ' + GREEN + ';">' +
      '<tr><td style="padding:22px 22px 18px;">' +
        '<div style="' + F + 'font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' +
          GREEN + ';font-weight:bold;">' + esc_(pkg.tag || 'Санал болгож буй шийдэл') + '</div>' +
        '<div style="' + F + 'font-size:24px;font-weight:bold;color:' + INK + ';padding-top:8px;">' +
          esc_(pkg.name || '') + '</div>' +
        '<div style="' + F + 'font-size:32px;font-weight:bold;color:' + DEEP + ';padding-top:4px;line-height:1.1;">' +
          esc_(pkg.price || '') + '</div>' +
        '<div style="' + F + 'font-size:13px;color:' + MUTED + ';padding-top:6px;">Хүлээлгэн өгөх хугацаа: <b style="color:' +
          INK + ';">' + esc_(pkg.time || '') + '</b></div>' +
        (items ? '<div style="border-top:1px solid ' + LINE + ';margin-top:16px;padding-top:14px;">' +
          '<div style="' + F + 'font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' +
            MUTED + ';padding-bottom:8px;">Багцад юу багтах вэ</div>' +
          '<table cellpadding="0" cellspacing="0" border="0" width="100%">' + items + '</table></div>' : '') +
      '</td></tr></table></td></tr>';

  var terms =
    '<tr><td style="padding:16px 0 0;">' +
      '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid ' + LINE +
        ';background-color:#f4f9f7;"><tr><td style="padding:16px 18px;">' +
        row_('Нийт үнэ', pkg.price || '—', F, MUTED, INK) +
        row_('Хугацаа', pkg.time || '—', F, MUTED, INK) +
        row_('Төлбөр', '50% урьдчилгаа · 50% хүлээлгэн өгөхөд', F, MUTED, INK) +
        '<div style="' + F + 'font-size:12px;line-height:1.6;color:' + MUTED +
          ';padding-top:12px;border-top:1px solid ' + LINE + ';margin-top:12px;">' +
          'Хугацаа нь танаас материал (лого, зураг, текст, хандах эрх) бүрэн ирсэн өдрөөс эхэлнэ. ' +
          'Дээрх үнэ бол таны хариулт дээр тулгуурласан урьдчилсан тооцоо — үнэгүй уулзалтаар ажлын хүрээг ' +
          'тодруулж эцэслэнэ. Домэйн, хостингийн жилийн төлбөр, төлбөрийн системийн шимтгэл, ' +
          'сурталчилгааны төсөв багтаагүй.' +
        '</div>' +
      '</td></tr></table></td></tr>';

  var secBlock = sec ?
    '<tr><td style="padding:16px 0 0;"><table cellpadding="0" cellspacing="0" border="0" width="100%" ' +
      'style="border-left:3px solid ' + GREEN + ';"><tr><td style="padding:2px 0 2px 16px;">' +
      '<div style="' + F + 'font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' +
        MUTED + ';">Дараагийн үе шатанд</div>' +
      '<div style="' + F + 'font-size:15px;font-weight:bold;color:' + INK + ';padding-top:4px;">' +
        esc_(sec.name) + ' — ' + esc_(sec.price) + '</div>' +
      '<div style="' + F + 'font-size:13px;line-height:1.6;color:' + MUTED + ';padding-top:4px;">' +
        'Эхний шийдэл ажиллаж эхэлсний дараа нэмбэл зохимжтой.</div>' +
      '</td></tr></table></td></tr>' : '';

  return '' +
  '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f4f6f5;">' +
  '<tr><td align="center" style="padding:24px 12px;">' +
  '<table cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background-color:#ffffff;border:1px solid ' + LINE + ';">' +

    '<tr><td style="background-color:' + DEEP + ';padding:20px 26px;">' +
      '<div style="' + F + 'font-size:13px;font-weight:bold;letter-spacing:2px;color:#ffffff;">' + COMPANY + '</div>' +
      '<div style="' + F + 'font-size:12px;color:#a9d7c4;padding-top:3px;">Бизнесийн оношилгооны дүгнэлт</div>' +
    '</td></tr>' +

    '<tr><td style="padding:26px 26px 0;">' +
      '<table cellpadding="0" cellspacing="0" border="0" width="100%">' +

        '<tr><td style="' + F + 'font-size:15px;line-height:1.65;color:' + INK + ';">' +
          'Сайн байна уу, <b>' + esc_(name) + '</b>.<br>' +
          'Танай бөглөсөн 14 асуултын хариулт дээр үндэслэн дараах дүгнэлтийг гаргалаа.' +
        '</td></tr>' +

        h2('Танай бизнесээс илэрсэн зүйлс') +
        bullets(d.problems, '!', '#c2410c') +
        figRow +

        h2('Бид үүнийг ингэж шийдэж өгнө') +
        bullets(d.solutions, '→', GREEN) +

        (d.later && d.later.length ? h2('Энэ багцад багтахгүй — дараагийн үе шатанд') +
          bullets(d.later, '·', MUTED) : '') +

        h2('Санал болгож буй шийдэл') +
        offer +
        terms +
        secBlock +

        h2('Дараагийн алхам') +
        '<tr><td style="' + F + 'font-size:14px;line-height:1.7;color:' + INK + ';padding-top:4px;">' +
          'Бид ажлын 1 өдөрт багтаан <b>' + esc_(c.phone || '') + '</b> дугаараар холбогдоно. ' +
          'Эхний уулзалт үнэ төлбөргүй — танай үйл ажиллагааг судалж, ажлын хүрээ, хугацааг эцэслэнэ.<br><br>' +
          'Хүсвэл өөрөө шууд залгаарай: <b>' + PHONE_1 + '</b> · ' + PHONE_2 +
        '</td></tr>' +

        '<tr><td style="padding:26px 0 0;"><div style="border-top:1px solid ' + LINE + ';"></div></td></tr>' +

        '<tr><td style="padding:16px 0 26px;">' +
          '<div style="' + F + 'font-size:14px;font-weight:bold;color:' + INK + ';letter-spacing:0.5px;">' + COMPANY + '</div>' +
          '<div style="' + F + 'font-size:13px;color:' + MUTED + ';padding-top:4px;line-height:1.6;">' +
            PHONE_1 + ' | ' + PHONE_2 + '<br>' +
            '<a href="mailto:' + NOTIFY_TO + '" style="color:' + GREEN + ';text-decoration:none;">' + NOTIFY_TO + '</a> · ' +
            '<a href="https://' + SITE + '" style="color:' + GREEN + ';text-decoration:none;">' + SITE + '</a><br>' +
            'MN TOWER 1201, J.Sambuu St, 5th khoroo,<br>Chingeltei district, Ulaanbaatar 15141, Mongolia' +
          '</div>' +
        '</td></tr>' +

      '</table></td></tr>' +
  '</table>' +

  '<div style="' + F + 'font-size:11px;line-height:1.6;color:#8a9c95;padding:14px 12px 0;max-width:600px;">' +
    'Энэ захидлыг та scorej.biz/shalgalt/ дээр оношилгоо бөглөж, хариугаа авахыг зөвшөөрсөн тул илгээв. ' +
    'Таны мэдээллийг зөвхөн танайхтай холбоо барихад ашиглана.' +
  '</div>' +

  '</td></tr></table>';
}

function row_(k, val, F, MUTED, INK) {
  return '<div style="padding:3px 0;"><span style="' + F + 'font-size:12px;letter-spacing:1px;' +
    'text-transform:uppercase;color:' + MUTED + ';">' + esc_(k) + ': </span>' +
    '<span style="' + F + 'font-size:14px;font-weight:bold;color:' + INK + ';">' + esc_(val) + '</span></div>';
}

/** HTML уншдаггүй клиентэд зориулсан энгийн текст хувилбар */
function reportText_(c, d, name) {
  var pkg = d.pkg || {}, sec = d.secondPkg;
  var L = [];
  L.push(COMPANY + ' — Бизнесийн оношилгооны дүгнэлт');
  L.push('');
  L.push('Сайн байна уу, ' + name + '.');
  L.push('Танай бөглөсөн 14 асуултын хариулт дээр үндэслэн дараах дүгнэлтийг гаргалаа.');
  L.push('');
  L.push('ТАНАЙ БИЗНЕСЭЭС ИЛЭРСЭН ЗҮЙЛС');
  (d.problems || []).forEach(function (x) { L.push('  ! ' + x); });
  if (d.monthlyHrs) L.push('  · Сард мессежид: ~' + d.monthlyHrs + ' цаг');
  if (d.saveHrs)    L.push('  · Хэмнэж болох: ~' + d.saveHrs + ' цаг / сард');
  if (d.lostOrders) L.push('  · Алдагдсан захиалга: ~' + d.lostOrders + ' / сард');
  L.push('');
  L.push('БИД ҮҮНИЙГ ИНГЭЖ ШИЙДЭЖ ӨГНӨ');
  (d.solutions || []).forEach(function (x) { L.push('  -> ' + x); });
  if (d.later && d.later.length) {
    L.push('');
    L.push('ЭНЭ БАГЦАД БАГТАХГҮЙ — ДАРААГИЙН ҮЕ ШАТАНД');
    d.later.forEach(function (x) { L.push('  · ' + x); });
  }
  L.push('');
  L.push('САНАЛ БОЛГОЖ БУЙ ШИЙДЭЛ');
  L.push('  ' + (pkg.name || '') + ' — ' + (pkg.price || ''));
  L.push('  Хүлээлгэн өгөх хугацаа: ' + (pkg.time || ''));
  L.push('  Төлбөр: 50% урьдчилгаа, 50% хүлээлгэн өгөхөд');
  (pkg.items || []).forEach(function (x) { L.push('    • ' + x); });
  if (sec) {
    L.push('');
    L.push('ДАРААГИЙН ҮЕ ШАТАНД: ' + sec.name + ' — ' + sec.price);
  }
  L.push('');
  L.push('Хугацаа нь танаас материал бүрэн ирсэн өдрөөс эхэлнэ. Дээрх үнэ бол урьдчилсан');
  L.push('тооцоо — үнэгүй уулзалтаар ажлын хүрээг тодруулж эцэслэнэ.');
  L.push('');
  L.push('ДАРААГИЙН АЛХАМ');
  L.push('Бид ажлын 1 өдөрт багтаан ' + (c.phone || '') + ' дугаараар холбогдоно.');
  L.push('Эсвэл шууд залгаарай: ' + PHONE_1 + ' | ' + PHONE_2);
  L.push('');
  L.push(COMPANY);
  L.push(NOTIFY_TO + ' · ' + SITE);
  L.push('MN TOWER 1201, J.Sambuu St, 5th khoroo, Chingeltei district, Ulaanbaatar 15141, Mongolia');
  return L.join('\n');
}

/* ============================================================
   ДОТООД МЭДЭГДЭЛ
   ============================================================ */
function notify_(c, d, answersText, sent) {
  var subject = 'Шинэ оношилгоо — ' + (c.company || c.name) + ' · ' + (d.recommended || '');
  var body =
    'ХОЛБОО БАРИХ\n' +
    'Нэр:          ' + (c.name || '-') + '\n' +
    'Утас:         ' + (c.phone || '-') + '\n' +
    'И-мэйл:       ' + (c.email || '-') + '\n' +
    'Байгууллага:  ' + (c.company || '-') + '\n' +
    'Албан тушаал: ' + (c.role || '-') + '\n' +
    'Вэб:          ' + (c.website || '-') + '\n' +
    'Facebook:     ' + (c.facebook || '-') + '\n' +
    'Instagram:    ' + (c.instagram || '-') + '\n\n' +
    'ҮР ДҮН\n' +
    'Санал болгосон:  ' + (d.recommended || '-') + '\n' +
    'Дараагийн алхам: ' + (d.second || '-') + '\n' +
    'Сард мессежид:   ~' + (d.monthlyHrs || 0) + ' цаг\n' +
    'Хэмнэж болох:    ~' + (d.saveHrs || 0) + ' цаг\n' +
    'Алдагдсан захиалга: ~' + (d.lostOrders || 0) + ' / сард\n\n' +
    'Дүгнэлт үйлчлүүлэгч рүү илгээсэн эсэх: ' + sent + '\n\n' +
    'БҮХ ХАРИУЛТ\n' + answersText + '\n\n' +
    'Хүснэгт: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject: subject,
    body: body,
    replyTo: c.email || undefined,
    name: 'scorej.biz оношилгоо'
  });
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
