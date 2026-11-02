/**
 * Pink Liliy Deer — orders + early-access backend (Google Apps Script).
 * Bound to a Google Sheet. Deploy as Web app: Execute as "Me", access "Anyone".
 */
const OWNER_EMAIL = 'YOUR_EMAIL@gmail.com'; // receives new-order notifications
const BRAND       = 'Pink Liliy Deer';
const SITE_BASE   = 'https://pink-liliy-deer.web.app/';
const EARLY_ACCESS_PASSWORD = 'pink2027'; // must match ACCESS_PW in lib/index.html

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.type === 'order')  handleOrder(data);
    if (data.type === 'signup') result = handleSignup(data);
    return json({ ok: true, result });
  } catch (err) {
    console.error(err);
    return json({ ok: false });
  }
}

function doGet() { return json({ ok: true, service: BRAND, version: 'access-email-v4' }); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const isEmail = s => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || ''));

/* ── Early access ── */
function handleSignup(d) {
  const email = String(d.email || '').trim().toLowerCase();
  if (!isEmail(email)) return 'invalid';
  const sh = sheet('Signups', ['Date', 'Email', 'Source', 'Last emailed']);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  const i = rows.findIndex(r => String(r[1]).trim().toLowerCase() === email);

  if (i !== -1) {
    // already signed up: resend the password, at most once per hour per address
    const last = rows[i][3];
    if (last && Date.now() - new Date(last).getTime() < 60 * 60 * 1000) return 'recent';
    sendAccessEmail(email, d.lang);
    sh.getRange(i + 2, 4).setValue(new Date());
    return 'resent';
  }
  sendAccessEmail(email, d.lang);
  sh.appendRow([new Date(), email, d.source || '', new Date()]);
  return 'mailed';
}

function sendAccessEmail(email, lang) {
  const ar = lang === 'ar';
  const pw = `<div style="font-size:22px;letter-spacing:2px;font-weight:bold;color:#dc768c;background:#fff;border:1px dashed #dc768c;padding:14px;text-align:center;margin:16px 0">${esc(EARLY_ACCESS_PASSWORD)}</div>`;
  const body = ar
    ? `<div dir="rtl" style="text-align:right"><p>شكراً لانضمامك إلى ${BRAND}! 🌸</p><p>كلمة مرور الوصول المبكر لكولكشن الشتاء / الخريف ٢٧:</p>${pw}<p><b>يفتح الوصول المبكر يوم الخميس ٥ نوفمبر.</b> وقتها اضغطي على «لديك كلمة مرور؟» في صفحتنا وأدخليها:</p><p><a href="${SITE_BASE}" style="color:#dc768c">${SITE_BASE}</a></p></div>`
    : `<p>Thank you for joining ${BRAND}! 🌸</p><p>Here is your early access password for the Winter / Fall ’27 collection:</p>${pw}<p><b>Early access opens on Thursday, 5 November.</b> From then, tap “Already have a password?” on our page and enter it:</p><p><a href="${SITE_BASE}" style="color:#dc768c">${SITE_BASE}</a></p>`;
  MailApp.sendEmail({
    to: email, name: BRAND, replyTo: OWNER_EMAIL,
    subject: ar ? `كلمة مرور الوصول المبكر — ${BRAND}` : `Your early access password — ${BRAND}`,
    htmlBody: wrap(ar ? 'وصول مبكر' : 'Early access', body),
  });
}

/* ── Orders ── */
function itemsHtml(items) {
  return items.map(i => `
    <table style="width:100%;margin-bottom:14px;padding-bottom:14px;border-bottom:1px solid #f5d5db;border-collapse:collapse"><tr>
      <td style="width:80px;vertical-align:top;padding-right:14px">
        <img src="${esc(SITE_BASE + i.img)}" width="70" height="88" style="display:block;object-fit:cover;border:1px solid #f5d5db"></td>
      <td style="vertical-align:top">
        <div style="font-weight:bold;color:#1a1218;margin-bottom:4px">${esc(i.name)}</div>
        <div style="color:#9c8790;font-size:12px;margin-bottom:2px">Size: ${esc(i.size)}${i.color ? ` &middot; Color: <span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:${esc(i.color)};border:1px solid #f5d5db"></span>` : ''}</div>
        <div style="color:#9c8790;font-size:12px;margin-bottom:6px">Qty: ${esc(i.qty)}</div>
        <div style="color:#dc768c;font-weight:bold">LE ${(Number(i.price) * Number(i.qty)).toLocaleString('en-US')}.00</div>
      </td></tr></table>`).join('');
}

function summaryHtml(d) {
  return `
    <div style="margin-top:8px;font-size:14px;color:#1a1218">
      <div>Subtotal: ${esc(d.subtotal)}</div>
      <div>Shipping: ${esc(d.shipping)}</div>
      <div style="font-weight:bold;margin-top:4px">Total: ${esc(d.total)}</div>
      <div style="color:#9c8790;font-size:12px;margin-top:4px">Payment: ${esc(d.payment_method)}</div>
    </div>`;
}

function wrap(title, body) {
  return `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;padding:24px;background:#fff5f7;color:#1a1218">
    <h2 style="color:#dc768c;font-weight:normal;margin:0 0 16px">${esc(title)}</h2>${body}</div>`;
}

function handleOrder(d) {
  const items = Array.isArray(d.items) ? d.items : [];
  const address = [d.customer_address, d.customer_city, d.customer_gov].filter(Boolean).join(', ');

  sheet('Orders', ['Date', 'Order #', 'Name', 'Email / Phone', 'Phone', 'Address', 'Items', 'Subtotal', 'Shipping', 'Total', 'Payment'])
    .appendRow([new Date(), d.order_num, d.customer_name, d.customer_email, d.customer_phone, address,
      items.map(i => `${i.name} (${i.size}) x${i.qty}`).join('; '), d.subtotal, d.shipping, d.total, d.payment_method]);

  const ownerBody =
    `<p><b>${esc(d.customer_name)}</b><br>${esc(d.customer_email)}<br>${esc(d.customer_phone)}<br>${esc(address)}</p>` +
    itemsHtml(items) + summaryHtml(d);
  MailApp.sendEmail({
    to: OWNER_EMAIL, subject: `New order ${d.order_num} — ${d.customer_name}`,
    htmlBody: wrap(`New order ${d.order_num}`, ownerBody), name: BRAND,
  });

  if (isEmail(d.customer_email)) {
    const custBody =
      `<p>Hi ${esc(d.customer_name)}, thank you for your order! We'll contact you shortly to confirm delivery.</p>
       <p style="color:#9c8790;font-size:12px">${esc(d.order_date)}</p>` +
      itemsHtml(items) + summaryHtml(d) + `<p style="margin-top:16px;font-size:13px">Delivering to: ${esc(address)}</p>`;
    MailApp.sendEmail({
      to: d.customer_email, subject: `Your ${BRAND} order ${d.order_num}`,
      htmlBody: wrap(`Order ${d.order_num} confirmed`, custBody), name: BRAND, replyTo: OWNER_EMAIL,
    });
  }
}

/**
 * Run once from the editor (choose it in the function dropdown, then Run).
 * Emails the early-access password to every Signups row that has not been emailed yet
 * (column D empty) and records the time in column D.
 */
function sendPendingAccessEmails() {
  const sh = sheet('Signups', ['Date', 'Email', 'Source', 'Last emailed']);
  const last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  let sent = 0;
  rows.forEach((r, i) => {
    const email = String(r[1]).trim().toLowerCase();
    if (r[3] || !isEmail(email)) return;
    try {
      sendAccessEmail(email, 'en');
      sh.getRange(i + 2, 4).setValue(new Date());
      sent++;
    } catch (err) {
      console.error(email, err);
    }
  });
  console.log('Sent ' + sent + ' email(s)');
}
