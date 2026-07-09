function foNowId_(prefix) {
  const ts = Utilities.formatDate(new Date(), FO_CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
  return prefix + '-' + ts;
}

function foGetActiveUser_() {
  try {
    return Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {
    return 'unknown';
  }
}

function foNum_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/[$,%\s,]/g, '');
  const n = Number(cleaned);
  return isNaN(n) ? '' : n;
}

function foGetVal_(row, headers, name) {
  const index = headers.indexOf(name);
  return index >= 0 ? row[index] : '';
}

function foTimestamp_() {
  return new Date();
}