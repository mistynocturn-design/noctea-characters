var TWITTER_ARCHIVE_API = {
  VERSION: '2026-07-27.1',
  SHEET_NAME: 'Twitter',
  READ_KEY_PROPERTY: 'BACKUP_READ_KEY',
  MAX_ROWS: 5000
};

function doGet() {
  return json_({
    ok: true,
    service: 'noctea-twitter-archive-reader',
    api_version: TWITTER_ARCHIVE_API.VERSION,
    time: new Date()
  });
}

function doPost(e) {
  try {
    var body = parseBody_(e);
    checkReadKey_(body.read_key);
    if (String(body.action || '').toLowerCase() !== 'read_twitter') {
      throw new Error('Unknown action.');
    }
    return json_(readTwitter_(body));
  } catch (error) {
    return json_({
      ok: false,
      api_version: TWITTER_ARCHIVE_API.VERSION,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function readTwitter_(body) {
  var month = normalizeMonth_(body.month);
  if (!month) throw new Error('Month must use YYYY-MM format.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TWITTER_ARCHIVE_API.SHEET_NAME);
  if (!sheet) throw new Error('Missing sheet: ' + TWITTER_ARCHIVE_API.SHEET_NAME);

  var values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Twitter sheet is empty.');
  var headers = values.shift().map(function (value) { return normalizeHeader_(value); });
  var columns = {
    date: findColumn_(headers, ['date', '작성일', '날짜']),
    author: findColumn_(headers, ['author', '작성자']),
    profile: findOptionalColumn_(headers, ['profileimage', 'profile_image', '프로필이미지', '프로필']),
    tags: findOptionalColumn_(headers, ['tags', 'tag', '태그']),
    content: findColumn_(headers, ['content', 'text', '내용', '본문']),
    images: findOptionalColumn_(headers, ['imageurls', 'image_urls', 'images', '이미지', '이미지링크']),
    threadId: findOptionalColumn_(headers, ['threadid', 'thread_id', '타래id', '타래아이디']),
    threadMonth: findOptionalColumn_(headers, ['threadstartmonth', 'thread_start_month', '타래시작월'])
  };

  if (values.length > TWITTER_ARCHIVE_API.MAX_ROWS) {
    throw new Error('Twitter sheet has too many rows.');
  }

  var rows = values.map(function (row, index) {
    return {
      sheet_row: index + 2,
      date: dateText_(row[columns.date]),
      author: cell_(row, columns.author),
      profile_image: cell_(row, columns.profile),
      tags: cell_(row, columns.tags),
      content: cell_(row, columns.content),
      image_urls: cell_(row, columns.images),
      thread_id: cell_(row, columns.threadId),
      thread_start_month: normalizeMonth_(cell_(row, columns.threadMonth))
    };
  }).filter(function (row) {
    return row.date || row.content || row.image_urls;
  });

  var blockNumber = 0;
  var previousImplicitId = '';
  var previousImplicitKey = '';
  var previousImplicitMonth = '';

  rows.forEach(function (row) {
    var writtenMonth = normalizeMonth_(row.date);
    if (!row.thread_id) {
      row.thread_owner_month = writtenMonth;
      row.thread_key = '';
      previousImplicitId = '';
      previousImplicitKey = '';
      previousImplicitMonth = '';
      return;
    }
    if (row.thread_start_month) {
      row.thread_owner_month = row.thread_start_month;
      row.thread_key = row.thread_start_month + '::' + row.thread_id;
      previousImplicitId = '';
      previousImplicitKey = '';
      previousImplicitMonth = '';
      return;
    }
    if (row.thread_id === previousImplicitId && writtenMonth === previousImplicitMonth && previousImplicitKey) {
      row.thread_key = previousImplicitKey;
    } else {
      blockNumber += 1;
      row.thread_key = writtenMonth + '::' + row.thread_id + '::' + blockNumber;
    }
    row.thread_owner_month = row.thread_key.slice(0, 7);
    previousImplicitId = row.thread_id;
    previousImplicitKey = row.thread_key;
    previousImplicitMonth = writtenMonth;
  });

  var output = rows.filter(function (row) {
    return row.thread_owner_month === month;
  }).map(function (row) {
    row.thread_start_month = row.thread_id ? row.thread_owner_month : '';
    delete row.thread_owner_month;
    delete row.thread_key;
    return row;
  });

  return {
    ok: true,
    api_version: TWITTER_ARCHIVE_API.VERSION,
    month: month,
    rows: output
  };
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/[\s_-]/g, '');
}

function findColumn_(headers, aliases) {
  var index = findOptionalColumn_(headers, aliases);
  if (index === -1) throw new Error('Missing column. Expected one of: ' + aliases.join(', '));
  return index;
}

function findOptionalColumn_(headers, aliases) {
  var normalizedAliases = aliases.map(normalizeHeader_);
  for (var index = 0; index < headers.length; index += 1) {
    if (normalizedAliases.indexOf(headers[index]) !== -1) return index;
  }
  return -1;
}

function cell_(row, index) {
  return index === -1 || !row ? '' : String(row[index] == null ? '' : row[index]).trim();
}

function dateText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var text = String(value || '').trim().replace(/\./g, '-').replace(/\//g, '-');
  var match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return match ? match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2) : text;
}

function normalizeMonth_(value) {
  var text = String(value || '').trim().replace(/\./g, '-').replace(/\//g, '-');
  var match = text.match(/^(\d{4})-(\d{1,2})/);
  return match ? match[1] + '-' + ('0' + match[2]).slice(-2) : '';
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Missing POST body.');
  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('POST body must be valid JSON.');
  }
}

function checkReadKey_(provided) {
  var expected = PropertiesService.getScriptProperties().getProperty(TWITTER_ARCHIVE_API.READ_KEY_PROPERTY);
  if (!expected) throw new Error('BACKUP_READ_KEY is not configured.');
  if (String(provided || '') !== expected) throw new Error('Invalid read key.');
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
