const LEDGER = Object.freeze({
  outgoingSheet: '출금관리',
  incomingSheet: '입금관리',
  sourceSheet: '원본대조',
  logSheet: '가져오기 기록',
  dashboardSheet: '대시보드',
  monthSheet: '월목록',
  naverSyncLogSheet: '네이버페이 동기화 기록',
  spreadsheetIdProperty: 'LEDGER_SPREADSHEET_ID',
  naverTokenProperty: 'NAVER_PAY_SYNC_TOKEN',
  uploadRoot: '가계부_거래내역_가져오기',
  uploadFolder: '스프레드시트_업로드_원본',
  maxFileBytes: 8 * 1024 * 1024,
  maxBatchBytes: 25 * 1024 * 1024,
  extensions: ['csv', 'xls', 'xlsx', 'pdf'],
  expenseCategories: [
    '식비', '카페·간식', '식료품·편의점', '교통비', '공과금',
    '문화·여가', '의료·건강', '미용', '구독·디지털', '쇼핑·생활',
    '지역상품권·온누리충전', '기타·확인필요', '내부이체'
  ],
  incomeCategories: [
    '이자', '환급·캐시백', '급여', '용돈', '정산금', '기타입금',
    '입금·확인필요', '내부이체'
  ]
});

function onOpen() {
  const ss = SpreadsheetApp.getActive();
  const menu = SpreadsheetApp.getUi().createMenu('가계부')
    .addItem('거래내역 여러 파일 업로드', 'showUploadDialog')
    .addItem('선택 거래를 같은 상호로 분류', 'applySelectedMerchantRule')
    .addItem('분류 규칙 보기', 'showCategoryRules')
    .addItem('새 월 등록', 'registerNewMonth')
    .addSeparator()
    .addItem('월별 대시보드 보기', 'showDashboard');
  if (!isSimpleLedger_(ss)) menu.addItem('네이버페이 자동 동기화 설정', 'showNaverPaySyncSetup');
  menu
    .addSeparator()
    .addItem('자동화 초기 설정', 'setupLedgerUploader')
    .addItem('가져오기 기록 보기', 'showImportLog')
    .addToUi();
}

function onEdit(event) {
  if (!event || !event.range) return;
  const sheet = event.range.getSheet();
  const ss = event.source;
  if (isSimpleLedger_(ss)) {
    handleSimpleLedgerEdit_(event);
    return;
  }
  handleLegacyLedgerEdit_(event);
  if (sheet.getName() === LEDGER.dashboardSheet && event.range.getA1Notation() === 'B2') {
    ensureDashboard_(ss, cleanText_(event.value));
    return;
  }
  if (sheet.getName() === LEDGER.outgoingSheet) {
    const dashboard = ss.getSheetByName(LEDGER.dashboardSheet);
    if (dashboard) refreshCalendarNotes_(ss, dashboard, selectedMonth_(ss));
  }
}

function showUploadDialog() {
  setupLedgerUploader_();
  const html = HtmlService.createHtmlOutputFromFile('Upload')
    .setWidth(760)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '은행 거래내역 여러 파일 업로드');
}

function setupLedgerUploader() {
  setupLedgerUploader_();
  SpreadsheetApp.getUi().alert('업로드 기능을 사용할 준비가 되었습니다.');
}

function setupLedgerUploader_() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty(LEDGER.spreadsheetIdProperty, ss.getId());
  if (isSimpleLedger_(ss)) {
    setupSimpleLedger_(ss);
    return;
  }
  requireLedgerSheets_(ss);
  ensureSimpleRuleSheet_(ss);
  ensureLogSheet_(ss);
  ensureMonthRegistry_(ss);
  ensureDashboard_(ss);
}

function showImportLog() {
  const ss = SpreadsheetApp.getActive();
  if (isSimpleLedger_(ss)) {
    ss.setActiveSheet(ensureSimpleImportLog_(ss));
    return;
  }
  const sheet = ensureLogSheet_(ss);
  ss.setActiveSheet(sheet);
}

function showDashboard() {
  const ss = SpreadsheetApp.getActive();
  if (isSimpleLedger_(ss)) {
    ss.setActiveSheet(ss.getSheetByName('대시보드'));
    return;
  }
  ensureMonthRegistry_(ss);
  const sheet = ensureDashboard_(ss);
  ss.setActiveSheet(sheet);
}

function registerNewMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('새 월 등록', '등록할 월을 YYYY-MM 형식으로 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const month = cleanText_(response.getResponseText());
  if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month)) {
    ui.alert('월은 YYYY-MM 형식으로 입력하세요.');
    return;
  }
  const ss = SpreadsheetApp.getActive();
  if (isSimpleLedger_(ss)) {
    const dashboard = ss.getSheetByName('대시보드');
    dashboard.getRange('B3')
      .setValue(new Date(month + '-01T00:00:00+09:00'))
      .setNumberFormat('yyyy-mm');
    ss.setActiveSheet(dashboard);
    ui.alert(month + ' 월이 선택되었습니다. 이제 거래내역 파일을 업로드할 수 있습니다.');
    return;
  }
  ensureMonthRegistry_(ss, month);
  const dashboard = ensureDashboard_(ss, month);
  dashboard.getRange('B2').setValue(month);
  ss.setActiveSheet(dashboard);
  ui.alert(month + ' 월이 등록되었습니다. 이제 거래내역 파일을 업로드할 수 있습니다.');
}

function showNaverPaySyncSetup() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty(LEDGER.spreadsheetIdProperty, ss.getId());
  ensureNaverPaySyncToken_();
  const html = HtmlService.createHtmlOutputFromFile('SyncSetup').setWidth(620).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '네이버페이 자동 동기화 설정');
}

function getNaverPaySyncSetup() {
  const properties = PropertiesService.getScriptProperties();
  return {
    spreadsheetName: SpreadsheetApp.getActive().getName(),
    token: ensureNaverPaySyncToken_(),
    spreadsheetConfigured: Boolean(properties.getProperty(LEDGER.spreadsheetIdProperty))
  };
}

function rotateNaverPaySyncToken() {
  const token = createNaverPaySyncToken_();
  PropertiesService.getScriptProperties().setProperty(LEDGER.naverTokenProperty, token);
  return token;
}

function getUploaderConfig() {
  const ss = SpreadsheetApp.getActive();
  const savedNames = PropertiesService.getUserProperties().getProperty('OWNER_LABELS') || '';
  return {
    spreadsheetName: ss.getName(),
    targetMonth: selectedMonth_(ss),
    ownerLabels: savedNames,
    extensions: LEDGER.extensions,
    maxFileBytes: LEDGER.maxFileBytes,
    maxBatchBytes: LEDGER.maxBatchBytes
  };
}

/**
 * Saves and analyzes one browser-selected file. The dialog calls this method
 * sequentially for every file in the multi-select list.
 */
function analyzeUploadedFile(payload) {
  validateUploadPayload_(payload);
  const ownerLabels = parseOwnerLabels_(payload.ownerLabels);
  if (!ownerLabels.length) throw new Error('본인 계좌이체를 구분할 이름을 한 개 이상 입력하세요.');

  const bytes = Utilities.base64Decode(payload.base64);
  if (bytes.length > LEDGER.maxFileBytes) {
    throw new Error('파일 하나는 8MB 이하만 업로드할 수 있습니다.');
  }

  const mime = payload.mimeType || mimeFromExtension_(extension_(payload.name));
  const blob = Utilities.newBlob(bytes, mime, safeFileName_(payload.name));
  const originalFolder = getOrCreateChildFolder_(getOrCreateFolder_(LEDGER.uploadRoot), LEDGER.uploadFolder);
  const original = originalFolder.createFile(blob);

  try {
    if (payload.pdfEncrypted && !payload.pdfText) {
      throw new Error('암호는 확인했지만 PDF에서 텍스트를 읽지 못했습니다. 스캔형 암호화 PDF는 현재 지원하지 않습니다.');
    }
    const raw = parseSavedFile_(original, payload.pdfText || '');
    if (payload.pdfEncrypted) raw.warnings = (raw.warnings || []).concat('암호화 PDF 해제 완료(암호는 저장하지 않음)');
    const transactions = raw.rows.map((row) => normalizeTransaction_(row, raw.bank, original, ownerLabels));
    if (!transactions.length) throw new Error('거래 행을 찾지 못했습니다. 은행명과 파일 기간을 확인하세요.');

    const periods = [...new Set(transactions.map((row) => row.date.slice(0, 7)))].sort();
    return {
      fileId: original.getId(),
      fileName: original.getName(),
      fileUrl: original.getUrl(),
      bank: raw.bank,
      periods: periods,
      warnings: raw.warnings || [],
      count: transactions.length,
      outgoingCount: transactions.filter((row) => row.direction === '출금').length,
      incomingCount: transactions.filter((row) => row.direction === '입금').length,
      spend: sum_(transactions.filter((row) => row.bucket === '지출').map((row) => row.amount)),
      incoming: sum_(transactions.filter((row) => row.bucket === '입금').map((row) => row.amount)),
      transactions: transactions
    };
  } catch (error) {
    const entry = {
      fileName: original.getName(), fileUrl: original.getUrl(), bank: '',
      found: 0, added: 0, duplicate: 0, skipped: 0,
      status: '분석실패', message: error.message
    };
    if (isSimpleLedger_(SpreadsheetApp.getActive())) appendSimpleImportLog_(SpreadsheetApp.getActive(), entry);
    else appendImportLog_(entry);
    throw error;
  }
}

/** Applies all analyzed files as one user-confirmed batch. */
function commitImportBatch(payload) {
  if (!payload || !Array.isArray(payload.files) || !payload.files.length) {
    throw new Error('반영할 분석 결과가 없습니다.');
  }
  if (!/^\d{4}-\d{2}$/.test(payload.targetMonth || '')) {
    throw new Error('대상 월을 YYYY-MM 형식으로 지정하세요.');
  }

  const ss = SpreadsheetApp.getActive();
  if (isSimpleLedger_(ss)) return commitSimpleLedgerBatch_(ss, payload);
  requireLedgerSheets_(ss);
  ensureLogSheet_(ss);
  PropertiesService.getUserProperties().setProperty('OWNER_LABELS', String(payload.ownerLabels || ''));

  const files = payload.files.map(validateAnalyzedFile_);
  const all = files.flatMap((file) => file.transactions.map((tx) => validateTransaction_(tx, file.fileId)));
  if (all.length > 10000) throw new Error('한 번에 반영할 수 있는 거래는 10,000건 이하입니다.');

  const inMonth = all.filter((row) => row.date.slice(0, 7) === payload.targetMonth);
  const autoTopUpMatches = reconcileTossCardTopUps_(inMonth);
  const skippedMonth = all.length - inMonth.length;
  const existing = readExistingTransactionKeys_(ss);
  const seen = new Set(existing);
  const fresh = [];
  let duplicateCount = 0;

  inMonth.forEach((row) => {
    const key = transactionKey_(row);
    if (seen.has(key)) duplicateCount += 1;
    else {
      seen.add(key);
      row.id = 'auto-' + sha256_(key).slice(0, 16);
      fresh.push(row);
    }
  });

  applyLegacyRulesToTransactions_(ss, fresh);

  const outgoing = fresh.filter((row) => row.direction === '출금');
  const incoming = fresh.filter((row) => row.direction === '입금');
  appendLedgerRows_(ss.getSheetByName(LEDGER.outgoingSheet), outgoing, true);
  appendLedgerRows_(ss.getSheetByName(LEDGER.incomingSheet), incoming, false);
  appendSourceRows_(ss.getSheetByName(LEDGER.sourceSheet), fresh);
  refreshSummaryFormulas_(ss);
  ensureMonthRegistry_(ss, payload.targetMonth);
  ensureDashboard_(ss, payload.targetMonth).getRange('B2').setValue(payload.targetMonth);
  expandNativeTablesBestEffort_(ss);

  files.forEach((file) => {
    const fileRows = inMonth.filter((row) => row.fileId === file.fileId);
    const added = fresh.filter((row) => row.fileId === file.fileId).length;
    appendImportLog_({
      fileName: file.fileName, fileUrl: file.fileUrl, bank: file.bank,
      found: file.transactions.length, added: added,
      duplicate: fileRows.length - added,
      skipped: file.transactions.length - fileRows.length,
      status: '완료', message: file.warnings.join(' / ')
    });
  });

  SpreadsheetApp.flush();
  return {
    added: fresh.length,
    outgoing: outgoing.length,
    incoming: incoming.length,
    duplicate: duplicateCount,
    skippedMonth: skippedMonth,
    targetMonth: payload.targetMonth,
    autoTopUpMatches: autoTopUpMatches
  };
}

function parseSavedFile_(file, pdfText) {
  const ext = extension_(file.getName());
  if (ext === 'csv') return parseCsvFile_(file);
  if (ext === 'xls' || ext === 'xlsx') return parseExcelFile_(file);
  if (ext === 'pdf') return pdfText ? parseWooriPdfText_(pdfText, file.getName()) : parsePdfFile_(file);
  throw new Error('지원하지 않는 파일 형식입니다: ' + ext);
}

function parseCsvFile_(file) {
  const blob = file.getBlob();
  let text = blob.getDataAsString('UTF-8');
  if ((text.match(/�/g) || []).length > 2) text = blob.getDataAsString('EUC-KR');
  text = text.replace(/^\uFEFF/, '');
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  const delimiters = [',', '\t', ';'];
  const delimiter = delimiters.sort((a, b) => firstLine.split(b).length - firstLine.split(a).length)[0];
  return parseMatrix_(Utilities.parseCsv(text, delimiter), file.getName());
}

function parseExcelFile_(file) {
  let tempId = '';
  try {
    const converted = Drive.Files.create(
      {name: 'tmp_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS},
      file.getBlob(),
      {fields: 'id'}
    );
    tempId = converted.id;
    const book = SpreadsheetApp.openById(tempId);
    const candidates = book.getSheets().map((sheet) => parseMatrix_(sheet.getDataRange().getDisplayValues(), file.getName(), true));
    const best = candidates.sort((a, b) => b.rows.length - a.rows.length)[0];
    if (!best || !best.rows.length) throw new Error('엑셀에서 거래 표를 찾지 못했습니다.');
    return best;
  } finally {
    if (tempId) DriveApp.getFileById(tempId).setTrashed(true);
  }
}

function parsePdfFile_(file) {
  let tempId = '';
  try {
    const converted = Drive.Files.create(
      {name: 'tmp_' + file.getName(), mimeType: MimeType.GOOGLE_DOCS},
      file.getBlob(),
      {ocrLanguage: 'ko', fields: 'id'}
    );
    tempId = converted.id;
    Utilities.sleep(800);
    const text = DocumentApp.openById(tempId).getBody().getText();
    return parseWooriPdfText_(text, file.getName());
  } finally {
    if (tempId) DriveApp.getFileById(tempId).setTrashed(true);
  }
}

function parseMatrix_(matrix, fileName, allowEmpty) {
  const headerIndex = matrix.findIndex((row) => {
    const joined = row.map(normalizeHeader_).join('|');
    return /적요|거래내용|가맹점/.test(joined) && /거래일시|거래일자|거래일/.test(joined) && /거래금액|출금액|지급금액/.test(joined);
  });
  if (headerIndex < 0) {
    if (allowEmpty) return {bank: bankFrom_(fileName, ''), rows: [], warnings: []};
    throw new Error('CSV/엑셀의 거래 헤더를 찾지 못했습니다.');
  }

  const headers = matrix[headerIndex].map(normalizeHeader_);
  const index = (names) => headers.findIndex((header) => names.indexOf(header) >= 0);
  const dateCol = index(['거래일시', '거래일자', '거래일', '일자', '날짜']);
  const descCol = index(['적요', '거래내용', '내용', '가맹점', '사용처']);
  const typeCol = index(['거래유형', '구분', '거래구분']);
  const amountCol = index(['거래금액', '금액']);
  const outCol = index(['출금액', '지급금액', '출금금액']);
  const inCol = index(['입금액', '입금금액']);
  const balanceCol = index(['거래후잔액', '잔액']);
  const institutionCol = index(['거래기관', '기관']);

  const bank = bankFrom_(fileName, matrix.slice(0, headerIndex + 2).flat().join(' '));
  const rows = [];
  matrix.slice(headerIndex + 1).forEach((row, offset) => {
    const dt = parseDateTime_(row[dateCol]);
    if (!dt) return;
    const type = cell_(row, typeCol);
    const desc = cell_(row, descCol);
    if (!desc && !type) return;
    let outgoing = parseAmount_(cell_(row, outCol));
    let incoming = parseAmount_(cell_(row, inCol));
    let rawAmount = 0;

    if (amountCol >= 0) {
      rawAmount = parseSignedAmount_(cell_(row, amountCol));
      const typeIsIncoming = /입금|환불|취소|캐시백|이자/.test(type);
      if (typeIsIncoming || rawAmount > 0) incoming = Math.abs(rawAmount);
      else outgoing = Math.abs(rawAmount);
    } else rawAmount = incoming - outgoing;

    if (!outgoing && !incoming) return;
    rows.push({
      date: dt.date, time: dt.time, description: desc, type: type,
      institution: cell_(row, institutionCol), outgoing: outgoing,
      incoming: incoming, rawAmount: rawAmount,
      balance: balanceCol >= 0 ? parseSignedAmount_(cell_(row, balanceCol)) : '',
      sourceRow: headerIndex + offset + 2
    });
  });
  return {bank: bank, rows: rows, warnings: []};
}

function parseWooriPdfText_(text, fileName) {
  const bank = bankFrom_(fileName, text);
  const pattern = /(20\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s+(\S+)\s+(.*?)\s+([\d,]+|-)\s+([\d,]+|-)\s+([\d,]+)/g;
  const rows = [];
  let match;
  while ((match = pattern.exec(text)) !== null) {
    const outgoing = parseAmount_(match[5]);
    const incoming = parseAmount_(match[6]);
    rows.push({
      date: match[1], time: match[2], type: match[3], description: match[4].trim(),
      institution: '', outgoing: outgoing, incoming: incoming,
      rawAmount: incoming - outgoing, balance: parseAmount_(match[7]), sourceRow: rows.length + 1
    });
  }
  if (!rows.length) throw new Error('PDF에서 거래 표를 찾지 못했습니다. 스캔 품질이나 비밀번호를 확인하세요.');
  return {bank: bank, rows: rows, warnings: ['PDF 변환 결과는 원본대조에서 확인하세요.']};
}

function normalizeTransaction_(raw, bank, file, ownerLabels) {
  const description = cleanText_(raw.description);
  const type = cleanText_(raw.type);
  const direction = raw.outgoing ? '출금' : '입금';
  const amount = Math.abs(raw.outgoing || raw.incoming);
  const ownerHit = ownerLabels.some((name) => description.indexOf(name) >= 0);
  const isTransferType = !/카드결제|체크카드/.test(type);
  const internal = /카드잔액\s*자동충전|토스.*자동충전|자동충전.*토스/.test(description) || (ownerHit && isTransferType);
  const bucket = internal ? '내부이체' : direction === '출금' ? '지출' : '입금';
  const method = paymentMethod_(description, type);
  const category = category_(description, type, direction, bucket, method);
  const row = {
    id: '', date: raw.date, time: raw.time || '미제공', bank: bank,
    method: method, description: description, type: type,
    direction: direction, amount: amount, rawAmount: Number(raw.rawAmount) || 0,
    balance: raw.balance === '' ? '' : Number(raw.balance),
    bucket: bucket, category: category,
    note: raw.time ? '' : '원본에 거래시간 없음',
    sourceRow: Number(raw.sourceRow) || 0,
    fileId: file.getId(), fileName: file.getName(), fileUrl: file.getUrl()
  };
  row.id = 'auto-' + sha256_(transactionKey_(row)).slice(0, 16);
  return row;
}

/**
 * OK저축은행 출금과 토스뱅크의 카드잔액 자동충전 입금을 한 쌍으로 표시합니다.
 * 두 거래는 출금/입금 시트에 남기되 지출 및 외부입금 합계에서는 제외합니다.
 */
function reconcileTossCardTopUps_(rows) {
  const usedOkRows = new Set();
  let matched = 0;
  const okOutgoing = rows.filter((row) =>
    row.bank === 'OK저축은행' && row.direction === '출금'
  );
  const tossTopUps = rows.filter((row) =>
    row.bank === '토스뱅크' && row.direction === '입금' && /카드잔액\s*자동충전/.test(row.description)
  );

  tossTopUps.forEach((tossRow) => {
    const okRow = okOutgoing.find((candidate) =>
      !usedOkRows.has(candidate) &&
      candidate.date === tossRow.date &&
      Number(candidate.amount) === Number(tossRow.amount)
    );
    if (!okRow) return;
    usedOkRows.add(okRow);
    [okRow, tossRow].forEach((row) => {
      row.bucket = '내부이체';
      row.category = '내부이체';
      row.method = '카드잔액 자동충전';
      if (!/OK저축은행→토스뱅크 자동충전/.test(row.note || '')) {
        row.note = cleanText_([row.note, 'OK저축은행→토스뱅크 자동충전 대응'].filter(Boolean).join(' / '));
      }
    });
    matched += 1;
  });
  return matched;
}

function paymentMethod_(description, type) {
  const text = description + ' ' + type;
  if (/체크카드|카드결제/.test(text)) return '체크카드';
  if (/네이버/.test(text)) return '네이버페이';
  if (/카카오페이/.test(text)) return '카카오페이';
  if (/온누리|상품권/.test(text) && /충전|상품권/.test(text)) return '상품권 충전';
  return '계좌이체';
}

function category_(description, type, direction, bucket, method) {
  if (bucket === '내부이체') return '내부이체';
  if (direction === '입금') {
    if (/이자/.test(description + type)) return '이자';
    if (/환불|취소|캐시백/.test(description + type)) return '환급·캐시백';
    return '입금·확인필요';
  }
  if (method === '상품권 충전') return '지역상품권·온누리충전';
  if (/전력|가스요금|수도|통신요금/.test(description)) return '공과금';
  if (/PC방|피시방|노래방|영화|유람선|넥슨캐시/i.test(description)) return '문화·여가';
  if (/의원|병원|약국/.test(description)) return '의료·건강';
  if (/헤어|미용/.test(description)) return '미용';
  if (/주유|교통|티머니|버스|택시|철도|한국자동차환경협회/.test(description)) return '교통비';
  if (/커피|카페|스타벅스|빽다방|제과|베이커리|설빙|카이막|모찌/.test(description)) return '카페·간식';
  if (/밥상|찌개|소바|어묵|타코야끼|식당|고기|분식|치킨|피자|떡볶이|카츠호|황양반|황올/.test(description)) return '식비';
  if (/마트|편의점|지에스25|GS25|세븐일레븐|CU|자연드림|초코에몽|이클립스|닭가슴살|혜자|비비고김치|드럼스틱버블/.test(description)) return '식료품·편의점';
  if (/ANTHROPIC|NETFLIX|YOUTUBE|APPLE SERVICES|구독/i.test(description)) return '구독·디지털';
  if (/다이소|에프알엘코리아/.test(description)) return '쇼핑·생활';
  return '기타·확인필요';
}

function appendLedgerRows_(sheet, rows, outgoing) {
  if (!rows.length) return;
  const start = sheet.getLastRow() + 1;
  const values = rows.map((row) => [
    row.id, toSheetDate_(row.date), row.time, safeCellText_(row.bank),
    safeCellText_(row.method), safeCellText_(row.description), row.amount,
    row.bucket, row.category, '', '', safeCellText_(row.note)
  ]);
  ensureSheetCapacity_(sheet, start + values.length - 1, 12);
  sheet.getRange(start, 1, values.length, 12).setValues(values);
  sheet.getRange(start, 10, values.length, 1).setFormulas(rows.map((_, i) => [
    '=IF(H' + (start + i) + '="' + (outgoing ? '지출' : '입금') + '",G' + (start + i) + ',0)'
  ]));
  sheet.getRange(start, 11, values.length, 1).setFormulas(rows.map((_, i) => [
    '=IF(OR(I' + (start + i) + '="기타·확인필요",I' + (start + i) + '="입금·확인필요"),"확인필요","분류완료")'
  ]));
  const list = outgoing ? LEDGER.expenseCategories : LEDGER.incomeCategories;
  const validation = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(start, 9, values.length, 1).setDataValidation(validation);
  sheet.getRange(start, 2, values.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(start, 7, values.length, 1).setNumberFormat('#,##0');
  sheet.getRange(start, 10, values.length, 1).setNumberFormat('#,##0');
}

function appendSourceRows_(sheet, rows) {
  if (!rows.length) return;
  const start = sheet.getLastRow() + 1;
  const sourceNames = {toss: '토스', woori: '우리', npay: '우리 N페이', ok: 'OK'};
  const values = rows.map((row) => [
    row.id, toSheetDate_(row.date), safeCellText_(row.bank), safeCellText_(row.description),
    safeCellText_(row.type), row.rawAmount, row.balance, row.sourceRow,
    sourceNames[bankCode_(row.bank)] || '자동업로드', safeCellText_(row.note), row.fileUrl
  ]);
  ensureSheetCapacity_(sheet, start + values.length - 1, 11);
  sheet.getRange(start, 1, values.length, 11).setValues(values);
  sheet.getRange(start, 2, values.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(start, 6, values.length, 2).setNumberFormat('#,##0');
}

function ensureSheetCapacity_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function refreshSummaryFormulas_(ss) {
  const sheet = ss.getSheetByName('월간요약');
  if (!sheet) return;
  const outEnd = ss.getSheetByName(LEDGER.outgoingSheet).getLastRow();
  const inEnd = ss.getSheetByName(LEDGER.incomingSheet).getLastRow();
  sheet.getRange('B5:B9').setFormulas([
    ["=SUM('출금관리'!J2:J" + outEnd + ')'],
    ["=SUM('입금관리'!J2:J" + inEnd + ')'],
    ["=SUM('출금관리'!G2:G" + outEnd + ')'],
    ["=SUM('입금관리'!G2:G" + inEnd + ')'],
    [`=COUNTIFS('출금관리'!I2:I${outEnd},"기타·확인필요",'출금관리'!H2:H${outEnd},"지출")`]
  ]);
  const categories = sheet.getRange('A12:A23').getDisplayValues().flat();
  sheet.getRange('B12:B23').setFormulas(categories.map((_, i) => [
    "=SUMIF('출금관리'!I$2:I$" + outEnd + ',A' + (i + 12) + ",'출금관리'!J$2:J$" + outEnd + ')'
  ]));
  for (let row = 5; row <= 35; row += 1) {
    sheet.getRange(row, 5).setFormula("=SUMIF('출금관리'!B$2:B$" + outEnd + ',D' + row + ",'출금관리'!J$2:J$" + outEnd + ')');
    sheet.getRange(row, 6).setFormula("=SUMIF('입금관리'!B$2:B$" + inEnd + ',D' + row + ",'입금관리'!J$2:J$" + inEnd + ')');
  }
}

function ensureMonthRegistry_(ss, explicitMonth) {
  let sheet = ss.getSheetByName(LEDGER.monthSheet);
  if (!sheet) {
    sheet = ss.insertSheet(LEDGER.monthSheet);
    sheet.getRange('A1').setValue('등록월').setFontWeight('bold').setBackground('#e8f0fe');
  }

  const months = new Set();
  if (explicitMonth) months.add(explicitMonth);
  months.add(monthFromTitle_(ss.getName()));
  [LEDGER.outgoingSheet, LEDGER.incomingSheet].forEach((name) => {
    const ledgerSheet = ss.getSheetByName(name);
    if (!ledgerSheet || ledgerSheet.getLastRow() < 2) return;
    ledgerSheet.getRange(2, 2, ledgerSheet.getLastRow() - 1, 1).getValues().forEach((row) => {
      const value = row[0];
      const month = value instanceof Date
        ? Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM')
        : normalizeDate_(value).slice(0, 7);
      if (/^20\d{2}-(0[1-9]|1[0-2])$/.test(month)) months.add(month);
    });
  });
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().forEach((month) => {
      if (/^20\d{2}-(0[1-9]|1[0-2])$/.test(month)) months.add(month);
    });
  }
  const values = [...months].sort().map((month) => [month]);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).clearContent();
  if (values.length) sheet.getRange(2, 1, values.length, 1).setValues(values);
  sheet.setColumnWidth(1, 110);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}

function ensureDashboard_(ss, preferredMonth) {
  let sheet = ss.getSheetByName(LEDGER.dashboardSheet);
  if (!sheet) sheet = ss.insertSheet(LEDGER.dashboardSheet, 0);
  const monthSheet = ensureMonthRegistry_(ss, preferredMonth);
  const current = /^20\d{2}-(0[1-9]|1[0-2])$/.test(sheet.getRange('B2').getDisplayValue())
    ? sheet.getRange('B2').getDisplayValue()
    : (preferredMonth || monthFromTitle_(ss.getName()));

  sheet.getRange('A1:G50').clearFormat();
  sheet.getRange('A1:G50').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.getRange('A1:G1').breakApart().merge().setValue('월별 가계부 대시보드')
    .setBackground('#174ea6').setFontColor('#ffffff').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue('기준 월').setFontWeight('bold');
  sheet.getRange('B2').setValue(current).setBackground('#e8f0fe').setFontWeight('bold');
  const monthValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(monthSheet.getRange('A2:A'), true).setAllowInvalid(false).build();
  sheet.getRange('B2').setDataValidation(monthValidation);

  sheet.getRange('A4:B4').setValues([['월 요약', '금액/건수']]).setBackground('#d2e3fc').setFontWeight('bold');
  sheet.getRange('A5:A8').setValues([['총지출'], ['총입금'], ['수지'], ['확인필요']]);
  const start = 'DATE(VALUE(LEFT($B$2,4)),VALUE(RIGHT($B$2,2)),1)';
  const end = 'EDATE(' + start + ',1)';
  sheet.getRange('B5:B8').setFormulas([
    ["=SUMIFS('출금관리'!$J:$J,'출금관리'!$B:$B,\">=\"&" + start + ",'출금관리'!$B:$B,\"<\"&" + end + ')'],
    ["=SUMIFS('입금관리'!$J:$J,'입금관리'!$B:$B,\">=\"&" + start + ",'입금관리'!$B:$B,\"<\"&" + end + ')'],
    ['=B6-B5'],
    ["=COUNTIFS('출금관리'!$I:$I,\"기타·확인필요\",'출금관리'!$H:$H,\"지출\",'출금관리'!$B:$B,\">=\"&" + start + ",'출금관리'!$B:$B,\"<\"&" + end + ')']
  ]);
  sheet.getRange('B5:B7').setNumberFormat('#,##0원');
  sheet.getRange('B8').setNumberFormat('#,##0건');

  const categories = LEDGER.expenseCategories.filter((value) => value !== '내부이체');
  sheet.getRange('A11:B11').setValues([['소비 항목', '금액']]).setBackground('#d2e3fc').setFontWeight('bold');
  sheet.getRange(12, 1, categories.length, 1).setValues(categories.map((value) => [value]));
  sheet.getRange(12, 2, categories.length, 1).setFormulas(categories.map((_, index) => {
    const row = index + 12;
    return [`=SUMIFS('출금관리'!$J:$J,'출금관리'!$I:$I,A${row},'출금관리'!$H:$H,"지출",'출금관리'!$B:$B,">="&${start},'출금관리'!$B:$B,"<"&${end})`];
  })).setNumberFormat('#,##0원');

  buildExpenseCalendar_(ss, sheet, current);
  sheet.setFrozenRows(2);
  sheet.setColumnWidths(1, 7, 125);

  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));
  const chart = sheet.newChart().asPieChart()
    .addRange(sheet.getRange(11, 1, categories.length + 1, 2))
    .setNumHeaders(1)
    .setPosition(2, 4, 0, 0)
    .setOption('title', '소비 항목 비율')
    .setOption('legend', {position: 'right', textStyle: {fontSize: 11}})
    .setOption('pieHole', 0.35)
    .setOption('pieSliceText', 'percentage')
    .setOption('backgroundColor', '#ffffff')
    .setOption('width', 560)
    .setOption('height', 330)
    .build();
  sheet.insertChart(chart);
  refreshCalendarNotes_(ss, sheet, current);
  return sheet;
}

function buildExpenseCalendar_(ss, sheet, month) {
  const year = Number(month.slice(0, 4));
  const monthIndex = Number(month.slice(5, 7)) - 1;
  const first = new Date(year, monthIndex, 1);
  const firstMonday = new Date(year, monthIndex, 1 - ((first.getDay() + 6) % 7));
  const weekdays = ['월', '화', '수', '목', '금', '토', '일'];

  sheet.getRange('D19:G50').breakApart().clearContent().clearFormat().clearNote().clearDataValidations();
  sheet.getRange('A25:G50').breakApart().clearContent().clearFormat().clearNote().clearDataValidations();
  sheet.getRange('A25:G25').merge().setValue('일자별 지출 달력')
    .setBackground('#174ea6').setFontColor('#ffffff').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A26:G26').setValues([weekdays]).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#e8f0fe');
  sheet.getRange('F26').setFontColor('#1a73e8');
  sheet.getRange('G26').setFontColor('#d93025');

  for (let week = 0; week < 6; week += 1) {
    const dateRow = 27 + week * 2;
    const amountRow = dateRow + 1;
    const dates = [];
    const formulas = [];
    for (let day = 0; day < 7; day += 1) {
      const date = new Date(firstMonday.getFullYear(), firstMonday.getMonth(), firstMonday.getDate() + week * 7 + day);
      const inMonth = date.getFullYear() === year && date.getMonth() === monthIndex;
      dates.push(inMonth ? date : '');
      const cell = columnLetter_(day + 1) + dateRow;
      formulas.push(inMonth
        ? `=IF(${cell}="",,SUMIF('출금관리'!$B:$B,${cell},'출금관리'!$J:$J))`
        : '');
    }
    sheet.getRange(dateRow, 1, 1, 7).setValues([dates]).setNumberFormat('d')
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left')
      .setBackground('#f8fafd').setBorder(true, true, false, true, false, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(amountRow, 1, 1, 7).setFormulas([formulas]).setNumberFormat('#,##0"원"')
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#ffffff').setBorder(false, true, true, true, false, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(dateRow, 24);
    sheet.setRowHeight(amountRow, 42);
    sheet.getRange(dateRow, 6, 2, 1).setBackground('#f3f8ff');
    sheet.getRange(dateRow, 7, 2, 1).setBackground('#fff5f5');
    sheet.getRange(dateRow, 6).setFontColor('#1a73e8');
    sheet.getRange(dateRow, 7).setFontColor('#d93025');
  }
}

function refreshCalendarNotes_(ss, dashboard, month) {
  if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month || '')) return;
  const byDate = new Map();
  const outgoing = ss.getSheetByName(LEDGER.outgoingSheet);
  if (outgoing && outgoing.getLastRow() >= 2) {
    outgoing.getRange(2, 1, outgoing.getLastRow() - 1, 12).getValues().forEach((row) => {
      if (cleanText_(row[7]) !== '지출') return;
      const date = row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Seoul', 'yyyy-MM-dd') : normalizeDate_(row[1]);
      if (date.slice(0, 7) !== month) return;
      const naverMarker = cleanText_(row[11]).match(/\[(?:네이버페이 사용처|네이버페이 결제):\s*([^\/\]]+)/);
      const time = cleanText_(row[2]) || '시간 미제공';
      const merchant = cleanText_(row[5]) || (naverMarker ? cleanText_(naverMarker[1]) : '') || '사용처 확인필요';
      const category = cleanText_(row[8]) || '기타·확인필요';
      const amount = Number(row[9]) || Number(row[6]) || 0;
      if (!byDate.has(date)) byDate.set(date, []);
      byDate.get(date).push({time: time, merchant: merchant, category: category, amount: amount});
    });
  }

  for (let week = 0; week < 6; week += 1) {
    const dateRow = 27 + week * 2;
    const amountRow = dateRow + 1;
    const dates = dashboard.getRange(dateRow, 1, 1, 7).getValues()[0];
    const notes = dates.map((value) => {
      if (!(value instanceof Date)) return '';
      const key = Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM-dd');
      const items = (byDate.get(key) || []).sort((left, right) => left.time.localeCompare(right.time));
      if (!items.length) return '지출 없음';
      const total = sum_(items.map((item) => item.amount));
      return ['총지출 ' + total.toLocaleString('ko-KR') + '원', '']
        .concat(items.map((item) => item.time + ' · ' + item.merchant + ' · ' + item.amount.toLocaleString('ko-KR') + '원 · ' + item.category))
        .join('\n');
    });
    dashboard.getRange(dateRow, 1, 1, 7).setNotes([notes]);
    dashboard.getRange(amountRow, 1, 1, 7).setNotes([notes]);
  }
}

function columnLetter_(column) {
  let result = '';
  while (column > 0) {
    column -= 1;
    result = String.fromCharCode(65 + column % 26) + result;
    column = Math.floor(column / 26);
  }
  return result;
}

function selectedMonth_(ss) {
  if (isSimpleLedger_(ss)) return simpleSelectedMonth_(ss);
  const sheet = ss.getSheetByName(LEDGER.dashboardSheet);
  const value = sheet ? sheet.getRange('B2').getDisplayValue() : '';
  return /^20\d{2}-(0[1-9]|1[0-2])$/.test(value) ? value : monthFromTitle_(ss.getName());
}

function doPost(event) {
  try {
    const payload = JSON.parse(event && event.postData && event.postData.contents || '{}');
    if (payload.action !== 'syncNaverPay') throw new Error('지원하지 않는 요청입니다.');
    const expected = PropertiesService.getScriptProperties().getProperty(LEDGER.naverTokenProperty) || '';
    if (!expected || !secureEquals_(String(payload.token || ''), expected)) throw new Error('동기화 토큰이 올바르지 않습니다.');
    if (!Array.isArray(payload.records) || !payload.records.length) throw new Error('동기화할 결제내역이 없습니다.');
    if (payload.records.length > 500) throw new Error('한 번에 500건까지만 동기화할 수 있습니다.');

    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(LEDGER.spreadsheetIdProperty);
    if (!spreadsheetId) throw new Error('대상 가계부가 설정되지 않았습니다.');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const result = importNaverPayDetails_(SpreadsheetApp.openById(spreadsheetId), payload.records);
      return jsonOutput_(Object.assign({ok: true}, result));
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return jsonOutput_({ok: false, error: error.message || String(error)});
  }
}

function importNaverPayDetails_(ss, records) {
  requireLedgerSheets_(ss);
  const log = ensureNaverSyncLog_(ss);
  const outgoing = ss.getSheetByName(LEDGER.outgoingSheet);
  const rowCount = Math.max(outgoing.getLastRow() - 1, 0);
  if (!rowCount) return {received: records.length, matched: 0, unmatched: records.length, duplicate: 0};

  const outgoingRows = outgoing.getRange(2, 1, outgoing.getLastRow() - 1, 12).getValues();
  const merchantRules = readSimpleRules_(ss);
  const claimed = new Set();
  const knownPaymentIds = new Set();
  const knownDetailKeys = new Set();
  const logRowByPaymentId = new Map();
  if (log.getLastRow() >= 2) {
    log.getRange(2, 6, log.getLastRow() - 1, 1).getDisplayValues().flat().forEach((paymentId, index) => {
      if (paymentId) logRowByPaymentId.set(paymentId, index + 2);
    });
  }
  outgoingRows.forEach((row) => {
    const note = cleanText_(row[11]);
    const matches = note.match(/결제번호:\s*([^\s|\/\]]+)/g) || [];
    matches.forEach((value) => knownPaymentIds.add(value.replace(/^결제번호:\s*/, '')));
    const marker = note.match(/\[네이버페이 결제:\s*([^\/\]]+)/);
    const shownAmount = note.match(/표시금액:\s*([\d,]+)원/);
    if (marker && shownAmount) {
      const date = row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Seoul', 'yyyy-MM-dd') : normalizeDate_(row[1]);
      knownDetailKeys.add(naverDetailKey_(date, row[2], Number(shownAmount[1].replace(/,/g, '')), marker[1]));
    }
  });

  let matched = 0;
  let unmatched = 0;
  let duplicate = 0;
  const unmatchedRows = [];
  records.map(normalizeNaverPayDetail_).forEach((detail) => {
    if (detail.paymentId && knownPaymentIds.has(detail.paymentId)) {
      duplicate += 1;
      return;
    }
    const detailKey = naverDetailKey_(detail.date, detail.time, detail.amount, detail.merchant || detail.item);
    if (!detail.paymentId && knownDetailKeys.has(detailKey)) {
      duplicate += 1;
      return;
    }
    const candidates = [];
    outgoingRows.forEach((row, index) => {
      if (claimed.has(index)) return;
      if (!cleanText_(row[0]) || cleanText_(row[7]) !== '지출') return;
      const date = row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Seoul', 'yyyy-MM-dd') : normalizeDate_(row[1]);
      const score = naverMatchScore_(date, row[6], row[2], detail);
      if (score === null) return;
      const naverHint = cleanText_(row[4]) === '네이버페이' || /네이버페이|네이버파이낸셜/.test(cleanText_(row[3]) + ' ' + cleanText_(row[5]));
      candidates.push({index: index, score: score + (naverHint ? 0 : 172800)});
    });
    candidates.sort((a, b) => a.score - b.score);
    if (!candidates.length) {
      unmatched += 1;
      if (!detail.paymentId || !logRowByPaymentId.has(detail.paymentId)) {
        unmatchedRows.push(naverSyncLogRow_(detail, '은행 거래와 미일치'));
      }
      return;
    }

    const index = candidates[0].index;
    claimed.add(index);
    const row = outgoingRows[index];
    const merchant = detail.merchant || detail.item || cleanText_(row[5]) || '사용처 확인필요';
    const description = cleanText_([detail.merchant, detail.item].filter(Boolean).join(' '));
    const matchedRule = matchSimpleRule_(merchant, merchantRules);
    const suggested = matchedRule ? matchedRule.category : category_(description, '', '출금', '지출', '네이버페이');
    const originalDescription = cleanText_(row[5]);
    row[5] = merchant;
    if ((!cleanText_(row[8]) || cleanText_(row[8]) === '기타·확인필요') && suggested !== '기타·확인필요') row[8] = suggested;
    row[11] = naverLedgerNote_(row[11], detail, originalDescription);
    outgoing.getRange(index + 2, 6).setValue(safeCellText_(row[5]));
    outgoing.getRange(index + 2, 9).setValue(safeCellText_(row[8] || '기타·확인필요'));
    outgoing.getRange(index + 2, 12).setValue(safeCellText_(row[11]));
    if (detail.paymentId && logRowByPaymentId.has(detail.paymentId)) {
      log.getRange(logRowByPaymentId.get(detail.paymentId), 7).setValue('추후 매칭완료');
    }
    if (detail.paymentId) knownPaymentIds.add(detail.paymentId);
    knownDetailKeys.add(detailKey);
    matched += 1;
  });

  if (unmatchedRows.length) {
    const startRow = log.getLastRow() + 1;
    ensureSheetCapacity_(log, startRow + unmatchedRows.length - 1, 8);
    log.getRange(startRow, 1, unmatchedRows.length, 8).setValues(unmatchedRows);
    log.getRange(log.getLastRow() - unmatchedRows.length + 1, 1, unmatchedRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    log.getRange(log.getLastRow() - unmatchedRows.length + 1, 2, unmatchedRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    log.getRange(log.getLastRow() - unmatchedRows.length + 1, 3, unmatchedRows.length, 1).setNumberFormat('#,##0원');
  }
  ensureDashboard_(ss);
  SpreadsheetApp.flush();
  return {received: records.length, matched: matched, unmatched: unmatched, duplicate: duplicate};
}

function normalizeNaverPayDetail_(record) {
  const date = cleanText_(record && record.date);
  const time = cleanText_(record && record.time) || '00:00:00';
  const amount = Number(record && record.amount);
  if (!/^20\d{2}-\d{2}-\d{2}$/.test(date)) throw new Error('네이버페이 결제 날짜 형식이 잘못되었습니다.');
  if (!/^\d{2}:\d{2}:\d{2}$/.test(time)) throw new Error('네이버페이 결제 시간 형식이 잘못되었습니다.');
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('네이버페이 결제 금액이 잘못되었습니다.');
  const detailUrl = cleanText_(record.detailUrl);
  if (detailUrl && !/^https:\/\/(?:orders\.pay\.naver\.com|pay\.naver\.com)\//.test(detailUrl)) throw new Error('네이버페이 상세 주소가 올바르지 않습니다.');
  return {
    paymentId: cleanText_(record.paymentId).slice(0, 100), date: date, time: time, amount: amount,
    merchant: cleanText_(record.merchant).slice(0, 300), item: cleanText_(record.item).slice(0, 500),
    detailUrl: detailUrl.slice(0, 1000)
  };
}

function naverLedgerNote_(previous, detail, originalDescription) {
  if (detail.paymentId && cleanText_(previous).indexOf('결제번호: ' + detail.paymentId) >= 0) return cleanText_(previous);
  const merchant = detail.merchant || detail.item || '사용처 확인필요';
  const item = detail.item ? '상품: ' + detail.item : '';
  const payment = detail.paymentId ? '결제번호: ' + detail.paymentId : '';
  const shownAmount = detail.amount ? '표시금액: ' + Number(detail.amount).toLocaleString('ko-KR') + '원' : '';
  const source = originalDescription && originalDescription !== merchant ? '은행 표시내용: ' + originalDescription : '';
  const marker = '[네이버페이 결제: ' + [merchant, payment, item, shownAmount, source, detail.detailUrl].filter(Boolean).join(' / ') + ']';
  const cleaned = cleanText_(previous).replace(/(?:^|\s)\[(?:NPay 자동동기화|네이버페이 사용처|네이버페이 결제):?[^\]]*\]/g, '').trim();
  return cleanText_((cleaned ? cleaned + ' ' : '') + marker).slice(0, 5000);
}

function naverMatchScore_(date, amount, time, detail) {
  if (normalizeDate_(date) !== detail.date) return null;
  const distance = timeDistance_(cleanText_(time), detail.time);
  if (Number(amount) === Number(detail.amount)) return distance;
  return distance <= 90 ? 86400 + distance : null;
}

function naverDetailKey_(date, time, amount, merchant) {
  return [normalizeDate_(date), cleanText_(time), Number(amount) || 0, cleanText_(merchant)].join('|');
}

function ensureNaverSyncLog_(ss) {
  let sheet = ss.getSheetByName(LEDGER.naverSyncLogSheet);
  if (!sheet) {
    sheet = ss.insertSheet(LEDGER.naverSyncLogSheet);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '처리시각', '결제일시', '금액', '사용처', '상품', '결제번호', '상태', '상세링크'
    ]]).setBackground('#eeeeee').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 3, 120);
    sheet.setColumnWidths(4, 2, 220);
    sheet.setColumnWidth(6, 180);
    sheet.setColumnWidth(7, 140);
    sheet.setColumnWidth(8, 260);
  }
  return sheet;
}

function naverSyncLogRow_(detail, status) {
  return [
    new Date(), new Date(detail.date + 'T' + detail.time + '+09:00'), detail.amount,
    safeCellText_(detail.merchant), safeCellText_(detail.item), safeCellText_(detail.paymentId),
    safeCellText_(status), detail.detailUrl
  ];
}

function timeDistance_(left, right) {
  const seconds = (value) => {
    const match = cleanText_(value).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    return match ? Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3] || 0) : 0;
  };
  return Math.abs(seconds(left) - seconds(right));
}

function ensureNaverPaySyncToken_() {
  const properties = PropertiesService.getScriptProperties();
  let token = properties.getProperty(LEDGER.naverTokenProperty);
  if (!token) {
    token = createNaverPaySyncToken_();
    properties.setProperty(LEDGER.naverTokenProperty, token);
  }
  return token;
}

function createNaverPaySyncToken_() {
  return sha256_(Utilities.getUuid() + '|' + Utilities.getUuid() + '|' + new Date().getTime());
}

function secureEquals_(left, right) {
  if (left.length !== right.length) return false;
  let result = 0;
  for (let index = 0; index < left.length; index += 1) result |= left.charCodeAt(index) ^ right.charCodeAt(index);
  return result === 0;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function readExistingTransactionKeys_(ss) {
  const keys = [];
  [[LEDGER.outgoingSheet, '출금'], [LEDGER.incomingSheet, '입금']].forEach(([name, direction]) => {
    const sheet = ss.getSheetByName(name);
    if (sheet.getLastRow() < 2) return;
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 6).getDisplayValues().forEach((row) => {
      keys.push([normalizeDate_(row[0]), row[1] || '미제공', cleanText_(row[2]), cleanText_(row[4]), parseAmount_(row[5]), direction].join('|'));
    });
  });
  return keys;
}

function transactionKey_(row) {
  return [row.date, row.time || '미제공', cleanText_(row.bank), cleanText_(row.description), Number(row.amount), row.direction].join('|');
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(LEDGER.logSheet);
  if (!sheet) {
    sheet = ss.insertSheet(LEDGER.logSheet);
    sheet.getRange(1, 1, 1, 11).setValues([[
      '처리시각', '파일명', '기록은행', '발견건수', '추가건수', '중복건수',
      '다른월 제외', '상태', '메시지', '원본링크', '실행사용자'
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:K1').setBackground('#eeeeee').setFontWeight('bold');
  }
  return sheet;
}

function appendImportLog_(entry) {
  const sheet = ensureLogSheet_(SpreadsheetApp.getActive());
  sheet.appendRow([
    new Date(), safeCellText_(entry.fileName), safeCellText_(entry.bank), entry.found || 0,
    entry.added || 0, entry.duplicate || 0, entry.skipped || 0,
    safeCellText_(entry.status), safeCellText_(entry.message), entry.fileUrl || '',
    Session.getActiveUser().getEmail() || '현재 사용자'
  ]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function expandNativeTablesBestEffort_(ss) {
  try {
    const meta = Sheets.Spreadsheets.get(ss.getId(), {fields: 'sheets(properties(sheetId,title),tables(tableId,range))'});
    const requests = [];
    (meta.sheets || []).forEach((item) => {
      const title = item.properties.title;
      if (title !== LEDGER.outgoingSheet && title !== LEDGER.incomingSheet) return;
      const lastRow = ss.getSheetByName(title).getLastRow();
      (item.tables || []).forEach((table) => {
        const next = Object.assign({}, table.range, {endRowIndex: lastRow, endColumnIndex: 12});
        requests.push({updateTable: {table: {tableId: table.tableId, range: next}, fields: 'range'}});
      });
    });
    if (requests.length) Sheets.Spreadsheets.batchUpdate({requests: requests}, ss.getId());
  } catch (error) {
    console.warn('표 범위 확장 보류: ' + error.message);
  }
}

/**
 * New template adapter. The new household ledger intentionally keeps one
 * transaction table instead of the legacy incoming/outgoing split.
 */
function isSimpleLedger_(ss) {
  return Boolean(ss && ss.getSheetByName('거래내역') && ss.getSheetByName('분류설정'));
}

function setupSimpleLedger_(ss) {
  const transactions = ss.getSheetByName('거래내역');
  const headers = transactions.getRange(5, 1, 1, 12).getDisplayValues()[0];
  if (headers.join('|') !== '거래일|구분|대분류|소분류|금액|결제방식|출금계좌|입금계좌|상호명·상대방|메모|기준월|연결번호') {
    throw new Error('새 가계부의 거래내역 헤더를 찾지 못했습니다. A5:L5를 확인하세요.');
  }
  ensureSimpleRuleSheet_(ss);
  ensureSimpleSourceSheet_(ss);
  ensureSimpleImportLog_(ss);
  expandSimpleDashboardFormulaRanges_(ss);
}

function simpleSelectedMonth_(ss) {
  const value = ss.getSheetByName('대시보드').getRange('B3').getValue();
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM');
  const match = cleanText_(value).match(/^(20\d{2})-(\d{2})/);
  return match ? match[0] : Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
}

function expandSimpleDashboardFormulaRanges_(ss) {
  const dashboard = ss.getSheetByName('대시보드');
  if (!dashboard) return;
  dashboard.getDataRange().getFormulas().forEach((row, rowIndex) => {
    row.forEach((formula, columnIndex) => {
      if (!formula) return;
      const updated = formula.replace(/\$205\b/g, '$5000');
      if (updated !== formula) dashboard.getRange(rowIndex + 1, columnIndex + 1).setFormula(updated);
    });
  });
}

function commitSimpleLedgerBatch_(ss, payload) {
  setupSimpleLedger_(ss);
  PropertiesService.getUserProperties().setProperty('OWNER_LABELS', String(payload.ownerLabels || ''));
  const files = payload.files.map(validateAnalyzedFile_);
  const all = files.flatMap((file) => file.transactions.map((tx) => validateTransaction_(tx, file.fileId)));
  if (all.length > 10000) throw new Error('한 번에 반영할 수 있는 거래는 10,000건 이하입니다.');

  const inMonth = all.filter((row) => row.date.slice(0, 7) === payload.targetMonth);
  const autoTopUpMatches = reconcileTossCardTopUps_(inMonth);
  const existing = readSimpleTransactionKeys_(ss);
  const seen = new Set(existing);
  const fresh = [];
  let duplicateCount = 0;
  inMonth.forEach((row) => {
    const key = transactionKey_(row);
    if (seen.has(key)) duplicateCount += 1;
    else {
      seen.add(key);
      row.id = 'auto-' + sha256_(key).slice(0, 16);
      fresh.push(row);
    }
  });

  const rules = readSimpleRules_(ss);
  appendSimpleLedgerRows_(ss, fresh.map((row) => simpleLedgerRow_(row, rules)));
  appendSimpleSourceRows_(ss, fresh);
  files.forEach((file) => {
    const fileRows = inMonth.filter((row) => row.fileId === file.fileId);
    const added = fresh.filter((row) => row.fileId === file.fileId).length;
    appendSimpleImportLog_(ss, {
      fileName: file.fileName, fileUrl: file.fileUrl, bank: file.bank,
      found: file.transactions.length, added: added, duplicate: fileRows.length - added,
      skipped: file.transactions.length - fileRows.length, status: '완료',
      message: file.warnings.join(' / ')
    });
  });
  expandSimpleTransactionTable_(ss);
  SpreadsheetApp.flush();
  return {
    added: fresh.length,
    outgoing: fresh.filter((row) => row.direction === '출금').length,
    incoming: fresh.filter((row) => row.direction === '입금').length,
    duplicate: duplicateCount,
    skippedMonth: all.length - inMonth.length,
    targetMonth: payload.targetMonth,
    autoTopUpMatches: autoTopUpMatches
  };
}

function readSimpleTransactionKeys_(ss) {
  const sheet = ensureSimpleSourceSheet_(ss);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean);
}

function simpleLedgerRow_(row, rules) {
  const transfer = row.bucket === '내부이체';
  const kind = transfer ? '이체' : row.direction === '출금' ? '지출' : '수입';
  const merchant = cleanText_(row.description) || '상호 확인 필요';
  const matchedRule = matchSimpleRule_(merchant, rules);
  const category = transfer ? '이체' : (matchedRule ? matchedRule.category : simpleFallbackCategory_(row));
  const sourceAccount = simpleAccountFromBank_(row.bank);
  const topUp = /OK저축은행→토스뱅크 자동충전 대응/.test(row.note || '');
  const fromAccount = transfer && topUp ? 'OK저축은행' : row.direction === '출금' ? sourceAccount : '';
  const toAccount = transfer && topUp ? '토스' : row.direction === '입금' ? sourceAccount : '';
  const memo = [
    '원본 ' + row.bank + (row.time ? ' ' + row.time : ''),
    row.type ? '유형: ' + row.type : '',
    matchedRule ? '분류규칙: ' + matchedRule.example : '',
    row.note || '',
    !transfer && !matchedRule && category === '기타' ? '분류 확인 필요' : ''
  ].filter(Boolean).join(' / ');
  return [
    toSheetDate_(row.date), kind, category, '', row.amount,
    simplePaymentMethod_(row, sourceAccount), fromAccount, toAccount,
    merchant, memo, row.date.slice(0, 7), row.id
  ];
}

function simpleFallbackCategory_(row) {
  if (row.direction === '입금') {
    if (/이자/.test(row.description + row.type)) return '기타수입';
    if (/환불|취소|캐시백/.test(row.description + row.type)) return '환급';
    return '기타수입';
  }
  const legacy = category_(row.description, row.type, row.direction, row.bucket, paymentMethod_(row.description, row.type));
  const map = {
    '식비': '식비', '카페·간식': '카페·간식', '교통비': '교통', '공과금': '주거·공과금',
    '의료·건강': '건강', '문화·여가': '취미·여가', '미용': '생활용품',
    '식료품·편의점': '식료품·편의점', '쇼핑·생활': '쇼핑', '지역상품권·온누리충전': '기타',
    '구독·디지털': '기타'
  };
  return map[legacy] || '기타';
}

function simplePaymentMethod_(row, sourceAccount) {
  const method = paymentMethod_(row.description, row.type);
  if (method === '체크카드') return sourceAccount === '토스' ? '토스카드' : '카드';
  if (method === '네이버페이') return '네이버페이 QR';
  if (method === '상품권 충전') return /온누리/.test(row.description + row.type) ? '온누리 앱' : '지역상품권 카드';
  return '계좌이체';
}

function simpleAccountFromBank_(bank) {
  if (bank === '토스뱅크') return '토스';
  if (bank === '우리은행 일반' || bank === '우리은행 N페이') return '우리은행';
  if (bank === 'OK저축은행') return 'OK저축은행';
  return '';
}

function firstSimpleTransactionRow_(sheet) {
  const start = 6;
  const length = Math.max(sheet.getLastRow() - start + 1, 1);
  const values = sheet.getRange(start, 1, length, 1).getDisplayValues().flat();
  const empty = values.findIndex((value) => !cleanText_(value));
  return empty >= 0 ? start + empty : sheet.getLastRow() + 1;
}

function appendSimpleLedgerRows_(ss, rows) {
  if (!rows.length) return;
  const sheet = ss.getSheetByName('거래내역');
  const start = firstSimpleTransactionRow_(sheet);
  ensureSheetCapacity_(sheet, start + rows.length - 1, 12);
  sheet.getRange(start, 1, rows.length, 12).setValues(rows);
  sheet.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(start, 5, rows.length, 1).setNumberFormat('#,##0');
}

function ensureSimpleRuleSheet_(ss) {
  let sheet = ss.getSheetByName('분류규칙');
  if (!sheet) {
    sheet = ss.insertSheet('분류규칙');
    sheet.getRange(1, 1, 1, 6).setValues([['상호키', '예시 상호', '대분류', '적용횟수', '마지막 적용일', '메모']]);
    sheet.getRange('A1:F1').setBackground('#e8f0fe').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 2, 180);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidths(4, 2, 110);
    sheet.setColumnWidth(6, 220);
  }
  return sheet;
}

function ensureSimpleSourceSheet_(ss) {
  let sheet = ss.getSheetByName('원본대조');
  if (!sheet) {
    sheet = ss.insertSheet('원본대조');
    sheet.getRange(1, 1, 1, 10).setValues([['거래키', '거래일', '기록은행', '적요', '거래유형', '원거래금액', '거래후잔액', '원본행', '메모', '원본링크']]);
    sheet.getRange('A1:J1').setBackground('#eeeeee').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function ensureSimpleImportLog_(ss) {
  let sheet = ss.getSheetByName('가져오기 기록');
  if (!sheet) {
    sheet = ss.insertSheet('가져오기 기록');
    sheet.getRange(1, 1, 1, 10).setValues([['처리시각', '파일명', '기록은행', '발견건수', '추가건수', '중복건수', '다른월 제외', '상태', '메시지', '원본링크']]);
    sheet.getRange('A1:J1').setBackground('#eeeeee').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function appendSimpleSourceRows_(ss, rows) {
  if (!rows.length) return;
  const sheet = ensureSimpleSourceSheet_(ss);
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, 10).setValues(rows.map((row) => [
    transactionKey_(row), toSheetDate_(row.date), row.bank, row.description, row.type,
    row.rawAmount, row.balance, row.sourceRow, row.note || '', row.fileUrl
  ]));
  sheet.getRange(start, 2, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(start, 6, rows.length, 2).setNumberFormat('#,##0');
}

function appendSimpleImportLog_(ss, entry) {
  const sheet = ensureSimpleImportLog_(ss);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 10).setValues([[new Date(), entry.fileName, entry.bank, entry.found || 0, entry.added || 0, entry.duplicate || 0, entry.skipped || 0, entry.status, entry.message || '', entry.fileUrl || '']]);
  sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function expandSimpleTransactionTable_(ss) {
  try {
    const sheet = ss.getSheetByName('거래내역');
    const meta = Sheets.Spreadsheets.get(ss.getId(), {fields: 'sheets(properties(sheetId,title),tables(tableId,range))'});
    const item = (meta.sheets || []).find((candidate) => candidate.properties.title === '거래내역');
    const table = item && (item.tables || [])[0];
    if (!table) return;
    const range = Object.assign({}, table.range, {endRowIndex: sheet.getLastRow(), endColumnIndex: 12});
    Sheets.Spreadsheets.batchUpdate({requests: [{updateTable: {table: {tableId: table.tableId, range: range}, fields: 'range'}}]}, ss.getId());
  } catch (error) {
    console.warn('거래내역 표 범위 확장 보류: ' + error.message);
  }
}

function merchantKey_(value) {
  const normalized = cleanText_(value).toUpperCase()
    .replace(/\(주\)|㈜|주식회사|\[주\]/g, '')
    .replace(/지에스\s*25/g, 'GS25')
    .replace(/[^0-9A-Z가-힣]/g, '');
  return normalized.indexOf('GS25') >= 0 ? 'GS25' : normalized;
}

function readSimpleRules_(ss) {
  const sheet = ensureSimpleRuleSheet_(ss);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues()
    .filter((row) => cleanText_(row[0]) && cleanText_(row[2]))
    .map((row) => ({key: cleanText_(row[0]), example: cleanText_(row[1]), category: cleanText_(row[2])}));
}

function matchSimpleRule_(merchant, rules) {
  const key = merchantKey_(merchant);
  return (rules || []).find((rule) => rule.key === key || (rule.key.length >= 4 && (key.indexOf(rule.key) >= 0 || rule.key.indexOf(key) >= 0))) || null;
}

function handleSimpleLedgerEdit_(event) {
  const range = event.range;
  if (range.getSheet().getName() !== '거래내역' || range.getRow() < 6 || range.getColumn() !== 3 || range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  const category = cleanText_(event.value);
  if (!category || category === '이체') return;
  applyMerchantRuleForRow_(event.source, range.getRow(), category, false);
}

function handleLegacyLedgerEdit_(event) {
  const range = event.range;
  if (range.getSheet().getName() !== LEDGER.outgoingSheet || range.getRow() < 2 || range.getColumn() !== 9 || range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  const category = cleanText_(event.value);
  if (!category || category === '기타·확인필요' || category === '내부이체') return;
  applyLegacyMerchantRuleForRow_(event.source, range.getRow(), category, false);
}

function applySelectedMerchantRule() {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getActiveRange();
  if (!isSimpleLedger_(ss)) {
    if (!range || range.getSheet().getName() !== LEDGER.outgoingSheet || range.getRow() < 2) throw new Error('출금관리에서 분류할 거래 행을 선택하세요.');
    const category = cleanText_(range.getSheet().getRange(range.getRow(), 9).getDisplayValue());
    if (!category || category === '기타·확인필요' || category === '내부이체') throw new Error('먼저 선택한 행의 지출분류를 확정하세요.');
    const result = applyLegacyMerchantRuleForRow_(ss, range.getRow(), category, true);
    SpreadsheetApp.getUi().alert('“' + result.merchant + '” 규칙을 저장하고 ' + result.count + '건에 적용했습니다.');
    return;
  }
  if (!range || range.getSheet().getName() !== '거래내역' || range.getRow() < 6) throw new Error('거래내역에서 분류할 거래 행을 선택하세요.');
  const category = cleanText_(range.getSheet().getRange(range.getRow(), 3).getDisplayValue());
  if (!category || category === '이체') throw new Error('먼저 선택한 행의 대분류를 확정하세요.');
  const result = applyMerchantRuleForRow_(ss, range.getRow(), category, true);
  SpreadsheetApp.getUi().alert('“' + result.merchant + '” 규칙을 저장하고 ' + result.count + '건에 적용했습니다.');
}

function applyMerchantRuleForRow_(ss, rowNumber, category, notify) {
  const sheet = ss.getSheetByName('거래내역');
  const row = sheet.getRange(rowNumber, 1, 1, 12).getDisplayValues()[0];
  const merchant = cleanText_(row[8]);
  const key = merchantKey_(merchant);
  if (!key) throw new Error('선택한 거래에 상호명·상대방이 없습니다.');
  const rules = ensureSimpleRuleSheet_(ss);
  const ruleValues = rules.getLastRow() >= 2 ? rules.getRange(2, 1, rules.getLastRow() - 1, 6).getValues() : [];
  const existing = ruleValues.findIndex((value) => cleanText_(value[0]) === key);
  const now = new Date();
  const ruleRow = existing >= 0 ? existing + 2 : rules.getLastRow() + 1;
  if (existing >= 0) rules.getRange(ruleRow, 1, 1, 6).setValues([[key, merchant, category, Number(ruleValues[existing][3]) || 0, now, '거래내역에서 확정']]);
  else rules.appendRow([key, merchant, category, 0, now, '거래내역에서 확정']);

  const start = 6;
  const length = Math.max(sheet.getLastRow() - start + 1, 1);
  const values = sheet.getRange(start, 1, length, 12).getValues();
  let count = 0;
  values.forEach((value) => {
    if (cleanText_(value[1]) !== '지출' || !matchSimpleRule_(cleanText_(value[8]), [{key: key}])) return;
    if (cleanText_(value[2]) !== category) count += 1;
    value[2] = category;
  });
  sheet.getRange(start, 3, length, 1).setValues(values.map((value) => [value[2]]));
  rules.getRange(ruleRow, 4).setValue(count);
  rules.getRange(ruleRow, 5).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
  if (notify) SpreadsheetApp.getActive().toast('같은·유사 상호 ' + count + '건을 ' + category + '로 분류했습니다.', '가계부');
  return {merchant: merchant, count: count};
}

function applyLegacyMerchantRuleForRow_(ss, rowNumber, category, notify) {
  const sheet = ss.getSheetByName(LEDGER.outgoingSheet);
  const row = sheet.getRange(rowNumber, 1, 1, 12).getDisplayValues()[0];
  if (cleanText_(row[7]) !== '지출') throw new Error('지출 거래만 상호 규칙으로 분류할 수 있습니다.');
  const merchant = cleanText_(row[5]);
  const key = merchantKey_(merchant);
  if (!key) throw new Error('선택한 거래에 사용처·적요가 없습니다.');

  const rules = ensureSimpleRuleSheet_(ss);
  const ruleValues = rules.getLastRow() >= 2 ? rules.getRange(2, 1, rules.getLastRow() - 1, 6).getValues() : [];
  const existing = ruleValues.findIndex((value) => cleanText_(value[0]) === key);
  const now = new Date();
  const ruleRow = existing >= 0 ? existing + 2 : rules.getLastRow() + 1;
  if (existing >= 0) rules.getRange(ruleRow, 1, 1, 6).setValues([[key, merchant, category, Number(ruleValues[existing][3]) || 0, now, '출금관리에서 확정']]);
  else rules.appendRow([key, merchant, category, 0, now, '출금관리에서 확정']);

  const length = Math.max(sheet.getLastRow() - 1, 1);
  const values = sheet.getRange(2, 1, length, 12).getValues();
  let count = 0;
  values.forEach((value) => {
    if (cleanText_(value[7]) !== '지출' || !matchSimpleRule_(cleanText_(value[5]), [{key: key}])) return;
    if (cleanText_(value[8]) !== category) count += 1;
    value[8] = category;
  });
  sheet.getRange(2, 9, length, 1).setValues(values.map((value) => [value[8]]));
  rules.getRange(ruleRow, 4).setValue(count);
  rules.getRange(ruleRow, 5).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
  if (notify) ss.toast('같은·유사 상호 ' + count + '건을 ' + category + '로 분류했습니다.', '가계부');
  return {merchant: merchant, count: count};
}

function applyLegacyRulesToTransactions_(ss, transactions) {
  const rules = readSimpleRules_(ss);
  if (!rules.length) return;
  (transactions || []).forEach((row) => {
    if (row.direction !== '출금' || row.bucket !== '지출') return;
    const matched = matchSimpleRule_(row.description, rules);
    if (!matched) return;
    row.category = matched.category;
    row.note = cleanText_([row.note, '분류규칙: ' + matched.example].filter(Boolean).join(' / '));
  });
}

function showCategoryRules() {
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ensureSimpleRuleSheet_(ss));
}

function requireLedgerSheets_(ss) {
  [LEDGER.outgoingSheet, LEDGER.incomingSheet, LEDGER.sourceSheet].forEach((name) => {
    if (!ss.getSheetByName(name)) throw new Error('필수 시트가 없습니다: ' + name);
  });
}

function validateUploadPayload_(payload) {
  if (!payload || !payload.name || !payload.base64) throw new Error('업로드 데이터가 비어 있습니다.');
  const ext = extension_(payload.name);
  if (LEDGER.extensions.indexOf(ext) < 0) throw new Error('CSV, XLS, XLSX, PDF만 업로드할 수 있습니다.');
}

function validateAnalyzedFile_(file) {
  if (!file || !file.fileId || !Array.isArray(file.transactions)) throw new Error('분석 결과가 손상되었습니다.');
  const driveFile = DriveApp.getFileById(file.fileId);
  return {
    fileId: file.fileId, fileName: driveFile.getName(), fileUrl: driveFile.getUrl(),
    bank: cleanText_(file.bank), warnings: Array.isArray(file.warnings) ? file.warnings.map(cleanText_) : [],
    transactions: file.transactions
  };
}

function validateTransaction_(row, fileId) {
  if (!row || !/^20\d{2}-\d{2}-\d{2}$/.test(row.date || '')) throw new Error('거래 날짜 형식이 잘못되었습니다.');
  if (row.direction !== '출금' && row.direction !== '입금') throw new Error('거래 방향이 잘못되었습니다.');
  const amount = Number(row.amount);
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('거래 금액이 잘못되었습니다.');
  row.fileId = fileId;
  row.amount = amount;
  return row;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function bankFrom_(fileName, content) {
  const name = String(fileName).toLowerCase();
  const text = String(content).toLowerCase();
  if (/우리/.test(name)) return /n페이|네이버|n_pay|npay/.test(name) ? '우리은행 N페이' : '우리은행 일반';
  if (/ok저축|오케이저축/.test(name)) return 'OK저축은행';
  if (/토스|toss/.test(name)) return '토스뱅크';
  if (/ok저축|오케이저축|064-61/.test(text)) return 'OK저축은행';
  if (/우리/.test(text)) return /n페이|네이버|n_pay|npay/.test(name + ' ' + text.slice(0, 1000)) ? '우리은행 N페이' : '우리은행 일반';
  if (/토스뱅크|toss/.test(text)) return '토스뱅크';
  return '은행 확인필요';
}

function bankCode_(bank) {
  if (bank === '토스뱅크') return 'toss';
  if (bank === 'OK저축은행') return 'ok';
  if (bank === '우리은행 N페이') return 'npay';
  if (bank === '우리은행 일반') return 'woori';
  return 'upload';
}

function parseDateTime_(value) {
  const text = cleanText_(value).replace(/\./g, '-').replace(/\//g, '-');
  const match = text.match(/(20\d{2}-\d{1,2}-\d{1,2})(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?/);
  if (!match) return null;
  const date = match[1].split('-').map((part, i) => i ? part.padStart(2, '0') : part).join('-');
  let time = match[2] || '';
  if (time && time.split(':').length === 2) time += ':00';
  return {date: date, time: time};
}

function normalizeDate_(value) {
  const parsed = parseDateTime_(value);
  return parsed ? parsed.date : cleanText_(value);
}

function normalizeHeader_(value) {
  return cleanText_(value).replace(/[\s_·()\[\]\/]/g, '');
}

function parseAmount_(value) {
  if (value === null || value === undefined || value === '' || value === '-') return 0;
  const number = Number(String(value).replace(/[^0-9.-]/g, ''));
  return Number.isFinite(number) ? Math.abs(number) : 0;
}

function parseSignedAmount_(value) {
  if (value === null || value === undefined || value === '' || value === '-') return 0;
  const number = Number(String(value).replace(/[^0-9.-]/g, ''));
  return Number.isFinite(number) ? number : 0;
}

function parseOwnerLabels_(value) {
  return String(value || '').split(/[,\n]/).map(cleanText_).filter((name) => name.length >= 2).slice(0, 10);
}

function extension_(name) {
  const match = String(name).toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : '';
}

function mimeFromExtension_(ext) {
  return {
    csv: 'text/csv', xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pdf: 'application/pdf'
  }[ext] || 'application/octet-stream';
}

function monthFromTitle_(title) {
  const match = String(title).match(/(20\d{2})년\s*(\d{1,2})월/);
  return match ? match[1] + '-' + match[2].padStart(2, '0') : Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
}

function toSheetDate_(iso) {
  return new Date(iso + 'T00:00:00+09:00');
}

function sha256_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map((byte) => (byte + 256).toString(16).slice(-2)).join('');
}

function safeCellText_(value) {
  const text = cleanText_(value).slice(0, 5000);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function safeFileName_(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, '_').slice(0, 180);
}

function cleanText_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/\s+/g, ' ').trim();
}

function cell_(row, index) {
  return index < 0 ? '' : row[index];
}

function sum_(numbers) {
  return numbers.reduce((total, value) => total + Number(value || 0), 0);
}
