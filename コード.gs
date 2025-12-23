// ============================================================
// ミルシート一覧アプリ - バックエンド (GAS)
// ============================================================

const STORAGE_KEY = 'MILLSHEET_RECORDS';
const FOLDER_NAME = 'ミルシート添付ファイル';
const ADMIN_EMAILS = ['konno@race-tech.co.jp']; // 管理者メールアドレス

// ------------------------------------------------------------
// HTML表示
// ------------------------------------------------------------
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = Session.getActiveUser().getEmail();
  template.userRole = getUserRole();
  
  return template.evaluate()
    .setTitle('ミルシート一覧アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ------------------------------------------------------------
// ユーザー権限管理
// ------------------------------------------------------------
function getUserRole() {
  const email = Session.getActiveUser().getEmail();
  if (ADMIN_EMAILS.includes(email)) return 'admin';
  // 実運用では User Properties や Spreadsheet で管理
  return 'editor'; // 'admin', 'editor', 'viewer'
}

function getUserInfo() {
  return {
    email: Session.getActiveUser().getEmail(),
    role: getUserRole()
  };
}

// ------------------------------------------------------------
// データストレージ（Drive JSON方式）
// ------------------------------------------------------------
function getStorageFile() {
  const folder = getStorageFolder();
  const files = folder.getFilesByName('millsheet_data.json');
  
  if (files.hasNext()) {
    return files.next();
  } else {
    const file = folder.createFile('millsheet_data.json', '[]', MimeType.PLAIN_TEXT);
    return file;
  }
}

function getStorageFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(FOLDER_NAME);
  }
}

function loadRecords() {
  try {
    const file = getStorageFile();
    const content = file.getBlob().getDataAsString();
    return JSON.parse(content || '[]');
  } catch (e) {
    Logger.log('Load error: ' + e);
    return [];
  }
}

function saveRecords(records) {
  try {
    const file = getStorageFile();
    file.setContent(JSON.stringify(records, null, 2));
    return { success: true };
  } catch (e) {
    Logger.log('Save error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ------------------------------------------------------------
// CRUD操作
// ------------------------------------------------------------
function getAllRecords(filter) {
  let records = loadRecords();
  
  // 削除済みを除外
  records = records.filter(r => !r.deleted);
  
  // フィルタリング
  if (filter) {
    if (filter.keyword) {
      const kw = filter.keyword.toLowerCase();
      records = records.filter(r => 
        (r.deliveryNo && r.deliveryNo.toLowerCase().includes(kw)) ||
        (r.contractNo && r.contractNo.toLowerCase().includes(kw)) ||
        (r.note && r.note.toLowerCase().includes(kw))
      );
    }
    
    if (filter.dateFrom) {
      records = records.filter(r => r.arrivalDate >= filter.dateFrom);
    }
    
    if (filter.dateTo) {
      records = records.filter(r => r.arrivalDate <= filter.dateTo);
    }
    
    if (filter.diameter) {
      records = records.filter(r => r.diameter === filter.diameter);
    }
    
    if (filter.submitStatus) {
      records = records.filter(r => r.submitStatus === filter.submitStatus);
    }
    
    if (filter.hasAttachment !== undefined) {
      records = records.filter(r => 
        filter.hasAttachment ? (r.attachUrl || r.attachFileId) : !(r.attachUrl || r.attachFileId)
      );
    }
  }
  
  // ソート（到着日降順→納品書番号降順）
  records.sort((a, b) => {
    if (b.arrivalDate !== a.arrivalDate) {
      return b.arrivalDate.localeCompare(a.arrivalDate);
    }
    return b.deliveryNo.localeCompare(a.deliveryNo);
  });
  
  return records;
}

function getRecordById(id) {
  const records = loadRecords();
  return records.find(r => r.id === id && !r.deleted);
}

function createRecord(data) {
  const user = Session.getActiveUser().getEmail();
  const records = loadRecords();
  
  // 納品書番号重複チェック
  if (records.some(r => r.deliveryNo === data.deliveryNo && !r.deleted)) {
    return { success: false, error: '納品書番号が重複しています' };
  }
  
  const now = new Date().toISOString();
  const record = {
    id: Utilities.getUuid(),
    arrivalDate: data.arrivalDate,
    deliveryNo: data.deliveryNo,
    contractNo: data.contractNo,
    diameter: data.diameter,
    strength: data.strength || '',
    maker: data.maker || '',
    quantity: parseFloat(data.quantity),
    note: data.note || '',
    attachType: data.attachType || '',
    attachUrl: data.attachUrl || '',
    attachFileId: data.attachFileId || '',
    millSentDate: data.millSentDate || '',
    tagSentDate: data.tagSentDate || '',
    pdfSentDate: data.pdfSentDate || '',
    netSentDate: data.netSentDate || '',
    submitStatus: calculateSubmitStatus(data),
    createdAt: now,
    createdBy: user,
    updatedAt: now,
    updatedBy: user,
    deleted: false
  };
  
  records.push(record);
  const result = saveRecords(records);
  
  if (result.success) {
    return { success: true, record: record };
  } else {
    return result;
  }
}

function updateRecord(id, data) {
  const user = Session.getActiveUser().getEmail();
  const records = loadRecords();
  const index = records.findIndex(r => r.id === id);
  
  if (index === -1) {
    return { success: false, error: 'レコードが見つかりません' };
  }
  
  // 納品書番号重複チェック（自分以外）
  if (records.some(r => r.id !== id && r.deliveryNo === data.deliveryNo && !r.deleted)) {
    return { success: false, error: '納品書番号が重複しています' };
  }
  
  const record = records[index];
  record.arrivalDate = data.arrivalDate;
  record.deliveryNo = data.deliveryNo;
  record.contractNo = data.contractNo;
  record.diameter = data.diameter;
  record.strength = data.strength || '';
  record.maker = data.maker || '';
  record.quantity = parseFloat(data.quantity);
  record.note = data.note || '';
  record.attachType = data.attachType || '';
  record.attachUrl = data.attachUrl || '';
  record.attachFileId = data.attachFileId || '';
  record.millSentDate = data.millSentDate || '';
  record.tagSentDate = data.tagSentDate || '';
  record.pdfSentDate = data.pdfSentDate || '';
  record.netSentDate = data.netSentDate || '';
  record.submitStatus = calculateSubmitStatus(data);
  record.updatedAt = new Date().toISOString();
  record.updatedBy = user;
  
  const result = saveRecords(records);
  
  if (result.success) {
    return { success: true, record: record };
  } else {
    return result;
  }
}

function deleteRecord(id) {
  const role = getUserRole();
  if (role !== 'admin') {
    return { success: false, error: '削除権限がありません' };
  }
  
  const records = loadRecords();
  const record = records.find(r => r.id === id);
  
  if (!record) {
    return { success: false, error: 'レコードが見つかりません' };
  }
  
  // 論理削除
  record.deleted = true;
  record.updatedAt = new Date().toISOString();
  record.updatedBy = Session.getActiveUser().getEmail();
  
  return saveRecords(records);
}

function markAsSubmitted(id) {
  const records = loadRecords();
  const record = records.find(r => r.id === id);
  
  if (!record) {
    return { success: false, error: 'レコードが見つかりません' };
  }
  
  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  
  // PDF送付日を代表として設定
  if (!record.pdfSentDate) {
    record.pdfSentDate = today;
  }
  
  record.submitStatus = 'submitted';
  record.updatedAt = new Date().toISOString();
  record.updatedBy = Session.getActiveUser().getEmail();
  
  const result = saveRecords(records);
  
  if (result.success) {
    return { success: true, record: record };
  } else {
    return result;
  }
}

function markAsUnsubmitted(id) {
  const role = getUserRole();
  if (role !== 'admin') {
    return { success: false, error: '権限がありません' };
  }
  
  const records = loadRecords();
  const record = records.find(r => r.id === id);
  
  if (!record) {
    return { success: false, error: 'レコードが見つかりません' };
  }
  
  record.submitStatus = 'unsubmitted';
  record.updatedAt = new Date().toISOString();
  record.updatedBy = Session.getActiveUser().getEmail();
  
  return saveRecords(records);
}

// ------------------------------------------------------------
// ヘルパー関数
// ------------------------------------------------------------
function calculateSubmitStatus(data) {
  if (data.millSentDate || data.pdfSentDate || data.netSentDate) {
    return 'submitted';
  }
  return 'unsubmitted';
}

// ------------------------------------------------------------
// ファイルアップロード
// ------------------------------------------------------------
function uploadAttachment(base64Data, fileName, mimeType) {
  try {
    const folder = getStorageFolder();
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      fileName
    );
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      fileId: file.getId(),
      url: file.getUrl()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ------------------------------------------------------------
// CSV出力
// ------------------------------------------------------------
function exportToCSV(filter) {
  const records = getAllRecords(filter);
  
  const headers = [
    'ID', '到着日', '納品書番号', '契約番号', '径', '強度', 'メーカー',
    '数量(t)', 'メモ', 'ミル送付日', 'タグ送付日', 'PDF送付日', 'NET登録日',
    '提出ステータス', '添付URL', '作成日時', '作成者', '更新日時', '更新者'
  ];
  
  let csv = headers.join(',') + '\n';
  
  records.forEach(r => {
    const row = [
      r.id,
      r.arrivalDate,
      r.deliveryNo,
      r.contractNo,
      r.diameter,
      r.strength,
      r.maker,
      r.quantity,
      `"${(r.note || '').replace(/"/g, '""')}"`,
      r.millSentDate,
      r.tagSentDate,
      r.pdfSentDate,
      r.netSentDate,
      r.submitStatus === 'submitted' ? '提出済' : '未提出',
      r.attachUrl,
      r.createdAt,
      r.createdBy,
      r.updatedAt,
      r.updatedBy
    ];
    csv += row.join(',') + '\n';
  });
  
  return csv;
}

function exportMonthlySummary(year, month) {
  const records = loadRecords().filter(r => !r.deleted);
  
  const target = records.filter(r => {
    const arrivalMonth = r.arrivalDate.substring(0, 7); // yyyy-mm
    return arrivalMonth === `${year}-${String(month).padStart(2, '0')}`;
  });
  
  // 径別集計
  const summary = {};
  
  target.forEach(r => {
    if (!summary[r.diameter]) {
      summary[r.diameter] = {
        submitted: 0,
        unsubmitted: 0
      };
    }
    
    if (r.submitStatus === 'submitted') {
      summary[r.diameter].submitted += r.quantity;
    } else {
      summary[r.diameter].unsubmitted += r.quantity;
    }
  });
  
  let csv = '径,提出済(t),未提出(t),合計(t)\n';
  
  Object.keys(summary).sort().forEach(dia => {
    const s = summary[dia];
    const total = s.submitted + s.unsubmitted;
    csv += `${dia},${s.submitted.toFixed(3)},${s.unsubmitted.toFixed(3)},${total.toFixed(3)}\n`;
  });
  
  const totalSubmitted = Object.values(summary).reduce((sum, s) => sum + s.submitted, 0);
  const totalUnsubmitted = Object.values(summary).reduce((sum, s) => sum + s.unsubmitted, 0);
  const grandTotal = totalSubmitted + totalUnsubmitted;
  
  csv += `合計,${totalSubmitted.toFixed(3)},${totalUnsubmitted.toFixed(3)},${grandTotal.toFixed(3)}\n`;
  
  return csv;
}

// ------------------------------------------------------------
// マスタデータ
// ------------------------------------------------------------
function getDiameterOptions() {
  return [
    'D10', 'D13', 'D16', 'D19', 'D22', 'D25', 'D29', 'D32', 'D35', 'D38', 'D41'
  ];
}

function getStrengthOptions() {
  return ['SD295A', 'SD295B', 'SD345', 'SD390'];
}
