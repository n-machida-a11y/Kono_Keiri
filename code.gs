/**
 * @OnlyCurrentDoc
 */

// スプレッドシートのIDとシート名を指定
const SPREADSHEET_ID = '1uU5wjvnKXklKG9GhG1yHq55QqZkCrjlekKvRUU4HFOI'; // スプレッドシートのID

// シート名を定数として定義
const RECEIPTS_SHEET_NAME = '領収書';
const DETAILS_SHEET_NAME = '明細';
const MASTER_SHEET_NAME = '勘定科目マスタ';
const SETTINGS_SHEET_NAME = '設定';
const USER_MASTER_SHEET_NAME = '使用者マスタ';
const FOLDER_NAME = '領収書';
const LEARNING_DATA_SHEET_NAME = '学習データ'; // 品目キーワードマスタとして使用

/**
 * Webアプリにアクセスした際にHTMLを返す
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('領収書登録アプリ');
}

/**
 * index.htmlから他のgsファイルやcssファイルを読み込めるようにする
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 指定された名前のフォルダを取得または作成する
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/**
 * フォームからデータを受け取り、AI処理、グルーピング、ファイル保存までを一貫して行う
 */
function startUploadAndProcess(formObject) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { status: 'error', message: '現在、他の処理を実行中です。しばらくしてから再度お試しください。' };
  }

  try {
    const historicalData = getHistoricalData(null);
    const { extractedReceipts, usageMetadata } = extractDataFromFile(formObject.fileData, formObject.mimeType, historicalData);

    if (extractedReceipts.length === 0) {
      return { status: 'error', message: 'AIが領収書を検出できませんでした。' };
    }

    const decodedFile = Utilities.base64Decode(formObject.fileData, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decodedFile, formObject.mimeType, formObject.fileName);
    const folder = getOrCreateFolder(FOLDER_NAME);
    const savedFile = folder.createFile(blob);
    
    // ファイルの共有設定を変更し、プレビューできるようにする
    savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileId = savedFile.getId();
    const fileName = savedFile.getName();

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const masterSheet = spreadsheet.getSheetByName(LEARNING_DATA_SHEET_NAME);
    let masterData = [];
    if (masterSheet && masterSheet.getLastRow() > 1) {
      masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
    }
    const categoryMap = new Map(masterData.map(row => [row[0], row[1]]));

    const placeholderRows = [];
    const timestamp = new Date();

    const initialDataArray = extractedReceipts.map((receipt, index) => {
      const groupedDetails = groupDetailsByCategory(receipt, categoryMap);
      const receiptId = new Date().getTime() + index;

      // ★★★ 修正: スプレッドシートの列構成に合わせてプレースホルダを作成
      // A:登録ID, B:登録日時, C:申請日, D:使用者, E:支払先, F:合計金額, G:メモ, 
      // H:ファイル名, I:ファイルID, J:入力トークン, K:出力トークン, L:インボイス番号(あり/なし)
      placeholderRows.push([
        receiptId, timestamp, 
        formObject.useDate || "処理中...", // C: 申請日 (yyyy-MM-dd)
        formObject.user || "", // D: 使用者
        "", // E: 支払先 (AI待ち)
        "", // F: 合計金額 (AI待ち)
        "", // G: メモ
        fileName, // H: ファイル名
        fileId, // I: ファイルID
        (usageMetadata.promptTokenCount / extractedReceipts.length) || 0, // J: 入力トークン
        (usageMetadata.candidatesTokenCount / extractedReceipts.length) || 0, // K: 出力トークン
        receipt.has_invoice ? 'あり' : 'なし' // L: インボイス番号
      ]);

      return {
        parentData: {
          '登録ID': receiptId,
          fileId: fileId,
          mimeType: formObject.mimeType, // mimeTypeも返す（プレビューに必要）
          // ★★★ 修正: AI抽出(use_date)よりフォーム入力(formObject.useDate)を優先
          // ★★★ 修正: 日付フォーマットを yyyy/MM/dd に統一
          useDate: formatDateToSlash(formObject.useDate || receipt.use_date), 
          user: formObject.user || '', 
          storeName: receipt.store_name || '',
          totalAmount: receipt.total_amount || 0,
          memo: '', 
          // ★★★ 修正: AIが抽出した `has_invoice` (true/false) を `hasInvoice` として追加
          hasInvoice: receipt.has_invoice || false
        },
        detailsData: groupedDetails
      };
    });
    
    if (placeholderRows.length > 0) {
      const receiptSheet = spreadsheet.getSheetByName(RECEIPTS_SHEET_NAME);
      // ★★★ 修正: プレースホルダ行の列数を 12 に変更 (L列まで)
      receiptSheet.getRange(receiptSheet.getLastRow() + 1, 1, placeholderRows.length, 12).setValues(placeholderRows);
    }

    return { status: 'success', data: initialDataArray };

  } catch (e) {
    console.error(`startUploadAndProcessでエラー: ${e.message} (Stack: ${e.stack})`);
    return { status: 'error', message: `サーバー側でエラーが発生しました: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ★★★ 新規関数: 日付文字列を yyyy/MM/dd に変換
 */
function formatDateToSlash(dateString) {
    if (!dateString) return '';
    try {
        // yyyy-MM-dd または yyyy/MM/dd の両方に対応
        const date = new Date(dateString.replace(/-/g, '/'));
        if (isNaN(date.getTime())) return dateString; // 不正な日付はそのまま返す
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    } catch(e) {
        return dateString; // 変換失敗時は元の日付文字列を返す
    }
}


/**
 * AIが抽出した品目を、勘定科目ごとに集計する関数
 */
function groupDetailsByCategory(extractedData, categoryMap) {
    const findCategory = (itemName) => {
        if (!itemName) return null;
        for (const [keyword, category] of categoryMap.entries()) {
            // キーワードが空でないことを確認
            if (keyword && itemName.includes(keyword)) {
                return category;
            }
        }
        return null;
    };

    const grouped = {};
    const items = extractedData.items || [];
    const receiptCategoryGuess = extractedData.category ? extractedData.category.trim() : '';

    
    if (items.length > 0) {
        items.forEach(item => {
            let itemCategory = findCategory(item.name);
            let category = itemCategory || receiptCategoryGuess;
            if (!category) {
              category = '雑費';
            }

            if (!grouped[category]) {
                grouped[category] = {
                    category: category,
                    item: [],
                    totalAmount: 0,
                    subtotal: 0,
                    tax: 0,
                    client: extractedData.client || '',
                    participants: extractedData.participants || '',
                    memo: ''
                };
            }
            
            if (item.name) {
              grouped[category].item.push(item.name);
            }
            grouped[category].totalAmount += Number(item.price || item.total_price || 0);
            grouped[category].subtotal += Number(item.subtotal || 0);
            grouped[category].tax += Number(item.tax || 0);
        });
    } else {
        const category = receiptCategoryGuess || '雑費';
        grouped[category] = {
            category: category,
            item: [],
            totalAmount: extractedData.total_amount || 0,
            subtotal: extractedData.subtotal || 0,
            tax: extractedData.tax || 0,
            client: extractedData.client || '',
            participants: extractedData.participants || '',
            memo: ''
        };
    }
    
    return Object.values(grouped).map(group => {
        group.item = group.item.join(', ');
        if (group.totalAmount === 0 && Object.keys(grouped).length === 1) {
          group.totalAmount = extractedData.total_amount || 0;
          group.subtotal = extractedData.subtotal || 0;
          group.tax = extractedData.tax || 0;
        }
        return group;
    });
}


/**
 * ユーザーが編集した最終的なデータでスプレッドシートを更新する
 */
function updateFinalReceiptRow(data) {
  const { parentData, detailsData } = data;
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const receiptSheet = spreadsheet.getSheetByName(RECEIPTS_SHEET_NAME);
    const detailSheet = spreadsheet.getSheetByName(DETAILS_SHEET_NAME);
    
    // 学習データシートの準備
    const learningSheet = spreadsheet.getSheetByName(LEARNING_DATA_SHEET_NAME);
    let existingKeywords = [];
    if (learningSheet.getLastRow() > 1) {
       existingKeywords = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, 1).getValues().flat().filter(String);
    }
    
    const receiptData = receiptSheet.getDataRange().getValues();
    const rowIndex = receiptData.findIndex(row => row[0].toString() === parentData['登録ID'].toString());

    if (rowIndex === -1) {
      throw new Error(`更新失敗: ID ${parentData['登録ID']} が見つかりません。`);
    }

    const totalAmount = detailsData.reduce((sum, detail) => sum + Number(detail.totalAmount || 0), 0);

    // C列(申請日)〜G列(メモ)までをセット
    receiptSheet.getRange(rowIndex + 1, 3, 1, 5).setValues([[
      parentData.useDate, // ★★★ 修正: ここに来る日付は yyyy-MM-dd (HTMLの<input type="date">の値)
      parentData.user,
      parentData.storeName,
      totalAmount,
      parentData.memo
    ]]);
    
    // ★★★ 修正: L列(12列目)にインボイス情報（あり/なし）をセット
    const invoiceStatus = parentData.hasInvoice ? 'あり' : 'なし';
    receiptSheet.getRange(rowIndex + 1, 12, 1, 1).setValue(invoiceStatus);


    const detailData = detailSheet.getDataRange().getValues();
    for (let i = detailData.length - 1; i > 0; i--) {
      if (detailData[i][1].toString() === parentData['登録ID'].toString()) {
        detailSheet.deleteRow(i + 1);
      }
    }

    const detailsToAppend = [];
    const learningDataToAppend = [];

    detailsData.forEach(detail => {
        detailsToAppend.push([
            new Date().getTime() + Math.random(),
            parentData['登録ID'],
            detail.category,
            detail.item,
            detail.totalAmount,
            detail.client,
            detail.participants,
            detail.subtotal,
            detail.tax,
            detail.memo
        ]);

        // 学習データを蓄積する処理
        if (detail.item && detail.category) {
            const keywords = detail.item.split(',').map(k => k.trim()).filter(String);
            keywords.forEach(keyword => {
                // まだ学習データシートに存在しないキーワードの場合
                if (!existingKeywords.includes(keyword)) {
                    learningDataToAppend.push([keyword, detail.category]);
                    existingKeywords.push(keyword); // 重複追加を防ぐ
                }
            });
        }
    });

    if (detailsToAppend.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailsToAppend.length, detailsToAppend[0].length).setValues(detailsToAppend);
    }
    
    // 学習データシートに新しいキーワードを追加
    if (learningDataToAppend.length > 0) {
      learningSheet.getRange(learningSheet.getLastRow() + 1, 1, learningDataToAppend.length, learningDataToAppend[0].length).setValues(learningDataToAppend);
    }
   
    return { status: 'success', message: '領収書の登録が完了しました。' };
  } catch (e) {
    console.error(e);
    return { status: 'error', message: '登録の保存中にエラーが発生しました: ' + e.message };
  }
}

/**
 * Gemini APIを呼び出して、画像やPDFから構造化されたデータを抽出する
 */
function extractDataFromFile(base64Data, mimeType, historicalData) {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!GEMINI_API_KEY) throw new Error("Gemini APIキーが設定されていません。");

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  let categories = [];
  if (masterSheet && masterSheet.getLastRow() > 1) {
    categories = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues().flat().filter(String);
  }
  if (categories.length === 0) {
    categories = ["交通費", "会議費", "接待交際費", "少額接待交際費", "消耗品費", "通信費", "雑費", "その他"];
  }

  const settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  let promptTemplate = settingsSheet ? settingsSheet.getRange('B1').getValue() : '';
  if (!promptTemplate) throw new Error("「設定」シートからプロンプトを取得できませんでした。");

  let prompt = promptTemplate.replace('{categories}', categories.join(', '));
  
  // 過去の履歴データ(historicalData)もプロンプトに追加する
  if (historicalData) {
    prompt += "\n\n以下は過去の登録履歴です。特に支払先が同じ場合は、これを参考にして勘定科目を推測してください。\n" + historicalData;
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64Data } }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resultText = response.getContentText();
  if (!resultText) throw new Error("AIからのレスポンスが空です。");
  
  const result = JSON.parse(resultText);

  if (result.candidates && result.candidates[0].content) {
    const jsonString = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
    try {
      let jsonData = JSON.parse(jsonString);
      const extractedReceipts = Array.isArray(jsonData) ? jsonData : [jsonData];
      const usageMetadata = result.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
      return { extractedReceipts, usageMetadata };
    } catch (e) {
      console.error("AIが返したJSONの解析に失敗:", jsonString);
      throw new Error("AIが有効な形式で応答しませんでした。");
    }
  } else {
    const errorMessage = result.error ? result.error.message : "不明なエラーです。";
    console.error("AIからのデータ抽出に失敗:", resultText);
    throw new Error("AIからのデータ抽出に失敗: " + errorMessage);
  }
}

/**
 * 過去の登録履歴から参考データを文字列として取得する
 */
function getHistoricalData(storeNameFilter) {
  try {
    const receiptSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RECEIPTS_SHEET_NAME);
    const detailSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DETAILS_SHEET_NAME);
    if (receiptSheet.getLastRow() < 2 || detailSheet.getLastRow() < 2) return "";
    
    // 履歴データは「領収書」シートの支払先(5列目)と「明細」シートの勘定科目(3列目)を結合して作成する
    const receiptData = receiptSheet.getRange(2, 1, receiptSheet.getLastRow() - 1, 5).getValues();
    const detailData = detailSheet.getRange(2, 1, detailSheet.getLastRow() - 1, 3).getValues();

    const receiptMap = new Map(receiptData.map(r => [r[0].toString(), r[4]])); // Map<receiptId, storeName>

    const merged = detailData.map(d => {
        const store = receiptMap.get(d[1].toString());
        return {
            store: store || null,
            category: d[2]
        }
    }).filter(m => m.store && m.category); // 支払先と勘定科目が両方あるものだけ

    let recentData = merged;
    if (storeNameFilter) {
      recentData = merged.filter(m => m.store === storeNameFilter);
    }
    
    // 直近5件の履歴を文字列化して返す
    return recentData.slice(-5).map(row => `- 支払先: ${row.store}, 勘定科目: ${row.category}`).join("\n");
  } catch (e) {
    console.error("履歴データの取得に失敗:", e);
    return "";
  }
}


/**
 * スプレッドシートから登録済みの領収書データを取得する
 */
function getReceipts(startDate = null, endDate = null, storeName = null) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RECEIPTS_SHEET_NAME);
    if (sheet.getLastRow() < 2) return [];

    // ★★★ 修正: ご提示の列構成 (L列まで = 12列) を読み込む
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

    if (startDate) data = data.filter(row => row[2] && new Date(row[2]) >= new Date(startDate));
    if (endDate) data = data.filter(row => row[2] && new Date(row[2]) <= new Date(endDate));
    if (storeName) data = data.filter(row => row[4] && row[4].toString().includes(storeName));
    
    return data.map(row => ({
      receiptId: row[0],
      // C列 (インデックス2)
      // ★★★ 修正: 日付フォーマットを yyyy/MM/dd に変更
      useDate: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy/MM/dd') : (row[2] ? formatDateToSlash(row[2]) : ''),
      user: row[3], // D列
      storeName: row[4], // E列
      totalAmount: row[5], // F列
      // ★★★ 修正: I列(インデックス8)からFileID、L列(インデックス11)からインボイス情報を取得
      fileId: row[8],
      hasInvoice: row[11] === 'あり' ? true : false // L列
    }));
  } catch (e) {
    console.error(e);
    return [];
  }
}

/**
 * IDから領収書の詳細データを取得する
 */
function getReceiptDetails(receiptId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const receiptSheet = spreadsheet.getSheetByName(RECEIPTS_SHEET_NAME);
    const detailSheet = spreadsheet.getSheetByName(DETAILS_SHEET_NAME);

    const receiptDataRange = receiptSheet.getDataRange().getValues();
    const rowIndex = receiptDataRange.findIndex(row => row[0].toString() === receiptId.toString());
    if (rowIndex === -1) throw new Error(`ID ${receiptId} が見つかりません。`);
    const parentRow = receiptDataRange[rowIndex];
    
    // ★★★ 修正: FileIDはI列(インデックス8)
    const fileId = parentRow[8]; 
    let mimeType = null; 

    if (fileId) {
      try {
        // 履歴からの編集時もプレビューできるように共有設定を変更する
        const file = DriveApp.getFileById(fileId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        mimeType = file.getMimeType();
      } catch (e) {
        console.error(`ファイル(ID: ${fileId})の取得または共有設定の変更に失敗: ${e.message}`);
        // ファイルが見つからない場合もエラーにせず、mimeType = null のまま続行
      }
    }

    const parentData = {
      '登録ID': parentRow[0], // A列
      // ★★★ 修正: 日付フォーマットを yyyy/MM/dd に変更
      useDate: parentRow[2] instanceof Date ? Utilities.formatDate(parentRow[2], Session.getScriptTimeZone(), 'yyyy/MM/dd') : (parentRow[2] ? formatDateToSlash(parentRow[2]) : ''), // C列
      user: parentRow[3], // D列
      storeName: parentRow[4], // E列
      totalAmount: parentRow[5], // F列
      memo: parentRow[6], // G列
      fileName: parentRow[7], // H列
      fileId: fileId, // I列
      // ★★★ 修正: インボイス情報はL列(インデックス11)から
      hasInvoice: parentRow[11] === 'あり' ? true : false,
      mimeType: mimeType
    };
    
    const allDetails = detailSheet.getDataRange().getValues();
    const detailsData = allDetails.filter((row, index) => {
        return index > 0 && row[1].toString() === receiptId.toString();
    }).map(row => ({
        category: row[2],
        item: row[3],
        totalAmount: row[4],
        client: row[5],
        participants: row[6],
        subtotal: row[7],
        tax: row[8],
        memo: row[9]
    }));

    return { status: 'success', data: { parentData, detailsData } };
  } catch (e) {
    console.error("詳細データの取得に失敗:", e);
    return { status: 'error', message: '詳細データの取得に失敗しました: ' + e.message };
  }
}


/**
 * 絞り込んだ履歴データをCSV形式で返す
 */
function getReceiptsAsCsv(startDate = null, endDate = null, storeName = null) {
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const receiptSheet = spreadsheet.getSheetByName(RECEIPTS_SHEET_NAME);
        const detailSheet = spreadsheet.getSheetByName(DETAILS_SHEET_NAME);

        if (receiptSheet.getLastRow() < 2) return "データがありません";
        
        // ★★★ 修正: 12列 (L列、インボイス番号) まで読み込む
        let receiptData = receiptSheet.getRange(2, 1, receiptSheet.getLastRow() - 1, 12).getValues();
        const detailData = detailSheet.getRange(2, 1, detailSheet.getLastRow() - 1, 10).getValues();

        if (startDate) receiptData = receiptData.filter(row => row[2] && new Date(row[2]) >= new Date(startDate));
        if (endDate) receiptData = receiptData.filter(row => row[2] && new Date(row[2]) <= new Date(endDate));
        if (storeName) receiptData = receiptData.filter(row => row[4] && row[4].toString().includes(storeName));

        // ★★★ 修正: CSVヘッダーを「申請日」に変更、および「インボイス(あり/なし)」を追加
        const header = ['登録ID', '登録日時', '申請日', '使用者', '支払先', '合計金額', 'メモ', 'ファイル名', 'ファイルID', 'インボイス(あり/なし)', '勘定科目', '項目', '明細金額', '取引先', '参加人数', '明細_税抜合計', '明細_消費税', '明細_メモ'];
        let csvContent = header.join(',') + '\n';
        
        const filteredReceiptIds = receiptData.map(r => r[0].toString());

        detailData.forEach(detailRow => {
            const receiptId = detailRow[1].toString();
            if(filteredReceiptIds.includes(receiptId)){
                const receiptRow = receiptData.find(r => r[0].toString() === receiptId);
                const csvRowData = [
                    receiptRow[0], // 登録ID
                    receiptRow[1] instanceof Date ? Utilities.formatDate(receiptRow[1], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss') : receiptRow[1], // 登録日時
                    // ★★★ 修正: 申請日 (yyyy/MM/dd)
                    receiptRow[2] instanceof Date ? Utilities.formatDate(receiptRow[2], Session.getScriptTimeZone(), 'yyyy/MM/dd') : (receiptRow[2] ? formatDateToSlash(receiptRow[2]) : ''), // 申請日
                    receiptRow[3], // 使用者
                    receiptRow[4], // 支払先
                    receiptRow[5], // 合計金額
                    receiptRow[6], // メモ
                    receiptRow[7], // ファイル名
                    receiptRow[8], // ファイルID
                    // ★★★ 修正: インボイス情報(L列=インデックス11)を追加
                    receiptRow[11], // インボイス(あり/なし)
                    detailRow[2], // 勘定科目
                    detailRow[3], // 項目
                    detailRow[4], // 明細金額
                    detailRow[5], // 取引先
                    detailRow[6], // 参加人数
                    detailRow[7], // 明細_税抜合計
                    detailRow[8], // 明細_消費税
                    detailRow[9]  // 明細_メモ
                ];
                
                const csvRow = csvRowData.map(cell => {
                    const value = cell != null ? cell.toString() : "";
                    if (value.includes('"') || value.includes(',')) {
                        return `"${value.replace("/", '""')}"`;
                    }
                    return value;
                });
                csvContent += csvRow.join(',') + '\n';
            }
        });

        return csvContent;
    } catch (e) {
        console.error(e);
        return "エラー: CSVの生成に失敗しました。";
    }
}

/**
 * 使用者マスタから使用者リストを取得する
 */
function getUsers() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(String);
  } catch (e) {
    console.error("使用者マスタの読み込みに失敗: ", e);
    return [];
  }
}


/**
 * 勘定科目マスタからリストを取得する
 */
function getCategories() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return ["交通費", "会議費", "接待交際費", "少額接待交際費", "消耗品費", "通信費", "雑費", "その他"];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(String);
  } catch (e) {
    console.error("勘定科目マスタの読み込みに失敗: ", e);
    return [];
  }
}