// ===============================================
// BOOTH売上通知メール自動収集 - Google Apps Script版
// 販売者向け: 商品が売れた時の売上データを自動でスプレッドシートに記録
// ===============================================

// ===== 🔧 設定項目（最上位） =====
const CONFIG = {
  // 必須設定
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE', // スプレッドシートのIDを設定

  // 基本設定
  SHEET_NAME: 'BOOTH売上履歴',
  SEARCH_LABEL: 'booth-processed', // 処理済みメールのラベル名
  PROCESSING_LABEL: 'booth-processing', // 処理中ラベル名（タイムアウト対策）

  // 処理制御
  MAX_SCAN_COUNT: 200, // 最大スキャン件数
  ENABLE_VARIANT_EXTRACTION: true, // 版情報抽出を有効にする

  // パフォーマンス設定
  BATCH_SIZE: 100, // ラベル追加のバッチサイズ
  LOG_INTERVAL: 10, // ログ出力間隔（件数）
  PROGRESS_INTERVAL: 100, // 進捗表示間隔（スレッド数）
  TIME_LIMIT_MINUTES: 4.5, // 実行時間制限（分）
  TIME_CHECK_INTERVAL: 50, // 時間チェック間隔（スレッド数）

  // 表機能設定（正式API準拠）
  DISABLE_TABLE_FEATURE: false,        // 表機能を無効化する場合はtrue
  ENABLE_BASIC_FORMATTING: true,       // フォールバックとして基本書式を有効にする
  TABLE_SETUP_DELAY: 500,              // 初回待機（任意）
  TABLE_FILTER_FULL_COLUMNS: true,     // A:E をシート全行に適用し自動追従
  BANDING_THEME: 'LIGHT_GREY'          // 交互色テーマ
};

// ===== 📋 メイン実行関数 =====

/**
 * メイン処理関数 - BOOTHの売上通知メールを検索・処理
 * トリガーで定期実行またはスクリプトエディタから手動実行
 */
function processBOOTHSalesEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    addLogFromServer('WARN','同時実行スキップ');
    return { processed: 0, skipped: true, message: 'locked' };
  }

  try {
    // 準備確認
    ensureConfigLoaded();
    if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      addLogFromServer('WARN','未セットアップのためスキップ',{reason:'SPREADSHEET_ID missing'});
      return { processed: 0, skipped: true, message: 'SPREADSHEET_ID not set' };
    }

    // シート確保（存在しなければ作成）。失敗時はスキップ。
    let sheet;
    try {
      sheet = getOrCreateSheet();
    } catch (e) {
      addLogFromServer('ERROR','シート初期化失敗',{error:String(e)});
      return { processed: 0, skipped: true, message: 'sheet init failed' };
    }

    addLogFromServer('INFO','処理開始',{fn:'processBOOTHSalesEmails'});

    // 検索
    const query = 'from:noreply@booth.pm (subject:商品が購入されました OR subject:ご注文が確定しました) -label:' + CONFIG.SEARCH_LABEL + ' -label:' + CONFIG.PROCESSING_LABEL;
    let allThreads = [];
    try {
      allThreads = GmailApp.search(query);
    } catch (e) {
      addLogFromServer('ERROR','Gmail検索失敗',{error:String(e)});
      return { processed: 0, skipped: true, message: 'gmail search failed' };
    }

    const threadsToProcess = allThreads.slice(0, CONFIG.MAX_SCAN_COUNT || 200);
    addLogFromServer('INFO','検索完了',{found: allThreads.length, target: threadsToProcess.length});
    if (threadsToProcess.length === 0) {
      addLogFromServer('INFO','新規なし');
      return { processed: 0, message: 'no new threads' };
    }

    const startTime = new Date();
    let timeoutOccurred = false;
    const existingOrderNumbers = getExistingOrderNumbers(sheet);

    let processedCount = 0;
    const results = [];
    const processedThreads = [];
    let lastProcessedIndex = -1;

    for (let i = 0; i < threadsToProcess.length; i++) {
      const thread = threadsToProcess[i];
      lastProcessedIndex = i;

      // 時間制限
      if (i > 0 && i % (CONFIG.TIME_CHECK_INTERVAL || 50) === 0) {
        const elapsedMinutes = (new Date() - startTime) / 60000;
        if (elapsedMinutes > (CONFIG.TIME_LIMIT_MINUTES || 4.5)) {
          addLogFromServer('WARN','時間制限で中断',{i, total:threadsToProcess.length, processedCount});
          timeoutOccurred = true;
          break;
        }
      }

      try {
        const messages = thread.getMessages();
        let threadProcessed = false;

        for (const message of messages) {
          try {
            const salesData = parseBOOTHSalesEmail(message.getPlainBody());
            if (!salesData) continue;

            if (existingOrderNumbers.has(salesData.orderNumber)) continue;

            // 書き込み（シート喪失時は復旧→再試行）
            addSalesRecord(sheet, salesData);
            existingOrderNumbers.add(salesData.orderNumber);

            processedCount++;
            threadProcessed = true;
            results.push({
              orderNumber: salesData.orderNumber,
              productName: salesData.productName,
              amount: salesData.amount,
              date: salesData.orderDateTime,
              paymentType: salesData.paymentType
            });

            if (processedCount % (CONFIG.LOG_INTERVAL || 10) === 0) {
              addLogFromServer('INFO','進捗',{processedCount, lastProduct:salesData.productName, lastAmount:salesData.amount});
            }
          } catch (errMsg) {
            addLogFromServer('ERROR','メール処理エラー',{index:i+1, error:String(errMsg)});
          }
        }

        if (threadProcessed || !timeoutOccurred) {
          processedThreads.push(thread);
        }

        if ((i + 1) % (CONFIG.PROGRESS_INTERVAL || 100) === 0) {
          addLogFromServer('INFO','進捗(スレッド)',{done:i+1, total:threadsToProcess.length, recorded:processedCount});
          if (processedThreads.length > 0) {
            addLabelBatch(processedThreads, CONFIG.SEARCH_LABEL);
            processedThreads.length = 0;
          }
          Utilities.sleep(100);
        }

      } catch (e) {
        addLogFromServer('ERROR','スレッド処理エラー',{index:i+1, error:String(e)});
      }
    }

    if (processedThreads.length > 0) addLabelBatch(processedThreads, CONFIG.SEARCH_LABEL);

    if (timeoutOccurred && lastProcessedIndex < threadsToProcess.length - 1) {
      const remainingThreads = threadsToProcess.slice(lastProcessedIndex + 1);
      if (remainingThreads.length > 0) {
        addLabelBatch(remainingThreads, CONFIG.PROCESSING_LABEL);
        addLogFromServer('INFO','処理中ラベル付与',{remaining: remainingThreads.length});
      }
    }

    const completionMessage = timeoutOccurred
      ? `時間制限により中断。${processedCount}件処理。残り${threadsToProcess.length - lastProcessedIndex - 1}件は次回`
      : `処理完了: ${processedCount}件の売上を記録`;

    addLogFromServer(timeoutOccurred ? 'WARN' : 'INFO','完了',{processed:processedCount, timeout:timeoutOccurred});

    return {
      processed: processedCount,
      total: threadsToProcess.length,
      totalFound: allThreads.length,
      results,
      timeoutOccurred,
      remaining: timeoutOccurred ? threadsToProcess.length - lastProcessedIndex - 1 : 0,
      message: completionMessage
    };

  } catch (error) {
    addLogFromServer('ERROR','メイン処理エラー',{error:String(error)});
    return { processed: 0, error: true, message: String(error) };
  } finally {
    lock.releaseLock();
  }
}
/**
 * 手動実行用のテスト関数
 */
function testRun() {
  console.log('=== テスト実行開始 ===');

  if (CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('CONFIG.SPREADSHEET_ID を設定してください');
  }

  const result = processBOOTHSalesEmails();
  console.log('実行結果:', result);

  return result;
}

// ===== 📧 メール解析関数 =====

/**
 * BOOTHメールの解析 - 売上通知（即時決済・後払い決済両方）に対応
 */
function parseBOOTHSalesEmail(emailBody) {
  try {
    const decodedBody = decodeQuotedPrintable(emailBody);

    const isInstantPayment = decodedBody.includes('商品が購入されました') || decodedBody.includes('の商品が購入されました');
    const isDelayedPayment = decodedBody.includes('ご注文が確定しました') || decodedBody.includes('購入者のお支払いを確認しました');

    if (!isInstantPayment && !isDelayedPayment) {
      return null;
    }

    const orderDateTime = extractOrderDateTime(decodedBody);

    const orderNumber = extractOrderNumber(decodedBody);
    if (!orderNumber) {
      console.warn('注文番号が見つかりませんでした');
      return null;
    }

    const productInfo = extractProductInfo(decodedBody, isInstantPayment, isDelayedPayment);

    if (!productInfo.amount || isNaN(productInfo.amount) || productInfo.amount <= 0) {
      console.warn('無効な金額が検出されました:', productInfo.amount);
      return null;
    }

    const processedProduct = processProductName(productInfo.productName);

    return {
      orderDateTime: orderDateTime,
      orderNumber: orderNumber,
      productName: processedProduct.cleanProductName || 'BOOTH商品',
      productVariant: processedProduct.productVariant,
      amount: productInfo.amount,
      paymentType: isInstantPayment ? 'instant' : 'delayed'
    };

  } catch (error) {
    console.error('メール解析エラー:', error);
    return null;
  }
}

/**
 * 注文日時の抽出
 */
function extractOrderDateTime(decodedBody) {
  let dateMatch = decodedBody.match(/(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2})時(\d{1,2})分/);

  if (dateMatch) {
    const [, year, month, day, hour, minute] = dateMatch;
    return new Date(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day),
      parseInt(hour),
      parseInt(minute)
    );
  }

  const encodedMatch = decodedBody.match(/(\d{4})=E5=B9=B4(\d{1,2})=E6=9C=88(\d{1,2})=E6=97=A5\s+(\d{1,2})=E6=99=82(\d{1,2})=E5=88=86/);
  if (encodedMatch) {
    const [, year, month, day, hour, minute] = encodedMatch;
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hour), parseInt(minute));
  }

  return new Date();
}

/**
 * 注文番号の抽出
 */
function extractOrderNumber(decodedBody) {
  const orderNumberMatch = decodedBody.match(/注文番号[^\d]*(\d+)/) ||
                         decodedBody.match(/(\d{8,})/);

  return orderNumberMatch ? parseInt(orderNumberMatch[1]) : null;
}

/**
 * 商品情報（商品名・金額）の抽出
 */
function extractProductInfo(decodedBody, isInstantPayment, isDelayedPayment) {
  let productName = '';
  let amount = 0;

  if (isInstantPayment) {
    const orderContentMatch = decodedBody.match(/注文内容[^】]*】([^\n¥]+)[\s\S]*?¥\s*([\d,]+)/);

    if (orderContentMatch) {
      const fullNameMatch = decodedBody.match(/】([^¥\n]+)/);
      if (fullNameMatch) {
        productName = fullNameMatch[1].trim();
        amount = parseInt(orderContentMatch[2].replace(/,/g, ''));
      }
    }

    if (!productName) {
      const encodedMatch = decodedBody.match(/=E3=80=90([^=]+)=E3=80=91([^=]+).*?=C2=A5\s*([\d,]+)/);
      if (encodedMatch) {
        const category = decodeQuotedPrintable('=E3=80=90' + encodedMatch[1] + '=E3=80=91');
        const name = decodeQuotedPrintable(encodedMatch[2]);
        productName = (category + name).trim();
        amount = parseInt(encodedMatch[3].replace(/,/g, ''));
      }
    }
  }

  if (isDelayedPayment) {
    const delayedProductMatch = decodedBody.match(/=E3=80=90([^=]+)=E3=80=91([^=\n¥]+)[\s\S]*?=C2=A5\s*([\d,]+)/);
    if (delayedProductMatch) {
      const category = decodeQuotedPrintable('=E3=80=90' + delayedProductMatch[1] + '=E3=80=91');
      const name = decodeQuotedPrintable(delayedProductMatch[2]);
      productName = (category + name).trim();
      amount = parseInt(delayedProductMatch[3].replace(/,/g, ''));
    }

    if (!productName) {
      const plainMatch = decodedBody.match(/【([^】]+)】([^¥\n]+)[\s\S]*?¥\s*([\d,]+)/);
      if (plainMatch) {
        productName = `【${plainMatch[1]}】${plainMatch[2]}`.trim();
        amount = parseInt(plainMatch[3].replace(/,/g, ''));
      }
    }
  }

  if (!productName || !amount) {
    const flexibleMatch = decodedBody.match(/([^\n]{10,100}?)[^¥]*¥\s*([\d,]+)/);
    if (flexibleMatch) {
      if (!productName) {
        productName = flexibleMatch[1].trim();
        if (productName.includes('注文') || productName.includes('番号') || productName.includes('日時') || productName.includes('支払')) {
          productName = '';
        }
      }
      if (!amount) {
        amount = parseInt(flexibleMatch[2].replace(/,/g, ''));
      }
    }
  }

  if (!amount) {
    const amountMatch = decodedBody.match(/¥\s*([\d,]+)/) ||
                       decodedBody.match(/([\d,]+)\s*円/) ||
                       decodedBody.match(/お支払金額[^¥]*¥\s*([\d,]+)/);
    if (amountMatch) {
      amount = parseInt(amountMatch[1].replace(/,/g, ''));
    }
  }

  if (!productName && amount) {
    productName = 'BOOTH商品（商品名取得失敗）';
  }

  return { productName: productName, amount: amount };
}

/**
 * 商品名のクリーンアップと版情報の分離
 */
function processProductName(productName) {
  let productVariant = '';

  if (productName && CONFIG.ENABLE_VARIANT_EXTRACTION) {
    productName = productName
      .replace(/\s+/g, ' ')
      .replace(/^\s+|\s+$/g, '')
      .replace(/\n/g, ' ');

    let variantMatch = productName.match(/^(.+?)（([^）]+)）\s*$/);

    if (!variantMatch) {
      variantMatch = productName.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
    }

    if (variantMatch) {
      const mainName = variantMatch[1].trim();
      const lastPart = variantMatch[2].trim();

      productName = mainName;
      productVariant = lastPart;
    }
  } else if (productName) {
    productName = productName
      .replace(/\s+/g, ' ')
      .replace(/^\s+|\s+$/g, '')
      .replace(/\n/g, ' ');
  }

  return { cleanProductName: productName, productVariant: productVariant };
}

/**
 * Quoted-Printableエンコーディングのデコード
 */
function decodeQuotedPrintable(str) {
  try {
    return str
      .replace(/=([0-9A-F]{2})/g, (match, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
      })
      .replace(/=\r?\n/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  } catch (error) {
    console.error('デコードエラー:', error);
    return str;
  }
}

// ===== 📊 スプレッドシート操作関数 =====

/**
 * スプレッドシートの取得または作成
 */
function getOrCreateSheet() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID not set');
  }

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Spreadsheet open failed: ' + String(e));
  }

  const name = CONFIG.SHEET_NAME || 'BOOTH売上履歴';
  let sheet = spreadsheet.getSheetByName(name);

  // 既存が壊れているケース: 取得できなければ新規作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    const headers = ['注文日時', '注文番号', '商品名', '版情報', '金額'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

/**
 * 既存の注文番号を取得（重複チェック用）
 */
function getExistingOrderNumbers(sheet) {
  const existingNumbers = new Set();

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const orderNumbers = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

      for (const [orderNumber] of orderNumbers) {
        if (orderNumber && typeof orderNumber === 'number') {
          existingNumbers.add(orderNumber);
        }
      }
    }
  } catch (error) {
    console.error('既存データ取得エラー:', error);
  }

  return existingNumbers;
}

/**
 * スプレッドシートに売上記録を追加（正式API準拠の表機能適用）
 */
function addSalesRecord(sheet, salesData) {
  const writeOnce = (sh) => {
    const newRow = [
      salesData.orderDateTime,
      salesData.orderNumber,
      salesData.productName,
      salesData.productVariant,
      salesData.amount
    ];
    sh.appendRow(newRow);

    const lastRow = sh.getLastRow();
    sh.getRange(lastRow, 1).setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange(lastRow, 2).setNumberFormat('0');
    sh.getRange(lastRow, 5).setNumberFormat('¥#,##0');

    if (!CONFIG.DISABLE_TABLE_FEATURE) {
      setupTableStructure(sh, lastRow);
    } else if (CONFIG.ENABLE_BASIC_FORMATTING) {
      applyBasicFormatting(sh, lastRow);
    }
  };

  try {
    writeOnce(sheet);
  } catch (e) {
    const msg = String(e);
    if (msg.includes('Sheet') && msg.includes('not found')) {
      // タブが消えている等。再取得→1回だけリトライ。
      const sh = getOrCreateSheet();
      writeOnce(sh);
      addLogFromServer('WARN','シート再取得してリトライ',{error: msg});
    } else {
      throw e;
    }
  }
}
// ===== 📊 表機能（Google Sheets 正式API準拠）=====

/**
 * 互換エントリポイント（旧API呼び出しの置換）
 */
function setupTableStructure(sheet, _currentRow) {
  try {
    if (CONFIG.TABLE_SETUP_DELAY > 0 && sheet.getLastRow() === 2) {
      Utilities.sleep(CONFIG.TABLE_SETUP_DELAY);
    }
    ensureTableFeatures(sheet);
  } catch (e) {
    console.log('⚠️ 表機能設定をスキップ: ' + e.message);
    if (CONFIG.ENABLE_BASIC_FORMATTING) applyBasicFormatting(sheet, sheet.getLastRow());
  }
}

/**
 * 既存の updateTableRange を互換提供（内部は同じ処理）
 */
function updateTableRange(sheet, _currentRow) {
  ensureTableFeatures(sheet);
}

/**
 * 表機能の実体:
 * - ヘッダー整形 + 固定
 * - フィルタの作成/更新（全行追従オプション）
 * - 交互色バンディングの作成/伸長
 * - 列幅調整
 */
function ensureLogSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(LOG_SHEET_NAME || 'BOOTH_LOGS');
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME || 'BOOTH_LOGS');
    sh.getRange(1,1,1,4).setValues([['日時','レベル','メッセージ','コンテキスト']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * フォールバックの基本書式
 */
function applyBasicFormatting(sheet, currentRow) {
  if (currentRow < 2) return;
  const header = sheet.getRange(1, 1, 1, 5);
  header.setFontWeight('bold').setBackground('#f1f3f4').setBorder(true, true, true, true, false, false);
  sheet.getRange(currentRow, 1, 1, 5).setBorder(true, true, true, true, false, false);
}

/**
 * 手動で表機能を設定（互換API）
 */
function manuallySetupTable() {
  console.log('=== 手動表機能設定開始 ===');
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('データがありません。データを追加してから実行してください。');
      return { success: false, message: 'データがありません' };
    }
    ensureTableFeatures(sheet);
    console.log(`✅ 表機能を手動設定しました（${lastRow}行, 5列）`);
    return {
      success: true,
      message: `表機能を手動設定しました（${lastRow}行のデータ）`,
      rows: lastRow,
      columns: 5
    };
  } catch (error) {
    console.error('手動表機能設定エラー:', error);
    if (CONFIG.ENABLE_BASIC_FORMATTING) {
      try {
        const sheet = getOrCreateSheet();
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const headerRange = sheet.getRange(1, 1, 1, 5);
          headerRange.setFontWeight('bold').setBackground('#f1f3f4');
          const allDataRange = sheet.getRange(1, 1, lastRow, 5);
          allDataRange.setBorder(true, true, true, true, true, true);
          sheet.autoResizeColumns(1, 5);
          console.log('✅ 基本書式を適用しました（表機能の代替）');
          return {
            success: true,
            message: '表機能は設定できませんでしたが、基本書式を適用しました',
            alternative: true
          };
        }
      } catch (backupError) {
        console.error('基本書式適用も失敗:', backupError);
      }
    }
    return {
      success: false,
      message: `表機能設定に失敗しました: ${error.message}`,
      suggestion: 'Googleスプレッドシートで手動で交互色またはフィルタを設定してください'
    };
  }
}

// ===== 🏷️ Gmail操作関数 =====

/**
 * バッチでラベルを追加する関数（パフォーマンス最適化）
 */
function addLabelBatch(threads, labelName) {
  if (threads.length === 0) return;

  try {
    const label = getOrCreateLabel(labelName || CONFIG.SEARCH_LABEL);

    for (let i = 0; i < threads.length; i += CONFIG.BATCH_SIZE) {
      const batch = threads.slice(i, i + CONFIG.BATCH_SIZE);

      for (const thread of batch) {
        thread.addLabel(label);
      }

      if (i + CONFIG.BATCH_SIZE < threads.length) {
        Utilities.sleep(50);
      }
    }
  } catch (error) {
    console.error('ラベル追加エラー:', error);
  }
}

/**
 * Gmailラベルの取得または作成
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
    console.log('新しいラベルを作成しました: ' + labelName);
  }

  return label;
}

// ===== 🔧 管理機能 =====

/** セットアップ後のみ実体を希望に同期し、最終状態(true=ON)を返す */
function setupTriggers(wantOn) {
  const sp = PropertiesService.getScriptProperties();
  const setupDone = sp.getProperty('SETUP_DONE') === '1';
  if (!setupDone) {
    // 前: 実体は作らない。常にOFF。
    ScriptApp.getProjectTriggers().forEach(t=>{
      if (t.getHandlerFunction && t.getHandlerFunction() === 'processBOOTHSalesEmails') {
        ScriptApp.deleteTrigger(t);
      }
    });
    addLogFromServer && addLogFromServer('INFO','未セットアップのためトリガー未作成',{wantOn:!!wantOn});
    return false;
  }
  const state = triggerSync_(!!wantOn);
  addLogFromServer && addLogFromServer('INFO','トリガー同期',{desired:!!wantOn, actual:state});
  return state;
}


function _isTriggerOn_() {
  return ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction && t.getHandlerFunction() === 'processBOOTHSalesEmails');
}

/** 実体トリガーを希望に同期し、最終状態を返す（唯一の作成/削除点） */
function triggerSync_(wantOn) {
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction && t.getHandlerFunction() === 'processBOOTHSalesEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });
  if (wantOn) {
    ScriptApp.newTrigger('processBOOTHSalesEmails').timeBased().everyMinutes(5).create();
    addLogFromServer && addLogFromServer('INFO','トリガー作成',{everyMinutes:5});
  } else {
    addLogFromServer && addLogFromServer('INFO','トリガー未作成',{reason:'OFF'});
  }
  return _isTriggerOn_();
}

/** トリガー状態取得
 *  セットアップ前は常にOFFを返す。セットアップ後は実体状態を返す。 */
function getTriggerStatus() {
  const sp = PropertiesService.getScriptProperties();
  const setupDone = sp.getProperty('SETUP_DONE') === '1';
  if (!setupDone) return false;            // 前: UI関係なくOFF
  return _isTriggerOn_();                  // 後: 実体を返す
}

/** トリガーON/OFF
 *  希望値は常に保存。セットアップ前は実体OFF維持。後は即同期。 */
function setTriggerEnabled(enabled) {
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('TRIGGER_ENABLED', enabled ? '1' : '0');
  const setupDone = sp.getProperty('SETUP_DONE') === '1';

  // セットアップ前は常に実体OFF（強制削除）
  if (!setupDone) {
    ScriptApp.getProjectTriggers().forEach(t=>{
      if (t.getHandlerFunction && t.getHandlerFunction() === 'processBOOTHSalesEmails') {
        ScriptApp.deleteTrigger(t);
      }
    });
    addLogFromServer && addLogFromServer('INFO','未セットアップのため実体OFF',{desired:enabled});
    return false;
  }

  // セットアップ後は希望に同期
  const actual = triggerSync_(!!enabled);
  addLogFromServer && addLogFromServer('INFO','トリガー更新',{desired:!!enabled, actual});
  return actual;
}


/**
 * 処理済みラベルをクリアする関数（再処理用）
 */
function clearProcessedFlags() {
  console.log('=== 処理済みフラグのクリア開始 ===');

  try {
    let totalCleared = 0;

    const processedLabel = GmailApp.getUserLabelByName(CONFIG.SEARCH_LABEL);
    if (processedLabel) {
      const processedThreads = processedLabel.getThreads();
      for (const thread of processedThreads) {
        thread.removeLabel(processedLabel);
        totalCleared++;
      }
      console.log(`処理済みラベル: ${processedThreads.length}件をクリア`);
    }

    const processingLabel = GmailApp.getUserLabelByName(CONFIG.PROCESSING_LABEL);
    if (processingLabel) {
      const processingThreads = processingLabel.getThreads();
      for (const thread of processingThreads) {
        thread.removeLabel(processingLabel);
        totalCleared++;
      }
      console.log(`処理中ラベル: ${processingThreads.length}件をクリア`);
    }

    if (totalCleared === 0) {
      return { cleared: 0, message: '処理済み・処理中ラベルが存在しません' };
    }

    console.log(`=== 処理済みフラグクリア完了: ${totalCleared}件 ===`);

    return {
      cleared: totalCleared,
      message: `${totalCleared}件のメールから処理済みフラグを削除しました`
    };

  } catch (error) {
    console.error('処理済みフラグクリアエラー:', error);
    throw error;
  }
}

/**
 * 全ての処理済みメールを再処理する関数
 */
function reprocessAllEmails() {
  console.log('=== 全メール再処理開始 ===');

  try {
    const clearResult = clearProcessedFlags();
    console.log(`フラグクリア結果: ${clearResult.message}`);

    Utilities.sleep(1000);

    const processResult = processBOOTHSalesEmails();

    console.log('=== 全メール再処理完了 ===');

    return {
      flagsCleared: clearResult.cleared,
      emailsProcessed: processResult.processed,
      message: `${clearResult.cleared}件のフラグをクリアし、${processResult.processed}件のメールを再処理しました`
    };

  } catch (error) {
    console.error('全メール再処理エラー:', error);
    throw error;
  }
}

/**
 * 既存データの金額列を数値形式に修正する関数（メンテナンス用）
 */
function fixAmountColumnFormat() {
  console.log('=== 金額列フォーマット修正開始 ===');

  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('データがありません');
      return;
    }

    const amountRange = sheet.getRange(2, 5, lastRow - 1, 1);
    const amountValues = amountRange.getValues();

    let fixedCount = 0;

    for (let i = 0; i < amountValues.length; i++) {
      const currentValue = amountValues[i][0];

      if (typeof currentValue === 'string') {
        const numericValue = parseInt(currentValue.replace(/[^0-9]/g, ''));
        if (!isNaN(numericValue) && numericValue > 0) {
          amountValues[i][0] = numericValue;
          fixedCount++;
        }
      }
    }

    if (fixedCount > 0) {
      amountRange.setValues(amountValues);
      amountRange.setNumberFormat('¥#,##0');

      console.log(`金額列修正完了: ${fixedCount}件のデータを数値形式に変換`);
    } else {
      console.log('修正が必要なデータはありませんでした');
    }

    return {
      fixed: fixedCount,
      message: `${fixedCount}件の金額データを修正しました`
    };

  } catch (error) {
    console.error('金額列修正エラー:', error);
    throw error;
  }
}

// ===== 📊 分析機能 =====

/**
 * 売上データの分析・集計用関数
 */
function analyzeSalesData() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return '売上データがありません';
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  let totalSales = 0;
  let totalOrders = data.length;
  const productSales = {};
  const variantSales = {};
  const monthlySales = {};

  data.forEach(([date, orderNumber, productName, productVariant, amount]) => {
    totalSales += amount;

    if (productSales[productName]) {
      productSales[productName] += amount;
    } else {
      productSales[productName] = amount;
    }

    if (productVariant) {
      if (variantSales[productVariant]) {
        variantSales[productVariant] += amount;
      } else {
        variantSales[productVariant] = amount;
      }
    }

    const monthKey = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM');
    if (monthlySales[monthKey]) {
      monthlySales[monthKey] += amount;
    } else {
      monthlySales[monthKey] = amount;
    }
  });

  const averageOrderValue = Math.round(totalSales / totalOrders);

  console.log('=== 売上分析結果 ===');
  console.log(`総売上: ¥${totalSales.toLocaleString()}`);
  console.log(`注文数: ${totalOrders}件`);
  console.log(`平均注文額: ¥${averageOrderValue.toLocaleString()}`);

  return {
    totalSales: totalSales,
    totalOrders: totalOrders,
    averageOrderValue: averageOrderValue,
    productSales: productSales,
    variantSales: variantSales,
    monthlySales: monthlySales
  };
}

/**
 * 処理中ラベルをクリアして続きから実行する関数
 */
function continueProcessing() {
  console.log('=== 処理継続開始 ===');

  try {
    const processingLabel = GmailApp.getUserLabelByName(CONFIG.PROCESSING_LABEL);

    if (!processingLabel) {
      console.log('処理中のメールはありません。通常処理を実行します。');
      return processBOOTHSalesEmails();
    }

    const processingThreads = processingLabel.getThreads();
    console.log(`処理中ラベル: ${processingThreads.length}件のメールを発見`);

    if (processingThreads.length === 0) {
      console.log('処理中のメールはありません。通常処理を実行します。');
      return processBOOTHSalesEmails();
    }

    for (const thread of processingThreads) {
      thread.removeLabel(processingLabel);
    }

    console.log('処理中ラベルを削除しました。メイン処理を再実行します。');

    const result = processBOOTHSalesEmails();

    return {
      ...result,
      continued: true,
      message: `継続処理完了: ${result.processed}件を処理（処理中だった${processingThreads.length}件を含む）`
    };

  } catch (error) {
    console.error('継続処理エラー:', error);
    return {
      processed: 0,
      error: true,
      message: `継続処理エラー: ${error.message}`
    };
  }
}

// ===== 🎛️ 追加の管理機能 =====

/**
 * 表機能を完全に無効化する設定関数
 */
function disableTableFeature() {
  console.log('=== 表機能無効化 ===');
  return {
    disabled: true,
    message: '表機能を無効化しました。CONFIG.DISABLE_TABLE_FEATURE = true を設定すると永続化されます。',
    instruction: 'CONFIG オブジェクトで DISABLE_TABLE_FEATURE: true を設定してください'
  };
}

/**
 * スクリプトの設定状況を確認する関数
 */
function checkConfiguration() {
  console.log('=== 設定確認開始 ===');

  const results = {
    spreadsheetId: CONFIG.SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE',
    sheetExists: false,
    triggersSet: false,
    labelsExist: {},
    configuration: CONFIG
  };

  try {
    if (results.spreadsheetId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
        results.sheetExists = sheet !== null;

        if (sheet) {
          results.sheetInfo = {
            lastRow: sheet.getLastRow(),
            lastColumn: sheet.getLastColumn(),
            dataRows: Math.max(0, sheet.getLastRow() - 1)
          };
        }
      } catch (error) {
        results.spreadsheetError = error.message;
      }
    }

    const triggers = ScriptApp.getProjectTriggers();
    results.triggersSet = triggers.some(trigger =>
      trigger.getHandlerFunction() === 'processBOOTHSalesEmails'
    );
    results.triggerCount = triggers.length;

    try {
      results.labelsExist.processed = GmailApp.getUserLabelByName(CONFIG.SEARCH_LABEL) !== null;
      results.labelsExist.processing = GmailApp.getUserLabelByName(CONFIG.PROCESSING_LABEL) !== null;
    } catch (error) {
      results.labelsError = error.message;
    }

    console.log('✅ 設定確認完了:');
    console.log(`- スプレッドシートID設定: ${results.spreadsheetId ? '✅' : '❌'}`);
    console.log(`- シート存在: ${results.sheetExists ? '✅' : '❌'}`);
    console.log(`- トリガー設定: ${results.triggersSet ? '✅' : '❌'}`);
    console.log(`- 処理済みラベル: ${results.labelsExist.processed ? '✅' : '❌'}`);
    console.log(`- 処理中ラベル: ${results.labelsExist.processing ? '✅' : '❌'}`);

    if (results.sheetInfo) {
      console.log(`- データ行数: ${results.sheetInfo.dataRows}件`);
    }

    return results;

  } catch (error) {
    console.error('設定確認エラー:', error);
    results.error = error.message;
    return results;
  }
}



/**
 * スクリプトの初期セットアップを一括実行する関数
 */
function _setupCore_(caller) {
  addLogFromServer && addLogFromServer('INFO','初期セットアップ開始',{caller});
  const res = { ok:false, steps:[], trigger:false, stats:null };

  // 1) ID確認
  ensureConfigLoaded();
  const id = (CONFIG.SPREADSHEET_ID || '').trim();
  if (!id || id === 'YOUR_SPREADSHEET_ID_HERE') {
    res.steps.push('❌ SPREADSHEET_ID 未設定。設定タブで保存してください。');
    addLogFromServer && addLogFromServer('WARN','初期セットアップ拒否: SPREADSHEET_ID未設定',{caller});
    return res;
  }

  // 2) シート準備
  let sheet;
  try {
    sheet = getOrCreateSheet();
    res.steps.push(`✅ シート準備完了: ${sheet.getName()}`);
  } catch (e) {
    res.steps.push(`❌ シート準備失敗: ${String(e)}`);
    addLogFromServer && addLogFromServer('ERROR','シート準備失敗',{error:String(e), caller});
    return res;
  }

  // 3) Gmail疎通
  try {
    GmailApp.search('from:noreply@booth.pm (subject:商品が購入されました OR subject:ご注文が確定しました)');
    res.steps.push('✅ Gmail 検索テスト OK');
  } catch (e) {
    res.steps.push(`❌ Gmail 検索テスト失敗: ${String(e)}`);
    addLogFromServer && addLogFromServer('ERROR','Gmail検索テスト失敗',{error:String(e), caller});
    return res;
  }

  // 4) 既存フラグ削除（非致命）
  try {
    const r = clearProcessedFlags();
    res.steps.push(`✅ フラグ削除: ${r && typeof r.cleared==='number' ? r.cleared : 0} 件`);
  } catch (e) {
    res.steps.push(`⚠️ フラグ削除失敗: ${String(e)}`);
    addLogFromServer && addLogFromServer('WARN','フラグ削除失敗',{error:String(e), caller});
  }

  // 5) ラベル作成
  try {
    getOrCreateLabel(CONFIG.SEARCH_LABEL);
    getOrCreateLabel(CONFIG.PROCESSING_LABEL);
    res.steps.push(`✅ ラベル作成: ${CONFIG.SEARCH_LABEL}, ${CONFIG.PROCESSING_LABEL}`);
  } catch (e) {
    res.steps.push(`❌ ラベル作成失敗: ${String(e)}`);
    addLogFromServer && addLogFromServer('ERROR','ラベル作成失敗',{error:String(e), caller});
    return res;
  }

  // 6) セットアップ完了 → 強制ON（UI希望は無視してONに上書き）
  try {
    const sp = PropertiesService.getScriptProperties();
    sp.setProperty('SETUP_DONE','1');        // セットアップ完了フラグ
    sp.setProperty('TRIGGER_ENABLED','1');   // 希望値もONに強制上書き
    const actual = triggerSync_(true);       // 実体もONに同期
    res.steps.push(`✅ トリガー強制ON`);
    res.trigger = actual;
  } catch (e) {
    res.steps.push(`⚠️ トリガー同期失敗: ${String(e)}`);
    addLogFromServer && addLogFromServer('WARN','トリガー同期失敗',{error:String(e), caller});
  }

  // 7) 初回ワンショット収集（統計のみ反映）
  try {
    const once = processBOOTHSalesEmails();
    addLogFromServer && addLogFromServer(once && once.error ? 'ERROR' : 'INFO', '初回収集完了', {
      processed: once && once.processed, totalFound: once && once.totalFound
    });
  } catch (e) {
    addLogFromServer && addLogFromServer('ERROR','初回収集失敗',{error:String(e)});
  }

  // 8) 表示メトリクス
  let collectedRows = null, scannedThreads = null;
  try { collectedRows = Math.max(0, getOrCreateSheet().getLastRow() - 1); } catch(e){}
  try {
    scannedThreads = GmailApp.search(
      'from:noreply@booth.pm (subject:商品が購入されました OR subject:ご注文が確定しました)'
    ).length;
  } catch(e){}

  res.stats = { collectedRows, scannedThreads };
  res.ok = true;
  addLogFromServer && addLogFromServer('INFO','初期セットアップ完了',{caller, trigger:res.trigger, stats:res.stats});
  return res;
}

function initialSetup() {
  return _setupCore_('initialSetup');
}
