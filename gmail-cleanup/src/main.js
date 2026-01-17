/**
 * GmailCleanup_DANGER
 *
 * 180日以上前、かつスターなし、かつ重要マークなしのメールを完全削除し、
 * 自分のSlack DMに通知するスクリプト
 *
 * @version 1.2.0
 * 
 * 【事前準備】
 * 1. Gmail APIを有効化（サービス → Gmail API を追加）
 * 2. Slack Appを作成し、Bot Token Scopesに「chat:write」を追加
 * 3. スクリプトプロパティを設定（以下の手順）:
 * 
 *    [スクリプトプロパティの設定方法]
 *    a. GASエディタ左メニュー「⚙️ プロジェクトの設定」をクリック
 *    b. 下部「スクリプト プロパティ」セクションへスクロール
 *    c. 「スクリプト プロパティを追加」をクリック
 *    d. 以下の2つを追加:
 *       - プロパティ: SLACK_BOT_TOKEN  値: xoxb-xxxx-xxxx-xxxx
 *       - プロパティ: SLACK_USER_ID    値: U01XXXXXXXX
 *    e. 「スクリプト プロパティを保存」をクリック
 * 
 *    [ユーザーIDの確認方法]
 *    Slackで自分のプロフィール → 「︙」→「メンバーIDをコピー」
 */

const CONFIG = {
  RETENTION_DAYS: 180,
  DRY_RUN: false,
  SLACK_POST_URL: 'https://slack.com/api/chat.postMessage',
  SLACK_AUTH_URL: 'https://slack.com/api/auth.test',
};

let _credentials = null;

/**
 * メイン関数：古いメールを削除してSlackに通知
 * @param {Object} [options] - オプション
 * @param {boolean} [options.dryRun] - ドライランモード（デフォルト: CONFIG.DRY_RUN）
 */
function cleanupOldEmailsAndNotify(options = {}) {
  const isDryRun = options.dryRun ?? CONFIG.DRY_RUN;

  const searchQuery = buildSearchQuery();
  const targetThreads = GmailApp.search(searchQuery);
  const targetCount = targetThreads.length;

  console.log(`検索クエリ: ${searchQuery}`);
  console.log(`対象スレッド数: ${targetCount}`);

  if (targetCount === 0) {
    console.log('対象のメールはありませんでした。');
    return;
  }

  logTargetThreads(targetThreads);

  if (isDryRun) {
    console.log('【ドライランモード】削除はスキップされました。');
    notifySlack(targetCount, true);
    return;
  }

  const deletedCount = deleteThreadsPermanently(targetThreads);
  notifySlack(deletedCount, false);
}

function buildSearchQuery() {
  return `older_than:${CONFIG.RETENTION_DAYS}d -is:starred -is:important`;
}

function logTargetThreads(threads) {
  console.log('--- 削除対象一覧 ---');
  threads.forEach((thread, index) => {
    const subject = thread.getFirstMessageSubject();
    const date = thread.getLastMessageDate();
    console.log(`${index + 1}. [${date.toLocaleDateString()}] ${subject}`);
  });
  console.log('--------------------');
}

function deleteThreadsPermanently(threads) {
  let deletedCount = 0;

  threads.forEach(thread => {
    try {
      deleteAllMessagesInThread(thread);
      deletedCount++;
    } catch (error) {
      const subject = thread.getFirstMessageSubject();
      console.error(`削除失敗: ${subject} - ${error.message}`);
    }
  });

  console.log(`削除完了: ${deletedCount}/${threads.length} スレッド`);
  return deletedCount;
}

function deleteAllMessagesInThread(thread) {
  const messages = thread.getMessages();
  messages.forEach(message => {
    Gmail.Users.Messages.remove('me', message.getId());
  });
}

function notifySlack(count, isDryRun) {
  const credentials = getSlackCredentials();
  const validation = validateCredentials(credentials);
  if (!validation.valid) {
    console.error('Slack認証情報が設定されていません。verifySettings() を実行して確認してください。');
    return;
  }

  const message = buildSlackMessage(count, isDryRun);
  
  const fetchOptions = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${credentials.token}`,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      channel: credentials.userId,
      text: message,
      blocks: buildSlackBlocks(count, isDryRun),
    }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.SLACK_POST_URL, fetchOptions);
    const result = JSON.parse(response.getContentText());

    if (result.ok) {
      console.log('Slack DM通知完了');
    } else {
      console.error(`Slack通知失敗: ${result.error}`);
    }
  } catch (error) {
    console.error(`Slack通知エラー: ${error.message}`);
  }
}

function buildSlackBlocks(count, isDryRun) {
  const status = isDryRun ? 'Dry Run' : 'Done';
  const filter = `${CONFIG.RETENTION_DAYS}日以上前 \`AND\`\nスターなし \`AND\`\n重要マークなし`;

  const header = (text) => ({ type: 'header', text: { type: 'plain_text', text, emoji: true } });
  const section = (text) => ({ type: 'section', text: { type: 'mrkdwn', text } });
  const fields = (items) => ({ type: 'section', fields: items.map(text => ({ type: 'mrkdwn', text })) });

  return [
    header(':broom: Gmail自動整理レポート'),
    section('対象メールを完全に削除しました'),
    section(`*Status:*\n${status}`),
    fields([`*Filter:*\n${filter}`, `*Count:*\n${count} 件`]),
  ];
}

function buildSlackMessage(count, isDryRun) {
  const prefix = isDryRun ? '【ドライラン】' : '';
  const action = isDryRun ? '削除対象として検出' : '完全に削除しました';
  return `${prefix}【Gmail自動整理】${CONFIG.RETENTION_DAYS}日以上前の不要なメールを ${count} 件、${action}。`;
}

/**
 * スクリプトプロパティからSlack認証情報を取得（キャッシュ付き）
 */
function getSlackCredentials() {
  if (!_credentials) {
    const props = PropertiesService.getScriptProperties();
    _credentials = {
      token: props.getProperty('SLACK_BOT_TOKEN'),
      userId: props.getProperty('SLACK_USER_ID'),
    };
  }
  return _credentials;
}

/**
 * Credentials の検証
 * @returns {{ valid: boolean, errors: string[] }}
 */
function validateCredentials(credentials) {
  const errors = [];

  if (!credentials.token) {
    errors.push('SLACK_BOT_TOKEN: 未設定');
  } else if (!credentials.token.startsWith('xoxb-')) {
    errors.push('SLACK_BOT_TOKEN: 形式が不正（xoxb- で始まる必要があります）');
  }

  if (!credentials.userId) {
    errors.push('SLACK_USER_ID: 未設定');
  } else if (!credentials.userId.startsWith('U')) {
    errors.push('SLACK_USER_ID: 形式が不正（U で始まる必要があります）');
  }

  return { valid: errors.length === 0, errors };
}

function verifySettings() {
  const credentials = getSlackCredentials();
  const validation = validateCredentials(credentials);

  console.log('=== 設定確認 ===');
  if (validation.valid) {
    console.log('✅ SLACK_BOT_TOKEN: 設定済み');
    console.log(`✅ SLACK_USER_ID: ${credentials.userId}`);
    console.log('✅ testSlackConnection() で接続テストを実行してください');
  } else {
    validation.errors.forEach(e => console.error(`❌ ${e}`));
    console.log('\n【設定方法】プロジェクトの設定 → スクリプト プロパティ');
    console.log('  SLACK_BOT_TOKEN: xoxb-xxxx-xxxx-xxxx');
    console.log('  SLACK_USER_ID: U01XXXXXXXX');
  }
}

function testSlackConnection() {
  const credentials = getSlackCredentials();
  const validation = validateCredentials(credentials);
  if (!validation.valid) {
    validation.errors.forEach(e => console.error(`❌ ${e}`));
    return;
  }

  const response = UrlFetchApp.fetch(CONFIG.SLACK_AUTH_URL, {
    method: 'post',
    headers: { 'Authorization': `Bearer ${credentials.token}` },
    muteHttpExceptions: true,
  });
  const result = JSON.parse(response.getContentText());

  if (result.ok) {
    console.log(`✅ 接続成功: ${result.team} / Bot: ${result.user}`);
  } else {
    console.error(`❌ 接続失敗: ${result.error}`);
  }
}

function dryRun() {
  cleanupOldEmailsAndNotify({ dryRun: true });
}