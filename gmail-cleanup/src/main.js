/**
 * GmailCleanup_DANGER
 * 
 * 180日以上前、かつスターなし、かつ重要マークなしのメールを完全削除し、
 * 自分のSlack DMに通知するスクリプト
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

// ===================
// 設定
// ===================
const CONFIG = {
  RETENTION_DAYS: 180,
  DRY_RUN: false,
};

const SLACK_API_URL = 'https://slack.com/api/chat.postMessage';

// ===================
// メイン処理
// ===================

/**
 * メイン関数：古いメールを削除してSlackに通知
 */
function cleanupOldEmailsAndNotify() {
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

  if (CONFIG.DRY_RUN) {
    console.log('【ドライランモード】削除はスキップされました。');
    notifySlack(targetCount, true);
    return;
  }

  const deletedCount = deleteThreadsPermanently(targetThreads);
  notifySlack(deletedCount, false);
}

// ===================
// ヘルパー関数
// ===================

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

// ===================
// Slack通知（自分のDMに送信）
// ===================

/**
 * 自分のSlack DMに削除結果を通知
 */
function notifySlack(count, isDryRun) {
  const credentials = getSlackCredentials();
  if (!credentials.token || !credentials.userId) {
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
    const response = UrlFetchApp.fetch(SLACK_API_URL, fetchOptions);
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

/**
 * Slack Block Kit形式のメッセージを構築
 */
function buildSlackBlocks(count, isDryRun) {
  const status = isDryRun ? 'Dry Run' : 'Done';
  const condition = `${CONFIG.RETENTION_DAYS}日以上前 \`AND\`\nスターなし \`AND\`\n重要マークなし`;
  
  return [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: ':broom: Gmail自動整理レポート',
        emoji: true,
      },
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '対象メールを完全に削除しました',
      },
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*Status:*\n${status}`,
      },
    },
    {
      type: 'section',
      fields: [
        {
          type: 'mrkdwn',
          text: `*Filter:*\n${condition}`,
        },
        {
          type: 'mrkdwn',
          text: `*Count:*\n${count} 件`,
        },
      ],
    },
  ];
}

function buildSlackMessage(count, isDryRun) {
  const prefix = isDryRun ? '【ドライラン】' : '';
  const action = isDryRun ? '削除対象として検出' : '完全に削除しました';
  return `${prefix}【Gmail自動整理】${CONFIG.RETENTION_DAYS}日以上前の不要なメールを ${count} 件、${action}。`;
}

/**
 * スクリプトプロパティからSlack認証情報を取得
 */
function getSlackCredentials() {
  const props = PropertiesService.getScriptProperties();
  return {
    token: props.getProperty('SLACK_BOT_TOKEN'),
    userId: props.getProperty('SLACK_USER_ID'),
  };
}

// ===================
// セットアップ用ユーティリティ
// ===================

/**
 * 設定確認（スクリプトプロパティが正しく設定されているか確認）
 * 初回セットアップ後に実行してください
 */
function verifySettings() {
  console.log('=== 設定確認 ===');
  
  const credentials = getSlackCredentials();
  let hasError = false;
  
  // SLACK_BOT_TOKEN の確認
  if (!credentials.token) {
    console.error('❌ SLACK_BOT_TOKEN: 未設定');
    hasError = true;
  } else if (!credentials.token.startsWith('xoxb-')) {
    console.error('❌ SLACK_BOT_TOKEN: 形式が不正（xoxb- で始まる必要があります）');
    hasError = true;
  } else {
    console.log('✅ SLACK_BOT_TOKEN: 設定済み');
  }
  
  // SLACK_USER_ID の確認
  if (!credentials.userId) {
    console.error('❌ SLACK_USER_ID: 未設定');
    hasError = true;
  } else if (!credentials.userId.startsWith('U')) {
    console.error('❌ SLACK_USER_ID: 形式が不正（U で始まる必要があります）');
    hasError = true;
  } else {
    console.log(`✅ SLACK_USER_ID: ${credentials.userId}`);
  }
  
  console.log('================');
  
  if (hasError) {
    console.log('');
    console.log('【設定方法】');
    console.log('1. GASエディタ左メニュー「⚙️ プロジェクトの設定」をクリック');
    console.log('2. 下部「スクリプト プロパティ」で以下を追加:');
    console.log('   - SLACK_BOT_TOKEN: xoxb-xxxx-xxxx-xxxx');
    console.log('   - SLACK_USER_ID: U01XXXXXXXX');
  } else {
    console.log('✅ 全ての設定が完了しています。testSlackConnection() で接続テストを実行してください。');
  }
}

/**
 * Slack接続テスト
 */
function testSlackConnection() {
  const credentials = getSlackCredentials();
  
  if (!credentials.token) {
    console.error('SLACK_BOT_TOKEN が設定されていません。verifySettings() を実行して確認してください。');
    return;
  }
  
  const response = UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${credentials.token}`,
    },
    muteHttpExceptions: true,
  });
  
  const result = JSON.parse(response.getContentText());
  
  if (result.ok) {
    console.log(`✅ 接続成功: ${result.team} / Bot: ${result.user}`);
  } else {
    console.error(`❌ 接続失敗: ${result.error}`);
  }
}

/**
 * ドライランモードで実行（削除せず通知のみ）
 */
function dryRun() {
  const originalDryRun = CONFIG.DRY_RUN;
  CONFIG.DRY_RUN = true;
  
  try {
    cleanupOldEmailsAndNotify();
  } finally {
    CONFIG.DRY_RUN = originalDryRun;
  }
}