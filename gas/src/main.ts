const NOTIFIED_LABEL_NAME = "通知済み" as const;
const PROPERTY_NAME_PAST_MAILS = "pastMails" as const;
const MODES = ["Discord", "Slack"] as const;
type modes_type = (typeof MODES)[number];
// 特定ラベルでフィルタリングするためのプロパティ名
const PROPERTY_FILTER_LABEL = "FILTER_LABEL";

function main() {
    const mode = getProperty("MODE");
    if (!(mode === MODES[0] || mode === MODES[1])) {
        console.error("The mode is invalid. It must be Discord or Slack");
        return;
    }

    const webhook = new Webhook(getProperty("WEBHOOK_URL"), mode);
    // 通知済みラベルを取得または作成
    const notifiedLabel = getOrCreateLabel(NOTIFIED_LABEL_NAME);

    // フィルタリングラベルの設定を取得
    const filterLabelName = getProperty(PROPERTY_FILTER_LABEL);
    // フィルタリングラベルがある場合は取得、なければnullのまま
    const filterLabel = filterLabelName ? GmailApp.getUserLabelByName(filterLabelName) : null;

    const now = new Date();
    const myEmail = Session.getActiveUser().getEmail();

    let shouldContinue = true;
    for (let i = 0; shouldContinue && i < 10; i++) { // 最大10ループに制限
        // 2件ずつ取得(最大10件)
        const threads = GmailApp.getInboxThreads(i * 2, (i + 1) * 5);
        let foundNew = false;

        threads.forEach((thread) => {
            const threadLabels = thread.getLabels();

            // 通知済みラベルがついているかチェック
            const hasNotifiedLabel = threadLabels.some(label => label.getName() === NOTIFIED_LABEL_NAME);
            if (hasNotifiedLabel) {
                return; // 既に通知済みの場合はスキップ
            }

            // フィルタリングラベルが指定されていて、スレッドにそのラベルがない場合はスキップ
            if (filterLabel && !threadLabels.some(label => label.getName() === filterLabel.getName())) {
                return; // このスレッドを処理せずにスキップ
            }

            const messages = thread.getMessages();
            messages.forEach((message) => {
                const date = message.getDate();
                const from = message.getFrom();

                if (!from.includes(myEmail) &&
                    now.getTime() - date.getTime() < 1000 * 60 * 60 * 12) {
                    foundNew = true;
                    // 送信元が自分のメールアドレスでなく、かつメールが12時間以内に来たもの
                    try {
                        webhook.send(from, date, message.getSubject(), message.getPlainBody());
                        // 送信が成功した場合、通知済みラベルを付与
                        thread.addLabel(notifiedLabel);
                    } catch (e) {
                        console.error(`Failed to send a message: ${e}`);
                    }
                }
            });
        });

        shouldContinue = foundNew;
    }
}

// 指定した名前のラベルを取得、なければ作成する関数
function getOrCreateLabel(name: string): GoogleAppsScript.Gmail.GmailLabel {
    let label = GmailApp.getUserLabelByName(name);
    if (!label) {
        label = GmailApp.createLabel(name);
    }
    return label;
}

class Webhook {
    url = "";
    mode: modes_type;

    constructor(url: string, mode: modes_type) {
        this.url = url;
        this.mode = mode;
    }

    _callAPI(
        url: string,
        method: GoogleAppsScript.URL_Fetch.HttpMethod,
        body?: object,
    ): {
        response?: any;
        status?: number;
        error: boolean;
    } {
        let res: GoogleAppsScript.URL_Fetch.HTTPResponse;
        try {
            if (["post", "put", "patch"].includes(method)) {
                res = UrlFetchApp.fetch(url, {
                    method: method,
                    headers: {
                        "Content-Type": "application/json",
                    },
                    payload: JSON.stringify(body),
                });
            } else {
                res = UrlFetchApp.fetch(url, {
                    method: method,
                });
            }
        } catch (e) {
            console.error(e);
            return { error: true };
        }

        return {
            response: res.getContentText(),
            status: res.getResponseCode(),
            error: false,
        };
    }

    send(from: string, date: Date | GoogleAppsScript.Base.Date, subject: string, message: string) {
        // 文字数制限を適用
        const truncatedFrom = truncateString(from, 80);
        const truncatedSubject = truncateString(subject, 256);
        const truncatedBody = truncateString(message, 1024);

        let body = {};
        if (this.mode == "Slack") {
            const content = [
                `*件名*: ${truncatedSubject}`,
                `*送信元*: ${truncatedFrom}`,
                `*受信日時*: ${formatDate(date)}`,
            ].join("\n");
            body = {
                blocks: [
                    {
                        type: "section",
                        text: {
                            type: "mrkdwn",
                            text: content,
                        }
                    },
                    {
                        type: "rich_text",
                        elements: [
                            {
                                type: "rich_text_quote",
                                elements: [
                                    {
                                        type: "text",
                                        text: truncatedFrom,
                                        style: {
                                            bold: true,
                                        }
                                    }
                                ]
                            },
                            {
                                type: "rich_text_quote",
                                elements: [
                                    {
                                        type: "text",
                                        text: truncatedSubject,
                                        style: {
                                            bold: true,
                                        }
                                    }
                                ]
                            },
                            {
                                type: "rich_text_quote",
                                elements: [
                                    {
                                        type: "text",
                                        text: truncatedBody,
                                    }
                                ]
                            },
                            {
                                type: "rich_text_quote",
                                elements: [
                                    {
                                        type: "date",
                                        timestamp: date.getTime(),
                                        format: "{ago}"
                                    }
                                ]
                            }
                        ]
                    }
                ]
            };
        } else if (this.mode == "Discord") {
            body = {
                username: truncatedFrom,
                content: [
                    `**件名**: ${truncatedSubject}`,
                    `**送信元**: ${truncatedFrom}`,
                    `**受信日時**: ${formatDate(date)}`,
                ].join("\n"),
                embeds: [
                    {
                        title: truncatedSubject,
                        type: "rich",
                        timestamp: date.toISOString(),
                        color: 15559967, // #ed6d1f を10進数に変換したもの
                        description: truncatedBody,
                        author: {
                            name: truncatedFrom,
                        },
                    },
                ],
            };
        }
        const result = this._callAPI(this.url, "post", body);
        if (result.error) {
            throw new Error("Failed to call webhook");
        }
    }
}

function getProperty(key: string): string {
    const properties = PropertiesService.getScriptProperties();
    return properties.getProperty(key) ?? "";
}

function setProperty(key: string, value: string): GoogleAppsScript.Properties.Properties {
    const properties = PropertiesService.getScriptProperties();
    return properties.setProperty(key, value);
}

function formatDate(date: Date | GoogleAppsScript.Base.Date) {
    return (
        `${date.getFullYear()}/${zeroPadding(date.getMonth() + 1, 2)}/${zeroPadding(date.getDate(), 2)} ` +
        `${zeroPadding(date.getHours(), 2)}:${zeroPadding(date.getMinutes(), 2)}:${zeroPadding(date.getSeconds(), 2)}`
    );
}

function zeroPadding(value: number | string, diget: number): string {
    return String(value).padStart(diget, "0");
}

/**
 * 文字列を指定した最大長に制限する
 * @param str 対象の文字列
 * @param maxLength 最大文字数
 * @returns 制限された文字列
 */
function truncateString(str: string, maxLength: number): string {
    if (str.length <= maxLength) {
        return str;
    }
    return str.substring(0, maxLength);
}
