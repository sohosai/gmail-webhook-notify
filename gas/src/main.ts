const PROPERTY_NAME_PAST_MAILS = "pastMails" as const;
const MODES = ["Discord", "Slack"] as const;
type modes_type = (typeof MODES)[number];

function main() {
    const mode = getProperty("MODE");
    if (!(mode === MODES[0] || mode === MODES[1])) {
        console.error("The mode is invalid. It must be Discord or Slack");
        return;
    }

    const webhook = new Webhook(getProperty("WEBHOOK_URL"), mode);

    let pastMails = loadIDs(PROPERTY_NAME_PAST_MAILS);

    const now = new Date();
    const myEmail = Session.getActiveUser().getEmail();

    // pastMailsの掃除
    pastMails = pastMails.filter((id) => {
        const mail = GmailApp.getMessageById(id);
        const date = mail.getDate();
        // 12時間以内に来たものだけを抽出
        return now.getTime() - date.getTime() < 1000 * 60 * 60 * 12;
    });

    let shouldContinue = true;
    for (let i = 0; shouldContinue; i++) {
        // 2件ずつ取得
        const threads = GmailApp.getInboxThreads(i, i + 2);
        threads.forEach((thread) => {
            shouldContinue = false;
            const messages = thread.getMessages();
            messages.forEach((message) => {
                const id = message.getId();
                const date = message.getDate();
                const from = message.getFrom();

                if (
                    from != myEmail &&
                    now.getTime() - date.getTime() < 1000 * 60 * 60 * 12 &&
                    !pastMails.includes(id)
                ) {
                    shouldContinue = true;
                    // 送信元が自分のメールアドレスでなく、かつメールが12時間以内に来たもので、かつ通知済みでないとき
                    pastMails = pastMails.filter((i) => id !== i);

                    try {
                        webhook.send(from, date, message.getSubject());
                        // 送信部分でエラーが発生した場合、即座にtry内から抜けるため以下は実行されない
                        pastMails.push(id);
                    } catch (e) {
                        console.error("Failed to send a message: ", e);
                    }
                }
            });
        });
    }
    saveIDs(PROPERTY_NAME_PAST_MAILS, pastMails);
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

    send(from: string, date: Date | GoogleAppsScript.Base.Date, subject: string) {
        let body = {};
        if (this.mode == "Slack") {
            const content = `*送信元*: ${from} \n` + `*受信日時*: ${formatDate(date)}\n` + `*件名*: ${subject}`;

            body = {
                attachments: [
                    {
                        color: "#ed6d1f",
                        blocks: [
                            {
                                type: "section",
                                text: {
                                    type: "mrkdwn",
                                    text: content,
                                },
                            },
                        ],
                    },
                ],
            };
        } else if (this.mode == "Discord") {
            body = {
                embeds: [
                    {
                        title: subject,
                        type: "rich",
                        timestamp: date.toISOString(),
                        color: 15559967, // #ed6d1f を10進数に変換したもの
                        author: {
                            name: from,
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

function saveIDs(basePropertyName: string, content: string[]) {
    const properties = PropertiesService.getScriptProperties();
    const keys = properties.getKeys().filter((k) => k.includes(basePropertyName));
    keys.forEach((key) => {
        properties.deleteProperty(key);
    });

    // 1つのプロパティーにつき6KBまでで、IDひとつにつき17バイトなので、350個単位に分割する
    const num = Math.ceil(content.length / 350);
    for (let i = 0; i < num; i++) {
        setProperty(`${basePropertyName}${i}`, content.slice(i * 350, i * 350 + 349).join(","));
    }
}

function loadIDs(basePropertyName: string): string[] {
    const properties = PropertiesService.getScriptProperties();
    const keys = properties.getKeys().filter((k) => k.includes(basePropertyName));
    const value = keys
        .map((key) => {
            const v = getProperty(key);
            if (v) {
                return v.split(",");
            } else {
                return [];
            }
        })
        .flat();

    return value;
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
        `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()} ` +
        `${zeroPadding(date.getHours(), 2)}:${zeroPadding(date.getMinutes(), 2)}:${zeroPadding(date.getSeconds(), 2)}`
    );
}

function zeroPadding(value: number | string, diget: number): string {
    return String(value).padStart(diget, "0");
}
