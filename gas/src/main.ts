const PROPERTY_NAME_PAST_MAILS = "pastMails" as const;

function main() {
  const webhook = new Webhook(getProperty("WEBHOOK_URL"));

  let pastMails = loadIDs(PROPERTY_NAME_PAST_MAILS);

  const now = new Date();

  // pastMailsの掃除
  pastMails = pastMails.filter((id) => {
    const mail = GmailApp.getMessageById(id);
    const date = mail.getDate();
    // 24時間以内に来たものだけを抽出
    return now.getTime() - date.getTime() < 1000 * 60 * 60 * 24;
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

        if (
          now.getTime() - date.getTime() < 1000 * 60 * 60 * 24 &&
          !pastMails.includes(id)
        ) {
          shouldContinue = true;
          // メールが24時間以内に来たもので、かつ通知済みでないとき
          pastMails = pastMails.filter((i) => id !== i);
          const content =
            `送信元: ${message.getFrom()} \n` +
            `受信日時: ${formatDate(date)}\n` +
            `件名: ${message.getSubject()}`;

          try {
            webhook.send(content);
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

  constructor(url: string) {
    this.url = url;
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

  send(message: string) {
    const result = this._callAPI(this.url, "post", { text: message });
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
    setProperty(
      `${basePropertyName}${i}`,
      content.slice(i * 350, i * 350 + 349).join(","),
    );
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

function setProperty(
  key: string,
  value: string,
): GoogleAppsScript.Properties.Properties {
  const properties = PropertiesService.getScriptProperties();
  return properties.setProperty(key, value);
}

function formatDate(date: Date | GoogleAppsScript.Base.Date) {
  return `${date.getFullYear()}/${
    date.getMonth() + 1
  }/${date.getDate()} ${date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`;
}
