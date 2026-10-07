/* global Office, require */
const { scanEmailForPII } = require("../taskpane/dlpScanner");

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

async function onMessageSendHandler(event) {
  try {
    const subject = await getSubject();
    const body = await getBody();
    const result = checkDlp(subject, body);

    if (result.block) {
      event.completed({
        allowEvent: false,
        errorMessage: result.reason,
      });

      return;
    }

    event.completed({
      allowEvent: true,
    });
  } catch {
    event.completed({
      allowEvent: false,
      errorMessage: "Unable to validate email.",
    });
  }
}

function checkDlp(subject, body) {
  const report = scanEmailForPII(subject, body);
  return {
    block: report.findings.length > 0,
    reason: report.summary,
  };
}

function getBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

function getSubject() {
  return Promise.resolve(Office.context.mailbox.item.subject);
}
