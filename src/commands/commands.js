/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, require, clearTimeout, setTimeout, console */
const { scanMessageForPII } = require("../taskpane/scanMessage");
const { buildHelpMessage } = require("../taskpane/help.js");

Office.onReady(() => {
  Office.actions.associate("action", action);
  Office.actions.associate("showHelpAction", showHelp);
});

/**
 * Reads the compose message body and subject and reports DLP guidance.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  let completed = false;
  let timeoutId;
  const complete = () => {
    if (completed) {
      return;
    }

    completed = true;
    clearTimeout(timeoutId);
    event.completed();
  };

  timeoutId = setTimeout(complete, 30000);

  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
      complete();
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
      try {
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          complete();
          return;
        }

        getSubjectText(item)
          .then((subject) => {
            if (completed) {
              return;
            }

            try {
              scanMessageForPII(item, subject, bodyResult.value || "")
                .then((report) => {
                  if (completed) {
                    return;
                  }

                  const riskScore = report.totalRiskScore || 0;
                  const riskBand = report.riskBand;
                  const messageText = report.findings.length
                    ? `Potential PII/PHI found. Total risk score (${riskScore}) - ${riskBand}. Read full feedback in the task pane.`
                    : "No PII/PHI found. Read full feedback in the task pane.";
                  const message = {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: fitNotificationMessage(messageText),
                    icon: "Icon.80x80",
                    persistent: true,
                  };

                  item.notificationMessages.replaceAsync(
                    "DlpScanNotification",
                    message,
                    (result) => {
                      if (result.status !== Office.AsyncResultStatus.Succeeded) {
                        console.error("Could not display the DLP scan result.", result.error);
                      }
                      complete();
                    }
                  );
                })
                .catch((error) => {
                  console.error("Could not scan the email and its attachments.", error);
                  complete();
                });
            } catch (error) {
              console.error("Scan Email ribbon action failed.", error);
              complete();
            }
          })
          .catch((error) => {
            console.error("Could not read the email subject.", error);
            complete();
          });
      } catch (error) {
        console.error("Scan Email ribbon action failed.", error);
        complete();
      }
    });
  } catch (error) {
    console.error("Scan Email ribbon action failed.", error);
    complete();
  }
}

function fitNotificationMessage(message) {
  const characters = Array.from(message);
  const maxLength = 150;

  if (characters.length <= maxLength) {
    return message;
  }

  return `${characters.slice(0, maxLength - 3).join("")}...`;
}

function getSubjectText(item) {
  if (typeof item.subject === "string") {
    return Promise.resolve(item.subject);
  }

  if (!item.subject || typeof item.subject.getAsync !== "function") {
    return Promise.resolve("");
  }

  return new Promise((resolve) => {
    item.subject.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  });
}

/**
 * Displays the DLP add-in Help dialog with the documented Outlook limitations.
 * @param event {Office.AddinCommands.Event}
 */
function showHelp(event) {
  const helpUrl = "https://localhost:3000/help.html";

  Office.context.ui.displayDialogAsync(
    helpUrl,
    {
      height: 70,
      width: 35,
      displayInIframe: false,
    },
    (dialogResult) => {
      if (dialogResult.status === Office.AsyncResultStatus.Failed) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("DlpHelpNotification", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: buildHelpMessage(),
          icon: "Icon.80x80",
          persistent: true,
        });
      }

      event.completed();
    }
  );
}
