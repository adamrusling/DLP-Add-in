/* global require, module */
const { scanEmailForPII, scanAttachmentForPII, mergeScanReports } = require("./dlpScanner");
const { readSupportedAttachments } = require("./attachmentScanner");

async function scanMessageForPII(item, subject, body) {
  const emailReport = scanEmailForPII(subject, body);
  const attachments = await readSupportedAttachments(item);
  const attachmentReports = attachments
    .filter((attachment) => attachment.text)
    .map((attachment) => scanAttachmentForPII(attachment.name, attachment.text));
  const attachmentReport = mergeScanReports(attachmentReports);
  const report = mergeScanReports([emailReport, ...attachmentReports]);

  report.emailFindings = emailReport.findings;
  report.attachmentFindings = attachmentReport.findings;
  report.attachmentMessages = attachments
    .filter((attachment) => attachment.error)
    .map((attachment) => `${attachment.name}: ${attachment.error}`);

  return report;
}

module.exports = {
  scanMessageForPII,
};
