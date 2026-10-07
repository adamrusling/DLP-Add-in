/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global require, module */
const { getCategoryRiskScores } = require("./dlpScanner");
const { DEFAULT_SCORING_CONFIG } = require("./riskScoring");

// Returns each risk band with a display range such as "6-15" or "41+".
function getRiskBands() {
  return DEFAULT_SCORING_CONFIG.riskLevels.map((level) => ({
    name: level.name,
    range: Number.isFinite(level.max) ? `${level.min}-${level.max}` : `${level.min}+`,
  }));
}

// Returned as plain text, shown before the per-category risk score table.
function buildHelpIntro() {
  return [
    "1. Open the DLP Add-in tab in Outlook when reading or composing an email.",
    "2. Select or compose the email you want to review.",
    "3. Click the Scan Email button to inspect the contents of the email for potential Personally Identifiable Information (PII) or Protected Health Information (PHI) patterns.",
    "4. Review the report and recommendations, then remove or redact any risky content before sending the message.",
    "5. Use the total risk score and banding to prioritise review of high-risk content.",
    "",
    "How the risk score is calculated:",
    "- Each detected category has its own fixed risk score as shown in the table below.",
    "- Each category\u2019s contribution to the total is its risk score multiplied by the number of times it was found (Risk Score × Count).",
    "- The overall total risk score is the sum of every category\u2019s contribution, which is then used to determine the risk band below.",
    "- If two or more high-sensitivity categories (driving licence, passport, national insurance, NHS number, bank details, credit card) are found, 10 points are added.",
  ].join("\n");
}

// Returned as plain text, shown after the per-category risk score table.
function buildHelpOutro() {
  return [
    "Limitations:",
    "- The add-in may not detect all instances of sensitive information, particularly if it is embedded in images.",
    "- Encrypted or password-protected attachments cannot be scanned by the add-in, but are not considered as risks due to their encryption.",
  ].join("\n");
}

function buildHelpMessage() {
  const categoryLines = getCategoryRiskScores().map(
    (category) => `- ${category.label}: ${category.score}`
  );

  return [
    buildHelpIntro(),
    "",
    "Total risk scoring:",
    ...getRiskBands().map((band) => `- ${band.range} = ${band.name}`),
    "",
    "Category risk scores:",
    ...categoryLines,
    "",
    buildHelpOutro(),
  ].join("\n");
}

module.exports = {
  buildHelpMessage,
  buildHelpIntro,
  buildHelpOutro,
  getCategoryRiskScores,
  getRiskBands,
};
