/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global clearTimeout, console, document, Office, require, setTimeout */
const { scanMessageForPII } = require("./scanMessage");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("scan-email-btn").addEventListener("click", scanEmail);
    document.getElementById("back-btn").addEventListener("click", showHomeScreen);
  }
});

function showResultsScreen() {
  document.getElementById("app-header").style.display = "none";
  document.getElementById("back-bar").style.display = "flex";
  document.getElementById("scan-email-btn").textContent = "Rescan Email";
}

function showHomeScreen() {
  document.getElementById("app-header").style.display = "flex";
  document.getElementById("back-bar").style.display = "none";
  document.getElementById("output").innerHTML = "";
  document.getElementById("scan-email-btn").textContent = "Scan Email";
}

async function scanEmail() {
  const button = document.getElementById("scan-email-btn");
  const output = document.getElementById("output");

  output.innerHTML = "";
  const status = document.createElement("p");
  status.textContent = "Scanning...";
  output.appendChild(status);
  button.disabled = true;

  let timeoutId;
  try {
    const timeout = new Promise((resolve, reject) => {
      timeoutId = setTimeout(() => reject(new Error("The scan timed out.")), 60000);
    });
    await Promise.race([runScan(), timeout]);
  } catch (error) {
    const message = (error && error.message) || String(error);
    output.innerHTML = "";
    const errorText = document.createElement("p");
    errorText.textContent = `Scan failed: ${message}`;
    output.appendChild(errorText);
    showResultsScreen();
    console.error("Scan failed", error);
  } finally {
    clearTimeout(timeoutId);
    button.disabled = false;
  }
}

async function runScan() {
  const item = Office.context.mailbox.item;

  const bodyText = await new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
      if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(bodyResult.value);
      } else {
        reject(bodyResult.error);
      }
    });
  });

  const subjectText = item.subject || "";
  const report = await scanMessageForPII(item, subjectText, bodyText);
  showResultsScreen();
  renderReport(report);
}

function buildFindingsTable(findings) {
  const table = document.createElement("table");
  const header = document.createElement("tr");
  ["Type", "Examples", "Count", "Risk score"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    header.appendChild(th);
  });
  table.appendChild(header);

  findings.forEach((finding) => {
    const row = document.createElement("tr");

    const typeCell = document.createElement("td");
    typeCell.textContent = finding.label;
    row.appendChild(typeCell);

    const examplesCell = document.createElement("td");
    examplesCell.textContent = finding.examples.join(", ");
    row.appendChild(examplesCell);

    const countCell = document.createElement("td");
    countCell.textContent = String(finding.count);
    row.appendChild(countCell);

    const riskCell = document.createElement("td");
    riskCell.textContent = String(finding.score);
    row.appendChild(riskCell);

    table.appendChild(row);
  });

  return table;
}

function appendFindingsSection(output, headingText, findings) {
  const section = document.createElement("section");
  section.className = "detected-categories";

  const heading = document.createElement("div");
  heading.className = "output-section-heading";
  heading.textContent = headingText;
  section.appendChild(heading);

  if (findings.length === 0) {
    const clear = document.createElement("p");
    clear.textContent = "No confident PII/PHI patterns found.";
    section.appendChild(clear);
  } else {
    section.appendChild(buildFindingsTable(findings));
  }

  output.appendChild(section);
}

function renderReport(report) {
  const output = document.getElementById("output");
  const emailFindings = report.emailFindings || report.findings || [];
  const attachmentFindings = report.attachmentFindings || [];

  output.innerHTML = "";

  const summaryHeading = document.createElement("div");
  summaryHeading.className = "output-section-heading";
  summaryHeading.textContent = "Summary";
  output.appendChild(summaryHeading);

  const summary = document.createElement("p");
  const summaryMessages = [report.summary];
  (report.attachmentMessages || []).forEach((message) => {
    summaryMessages.push(`Attachment warning: ${message}`);
  });
  summary.textContent = summaryMessages.join("\n\n");
  output.appendChild(summary);

  if (report.totalRiskScore > 0) {
    const totalRisk = document.createElement("p");
    totalRisk.textContent = `Total data risk score: ${report.totalRiskScore} (${report.riskBand})`;
    totalRisk.className = `risk-band risk-band--${report.riskBand.toLowerCase().replace(/\s+/g, "-")}`;
    output.appendChild(totalRisk);
  }

  if (report.recommendations.length > 0) {
    const suggestionsHeading = document.createElement("div");
    suggestionsHeading.className = "output-section-heading";
    suggestionsHeading.textContent = "Suggestions";
    output.appendChild(suggestionsHeading);

    const hintList = document.createElement("ul");
    report.recommendations.forEach((recommendation) => {
      const li = document.createElement("li");
      li.textContent = recommendation;
      hintList.appendChild(li);
    });

    output.appendChild(hintList);
  }

  // Keep email and attachment findings in separate sections.
  if (emailFindings.length > 0) {
    appendFindingsSection(output, "✉ Detected categories in email subject and body", emailFindings);
  }
  if (attachmentFindings.length > 0) {
    appendFindingsSection(output, "🔗 Detected categories in attachments", attachmentFindings);
  }
}
