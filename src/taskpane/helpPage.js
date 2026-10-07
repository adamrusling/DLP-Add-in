/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */
const {
  buildHelpIntro,
  buildHelpOutro,
  getCategoryRiskScores,
  getRiskBands,
} = require("./help.js");

// Builds the colour-banded total risk score table, reusing the scan report's band colours.
function buildRiskBandTable() {
  const table = document.createElement("table");
  const header = document.createElement("tr");
  ["Total risk score", "Risk band"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    header.appendChild(th);
  });
  table.appendChild(header);

  getRiskBands().forEach((band) => {
    const row = document.createElement("tr");

    const rangeCell = document.createElement("td");
    rangeCell.textContent = band.range;
    rangeCell.style.textAlign = "center";
    row.appendChild(rangeCell);

    const bandCell = document.createElement("td");
    const badge = document.createElement("span");
    badge.className = `risk-band risk-band--${band.name.toLowerCase().replace(/\s+/g, "-")}`;
    badge.textContent = band.name;
    bandCell.appendChild(badge);
    row.appendChild(bandCell);

    table.appendChild(row);
  });

  return table;
}

// Builds the per-category risk score table shown between the intro and outro guide text.
function buildCategoryRiskTable() {
  const table = document.createElement("table");
  const header = document.createElement("tr");
  ["Category", "Risk score"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    header.appendChild(th);
  });
  table.appendChild(header);

  getCategoryRiskScores().forEach((category) => {
    const row = document.createElement("tr");

    const labelCell = document.createElement("td");
    labelCell.textContent = category.label;
    row.appendChild(labelCell);

    const scoreCell = document.createElement("td");
    scoreCell.textContent = String(category.score);
    scoreCell.style.textAlign = "center";
    row.appendChild(scoreCell);

    table.appendChild(row);
  });

  return table;
}

Office.onReady(() => {
  const output = document.getElementById("help-output");
  if (!output) {
    return;
  }

  output.innerHTML = "";

  const intro = document.createElement("pre");
  intro.textContent = buildHelpIntro();
  output.appendChild(intro);

  const bandHeading = document.createElement("pre");
  bandHeading.textContent = "Total risk scoring:";
  output.appendChild(bandHeading);
  output.appendChild(buildRiskBandTable());

  const categoryHeading = document.createElement("pre");
  categoryHeading.textContent = "Category risk scores:";
  output.appendChild(categoryHeading);
  output.appendChild(buildCategoryRiskTable());

  const outro = document.createElement("pre");
  outro.textContent = buildHelpOutro();
  output.appendChild(outro);
});
