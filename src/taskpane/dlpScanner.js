/* global module, require */

const { DEFAULT_SCORING_CONFIG, calculateRiskScore } = require("./riskScoring");

// Trim scanned text and ignore case when comparing it.
function normalise(text) {
  return (text || "").toString().trim().toLowerCase();
}

// Use the catalogue weights when scoring findings.
function scoreFindings(findings) {
  return calculateRiskScore(
    patternCatalogue,
    findings.map((finding) => ({ label: finding.label, count: finding.count })),
    DEFAULT_SCORING_CONFIG
  );
}

// Apply a pattern and discard matches rejected by its filter.
function findPatternMatches(text, pattern) {
  const matches = [];

  for (const match of text.matchAll(pattern.regex)) {
    const value = match[0];
    const matchIndex = match.index ?? 0;

    if (pattern.filter && !pattern.filter(text, matchIndex, value)) {
      continue;
    }

    matches.push(value);
  }

  return matches;
}

// Six-digit values that look like years are usually dates, not sort codes.
function isLikelyDateLike(value) {
  const digitsOnly = value.replace(/\D/g, "");

  if (digitsOnly.length !== 6) {
    return false;
  }

  return /^20\d{4}$/.test(digitsOnly) || /^19\d{4}$/.test(digitsOnly);
}

// Treat long digit runs as card numbers unless nearby text points to bank details.
function isLikelyCreditCardChunk(value, contextText) {
  const digitsOnly = value.replace(/\D/g, "");
  const cleanedValue = value.replace(/[\s:#-]+/g, "").replace(/^IBAN/i, "");
  const looksLikeIban =
    /^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/i.test(cleanedValue) &&
    cleanedValue.length >= 15 &&
    cleanedValue.length <= 34;

  if (looksLikeIban) {
    return false;
  }

  if (digitsOnly.length >= 13 && !/(bank|account|sort\s*code|iban)/.test(contextText)) {
    return true;
  }

  return false;
}

// A UK phone number has 10 or 11 digits starting with 0, even if it has a leading +.
function isLikelyPhoneNumber(value) {
  const trimmedValue = (value || "").trim();
  const digitsOnly = trimmedValue.replace(/\D/g, "");

  // Ignore empty or too short/long values before checking the UK number shape.
  if (!digitsOnly || digitsOnly.length < 9 || digitsOnly.length > 15) {
    return false;
  }

  // Require an explicit UK-style prefix and a local number length.
  return /^[+0]/.test(trimmedValue) && /^0\d{9,10}$/.test(digitsOnly);
}

// Reject numeric fragments, but keep them when adjacent to a plausible account number.
function isPartOfLongerNumericSequence(text, index, valueLength) {
  const before = text.slice(Math.max(0, index - 1), index);
  const after = text.slice(index + valueLength, Math.min(text.length, index + valueLength + 1));

  if (/\d/.test(before) || /\d/.test(after)) {
    const nextAccount = text.slice(index + valueLength).match(/^\s*(\d{8})/);
    const previousAccount = text.slice(0, index).match(/(\d{8})\s*$/);

    if (
      (nextAccount && nextAccount[1].length === 8) ||
      (previousAccount && previousAccount[1].length === 8)
    ) {
      return false;
    }

    return true;
  }

  return false;
}

// Return nearby text in lowercase for contextual checks such as "sort code".
function getNearbyContext(text, index, valueLength, padding = 40) {
  const start = Math.max(0, index - padding);
  const end = Math.min(text.length, index + valueLength + padding);

  return text.slice(start, end).toLowerCase();
}

// Check whether an eight-digit account number sits beside the match.
function hasAdjacentBankAccount(text, index, valueLength) {
  const beforeText = text.slice(Math.max(0, index - 16), index).replace(/\s+/g, "");
  const afterText = text
    .slice(index + valueLength, Math.min(text.length, index + valueLength + 16))
    .replace(/\s+/g, "");

  return /^\d{8}$/.test(beforeText) || /^\d{8}$/.test(afterText);
}

function hasBankContext(text, index, valueLength) {
  const nearbyText = getNearbyContext(text, index, valueLength);

  return /(sort\s*code|sort-code|account\s*(?:number|no\.?|no)|bank\s*account|bank\s*details|iban|international\s+bank\s+account\s+number)/.test(
    nearbyText
  );
}

// An account number can be identified by a sort code immediately before it.
function hasNearbySortCode(text, index) {
  const before = text.slice(Math.max(0, index - 30), index).toLowerCase();
  return /\b\d{2}[-\s]?\d{2}[-\s]?\d{2}\b/.test(before);
}

function isLikelySortCode(value) {
  const digitsOnly = value.replace(/\D/g, "");

  return /^\d{6}$/.test(digitsOnly);
}

// Check nearby digits and phone formats so part of a phone number is not flagged as a bank detail.
function isPartOfPhoneNumber(text, index, valueLength) {
  const start = Math.max(0, index - 20);
  const end = Math.min(text.length, index + valueLength + 20);
  const surroundingText = text.slice(start, end);
  const digitsOnly = surroundingText.replace(/\D/g, "");

  if (digitsOnly.length < 9) {
    return false;
  }

  return /(?:\(\d{2,5}\)[-.\s]?\d{3,4}[-.\s]?\d{3,4}|(?:\+\d{1,4}\s*(?:\(?0\)?\s*)?|0\s*)\d{2,5}[-.\s]?\d{3}[-.\s]?\d{3,6})/.test(
    surroundingText
  );
}

// Do not treat a value beside a VAT, NHS, or NI label as bank data unless bank wording is also nearby.
function hasSensitiveNumericContext(text, index, value) {
  const start = Math.max(0, index - 18);
  const end = Math.min(text.length, index + value.length + 18);
  const context = text.slice(start, end).toLowerCase();

  if (
    /(sort\s*code|sort-code|account\s*(?:number|no\.?|no)|bank\s*account|bank\s*details|iban)/.test(
      context
    )
  ) {
    return false;
  }

  return /(vat|nhs|national\s+insurance|ni\s*number|tax\s+identifier|number\s*:)/.test(context);
}

// Prevent VAT, NHS, or NI labels from making a neighbouring value look like bank details.
function hasSensitiveIdentifierLabelContext(text, index, valueLength) {
  const start = Math.max(0, index - 18);
  const end = Math.min(text.length, index + valueLength + 18);
  const context = text.slice(start, end).toLowerCase();

  return /(?:\b(?:vat|tax|nhs|national\s+insurance|ni)\b\s*(?:number|identifier)|\b(?:number|identifier)\b\s*(?:vat|tax|nhs|national\s+insurance|ni)\b)/.test(
    context
  );
}

// A line break can split a sort code into four and two digits (for example, "7654 12").
function isSplitSortCodeFragment(value) {
  const plain = value.replace(/\s+/g, " ").trim();

  return /^\d{4}\s+\d{2}$/.test(plain);
}

// Apply the shared false-positive checks for sort codes, account numbers, and IBANs.
function shouldRejectBankDetail(text, index, value) {
  const nearbyText = getNearbyContext(text, index, value.length);
  const digitsOnly = value.replace(/\D/g, "");

  const hasSortCodeLabel =
    /sort\s*-?\s*code/i.test(value) ||
    /sort\s*-?\s*code\W{0,3}$/.test(text.slice(Math.max(0, index - 15), index).toLowerCase());

  // Sort codes starting 19/20 look like years, so a "sort code" label directly before overrides the date check.
  if (
    (isLikelyDateLike(value) && !hasSortCodeLabel) ||
    isLikelyCreditCardChunk(value, nearbyText)
  ) {
    return true;
  }

  if (digitsOnly.length === 6) {
    // A sort code is only "explicit" when it uses the standard NN-NN-NN separators or is labelled as such;
    // otherwise a bare six-digit number needs the extra checks below before it is trusted as bank data.
    const isExplicitSortCode =
      /^\d{2}[-\s]\d{2}[-\s]\d{2}$/.test(value) || /(?:sort\s*code|sort-code)/i.test(nearbyText);

    if (isSplitSortCodeFragment(value)) {
      return true;
    }

    if (isExplicitSortCode) {
      if (isPartOfPhoneNumber(text, index, value.length)) {
        return true;
      }

      return (
        hasSensitiveNumericContext(text, index, value) &&
        !/(sort\s*code|sort-code|account\s*(?:number|no\.?|no)|bank\s*account|bank\s*details|iban)/i.test(
          nearbyText
        )
      );
    }

    if (
      isPartOfPhoneNumber(text, index, value.length) ||
      hasSensitiveNumericContext(text, index, value)
    ) {
      return true;
    }
  }

  if (
    hasSensitiveIdentifierLabelContext(text, index, value.length) &&
    !/\b(?:bank|account|sort\s*code|iban)/i.test(nearbyText)
  ) {
    return true;
  }

  return (
    isPartOfLongerNumericSequence(text, index, value.length) &&
    !hasAdjacentBankAccount(text, index, value.length)
  );
}

// Require a contiguous IBAN-shaped value without a neighbouring VAT or NHS label.
function isValidIbanCandidate(value) {
  const cleanedValue = value.replace(/[\s:#-]+/g, "").replace(/^IBAN/i, "");

  return (
    /^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/i.test(cleanedValue) &&
    cleanedValue.length >= 15 &&
    cleanedValue.length <= 34 &&
    !/\b(?:vat|nhs|number)\b/i.test(cleanedValue)
  );
}

// Remove markup, quoted history, headers, hidden links, and zero-width characters before scanning.
// Keep line breaks so digits on separate lines cannot be mistaken for one number.
function getVisibleTextOnly(text) {
  return (text || "")
    .replace(/\r\n?/g, "\n")
    .replace(/\u200B|\u200C|\u200D|\u2060|\uFEFF/g, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/(?:^|\n)\s*>.*$/gm, " ")
    .replace(/(?:^|\n)\s*(?:from|sent|to|cc|subject):.*$/gim, " ")
    .replace(/(?:^|\n)\s*(?:[-*_]{2,}|--|__|\.{3,})\s*$/gm, " ")
    .replace(/(?:^|[\s<"'()])(?:[a-z][a-z0-9+.-]*):[^\s<"'()]+/gi, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

// The scanner and help-page score table use this same set of rules.
const patternCatalogue = [
  {
    type: "name",
    label: "Name",
    weight: 1,
    // The word after the title must be a capitalised proper noun (e.g. "Dr Jones"), not any lowercase
    // word, otherwise generic uses of "client"/"patient"/"professor" (e.g. "the client would...") false-positive.
    // Only mr/mrs/ms/dr take the optional trailing full stop, since it's a genuine abbreviation for those;
    // allowing it after "client"/"patient"/"professor" would match a sentence-ending full stop followed by
    // the next capitalised sentence
    // Matches a title or role followed by one to three capitalised name words.
    regex:
      /\b(?:(?:[Mm]r|[Mm]rs|[Mm]s|[Mm]iss|[Dd]r)\.?|[Pp]rofessor|[Pp]atient|[Cc]lient)\s+[A-Z][a-z]+(?:[ \t]+[A-Z][a-z]+){0,2}\b/g,
    guidance:
      "Remove the individual’s name or generalise it to a role such as “the clinician” or “the patient”. Alternatively, consider using the individual’s initials or a pseudonym if necessary for context.",
  },
  {
    type: "email",
    label: "Email",
    weight: 2,
    // Matches a local part, domain, and a suffix of at least two letters.
    regex: /\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b/gi,
    guidance:
      "Remove the full email address or replace it with a role-based mailbox such as support@example.com.",
  },
  {
    type: "telephone",
    label: "Telephone Number",
    weight: 3,
    // The first alternative handles a UK area code wrapped entirely in parentheses (e.g. "(01204) 385990"),
    // which the general digit-grouping categories below cannot replicate because of the closing bracket.
    regex:
      /(?<!\w)(?:\(\d{2,5}\)[-.\s]?\d{3,4}[-.\s]?\d{3,4}|(?:(?:\+\d{1,4}\s*)|(?:\(?0\)?\s*))?(?:\d{2,5}[-.\s]?\d{3,6}[-.\s]?\d{3,4}|\d{3}[-.\s]?\d{4}[-.\s]?\d{4}|\d{2,5}[-.\s]?\d{5,6}))(?!\w)/g,
    filter: (text, index, value) => {
      const digitsOnly = value.replace(/\D/g, "");
      if (digitsOnly.length < 9 || digitsOnly.length > 15) {
        return false;
      }

      const hasPhonePrefix = /^[+0(]/.test(value.trim());
      if (!hasPhonePrefix) {
        return false;
      }

      const start = Math.max(0, index - 20);
      const end = Math.min(text.length, index + value.length + 20);
      const nearbyText = text.slice(start, end).toLowerCase();
      const hasSensitiveQualifier =
        /(nhs|vat|national\s+insurance|driving\s+licence|passport|bank\s+details|sort\s+code|account\s+number|iban)/.test(
          nearbyText
        );

      if (hasSensitiveQualifier && !/\+\d{1,4}/.test(value)) {
        return false;
      }

      return true;
    },
    guidance:
      "Replace the phone number with a generic contact method or remove it from the message.",
  },
  {
    type: "dob",
    label: "Date of birth",
    weight: 4,
    // Matches DD/MM/YYYY, DD-MM-YYYY, DD-MMM-YYYY, "DD Month YYYY" and "Month DD, YYYY" with years 1900-2099;
    // the day-first forms also accept a two-digit year (e.g. 1/10/90).
    regex:
      /\b(?:(?:0?[1-9]|[12]\d|3[01])(?:st|nd|rd|th)?[-/.\s]+(?:(?:0?[1-9]|1[0-2])|(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?))[-/.,\s]+(?:(?:19|20)\d{2}|\d{2})|(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?\s+(?:0?[1-9]|[12]\d|3[01])(?:st|nd|rd|th)?,?\s+(?:19|20)\d{2})\b/gi,
    // Only flag dates that sit next to birth-related wording, so ordinary dates are ignored.
    filter: (text, index, value) =>
      /\b(?:dob|d\.o\.b\.?|date\s+of\s+birth|birth\s*date|birthday|born|b\/d)\b/i.test(
        getNearbyContext(text, index, value.length, 30)
      ),
    guidance:
      "Remove the exact birth date and use an age range or a non-identifying code if required.",
  },
  {
    type: "address",
    label: "Address",
    weight: 4,
    // Looks for a building number followed by a recognised street suffix.
    regex:
      /\b\d{1,5}\s+[a-z0-9.'\-\s]+(?:street|st\.|road|rd\.|avenue|ave\.|lane|ln\.|drive|dr\.|close|cl\.|court|ct\.|way|place|pl\.|crescent|crs\.|high street|highway)\b/gi,
    guidance:
      "Remove the address details and keep only a broad regional or service area description.",
  },
  {
    type: "driving_licence",
    label: "Driving Licence",
    weight: 10,
    highSensitivity: true,
    // Matches the 16-character UK format (5 letters, 6 digits, 2 initials, 1 digit, 2 letters), optionally space-separated.
    regex: /\b[A-Z]{5}\s?\d{6}\s?[A-Z\d]{2}\s?\d\s?[A-Z]{2}\b/g,
    guidance:
      "Remove or mask the driving licence number and use a non-identifying reference instead.",
  },
  {
    type: "passport",
    label: "Passport Number",
    weight: 10,
    highSensitivity: true,
    // Matches a letter plus seven or eight digits, or a bare nine-digit UK passport number when "passport" is nearby.
    regex: /\b(?:[A-Z]\d{7,8}|\d{9})\b/g,
    filter: (text, index, value) =>
      /^[A-Z]/.test(value) || /passport/i.test(getNearbyContext(text, index, value.length, 30)),
    guidance:
      "Remove the passport number and replace it with a case or reference code that does not expose the document ID.",
  },
  {
    type: "ni_number",
    label: "National Insurance Number",
    weight: 10,
    highSensitivity: true,
    // Matches two letters, six digits (optionally paired) and an A-D suffix; prefix letters are not restricted so test values like QQ are caught.
    regex: /\b[A-Z]{2}\s?\d{2}\s?\d{2}\s?\d{2}\s?[A-D]\b/g,
    guidance:
      "Remove the national insurance number and use a pseudonymous study reference instead of the direct identifier.",
  },
  // Three separate regexes detect bank details: sort codes, account numbers, and IBANs. Each is scored equally and shares the same guidance.
  {
    type: "bank_details",
    label: "Bank Details",
    weight: 10,
    highSensitivity: true,
    // Matches a six-digit sort code, with optional separators or a preceding label.
    regex:
      /(?<![A-Z0-9])(?:\d{2}[-\t ]?\d{2}[-\t ]?\d{2}|(?:sort\s*code|sort-code)\s*[:#-]?\s*\d{2}[-\t ]?\d{2}[-\t ]?\d{2})(?![A-Z0-9])/gi,
    filter: (text, index, value) => {
      const looksLikeSortCode = isLikelySortCode(value);

      // An unlabelled NN-NN-NN beside birth wording is a two-digit-year date of birth, not a sort code.
      if (
        !/sort\s*-?\s*code/i.test(value) &&
        /\b(?:dob|d\.o\.b\.?|date\s+of\s+birth|birth\s*date|birthday|born|b\/d)\b/i.test(
          getNearbyContext(text, index, value.length, 30)
        )
      ) {
        return false;
      }
      if (shouldRejectBankDetail(text, index, value)) {
        return false;
      }

      return (
        looksLikeSortCode &&
        !hasSensitiveNumericContext(text, index, value) &&
        !/\n/.test(value) &&
        !isPartOfPhoneNumber(text, index, value.length)
      );
    },
    guidance:
      "Remove or mask the bank details and use a neutral payment reference or a redacted placeholder instead.",
  },
  {
    type: "bank_details",
    label: "Bank Details",
    weight: 10,
    highSensitivity: true,
    // Matches an eight-digit account number, optionally introduced by bank wording.
    regex:
      /(?<!\d)\d{8}(?!\d)|(?:account\s*(?:number|no\.)|bank\s*account|bank\s*details)\s*[:#-]?\s*\d{8}/gi,
    filter: (text, index, value) => {
      const digitsOnly = value.replace(/\D/g, "");

      if (shouldRejectBankDetail(text, index, value)) {
        return false;
      }

      const hasAccountContext =
        hasBankContext(text, index, value.length) ||
        /(?:account\s*(?:number|no\.)|bank\s*account|bank\s*details)/i.test(value) ||
        hasNearbySortCode(text, index);

      return digitsOnly.length === 8 && hasAccountContext;
    },
    guidance:
      "Remove or mask the bank details and use a neutral payment reference or a redacted placeholder instead.",
  },
  {
    type: "bank_details",
    label: "Bank Details",
    weight: 10,
    highSensitivity: true,
    // Matches a contiguous IBAN-shaped code; the filter checks length and context.
    regex: /(?<![A-Z0-9])[A-Z]{2}\d{2}[A-Z0-9]{11,30}(?![A-Z0-9])/gi,
    filter: (text, index, value) => {
      if (shouldRejectBankDetail(text, index, value)) {
        return false;
      }

      const nearbyText = getNearbyContext(text, index, value.length);
      if (
        hasSensitiveIdentifierLabelContext(text, index, value.length) &&
        !/\b(?:bank|account|sort\s*code|iban)/i.test(nearbyText)
      ) {
        return false;
      }

      return isValidIbanCandidate(value);
    },
    guidance:
      "Remove or mask the bank details and use a neutral payment reference or a redacted placeholder instead.",
  },
  {
    type: "credit_card",
    label: "Credit card number",
    weight: 15,
    highSensitivity: true,
    // Matches four groups of four digits, allowing spaces or hyphens between groups.
    regex: /\b(?:\d{4}[\s-]?){3}\d{4}\b/g,
    guidance:
      "Remove the full credit card number and use a token, payment reference, or a redacted placeholder instead.",
  },
  {
    type: "vat",
    label: "VAT / tax identifier",
    weight: 2,
    // Matches a two-letter country prefix followed by 8 to 12 digits.
    regex: /\b[A-Z]{2}\s?\d{8,12}\b/gi,
    filter: (text, index, value) => {
      const cleanedValue = value.replace(/\s+/g, "");
      return /^[A-Z]{2}\d{8,12}$/i.test(cleanedValue);
    },
    guidance:
      "Remove or mask the VAT or tax identifier and use a business reference that does not disclose the direct tax number.",
  },
  {
    type: "nhs_number",
    label: "NHS Number",
    weight: 10,
    highSensitivity: true,
    // Matches ten digits, optionally grouped into threes and fours.
    regex: /\b\d{3}\s?\d{3}\s?\d{4}\b/g,
    filter: (text, index, value) => {
      const digitsOnly = value.replace(/\D/g, "");
      if (digitsOnly.length !== 10) {
        return false;
      }

      const context = text.slice(Math.max(0, index - 12), index + value.length + 12).toLowerCase();
      const hasInternationalPrefix = /(?:\+\s*\d{1,4}|\+\d{1,4})/.test(context);

      return !isLikelyPhoneNumber(value) && !hasInternationalPrefix;
    },
    guidance:
      "Remove the NHS number or replace it with a non-identifying study reference to avoid exposing the patient identifier.",
  },
];

// Return one score per category in ascending order; the three bank patterns share a score.
function getCategoryRiskScores() {
  const seenLabels = new Set();
  const categories = [];

  patternCatalogue.forEach((pattern) => {
    if (seenLabels.has(pattern.label)) {
      return;
    }

    seenLabels.add(pattern.label);
    categories.push({ label: pattern.label, score: pattern.weight });
  });

  // Stable sort keeps catalogue order within equal scores.
  return categories.sort((a, b) => a.score - b.score);
}

function groupFindings(findings) {
  return Array.from(
    findings
      .reduce((map, finding) => {
        const existing = map.get(finding.label);

        if (existing) {
          existing.count += finding.count;
          existing.examples = [...new Set([...existing.examples, ...finding.examples])].slice(0, 3);
        } else {
          map.set(finding.label, { ...finding });
        }

        return map;
      }, new Map())
      .values()
  ).sort((a, b) => b.score - a.score || b.count - a.count);
}

function buildScanReport(text, sourceLabel, summaryLabel) {
  const findings = [];

  patternCatalogue.forEach((pattern) => {
    const matches = findPatternMatches(text, pattern);
    if (matches.length > 0) {
      const examples = matches
        .slice(0, 3)
        .map((example) => (sourceLabel ? `"${sourceLabel}": ${example}` : example));
      findings.push({
        type: pattern.type,
        label: pattern.label,
        score: pattern.weight,
        count: matches.length,
        examples,
        guidance: pattern.guidance,
      });
    }
  });

  // Merge findings with the same label, including the three bank-detail patterns, before sorting by risk.
  const groupedFindings = groupFindings(findings);

  const recommendations = groupedFindings.map((finding) => finding.guidance);
  const uniqueRecommendations = [...new Set(recommendations)];
  const scoring = scoreFindings(groupedFindings);
  const totalRiskScore = scoring.finalScore;
  const riskBand = scoring.riskLevel;

  const summary = groupedFindings.length
    ? `Potential DLP issue: ${groupedFindings.length} PII/PHI categor${groupedFindings.length === 1 ? "y" : "ies"} detected in ${summaryLabel}. Overall risk score: ${totalRiskScore}. Review the suggestions and anonymise before sending.`
    : `No obvious PII/PHI patterns detected in ${summaryLabel}.`;

  return {
    summary,
    findings: groupedFindings,
    recommendations: uniqueRecommendations,
    totalRiskScore,
    riskBand,
    scoring,
    risk: groupedFindings.length > 0 ? "review" : "clear",
    scannedTextLength: text.length,
    normalisedText: normalise(text),
  };
}

// Scan the subject and visible body together, then report any matching categories.
function scanEmailForPII(subject, body) {
  const visibleText = getVisibleTextOnly(body);
  const text = `${subject || ""}\n${visibleText}`;
  return buildScanReport(text, "", "the message subject/body");
}

function scanAttachmentForPII(name, text) {
  return buildScanReport(getVisibleTextOnly(text), name, `the attachment "${name}"`);
}

function mergeScanReports(reports) {
  const findings = groupFindings(reports.flatMap((report) => report.findings));

  const recommendations = [...new Set(findings.map((finding) => finding.guidance))];
  const scoring = scoreFindings(findings);
  const totalRiskScore = scoring.finalScore;
  const riskBand = scoring.riskLevel;
  const hasAttachments = reports.length > 1;
  const summaryLabel = hasAttachments
    ? "the message and/or supported attachments"
    : "the message subject/body";

  return {
    summary: findings.length
      ? `Potential DLP issue${findings.length === 1 ? "" : "s"} found: ${findings.length} sensitive categor${findings.length === 1 ? "y" : "ies"} detected in ${summaryLabel}. Review the suggestions and anonymise where appropriate before sending, or use message encryption if available.`
      : `No obvious PII/PHI patterns detected in ${summaryLabel}.`,
    findings,
    recommendations,
    totalRiskScore,
    riskBand,
    scoring,
    risk: findings.length > 0 ? "review" : "clear",
    scannedTextLength: reports.reduce((total, report) => total + report.scannedTextLength, 0),
    normalisedText: reports.map((report) => report.normalisedText).join("\n"),
  };
}

module.exports = {
  scanEmailForPII,
  scanAttachmentForPII,
  mergeScanReports,
  getCategoryRiskScores,
};
