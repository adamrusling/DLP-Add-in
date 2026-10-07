/* global Office, atob, DOMParser, TextDecoder, TextEncoder, require, module */

const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");
const mammoth = require("mammoth/mammoth.browser");
const XLSX = require("xlsx");
const JSZip = require("jszip");

const TEXT_EXTENSIONS = new Set(["txt", "csv", "tsv", "log", "json", "xml", "html", "htm"]);
const SUPPORTED_EXTENSIONS = new Set([
  ...TEXT_EXTENSIONS,
  "pdf",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "ppt",
  "pptx",
]);

// Use the filename suffix to choose a parser for each attachment.
function getExtension(name) {
  const parts = (name || "").toLowerCase().split(".");
  return parts.length > 1 ? parts.pop() : "";
}

function decodeBase64(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

// Decode plain-text attachments as UTF-8, replacing invalid byte sequences.
function decodeText(bytes) {
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}

function stripXml(xml) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  return document.documentElement
    ? document.documentElement.textContent
    : xml.replace(/<[^>]+>/g, " ");
}

// Read selectable text from each PDF page; image-only pages need OCR and return no text here.
async function extractPdfText(bytes) {
  const pdf = await pdfjsLib.getDocument({ data: bytes, disableWorker: true }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => item.str).join(" "));
  }

  return pages.join("\n");
}

async function extractPowerPointText(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  const slides = [];

  for (const slideName of slideNames) {
    slides.push(stripXml(await zip.files[slideName].async("text")));
  }

  return slides.join("\n");
}

// Legacy .doc and .ppt files have no browser parser here; keep readable ASCII and UTF-16LE runs.
function extractLegacyOfficeStrings(bytes) {
  const chunks = [];
  let asciiChunk = "";

  const flushAscii = () => {
    if (asciiChunk.trim().length >= 3) {
      chunks.push(asciiChunk.trim());
    }
    asciiChunk = "";
  };

  for (let index = 0; index < bytes.length; index += 1) {
    const value = bytes[index];
    if (value >= 32 && value <= 126) {
      asciiChunk += String.fromCharCode(value);
    } else {
      flushAscii();
    }
  }
  flushAscii();

  for (let index = 0; index + 1 < bytes.length; index += 2) {
    if (bytes[index] >= 32 && bytes[index] <= 126 && bytes[index + 1] === 0) {
      let value = "";
      while (
        index + 1 < bytes.length &&
        bytes[index] >= 32 &&
        bytes[index] <= 126 &&
        bytes[index + 1] === 0
      ) {
        value += String.fromCharCode(bytes[index]);
        index += 2;
      }
      if (value.trim().length >= 3) {
        chunks.push(value.trim());
      }
      index -= 2;
    }
  }

  return chunks.join("\n");
}

function extractSpreadsheetText(bytes) {
  const workbook = XLSX.read(bytes, { type: "array", cellText: true, cellDates: true });
  return workbook.SheetNames.map((sheetName) =>
    XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName])
  ).join("\n");
}

async function extractAttachmentText(name, bytes) {
  const extension = getExtension(name);

  if (TEXT_EXTENSIONS.has(extension)) {
    return decodeText(bytes);
  }

  if (extension === "pdf") {
    return extractPdfText(bytes);
  }

  if (extension === "docx") {
    const result = await mammoth.extractRawText({ arrayBuffer: bytes.buffer });
    return result.value;
  }

  if (extension === "xls" || extension === "xlsx") {
    return extractSpreadsheetText(bytes);
  }

  if (extension === "pptx") {
    return extractPowerPointText(bytes);
  }

  if (extension === "doc" || extension === "ppt") {
    return extractLegacyOfficeStrings(bytes);
  }

  return "";
}

function getAttachmentContent(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }

      const content = result.value;
      resolve(
        content.format === Office.MailboxEnums.AttachmentContentFormat.Base64
          ? decodeBase64(content.content)
          : new TextEncoder().encode(content.content)
      );
    });
  });
}

// Compose items use getAttachmentsAsync; read items expose their attachment array directly.
function getItemAttachments(item) {
  if (typeof item.getAttachmentsAsync !== "function") {
    return Promise.resolve(item.attachments || []);
  }

  return new Promise((resolve, reject) => {
    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(result.error);
      }
    });
  });
}

async function readSupportedAttachments(item) {
  let attachments;
  try {
    attachments = await getItemAttachments(item);
  } catch {
    return [
      {
        name: "Attachments",
        text: "",
        error: "Attachments could not be listed, so they were not scanned.",
      },
    ];
  }

  const results = [];

  for (const attachment of attachments) {
    const extension = getExtension(attachment.name);
    // Ignore cloud links, Outlook items, and file types without an extraction path.
    if (
      attachment.attachmentType !== Office.MailboxEnums.AttachmentType.File ||
      !SUPPORTED_EXTENSIONS.has(extension)
    ) {
      continue;
    }

    try {
      const bytes = await getAttachmentContent(item, attachment.id);
      const text = await extractAttachmentText(attachment.name, bytes);
      if (text.trim()) {
        results.push({ name: attachment.name, text });
      }
    } catch {
      results.push({ name: attachment.name, text: "", error: "The attachment could not be read." });
    }
  }

  return results;
}

module.exports = {
  readSupportedAttachments,
};
