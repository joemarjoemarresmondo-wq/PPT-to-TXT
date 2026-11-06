"use strict";

const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#file-input");
const fileRow = document.querySelector("#file-row");
const fileName = document.querySelector("#file-name");
const fileSize = document.querySelector("#file-size");
const removeFileButton = document.querySelector("#remove-file");
const convertButton = document.querySelector("#convert-button");
const downloadButton = document.querySelector("#download-button");
const copyButton = document.querySelector("#copy-button");
const status = document.querySelector("#status");
const statusMessage = document.querySelector("#status-message");
const spinner = document.querySelector("#spinner");
const errorMessage = document.querySelector("#error-message");
const preview = document.querySelector("#preview");
const textPreview = document.querySelector("#text-preview");
const slideCount = document.querySelector("#slide-count");

const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

let selectedFile = null;
let convertedText = "";
let convertedSlideCount = 0;

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} bytes`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function clearResult() {
  convertedText = "";
  convertedSlideCount = 0;
  preview.hidden = true;
  downloadButton.hidden = true;
  status.hidden = true;
  status.classList.remove("is-success");
  errorMessage.hidden = true;
  errorMessage.textContent = "";
  textPreview.textContent = "";
}

function setFile(file) {
  clearResult();

  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    selectedFile = null;
    fileInput.value = "";
    fileRow.hidden = true;
    convertButton.disabled = true;
    showError("Please choose a .pptx file. Older .ppt files are not supported.");
    return;
  }

  if (file.size === 0) {
    selectedFile = null;
    fileInput.value = "";
    fileRow.hidden = true;
    convertButton.disabled = true;
    showError("This file is empty. Choose a valid PowerPoint presentation.");
    return;
  }

  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatFileSize(file.size);
  fileRow.hidden = false;
  convertButton.disabled = false;
  errorMessage.hidden = true;
}

function showError(message) {
  errorMessage.textContent = message;
  errorMessage.hidden = false;
}

function parseXml(xmlText, description) {
  const document = new DOMParser().parseFromString(xmlText, "application/xml");
  if (document.querySelector("parsererror")) {
    throw new Error(`The presentation contains invalid ${description} data.`);
  }
  return document;
}

function normalizeZipPath(path) {
  const parts = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") {
      parts.pop();
    } else {
      parts.push(part);
    }
  }
  return parts.join("/");
}

function getSlidePaths(presentation, relationships) {
  const relationshipsById = new Map();
  for (const relationship of relationships.getElementsByTagNameNS("*", "Relationship")) {
    const type = relationship.getAttribute("Type");
    const target = relationship.getAttribute("Target");
    const id = relationship.getAttribute("Id");
    if (type.endsWith("/slide") && target && id) {
      relationshipsById.set(id, target);
    }
  }

  const slideIds = presentation.getElementsByTagNameNS(PRESENTATION_NS, "sldId");
  const paths = [];
  for (const slideId of slideIds) {
    const relationshipId =
      slideId.getAttributeNS(RELATIONSHIPS_NS, "id") || slideId.getAttribute("r:id");
    const target = relationshipsById.get(relationshipId);
    if (!target) {
      throw new Error("A slide is missing its presentation relationship.");
    }
    const path = normalizeZipPath(target.startsWith("/") ? target : `ppt/${target}`);
    paths.push(path);
  }

  if (paths.length === 0) {
    throw new Error("No slides were found in this presentation.");
  }
  return paths;
}

function getSlideText(slideDocument) {
  const paragraphs = slideDocument.getElementsByTagNameNS(DRAWING_NS, "p");
  const lines = [];
  for (const paragraph of paragraphs) {
    const textRuns = paragraph.getElementsByTagNameNS(DRAWING_NS, "t");
    const text = Array.from(textRuns, (run) => run.textContent || "").join("");
    if (text.trim()) lines.push(text);
  }
  return lines;
}

async function extractPresentationText(file) {
  if (typeof JSZip === "undefined") {
    throw new Error("The converter library could not be loaded. Check your internet connection and reload the page.");
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(file);
  } catch {
    throw new Error("This file could not be opened as a PowerPoint presentation. It may be damaged.");
  }

  const presentationFile = zip.file("ppt/presentation.xml");
  const relationshipsFile = zip.file("ppt/_rels/presentation.xml.rels");
  if (!presentationFile || !relationshipsFile) {
    throw new Error("This file is missing PowerPoint slide data. Make sure it is a valid .pptx file.");
  }

  const [presentationXml, relationshipsXml] = await Promise.all([
    presentationFile.async("string"),
    relationshipsFile.async("string"),
  ]);
  const presentation = parseXml(presentationXml, "slide");
  const relationships = parseXml(relationshipsXml, "slide relationship");
  const paths = getSlidePaths(presentation, relationships);
  const slides = await Promise.all(
    paths.map(async (path, index) => {
      const slideFile = zip.file(path);
      if (!slideFile) {
        throw new Error(`Slide ${index + 1} could not be found in this presentation.`);
      }
      const slideXml = await slideFile.async("string");
      const slideDocument = parseXml(slideXml, `slide ${index + 1}`);
      return getSlideText(slideDocument);
    }),
  );

  const text = slides
    .map((lines, index) => [`Slide ${index + 1}`, ...lines].join("\n"))
    .join("\n\n");
  return { text, slideCount: slides.length };
}

function resetFile() {
  selectedFile = null;
  fileInput.value = "";
  fileRow.hidden = true;
  convertButton.disabled = true;
  clearResult();
}

dropzone.addEventListener("click", (event) => {
  if (event.target === fileInput || event.target.closest("button")) return;
  fileInput.click();
});

dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", () => {
  setFile(fileInput.files?.[0] || null);
});

for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
  });
}

dropzone.addEventListener("drop", (event) => {
  setFile(event.dataTransfer?.files?.[0] || null);
});

removeFileButton.addEventListener("click", resetFile);

convertButton.addEventListener("click", async () => {
  if (!selectedFile) return;

  convertButton.disabled = true;
  downloadButton.hidden = true;
  preview.hidden = true;
  errorMessage.hidden = true;
  status.hidden = false;
  status.classList.remove("is-success");
  spinner.hidden = false;
  statusMessage.textContent = "Reading slides…";

  try {
    const result = await extractPresentationText(selectedFile);
    convertedText = result.text;
    convertedSlideCount = result.slideCount;
    textPreview.textContent = convertedText;
    slideCount.textContent = `${convertedSlideCount} ${convertedSlideCount === 1 ? "slide" : "slides"}`;
    preview.hidden = false;
    downloadButton.hidden = false;
    status.classList.add("is-success");
    spinner.hidden = true;
    statusMessage.textContent = "Your presentation is ready.";
  } catch (error) {
    showError(error instanceof Error ? error.message : "Something went wrong while reading this presentation.");
    status.hidden = true;
  } finally {
    convertButton.disabled = false;
  }
});

downloadButton.addEventListener("click", () => {
  if (!convertedText) return;

  const blob = new Blob([convertedText], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const baseName = selectedFile.name.replace(/\.pptx$/i, "");
  link.href = url;
  link.download = `${baseName}.txt`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

copyButton.addEventListener("click", async () => {
  if (!convertedText) return;

  try {
    await navigator.clipboard.writeText(convertedText);
    copyButton.lastChild.textContent = " Copied";
    window.setTimeout(() => {
      copyButton.lastChild.textContent = " Copy text";
    }, 1600);
  } catch {
    showError("Could not copy text. Select the preview text and copy it manually.");
  }
});
