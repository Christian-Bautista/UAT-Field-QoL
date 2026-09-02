// Runs the shared uat_document.py builder in the browser through Pyodide.

const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v314.0.6/full/";
const TEMPLATE_URL = "UAT%20Field%20Template.docx";
const MODULE_URL = "uat_document.py";
const DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const POSTAL_CODE_LENGTH = 6;

// Drives the shared builder one stage at a time so the page can repaint between stages.
const GLUE = `
import io
import uat_document


class Generation:
    def __init__(self, ap_count, details):
        self.steps = uat_document.build_document_steps("template.docx", ap_count, details)
        self.document = None
        self.message = ""
        self.fraction = 0.0

    def step(self):
        try:
            self.message, self.fraction = next(self.steps)
        except StopIteration as finished:
            self.document = finished.value
            return False
        return True

    def data(self):
        buffer = io.BytesIO()
        self.document.save(buffer)
        return buffer.getvalue()
`;

const fields = {
  apCount: document.getElementById("ap-count"),
  documentName: document.getElementById("document-name"),
  organisation: document.getElementById("organisation"),
  venueName: document.getElementById("venue-name"),
  venueAddress: document.getElementById("venue-address"),
  venuePostalCode: document.getElementById("venue-postal-code"),
  venueCategory: document.getElementById("venue-category"),
  block: document.getElementById("block"),
};
const form = document.getElementById("form");
const fieldset = document.getElementById("fields");
const locationButton = document.getElementById("file-location");
const destinationPath = document.getElementById("destination-path");
const browserNote = document.getElementById("browser-note");
const generateButton = document.getElementById("generate");
const resetButton = document.getElementById("reset");
const progress = document.getElementById("progress");
const progressFill = document.getElementById("progress-fill");
const status = document.getElementById("status");

const canPickFolder = typeof window.showDirectoryPicker === "function";
let pyodide = null;
let folderHandle = null;
let postalCodeFilled = false;
let lastPaint = 0;

function setStatus(message, isError = false) {
  status.textContent = message;
  status.classList.toggle("status-error", isError);
}

function setProgress(fraction) {
  const percent = Math.round(fraction * 100);
  progressFill.style.width = percent + "%";
  progress.setAttribute("aria-valuenow", String(percent));
}

function setFieldError(field, message) {
  const error = document.getElementById(field.id + "-error");
  if (error) {
    error.textContent = message || "";
    error.hidden = !message;
  }
  if (message) {
    field.setAttribute("aria-invalid", "true");
  } else {
    field.removeAttribute("aria-invalid");
  }
}

function clearFieldErrors() {
  setFieldError(fields.apCount, "");
  setFieldError(fields.documentName, "");
}

function setInputsEnabled(enabled) {
  fieldset.disabled = !enabled;
  generateButton.disabled = !enabled;
  resetButton.disabled = !enabled;
  locationButton.disabled = !enabled || !canPickFolder;
}

function showDestination() {
  if (folderHandle) {
    destinationPath.textContent = "Saving to " + folderHandle.name;
  } else {
    destinationPath.textContent = "Saving to your browser's download folder";
  }
}

function showBrowserNote() {
  if (canPickFolder) {
    browserNote.textContent =
      "Documents are saved into the folder you choose with File Location. Picking a folder needs Chrome or Edge — " +
      "in Firefox and Safari saving falls back to the browser's download folder.";
  } else {
    browserNote.classList.add("note-warn");
    browserNote.textContent =
      "Saving will not work properly in this browser. Firefox and Safari cannot save into a folder you choose, so " +
      "File Location is unavailable and every document goes to the browser's download folder, named by the browser " +
      "if one is already there. Use Chrome or Edge for the full behaviour.";
  }
}

// Fill the postal code once the address ends in a space followed by six digits.
function onVenueAddressInput() {
  if (postalCodeFilled) {
    return;
  }
  const tail = fields.venueAddress.value.slice(-(POSTAL_CODE_LENGTH + 1));
  if (tail.length === POSTAL_CODE_LENGTH + 1 && tail[0] === " " && /^[0-9]+$/.test(tail.slice(1))) {
    fields.venuePostalCode.value = tail.slice(1);
    postalCodeFilled = true;
  }
}

function resetFields() {
  postalCodeFilled = false;
  for (const field of Object.values(fields)) {
    field.value = "";
  }
  clearFieldErrors();
}

async function chooseFolder() {
  postalCodeFilled = false;
  try {
    folderHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    showDestination();
  } catch (error) {
    if (error && error.name !== "AbortError") {
      setStatus(String(error.message || error), true);
    }
  }
}

// Hand the page back to the browser so it can repaint. A message task is used rather than an
// animation frame because a tab that is not being rendered never runs animation frames.
function nextTask() {
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      resolve();
    };
    channel.port2.postMessage(0);
  });
}

async function repaint(force = false) {
  const now = performance.now();
  if (!force && now - lastPaint < 60) {
    return;
  }
  lastPaint = now;
  await nextTask();
}

async function loadPython() {
  const { loadPyodide } = await import(PYODIDE_URL + "pyodide.mjs");
  pyodide = await loadPyodide({ indexURL: PYODIDE_URL });

  setStatus("Installing python-docx…");
  await pyodide.loadPackage(["micropip", "lxml"]);
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("python-docx");
  micropip.destroy();

  setStatus("Loading the template…");
  const [moduleResponse, templateResponse] = await Promise.all([fetch(MODULE_URL), fetch(TEMPLATE_URL)]);
  if (!moduleResponse.ok || !templateResponse.ok) {
    throw new Error("Could not download the template.");
  }
  pyodide.FS.writeFile("uat_document.py", await moduleResponse.text());
  pyodide.FS.writeFile("template.docx", new Uint8Array(await templateResponse.arrayBuffer()));
  await pyodide.runPythonAsync(GLUE);
}

async function nameExists(handle, name) {
  try {
    await handle.getFileHandle(name);
    return true;
  } catch (error) {
    return false;
  }
}

// Name a clashing file the way Windows does: "Report - Copy.docx", then "Report - Copy (2).docx".
async function availableName(handle, name) {
  if (!(await nameExists(handle, name))) {
    return name;
  }
  const dot = name.lastIndexOf(".");
  const stem = dot === -1 ? name : name.slice(0, dot);
  const extension = dot === -1 ? "" : name.slice(dot);
  let candidate = stem + " - Copy" + extension;
  let copy = 2;
  while (await nameExists(handle, candidate)) {
    candidate = stem + " - Copy (" + copy + ")" + extension;
    copy += 1;
  }
  return candidate;
}

async function saveDocument(bytes, name) {
  const blob = new Blob([bytes], { type: DOCX_TYPE });
  if (folderHandle) {
    if ((await folderHandle.queryPermission({ mode: "readwrite" })) !== "granted") {
      if ((await folderHandle.requestPermission({ mode: "readwrite" })) !== "granted") {
        throw new Error("Permission to write to " + folderHandle.name + " was declined.");
      }
    }
    const finalName = await availableName(folderHandle, name);
    const fileHandle = await folderHandle.getFileHandle(finalName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return "Saved " + finalName;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  return "Downloaded " + name;
}

async function generate(event) {
  event.preventDefault();
  postalCodeFilled = false;
  clearFieldErrors();

  const apCountText = fields.apCount.value.trim();
  const apCount = Number(apCountText);
  if (!/^[+-]?\d+$/.test(apCountText)) {
    setFieldError(fields.apCount, "Please enter a whole number of APs.");
    fields.apCount.focus();
    return;
  }
  if (apCount < 1) {
    setFieldError(fields.apCount, "The number of APs must be at least 1.");
    fields.apCount.focus();
    return;
  }

  let name = fields.documentName.value.trim();
  if (!name) {
    setFieldError(fields.documentName, "Please enter a name for the document.");
    fields.documentName.focus();
    return;
  }
  if (!name.toLowerCase().endsWith(".docx")) {
    name += ".docx";
  }

  const details = pyodide.toPy({
    organisation: fields.organisation.value.trim(),
    venue_name: fields.venueName.value.trim(),
    venue_address: fields.venueAddress.value.trim(),
    venue_postal_code: fields.venuePostalCode.value.trim(),
    venue_category: fields.venueCategory.value.trim(),
    block: fields.block.value.trim(),
  });

  setInputsEnabled(false);
  setProgress(0);
  setStatus("Starting…");
  await repaint(true);

  let generation = null;
  try {
    generation = pyodide.globals.get("Generation")(apCount, details);
    while (generation.step()) {
      setStatus(generation.message);
      setProgress(generation.fraction);
      await repaint();
    }

    setStatus("Saving document…");
    setProgress(0.95);
    await repaint(true);

    const data = generation.data();
    const bytes = data.toJs();
    data.destroy();

    setStatus(await saveDocument(bytes, name));
    setProgress(1);
  } catch (error) {
    setStatus(String((error && error.message) || error), true);
    setProgress(0);
  } finally {
    if (generation) {
      generation.destroy();
    }
    details.destroy();
    setInputsEnabled(true);
  }
}

form.addEventListener("submit", generate);
resetButton.addEventListener("click", resetFields);
locationButton.addEventListener("click", chooseFolder);
fields.venueAddress.addEventListener("input", onVenueAddressInput);

showBrowserNote();
showDestination();
locationButton.disabled = true;

loadPython()
  .then(() => {
    setInputsEnabled(true);
    setStatus("Ready.");
    fields.apCount.focus();
  })
  .catch((error) => {
    setStatus("Could not start Python: " + String((error && error.message) || error), true);
  });
