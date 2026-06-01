const form         = document.querySelector("#uploadForm");
const reportInput  = document.querySelector("#report");
const templateInput= document.querySelector("#template");
const reportName   = document.querySelector("#reportName");
const templateName = document.querySelector("#templateName");
const statusText   = document.querySelector("#statusText");
const progressBar  = document.querySelector("#progress span");
const resultCard   = document.querySelector("#resultCard");
const downloadLink = document.querySelector("#downloadLink");
const submitButton = document.querySelector(".primary-action");

let latestDownloadUrl = "";
let latestOutputName  = "FAIRe_metadata_filled.xlsx";

const fields = {
  projectTitle:   document.querySelector("#projectTitle"),
  parsingMethod:  document.querySelector("#parsingMethod"),
  institution:    document.querySelector("#institution"),
  lakeCount:      document.querySelector("#lakeCount"),
  roundCount:     document.querySelector("#roundCount"),
  assayList:      document.querySelector("#assayList"),
  sampleRows:     document.querySelector("#sampleRows"),
  experimentRows: document.querySelector("#experimentRows"),
};

reportInput.addEventListener("change", () => {
  reportName.textContent = reportInput.files?.[0]?.name ?? "Choose a report file";
});
templateInput.addEventListener("change", () => {
  templateName.textContent = templateInput.files?.[0]?.name ?? "Choose the FAIRe full template";
});

function setStatus(message, progress) {
  statusText.textContent = message;
  progressBar.style.width = `${progress}%`;
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  resultCard.classList.add("is-hidden");
  submitButton.disabled = true;

  const hasApiKey = !!document.querySelector("#apiKey")?.value.trim();
  setStatus(
    hasApiKey
      ? "Uploading files and starting AI-powered extraction…"
      : "Uploading files and reading the report…",
    20
  );

  const data = new FormData(form);

  try {
    const pulse = setInterval(() => {
      const cur = parseFloat(progressBar.style.width) || 30;
      progressBar.style.width = `${Math.min(cur + 5, 85)}%`;
    }, 600);

    const response = await fetch("/api/fill", { method: "POST", body: data });
    clearInterval(pulse);

    const payload = await response.json();
    if (!response.ok || payload.error) {
      throw new Error(payload.error || "The workbook could not be filled.");
    }

    fields.projectTitle.textContent   = payload.projectTitle  || "-";
    fields.institution.textContent    = payload.institution   || "-";
    fields.parsingMethod.textContent  = payload.parsingMethod === "claude-ai"
      ? "Claude AI" : "Rule-based (regex)";
    fields.lakeCount.textContent      = payload.lakes         ?? "-";
    fields.roundCount.textContent     = payload.rounds        ?? "-";
    fields.assayList.textContent      = Array.isArray(payload.assays)
      ? payload.assays.join(", ") : "-";
    fields.sampleRows.textContent     = payload.sampleRows    ?? "-";
    fields.experimentRows.textContent = payload.experimentRows ?? "-";

    latestDownloadUrl = payload.downloadUrl;
    latestOutputName  = payload.outputName || "FAIRe_metadata_filled.xlsx";
    downloadLink.href     = latestDownloadUrl;
    downloadLink.download = latestOutputName;

    setStatus("Done — review the summary then download the workbook.", 100);
    resultCard.classList.remove("is-hidden");
  } catch (err) {
    setStatus(err.message, 100);
  } finally {
    submitButton.disabled = false;
  }
});

downloadLink.addEventListener("click", async (event) => {
  if (!latestDownloadUrl) return;
  event.preventDefault();
  setStatus("Preparing download…", 100);
  try {
    const response = await fetch(latestDownloadUrl);
    if (!response.ok) throw new Error("Could not retrieve the completed workbook.");
    const blob = await response.blob();
    const url  = URL.createObjectURL(blob);
    const a    = Object.assign(document.createElement("a"), { href: url, download: latestOutputName });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setStatus("Download started.", 100);
  } catch (err) {
    setStatus(err.message + " — opening link directly.", 100);
    window.location.href = latestDownloadUrl;
  }
});
