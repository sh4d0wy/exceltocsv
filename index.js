const fileInput = document.getElementById("fileInput");
const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const progressLabel = document.getElementById("progressLabel");
const downloadContainer = document.getElementById("downloadContainer");
const downloadLink = document.getElementById("downloadLink");
const dropContainer = document.getElementById("cloud");
const dropBox = document.getElementById("box");
const fileName = document.getElementById("name");
const below = document.getElementsByClassName("below")[0];

function handleFile(file) {
  below.style.display = "block";
  fileName.textContent = file.name;
  progressLabel.textContent = "Converting...";
  downloadContainer.style.display = "none";

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    let processedCount = 0;

    const csvData = XLSX.utils.sheet_to_csv(worksheet, { raw: true });

    const csvLines = csvData.split("\n");
    const totalLines = csvLines.length;

    const updateProgress = () => {
      const progress = (processedCount / totalLines) * 100;
      progressBar.style.width = Math.round(progress) + "%";
      progressLabel.textContent = "Converted " + Math.round(progress) + "%";
    };

    const processLine = (lineIndex) => {
      processedCount++;
      updateProgress();

      if (lineIndex < totalLines - 1) {
        setTimeout(() => processLine(lineIndex + 1), 0);
      }else {
        progressContainer.style.display = "none";
        downloadContainer.style.display = "block"
        const blob = new Blob([csvData], { type: "text/csv" });
        downloadLink.href = URL.createObjectURL(blob);
      }
    };

    updateProgress();
    processLine(0);
  };

  reader.readAsArrayBuffer(file);
}

function handleDrop(e) {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  handleFile(file);
}

function handleBoxClick() {
  fileInput.click();
}

function handleFileChange() {
  const file = fileInput.files[0];
  handleFile(file);
}

dropContainer.addEventListener("dragover", function (e) {
  e.preventDefault();
});

dropContainer.addEventListener("drop", handleDrop);

dropContainer.addEventListener("click", handleBoxClick);

fileInput.addEventListener("change", handleFileChange);

document.getElementById("downloadButton").addEventListener("click", () => {
  downloadLink.click();
});
