document.addEventListener("DOMContentLoaded", function () {
    const form = document.getElementById("rasch-form");
    const fileInput = document.getElementById("results_file");
    const fileLabelText = document.getElementById("file-label-text");
    const progress = document.getElementById("progress");

    if (fileInput) {
        fileInput.addEventListener("change", function () {
            if (fileInput.files && fileInput.files.length > 0) {
                fileLabelText.textContent = fileInput.files[0].name;
            } else {
                fileLabelText.textContent = "Excel faylni tanlang…";
            }
        });
    }

    if (form) {
        form.addEventListener("submit", function () {
            // Show simple "processing" indicator
            if (progress) {
                progress.classList.remove("hidden");
            }
        });
    }
});


