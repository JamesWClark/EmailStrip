function extractEmails(contents) {
    const regex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const matches = contents.match(regex);
    if (matches) {
        return [...new Set(matches)];
    }
    return null;
}

document.getElementById('file-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const contents = e.target.result;
        const emails = extractEmails(contents);
        document.getElementById('emails').innerText = emails.join('\n');
    };
    reader.readAsText(file);
});

document.getElementById('upload-area').addEventListener('dragover', function(e) {
    e.preventDefault();
});

document.getElementById('upload-area').addEventListener('drop', function(e) {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const contents = e.target.result;
        const emails = extractEmails(contents);
        document.getElementById('emails').innerText = emails.join('\n');
    };
    reader.readAsText(file);
});

document.getElementById('clipboard-icon').addEventListener('click', function() {
    const emails = document.getElementById('emails').innerText;
    navigator.clipboard.writeText(emails);
});