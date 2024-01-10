function extractEmails(contents) {
    const regex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const matches = contents.match(regex);
    if (matches) {
        return [...new Set(matches)];
    }
    return null;
}

function extractEmailsFromContents(contents) {
    const emails = [];
    const regex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    for (const row of contents) {
        for (const cell of row) {
            const cellValue = String(cell);
            const matches = cellValue.match(regex);
            if (matches) {
                emails.push(...matches);
            }
        }
    }
    return [...new Set(emails)]; // remove duplicates
}

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        let emails;
        if (fileExtension === 'xlsx' || fileExtension === 'xls') {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const contents = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            emails = extractEmailsFromContents(contents);
        } else {
            const contents = e.target.result;
            emails = extractEmails(contents);
        }
        document.getElementById('emails').innerText = emails.join('\n');
        navigator.clipboard.writeText(emails.join('\n')).then(function() {
            console.log('Emails successfully copied to clipboard');
        }, function() {
            console.error('Failed to copy emails to clipboard');
        });
    };
    reader.readAsArrayBuffer(file);
}

document.getElementById('upload-area').addEventListener('dragover', function(e) {
    e.preventDefault();
});

document.getElementById('upload-area').addEventListener('drop', function(e) {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    handleFile(file);
});

document.getElementById('file-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const contents = e.target.result;
        const emails = extractEmails(contents);
        document.getElementById('emails').innerText = emails.join('\n');
    }
    reader.readAsText(file);
});

document.getElementById('clipboard-icon').addEventListener('click', function() {
    const emails = document.getElementById('emails').innerText;
    navigator.clipboard.writeText(emails);
});
