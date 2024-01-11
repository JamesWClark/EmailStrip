function extractEmailsFromString(contents) {
    const regex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const matches = contents.match(regex);
    if (matches) {
        return [...new Set(matches)];
    }
    return null;
}

function extractEmailsFromSpreadsheet(contents) {
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
    const fileExtension = file.name.split('.').pop().toLowerCase();

    if (fileExtension === 'tar.gz') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const decompressedData = pako.inflate(data);
            untar(decompressedData).then(function(files) {
                files.forEach(function(file) {
                    const text = new TextDecoder("utf-8").decode(file.buffer);
                    handleFile(new File([text], file.name.replace('.tar.gz', '')));
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'tar') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            untar(data).then(function(files) {
                files.forEach(function(file) {
                    const text = new TextDecoder("utf-8").decode(file.buffer);
                    handleFile(new File([text], file.name));
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'gz') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const decompressedData = pako.inflate(data);
            const text = new TextDecoder("utf-8").decode(decompressedData);
            handleFile(new File([text], file.name.replace('.gz', '')));
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'rar') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const archive = unrar.createArchive(data);
            const entries = archive.getEntries();
            entries.forEach(function(entry) {
                const fileData = archive.extract(entry);
                handleFile(new File([fileData], entry.name));
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === '7z') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            SevenZip.open(data).then(function(archive) {
                archive.extractAll().then(function(files) {
                    files.forEach(function(file) {
                        handleFile(new File([file.data], file.name));
                    });
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'zip') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            JSZip.loadAsync(data).then(function(zip) {
                zip.forEach(function(relativePath, zipEntry) {
                    zipEntry.async('blob').then(function(content) {
                        handleFile(new File([content], zipEntry.name));
                    });
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'xlsx' || fileExtension === 'xls' || fileExtension === 'ods') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const emails = [];
            for (const sheetName of workbook.SheetNames) {
                const worksheet = workbook.Sheets[sheetName];
                const contents = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                const sheetEmails = extractEmailsFromSpreadsheet(contents);
                emails.push(...sheetEmails);
            }
            const uniqueEmails = [...new Set(emails)]; // remove duplicates
            document.getElementById('emails').innerText = uniqueEmails.join('\n');
            navigator.clipboard.writeText(uniqueEmails.join('\n')).then(function() {
                console.log('Emails successfully copied to clipboard');
            }, function() {
                console.error('Failed to copy emails to clipboard');
            });
        };
        reader.readAsArrayBuffer(file);
    } else if (fileExtension === 'pdf') {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            console.log(window.pdfjsLib);
            const loadingTask = window.pdfjsLib.getDocument({data: data});
            loadingTask.promise.then(function(pdf) {
                let emails = [];
                let totalPromises = [];
                for (let i = 1; i <= pdf.numPages; i++) {
                    let pagePromise = pdf.getPage(i);
                    totalPromises.push(pagePromise.then(function(page) {
                        return page.getTextContent().then(function(textContent) {
                            let pageEmails = extractEmailsFromString(textContent.items.map(item => item.str).join(' '));
                            emails.push(...pageEmails);
                        });
                    }));
                }
                Promise.all(totalPromises).then(function() {
                    const uniqueEmails = [...new Set(emails)]; // remove duplicates
                    document.getElementById('emails').innerText = uniqueEmails.join('\n');
                    navigator.clipboard.writeText(uniqueEmails.join('\n')).then(function() {
                        console.log('Emails successfully copied to clipboard');
                    }, function() {
                        console.error('Failed to copy emails to clipboard');
                    });
                });
            });
        };
        reader.readAsArrayBuffer(file);
    } else {
        reader.onload = function(e) {
            const contents = e.target.result;
            const emails = extractEmailsFromString(contents);
            document.getElementById('emails').innerText = emails.join('\n');
            navigator.clipboard.writeText(emails.join('\n')).then(function() {
                console.log('Emails successfully copied to clipboard');
            }, function() {
                console.error('Failed to copy emails to clipboard');
            });
        };
        reader.readAsText(file);
    }
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
        const emails = extractEmailsFromString(contents);
        document.getElementById('emails').innerText = emails.join('\n');
    }
    reader.readAsText(file);
});

document.getElementById('clipboard-icon').addEventListener('click', function() {
    const emails = document.getElementById('emails').innerText;
    navigator.clipboard.writeText(emails);
});

window.addEventListener('paste', function(event) {
    const clipboardData = event.clipboardData || window.clipboardData;

    for (let i = 0; i < clipboardData.items.length; i++) {
        const item = clipboardData.items[i];
        if (item.kind === 'file') { // Handle file data
            const file = item.getAsFile();
            handleFile(file);
        } else {     // Handle text data
            const pastedData = clipboardData.getData('Text');
            const emails = extractEmailsFromString(pastedData);
            document.getElementById('emails').innerText = emails.join('\n');
        }
    }
});
