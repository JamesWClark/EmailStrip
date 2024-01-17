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

async function routeFile(file) {
    const fileExtension = file.name.split('.').pop().toLowerCase();
    let emails = [];

    if (['zip', 'rar', '7z', 'tar', 'gz', 'tar.gz'].includes(fileExtension)) {
        emails = await handleArchive(file);
    } else if (['xlsx', 'xls', 'ods', 'docx', 'pdf'].includes(fileExtension)) {
        emails = await handleArrayBuffer(file);
    } else {
        emails = await handleTextFile(file);
    }

    return emails;
}

function writeEmails(emails) {
    emails = [...new Set(emails)]; // remove duplicates
    document.getElementById('emails').innerText = emails.join('\n');
    navigator.clipboard.writeText(emails.join('\n')).then(function() {
        console.log('Emails successfully copied to clipboard');
    }, function() {
        console.error('Failed to copy emails to clipboard');
    });
}

async function handleZipAndReturnEmails(e) {
    let allEmails = [];
    console.log('handling zip file');

    async function handleZip(data) {
        const zip = await JSZip.loadAsync(data);
        const handleZipEntry = async (relativePath, zipEntry) => {
            if (zipEntry.dir) {
                zip.folder(relativePath).forEach(handleZipEntry);
            } else {
                const content = await zipEntry.async('blob');
                const file = new File([content], relativePath);
                if (file.name.endsWith('.zip')) {
                    const nestedZipData = new Uint8Array(await file.arrayBuffer());
                    const nestedEmails = await handleZipAndReturnEmails({target: {result: nestedZipData}});
                    allEmails.push(...nestedEmails);
                } else {
                    const emails = await routeFile(file);
                    allEmails.push(...emails);
                }
            }
        };
        await Promise.all(Object.keys(zip.files).map(fileName => handleZipEntry(fileName, zip.files[fileName])));
    }

    await handleZip(new Uint8Array(e.target.result));
    return allEmails;
}

async function handlePlainText(contents) {
    return extractEmailsFromString(contents);;
}

async function handlePdf(e) {
    const data = new Uint8Array(e.target.result);
    const pdf = await window.pdfjsLib.getDocument({data: data}).promise;
    let emails = [];
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        let pageEmails = extractEmailsFromString(textContent.items.map(item => item.str).join(' '));
        emails.push(...pageEmails);
    }
    return emails;
}

async function handleDocx(e) {
    const arrayBuffer = e.target.result;
    return mammoth.extractRawText({arrayBuffer: arrayBuffer})
        .then(function(result) {
            const text = result.value;
            const emails = extractEmailsFromString(text);
            return emails;
        });
}

async function handleExcel(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type: 'array'});
    const emails = [];
    for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const contents = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        const sheetEmails = extractEmailsFromSpreadsheet(contents);
        emails.push(...sheetEmails);
    }
    return emails;
}

async function handle7z(e) {
    const data = new Uint8Array(e.target.result);
    SevenZip.open(data).then(function(archive) {
        archive.extractAll().then(function(files) {
            files.forEach(function(file) {
                handleFile(new File([file.data], file.name));
            });
        });
    });
}

async function handleRar(e) {
    const data = new Uint8Array(e.target.result);
    const archive = unrar.createArchive(data);
    const entries = archive.getEntries();
    entries.forEach(function(entry) {
        const fileData = archive.extract(entry);
        handleFile(new File([fileData], entry.name));
    });
}

async function handleGz(e) {
    const data = new Uint8Array(e.target.result);
    const decompressedData = pako.inflate(data);
    const text = new TextDecoder("utf-8").decode(decompressedData);
    handleFile(new File([text], file.name.replace('.gz', '')));
}

async function handleTar(e) {
    const data = new Uint8Array(e.target.result);
    untar(data).then(function(files) {
        files.forEach(function(file) {
            const text = new TextDecoder("utf-8").decode(file.buffer);
            handleFile(new File([text], file.name));
        });
    });
}

async function handleTarGz(e) {
    const data = new Uint8Array(e.target.result);
    const decompressedData = pako.inflate(data);
    untar(decompressedData).then(function(files) {
        files.forEach(function(file) {
            const text = new TextDecoder("utf-8").decode(file.buffer);
            handleFile(new File([text], file.name.replace('.tar.gz', '')));
        });
    });
}

function handleArchive(file) {
    return new Promise((resolve, reject) => {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();
        reader.onload = async(e) => {
            try {
                switch (fileExtension) {
                    case 'tar.gz':
                        resolve(await handleTarGz(e));
                        break;
                    case 'tar':
                        resolve(await handleTar(e));
                        break;
                    case 'gz':
                        resolve(await handleGz(e));
                        break;
                    case 'rar':
                        resolve(await handleRar(e));
                        break;
                    case '7z':
                        resolve(await handle7z(e));
                        break;
                    case 'zip':
                        resolve(await handleZipAndReturnEmails(e));
                        break;
                    default:
                        throw new Error('Unsupported file type: ' + fileExtension);
                }
            } catch (error) {
                reject(error);
            }
        }
        reader.onerror = (e) => {
            reject(new Error('FileReader error: ' + e.target.error));
        }
        reader.readAsArrayBuffer(file);
    });
}

function handleArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();
        reader.onload = async(e) => {
            try {
                switch (fileExtension) {
                    case 'xlsx':
                    case 'xls':
                    case 'ods':
                        resolve(await handleExcel(e) || []);
                        break;
                    case 'docx':
                        resolve(await handleDocx(e) || []);
                        break;
                    case 'pdf':
                        resolve(await handlePdf(e) || []);
                        break;
                    default:
                        throw new Error('Unsupported file type: ' + fileExtension);
                }
            } catch (error) {
                reject(error);
            }
        }
        reader.onerror = (e) => {
            reject(new Error('FileReader error: ' + e.target.error));
        }
        reader.readAsArrayBuffer(file);
    });
}

function handleTextFile(file) {
    return new Promise((resolve, reject) => {
        // const fileExtension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();
        reader.onload = async(e) => {
            try {
                resolve(await handlePlainText(e.target.result) || []);
            } catch (error) {
                reject(error);
            }
        }
        reader.onerror = (e) => {
            reject(new Error('FileReader error: ' + e.target.error));
        }
        reader.readAsText(file);
    });
}


document.getElementById('upload-area').addEventListener('dragover', function(e) {
    e.preventDefault();
});

document.getElementById('upload-area').addEventListener('drop', async function(e) {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    const emails = await routeFile(file);
    writeEmails(emails);
});

document.getElementById('file-input').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const emails = await routeFile(file);
    writeEmails(emails);
});

document.getElementById('clipboard-icon').addEventListener('click', function() {
    const emails = document.getElementById('emails').innerText;
    navigator.clipboard.writeText(emails);
});

window.addEventListener('paste', async function(event) {
    const clipboardData = event.clipboardData || window.clipboardData;

    for (let i = 0; i < clipboardData.items.length; i++) {
        const item = clipboardData.items[i];
        if (item.kind === 'file') { // Handle file data
            const file = item.getAsFile();
            const emails = await routeFile(file);
            writeEmails(emails);
        } else {     // Handle text data
            const pastedData = clipboardData.getData('Text');
            const emails = extractEmailsFromString(pastedData);
            writeEmails(emails);
        }
    }
});
