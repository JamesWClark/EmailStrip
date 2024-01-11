document.addEventListener('DOMContentLoaded', function() {
    function extractEmailsFromPlainText(contents) {
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

        if (fileExtension === 'xlsx' || fileExtension === 'xls') {
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
                                let pageEmails = extractEmailsFromPlainText(textContent.items.map(item => item.str).join(' '));
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
                const emails = extractEmailsFromPlainText(contents);
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
            const emails = extractEmailsFromPlainText(contents);
            document.getElementById('emails').innerText = emails.join('\n');
        }
        reader.readAsText(file);
    });

    document.getElementById('clipboard-icon').addEventListener('click', function() {
        const emails = document.getElementById('emails').innerText;
        navigator.clipboard.writeText(emails);
    });
});