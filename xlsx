<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX to VCF Converter</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
</head>
<body>
    <h1>XLSX to VCF Converter</h1>
    <input type="file" id="xlsxFile" />
    <input type="text" id="contactName" placeholder="Contact Name" />
    <input type="number" id="contactsPerFile" placeholder="Contacts per File" min="1" />
    <input type="text" id="outputName" placeholder="Output Name" />
    <input type="text" id="range" placeholder="Range (e.g., A1:A100)" />
    <button onclick="convertFile()">Convert</button>

    <script>
        function convertFile() {
            const fileInput = document.getElementById('xlsxFile');
            const contactName = document.getElementById('contactName').value;
            const contactsPerFile = parseInt(document.getElementById('contactsPerFile').value, 10);
            const outputName = document.getElementById('outputName').value;
            const range = document.getElementById('range').value.toUpperCase();

            if (!fileInput.files[0] || !contactName || isNaN(contactsPerFile) || !outputName || !range) {
                alert("Please fill in all fields correctly.");
                return;
            }

            const [startCell, endCell] = range.split(':');
            const startCol = startCell.match(/[A-Z]+/)[0];
            const startRow = parseInt(startCell.match(/[0-9]+/)[0], 10);
            const endCol = endCell.match(/[A-Z]+/)[0];
            const endRow = parseInt(endCell.match(/[0-9]+/)[0], 10);

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                const vcfRecords = [];
                const startColIndex = XLSX.utils.decode_col(startCol);

                // Adjust for 0-based index
                const adjustedStartRow = startRow - 1;
                const adjustedEndRow = endRow - 1;

                rows.slice(adjustedStartRow, adjustedEndRow + 1).forEach((row, index) => {
                    if (index % contactsPerFile === 0) {
                        vcfRecords.push([]);
                    }
                    const phone = row[startColIndex] || '';
                    const fileIndex = Math.floor(index / contactsPerFile) + 1;
                    const contactNumber = (index % contactsPerFile) + 1;
                    const vcfRecord = `BEGIN:VCARD\nVERSION:3.0\nFN:${contactName} ${contactNumber}\nTEL:+${phone}\nEND:VCARD`;
                    vcfRecords[vcfRecords.length - 1].push(vcfRecord);
                });

                vcfRecords.forEach((vcfChunk, index) => {
                    const blob = new Blob([vcfChunk.join('\n')], { type: 'text/vcard' });
                    saveAs(blob, `${outputName} - ${index + 1}.vcf`);
                });
            };

            reader.readAsArrayBuffer(fileInput.files[0]);
        }
    </script>
</body>
</html>
