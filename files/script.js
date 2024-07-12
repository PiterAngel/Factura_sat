document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);
document.getElementById('downloadExcel').addEventListener('click', downloadExcel, false);

function handleFileSelect(event) {
    const files = event.target.files;
    const tableBody = document.querySelector('#resultTable tbody');

    // Clear previous results
    tableBody.innerHTML = '';

    for (let file of files) {
        const reader = new FileReader();

        reader.onload = function(e) {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(e.target.result, "text/xml");

            const bienOServicio = xmlDoc.querySelector('[BienOServicio]')?.getAttribute('BienOServicio');
            const numero = xmlDoc.querySelector('[Numero]')?.getAttribute('Numero');
            const serie = xmlDoc.querySelector('[Serie]')?.getAttribute('Serie');
            const nitEmisor = xmlDoc.querySelector('[NITEmisor]')?.getAttribute('NITEmisor');

            const row = tableBody.insertRow();

            const cellName = row.insertCell(0);
            const cellService = row.insertCell(1);
            const cellNumero = row.insertCell(2);
            const cellSerie = row.insertCell(3);
            const cellNITEmisor = row.insertCell(4);

            cellName.textContent = file.name;
            cellService.textContent = bienOServicio ? bienOServicio : 'No encontrado';
            cellNumero.textContent = numero ? numero : 'No encontrado';
            cellSerie.textContent = serie ? serie : 'No encontrado';
            cellNITEmisor.textContent = nitEmisor ? nitEmisor : 'No encontrado';
        };

        reader.onerror = function() {
            const row = tableBody.insertRow();

            const cellName = row.insertCell(0);
            const cellService = row.insertCell(1);
            const cellNumero = row.insertCell(2);
            const cellSerie = row.insertCell(3);
            const cellNITEmisor = row.insertCell(4);

            cellName.textContent = file.name;
            cellService.textContent = 'Error';
            cellNumero.textContent = 'Error';
            cellSerie.textContent = 'Error';
            cellNITEmisor.textContent = 'Error';
        };

        reader.readAsText(file);
    }
}

function downloadExcel() {
    const table = document.getElementById('resultTable');
    const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'table.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}
