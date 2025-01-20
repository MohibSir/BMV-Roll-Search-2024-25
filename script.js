document.addEventListener('DOMContentLoaded', () => {
    const filePath = 'Data.xlsx'; // Make sure this path is correct

    fetch(filePath)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            window.workbook = workbook; // Store globally for search
        })
        .catch(error => console.error('Error reading the Excel file:', error));
});

function searchData() {
    const searchInput = document.getElementById('searchInput').value.trim();
    if (!window.workbook) {
        console.error('Workbook not loaded');
        return;
    }

    const sheet = window.workbook.Sheets[window.workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = '';

    const student = json.find(student => student["Student Code"] === searchInput);

    if (student) {
        resultDiv.innerHTML = `
            <p><strong>Student Code:</strong> ${student["Student Code"]}</p>
            <p><strong>Student Name:</strong> ${student["Student Name"]}</p>
            <p><strong>Application No:</strong> ${student["Application No"]}</p> 
            <p><strong>Semester:</strong> ${student["Year"]}</p>
            <p><strong>Roll No:</strong> ${student["RollNo"]}</p>
            <p><strong>Reg Form No:</strong> ${student["Reg Form No"]}</p>
            <p><strong>Reg No:</strong> ${student["Reg No"]}</p>
            <button onclick="window.print()">Print</button>
        `;
    } else {
        resultDiv.innerHTML = '<p>Student not found.</p>';
    }
}
