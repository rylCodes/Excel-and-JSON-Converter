const fileInput = document.getElementById("fileInput");
const convertButton = document.getElementById("convertButton");

function convertFile() {
    const selectedFile = fileInput.files[0];
    if (!selectedFile) {
        alert("Please select a file.");
        return;
    }

    const fileExtension = selectedFile.name.split('.').pop().toLowerCase();

    if (fileExtension === 'xlsx') {
        convertExcelToJson(selectedFile);
    } else if (fileExtension === 'json') {
        convertJsonToExcel(selectedFile);
    } else {
        alert("Unsupported file format. Please select a .xlsx or .json file.");
    }
}

function convertExcelToJson(excelFile) {
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        // Assume there's only one sheet in the Excel file
        const sheetName = workbook.SheetNames[0];
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // Convert the JSON data to a string for download
        const jsonString = JSON.stringify(jsonData, null, 2);
        downloadFile(jsonString, "json", "data_from_excel.json");
    };

    reader.readAsArrayBuffer(excelFile);
}

function convertJsonToExcel(jsonFile) {
    const reader = new FileReader();

    reader.onload = function(event) {
        const jsonContent = event.target.result;
        const jsonData = JSON.parse(jsonContent);
        const worksheet = XLSX.utils.json_to_sheet(jsonData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet 1");
        const excelData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelData], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        downloadFile(blob, "xlsx", "data_from_json.xlsx");
    };

    reader.readAsText(jsonFile);
}

function downloadFile(data, fileType, fileName) {
    const blob = new Blob([data], { type: `application/${fileType}` });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
}

convertButton.addEventListener("click", convertFile);