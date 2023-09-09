<!DOCTYPE html>
<html>
<head>
    <title>Upload and Display File</title>
    <meta charset="UTF-8">
    <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
    <style>
        table, th, td {
            border: 1px solid black;
            border-collapse: collapse;
        }
        th, td {
            padding: 5px;
            text-align: center;
        }
        .selected {
            background-color: lightblue;
        }
    </style>
</head>
<body>
    <h1>Upload and Display File</h1>
    <input type="file" id="file-upload" onchange="handleFileUpload(this)" />
    <button onclick="saveFile()">Save File</button>
    <button onclick="deleteSelectedRows()">Delete Selected Rows</button>
    <button onclick="deleteSelectedColumns()">Delete Selected Columns</button>
    <button onclick="addRows()">Add Rows</button>
    <button onclick="addColumns()">Add Columns</button>
    <button onclick="updateSelectedCells()">Update Selected Cells</button>
    <button onclick="removeDuplicateRows()">Remove Duplicate Rows</button>
    <button onclick="removeEmptyRows()">Remove Empty Rows</button>
    <button onclick="calculateColumnSum()">Calculate Column Sum</button>
    <button onclick="calculateColumnAverage()">Calculate Column Average</button>
    <button onclick="calculateColumnMax()">Calculate Column Max</button>
    <button onclick="calculateColumnMin()">Calculate Column Min</button>
    <button onclick="datatypeconversion()">Data type conversion</button>
    <div id="excel-data"></div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <label for="file-format">File Format:</label>
    <select id="file-format">
        <option value="xlsx">Excel (XLSX)</option>
        <option value="csv">CSV</option>
        <option value="txt">Text (TXT)</option>
    </select>
    <script>
        let excelData = [];

        function handleFileUpload(input) {
            const file = input.files[0];
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = e.target.result;
                const workbook = XLSX.read(data, {type: 'binary'});
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                excelData = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
                displayExcelData(excelData);
            };
            reader.readAsBinaryString(file);
        }

        function saveFile() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const fileFormat = document.getElementById("file-format").value;
    let fileExtension = "";

    switch (fileFormat) {
        case "xlsx":
            fileExtension = ".xlsx";
            break;
        case "csv":
            fileExtension = ".csv";
            break;
        case "txt":
            fileExtension = ".txt";
            break;
    }

    const fileName = "data" + fileExtension;
    const fileData = XLSX.write(wb, { bookType: fileFormat, type: "array" });
    const blob = new Blob([fileData], { type: "application/octet-stream" });

    saveAs(blob, fileName);
}

        function displayExcelData(data) {
            let html = "<table>";
            for (let i = 0; i < data.length; i++) {
                html += "<tr>";
                for (let j = 0; j < data[i].length; j++) {
                    if (i == 0) {
                        html += "<th contenteditable='true' onblur='updateCell(" + i + ", " + j + ", this)' onclick='selectColumn(this)'>" + data[i][j] + "</th>";
                    } else {
                        html += "<td contenteditable='true' onblur='updateCell(" + i + ", " + j + ", this)' onclick='selectRow(this)'>" + data[i][j] + "</td>";
                    }
                }
                if (i == 0) {
                    html += "<th>Add</th><th>Delete</th>";
                } else {
                    html += "<td><button onclick='addRow(" + i + ")'>Add</button></td><td><button onclick='deleteRow(" + i + ")'>Delete</button></td>";
                }
                html += "</tr>";
            }
            html += "<tr>";
            for (let i = 0; i < data[0].length; i++) {
                html += "<td><button onclick='addColumn()'>Add</button></td>";
            }
            html += "<td colspan='2'><button onclick='deleteColumn()'>Delete Last Column</button></td>";
            html += "</tr>";
            html += "</table>";
            const div = document.createElement("div");
            div.innerHTML = html;
            document.querySelector("#excel-data").appendChild(div);
        }

        function addRow(index) {
            const newRow = [];
            for (let i = 0; i < excelData[0].length; i++) {
                newRow.push("");
            }
            excelData.splice(index + 1, 0, newRow);
            displayExcelData(excelData);
        }

        function deleteRow(index) {
            excelData.splice(index, 1);
            displayExcelData(excelData);
        }

        function addColumn() {
         for (let i = 0; i < excelData.length; i++) {
        excelData[i].push("");
         }
         displayExcelData(excelData);
}

        function deleteColumn() {
             for (let i = 0; i < excelData.length; i++) {
                 excelData[i].pop();
             }
             displayExcelData(excelData);
         }

         function updateCell(row, col, cell) {
             excelData[row][col] = cell.innerText;
         }

       function selectRow(cell) {
           cell.parentElement.classList.toggle("selected");
       }

       function selectColumn(cell) {
           const index = cell.cellIndex;
           const rows = cell.closest("table").rows;
           for (let i = 0; i < rows.length; i++) {
               rows[i].cells[index].classList.toggle("selected");
           }
       }

       function deleteSelectedRows() {
            const selectedRows = document.querySelectorAll("tr.selected");
            for (let i = selectedRows.length - 1; i >= 0; i--) {
                const index = selectedRows[i].rowIndex - 0;
                excelData.splice(index, 1);
            }
            displayExcelData(excelData);
        }

        function deleteSelectedColumns() {
            const selectedColumns = document.querySelectorAll("th.selected, td.selected");
            const columnIndexes = [];
            for (let i = 0; i < selectedColumns.length; i++) {
                const index = selectedColumns[i].cellIndex;
                if (!columnIndexes.includes(index)) {
                    columnIndexes.push(index);
                }
            }
            columnIndexes.sort((a, b) => b - a);
            for (let i = 0; i < excelData.length; i++) {
                for (let j = 0; j < columnIndexes.length; j++) {
                    excelData[i].splice(columnIndexes[j], 1);
                }
            }
            displayExcelData(excelData);
        }

        function addRows() {
            const numRows = prompt("How many rows do you want to add?");
            if (numRows != null && !isNaN(numRows)) {
                for (let i = 0; i < numRows; i++) {
                    const newRow = [];
                    for (let j = 0; j < excelData[0].length; j++) {
                        newRow.push("");
                    }
                    excelData.push(newRow);
                }
                displayExcelData(excelData);
            }
        }

        function addColumns() {
            const numColumns = prompt("How many columns do you want to add?");
            if (numColumns != null && !isNaN(numColumns)) {
                for (let i = 0; i < excelData.length; i++) {
                    for (let j = 0; j < numColumns; j++) {
                        excelData[i].push("");
                    }
                }
                displayExcelData(excelData);
            }
        }

        function updateSelectedCells() {
            const newValue = prompt("Enter the new value for the selected cells:");
            if (newValue != null) {
                const selectedCells = document.querySelectorAll("td.selected, th.selected");
                for (let i = 0; i < selectedCells.length; i++) {
                    const row = selectedCells[i].parentElement.rowIndex - 1;
                    const col = selectedCells[i].cellIndex;
                    excelData[row][col] = newValue;
                    selectedCells[i].innerText = newValue;
                }
            }
        }

        function removeDuplicateRows() {
            const uniqueRows = [];
            for (let i = 0; i < excelData.length; i++) {
                let isDuplicate = false;
                for (let j = 0; j < uniqueRows.length; j++) {
                    if (arraysEqual(excelData[i], uniqueRows[j])) {
                        isDuplicate = true;
                        break;
                    }
                }
                if (!isDuplicate) {
                    uniqueRows.push(excelData[i]);
                }
            }
            excelData = uniqueRows;
            displayExcelData(excelData);
        }

        function arraysEqual(a, b) {
            if (a === b) return true;
            if (a == null || b == null) return false;
            if (a.length !== b.length) return false;

            for (let i = 0; i < a.length; ++i) {
                if (a[i] !== b[i]) return false;
            }
            return true;
        }

        function removeEmptyRows() {
             const nonEmptyRows = [];
             for (let i = 0; i < excelData.length; i++) {
                 let isEmpty = false;
                 for (let j = 0; j < excelData[i].length; j++) {
                     if (excelData[i][j] === "") {
                         isEmpty = true;
                         break;
                     }
                 }
                 if (!isEmpty) {
                     nonEmptyRows.push(excelData[i]);
                 }
             }
             excelData = nonEmptyRows;
             displayExcelData(excelData);
         }
function calculateColumnSum() {
    const selectedCells = document.querySelectorAll('.selected');
    const columnIndex = Array.from(selectedCells)[0].cellIndex; 
    const columnValues = Array.from(excelData).map(row => parseFloat(row[columnIndex]));

    if (columnValues.every(value => isNaN(value))) {
        alert("Không có giá trị nào trong cột được chọn.");
        return;
    }

    const sum = columnValues.reduce((acc, value) => acc + (isNaN(value) ? 0 : value), 0);
    alert(`Sum: ${sum}`);
}

function calculateColumnAverage() {
    const selectedCells = document.querySelectorAll('.selected');
    const columnIndex = Array.from(selectedCells)[0].cellIndex; 
    const columnValues = Array.from(excelData).map(row => parseFloat(row[columnIndex]));

    if (columnValues.every(value => isNaN(value))) {
        alert("Không có giá trị nào trong cột được chọn.");
        return;
    }

    const sum = columnValues.reduce((acc, value) => acc + (isNaN(value) ? 0 : value), 0);
    const average = sum / columnValues.length;
    alert(`Average: ${average}`);
}

function calculateColumnMax() {
    const selectedCells = document.querySelectorAll('.selected');
    const columnIndex = Array.from(selectedCells)[0].cellIndex; 
    const columnValues = Array.from(excelData).map(row => parseFloat(row[columnIndex]));

    const validValues = columnValues.filter(value => !isNaN(value));

    if (validValues.length === 0) {
        alert("Không có giá trị hợp lệ trong cột được chọn.");
        return;
    }

    const max = Math.max(...validValues);
    alert(`Max: ${max}`);
}

function calculateColumnMin() {
    const selectedCells = document.querySelectorAll('.selected');
    const columnIndex = Array.from(selectedCells)[0].cellIndex; 
    const columnValues = Array.from(excelData).map(row => parseFloat(row[columnIndex]));

    const validValues = columnValues.filter(value => !isNaN(value));

    if (validValues.length === 0) {
        alert("Không có giá trị hợp lệ trong cột được chọn.");
        return;
    }

    const min = Math.min(...validValues);
    alert(`Min: ${min}`);
}

displayExcelData(excelData);


</script>
</body>
</html>
