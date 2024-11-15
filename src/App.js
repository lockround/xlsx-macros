import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

const ExcelMacroHandler = () => {
  const [data, setData] = useState([]);
  const [workbook, setWorkbook] = useState(null);

  // Handle file import
  const handleFileImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const binaryStr = event.target.result;
      const wb = XLSX.read(binaryStr, { type: "binary", bookVBA: true });

      // Assume the first sheet for simplicity
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];

      // Convert sheet to JSON
      const jsonData = XLSX.utils.sheet_to_json(sheet);
      setData(jsonData);
      setWorkbook(wb);
    };

    reader.readAsBinaryString(file);
  };

  // Handle data modification
  const handleDataChange = (index, key, value) => {
    const updatedData = [...data];
    updatedData[index][key] = value;
    setData(updatedData);
  };

  // Handle file export
  const handleFileExport = () => {
    if (!workbook) {
      alert("No file loaded!");
      return;
    }

    // Update the workbook with modified data
    const updatedSheet = XLSX.utils.json_to_sheet(data);
    const sheetName = workbook.SheetNames[0];
    workbook.Sheets[sheetName] = updatedSheet;

    // Export the updated workbook
    const wbout = XLSX.write(workbook, { bookType: "xlsm", type: "binary" });
    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    saveAs(blob, "modified-file.xlsm");
  };

  // Helper function to convert binary string to array buffer
  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

  return (
    <div style={{ padding: "20px" }}>
      <h2>Excel Macro File Handler</h2>
      <input type="file" accept=".xlsm" onChange={handleFileImport} />
      {data.length > 0 && (
        <div>
          <h3>Data Preview</h3>
          <table border="1" style={{ width: "100%", marginBottom: "20px" }}>
            <thead>
              <tr>
                {Object.keys(data[0]).map((key) => (
                  <th key={key}>{key}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {Object.keys(row).map((key) => (
                    <td key={key}>
                      <input
                        type="text"
                        value={row[key]}
                        onChange={(e) =>
                          handleDataChange(rowIndex, key, e.target.value)
                        }
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <button onClick={handleFileExport}>Export File</button>
        </div>
      )}
    </div>
  );
};

export default ExcelMacroHandler;
