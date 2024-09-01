import React, { useRef, useEffect } from "react";
import {
  SpreadsheetComponent,
  SheetsDirective,
  SheetDirective,
} from "@syncfusion/ej2-react-spreadsheet";
import { registerLicense } from "@syncfusion/ej2-base";
import "@syncfusion/ej2-base/styles/material.css";
import "@syncfusion/ej2-inputs/styles/material.css";
import "@syncfusion/ej2-buttons/styles/material.css";
import "@syncfusion/ej2-splitbuttons/styles/material.css";
import "@syncfusion/ej2-lists/styles/material.css";
import "@syncfusion/ej2-navigations/styles/material.css";
import "@syncfusion/ej2-popups/styles/material.css";
import "@syncfusion/ej2-dropdowns/styles/material.css";
import "@syncfusion/ej2-grids/styles/material.css";
import "@syncfusion/ej2-react-spreadsheet/styles/material.css";

import CellData from "./CellData";

// Register the Syncfusion license key
registerLicense(
  "Ngo9BigBOggjHTQxAR8/V1NCaF1cVGhIfEx1RHxQdld5ZFRHallYTnNWUj0eQnxTdEFjUX5acXRRQmJYVEJ1Ww=="
);

const Inputs = ({
  pressure,
  temperature,
  distance,
  time,
  velocity,
  weight,
}) => {
  const spreadsheetRef = useRef(null);

  const editableRanges = [
    "A2",
    "A3",
    "A4",
    "C2",
    "C3",
    "C4",
    "D2",
    "D3",
    "D4",
    "F3",
    "G3",
    "H3",
    "I3",
    "J3",
    "K3",
    "L3",
    "M3",
    "N3",
    "O3",
    "P3",
    "A15",
    "B15",
    "C15",
    "D15",
    "E15",
    "F15",
    "A17",
    "B17",
    "C17",
    "D17",
    "E17",
    "F17",
    "A19",
    "B19",
    "C19",
    "D19",
    "E19",
    "F19",
    "S2",
    "S3",
    "C10",
    "F11",
    "H11",
    "J11",
    "L11",
    "N11",
    "P11",
    "R11",
    "H15",
    "J15",
    "L15",
    "N15",
    "P15",
    "R15",
    "H17",
    "J17",
    "L17",
    "N17",
    "P17",
    "R17",
  ];

  const findNextEditableCell = (rowIndex, colIndex) => {
    const flatEditableRanges = editableRanges
      .map((range) => {
        const col = range.charCodeAt(0) - 65;
        const row = parseInt(range.substring(1)) - 1;
        return { row, col, cell: range };
      })
      .sort((a, b) => a.row - b.row || a.col - b.col);

    for (let i = 0; i < flatEditableRanges.length; i++) {
      const { row, col, cell } = flatEditableRanges[i];
      if (row > rowIndex || (row === rowIndex && col > colIndex)) {
        return cell;
      }
    }

    return flatEditableRanges[0].cell; // Loop back to the first editable cell
  };

  const findPreviousEditableCell = (rowIndex, colIndex) => {
    const flatEditableRanges = editableRanges
      .map((range) => {
        const col = range.charCodeAt(0) - 65;
        const row = parseInt(range.substring(1)) - 1;
        return { row, col, cell: range };
      })
      .sort((a, b) => a.row - b.row || a.col - b.col);

    for (let i = flatEditableRanges.length - 1; i >= 0; i--) {
      const { row, col, cell } = flatEditableRanges[i];
      if (row < rowIndex || (row === rowIndex && col < colIndex)) {
        return cell;
      }
    }

    return flatEditableRanges[flatEditableRanges.length - 1].cell; // Loop back to the last editable cell
  };

  const onBeforeSelect = (args) => {
    // Check if indexes are defined and have the required elements
    if (args.indexes && args.indexes.length >= 2) {
      const { indexes } = args;
      const rowIndex = indexes[0];
      const colIndex = indexes[1];
      const cell = String.fromCharCode(65 + colIndex) + (rowIndex + 1);

      if (!editableRanges.includes(cell)) {
        args.cancel = true; // Cancel selection if cell is not in editableRanges
      }
    }
  };

  const onCellSelected = (args) => {
    const spreadsheet = spreadsheetRef.current;

    const { rowIndex, colIndex } = args;

    const cell = String.fromCharCode(65 + colIndex) + (rowIndex + 1);

    if (!editableRanges.includes(cell)) {
      args.cancel = true; // Prevent selection

      let nextEditableCell = findNextEditableCell(rowIndex, colIndex);

      if (nextEditableCell) {
        setTimeout(() => spreadsheet.selectRange(nextEditableCell), 0);
      }
    }
  };

  const onKeyDown = (e) => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet && (e.key === "Tab" || e.key.startsWith("Arrow"))) {
      e.preventDefault();
      const activeCell = spreadsheet.getActiveSheet().activeCell;
      let [col, row] = activeCell.match(/[A-Z]+|[0-9]+/g);
      row = parseInt(row, 10) - 1;
      col = col.charCodeAt(0) - 65;

      let nextEditableCell;
      if (e.key === "Tab" || e.key === "ArrowRight") {
        nextEditableCell = findNextEditableCell(row, col + 0);
      } else if (e.key === "ArrowLeft") {
        nextEditableCell = findPreviousEditableCell(row, col - 0);
      } else if (e.key === "ArrowDown") {
        nextEditableCell = findNextEditableCell(row + 0, col);
      } else if (e.key === "ArrowUp") {
        nextEditableCell = findPreviousEditableCell(row - 0, col);
      }

      if (nextEditableCell) {
        setTimeout(() => spreadsheet.selectRange(nextEditableCell), 0);
      }
    }
  };

  const insertImage = () => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      // Image details
      const image1 = {
        src: "https://i.postimg.cc/y8NYvCTk/block-1.png", // Use the provided image URL
        width: 320, // Adjust width to fit within B10:B14
        height: 95, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image1], "F6"); // Start from the upper-left corner of the range
      const image2 = {
        src: "https://i.postimg.cc/SNXkPs7B/Screenshot-177.png", // Use the provided image URL
        width: 1450, // Adjust width to fit within B10:B14
        height: 100, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image2], "K5"); // Start from the upper-left corner of the range

      const image3 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 280, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image3], "N13"); // Start from the upper-left corner of the range

      const image4 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 240, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image4], "R13"); // Start from the upper-left corner of the range

      const image5 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 240, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image5], "U13"); // Start from the upper-left corner of the range

      const image6 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 240, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image6], "X13"); // Start from the upper-left corner of the range

      const image7 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 240, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image7], "AA13"); // Start from the upper-left corner of the range

      const image8 = {
        src: "https://i.postimg.cc/sDdYNB58/brackets-1.png", // Use the provided image URL
        width: 240, // Adjust width to fit within B10:B14
        height: 20, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image8], "AD13"); // Start from the upper-left corner of the range

      const image9 = {
        src: "https://i.postimg.cc/RhvyBv8x/download.png", // Use the provided image URL
        width: 980, // Adjust width to fit within B10:B14
        height: 10, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image9], "A14"); // Start from the upper-left corner of the range

      const image10 = {
        src: "https://i.postimg.cc/RhvyBv8x/download.png", // Use the provided image URL
        width: 960, // Adjust width to fit within B10:B14
        height: 10, // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image10], "Q14"); // Start from the upper-left corner of the range
    }
  };

  const customizeSpreadsheet = () => {
    const spreadsheet = spreadsheetRef.current;

    if (spreadsheet) {
      let CSSForReadOnly = {
        fontWeight: "bold",
        textAlign: "center",
        verticalAlign: "middle",
        fontSize: "12pt",
        color: "black",
        border: "1.5px solid black",
      };

      let CSSForWhiteBG = {
        backgroundColor: "white", // Match background color
        border: "none", // Remove borders
        color: "transparent", // Hide text color
      };

      try {
        // ******************************************************************************************

        CellData.ReadOnlyDataObj.forEach((cell) => {
          spreadsheet.updateCell(
            {
              value: cell.value,
              style: {
                ...CSSForReadOnly,
                backgroundColor: cell.backgroundColor
                  ? cell.backgroundColor
                  : "rgb(241, 217, 201)",
              },
            },
            `${cell.cell}`
          );
        });

        CellData.CellsWithWhiteBg.forEach((cell) => {
          spreadsheet.updateCell(
            {
              style: CSSForWhiteBG,
            },
            `${cell.cell}`
          );
        });

        // Merge specified cell ranges
        spreadsheet.merge("I8:R8");
        spreadsheet.merge("A21:F21");
        spreadsheet.merge("G21:S21");
        spreadsheet.merge("A22:S100");
        spreadsheet.merge("I7:R7");
        spreadsheet.merge("I6:R6");
        spreadsheet.merge("I5:R5");
        spreadsheet.merge("G4:R4");
        spreadsheet.merge("C1:D1");
        spreadsheet.merge("F1:P1");
        spreadsheet.merge("R1:S1");
        spreadsheet.merge("G10:G11");
        spreadsheet.merge("I10:I11");
        spreadsheet.merge("K10:K11");
        spreadsheet.merge("M10:M11");
        spreadsheet.merge("O10:O11");
        spreadsheet.merge("Q10:Q11");
        spreadsheet.merge("A14:F14");
        spreadsheet.merge("H14:P14");
        spreadsheet.merge("R16:S16");
        spreadsheet.merge("R18:S18");
        spreadsheet.merge("R15:S15");
        spreadsheet.merge("R17:S17");
        spreadsheet.merge("A6:B6");
        spreadsheet.merge("A7:B7");
        spreadsheet.merge("F5:F8");
        spreadsheet.merge("G19:S20");
        spreadsheet.merge("C6:C8");

        spreadsheet.setColumnsWidth(160, ["A", "B", "C", "D", "E", "F"]);
        spreadsheet.setColumnsWidth(100, [
          "G",
          "H",
          "I",
          "J",
          "K",
          "L",
          "M",
          "N",
          "O",
          "P",
          "Q",
          "R",
          "S",
        ]);
        spreadsheet.setColumnsWidth(
          0,
          Array.from({ length: 163840 }, (_, i) => String.fromCharCode(84 + i))
        );
        spreadsheet.setColumnsHeight(150, ["1", "2", "3", "4", "5"]);
      } catch (error) {
        console.error("Error updating spreadsheet:", error);
      }
    }
  };

  useEffect(() => {
    if (spreadsheetRef.current) {
      const spreadsheet = spreadsheetRef.current;

      spreadsheet.protectSheet(0, {
        selectCells: true,
        formatCells: false,
        formatRows: false,
        formatColumns: false,
        insertLink: false,
        insertShape: true,
        insertImage: true,
      });

      spreadsheet.lockCells("A1:AB30", true); // Lock all cells initially
      editableRanges.forEach((range) => {
        spreadsheet.lockCells(range, false); // Unlock the editable ranges
      });

      // Add beforeSelect event handler
      spreadsheet.beforeSelect = onBeforeSelect;

      // Add event listeners
      spreadsheet.cellSelected = onCellSelected;
      spreadsheet.element.addEventListener("keydown", onKeyDown);

      // Customize spreadsheet after it's created
      spreadsheet.created = () => {
        insertImage(); // Insert image into spreadsheet
        customizeSpreadsheet(); // Apply additional customizations
      };
    }

    // Cleanup event listener on component unmount
    return () => {
      if (spreadsheetRef.current) {
        const spreadsheet = spreadsheetRef.current;
        spreadsheet.element.removeEventListener("keydown", onKeyDown);
      }
    };
  }, []);

  return (
    <div style={{ height: "81vh", width: "97vw" }}>
      <SpreadsheetComponent
        ref={spreadsheetRef}
        style={{ height: "90%", width: "90%" }}
        allowEditing={false} // Disable editing by default
        showSheetTabs={false}
        showFormulaBar={false}
        showRibbon={false}
        showHeaders={false}
      >
        <SheetsDirective>
          <SheetDirective showHeaders={false}>
            {/* Define your cells and other sheet content here */}
          </SheetDirective>
        </SheetsDirective>
      </SpreadsheetComponent>

      {/* Fixed Bottom Pane */}
      <div
        style={{
          position: "fixed",
          bottom: 0,
          width: "95vw", // Adjusted width to match container width
          backgroundColor: "#f0f8ff", // Light blue background color
          borderTop: "1px solid #ccc",
          padding: "8px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          backgroundColor: "#bfd5e9",
          boxShadow: "0px -1px 5px rgba(0,0,0,0.2)", // Added shadow for better visibility
        }}
      >
        <div style={{ display: "flex", gap: "15px" }}>
          <span>Pressure: {pressure}</span>
          <span>Temp: {temperature}</span>
          <span>Distance: {distance}</span>
          <span>Time: {time}</span>
          <span>Velocity: {velocity}</span>
          <span>Weight: {weight}</span>
        </div>
        <div style={{ display: "flex", gap: "10px" }}>
          <button
            style={{
              padding: "5px 10px",
              backgroundColor: "#e0e0e0",
              border: "none",
              borderRadius: "4px",
            }}
            onClick={() => window.print()}
          >
            Print
          </button>
          <button
            style={{
              padding: "5px 10px",
              backgroundColor: "#e0e0e0",
              border: "none",
              borderRadius: "4px",
            }}
            onClick={() => alert("Save functionality not implemented.")}
          >
            Save
          </button>
          <button
            style={{
              padding: "5px 10px",
              backgroundColor: "#e0e0e0",
              border: "none",
              borderRadius: "4px",
            }}
            onClick={() => alert("Close functionality not implemented.")}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

export default Inputs;
