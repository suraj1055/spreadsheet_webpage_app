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
      let ReadOnlyDataObj = [
        { value: "Mold Temp Notes", cell: "A1", backgroundColor: "white" },
        { value: "Sections", cell: "B1", backgroundColor: "white" },
        { value: "Mold Temp Inputs", cell: "C1", backgroundColor: "white" },
        { value: "", cell: "A2" },
        { value: "", cell: "A3" },
        { value: "", cell: "A4" },
        { value: "3", cell: "B2", backgroundColor: "white" },
        { value: "2", cell: "B3", backgroundColor: "white" },
        { value: "1", cell: "B4" },
        { value: "", cell: "C2" },
        { value: "", cell: "C3" },
        { value: "", cell: "C4" },
        { value: "", cell: "D2" },
        { value: "", cell: "D3" },
        { value: "", cell: "D4" },
        { value: "Nozzle", cell: "F2" },
        { value: "H1", cell: "G2" },
        { value: "H2", cell: "H2" },
        { value: "H3", cell: "I2" },
        { value: "H4", cell: "J2" },
        { value: "H5", cell: "K2" },
        { value: "H6", cell: "L2" },
        { value: "H7", cell: "M2" },
        { value: "H8", cell: "N2" },
        { value: "H9", cell: "O2" },
        { value: "H10", cell: "P2" },
        { value: "", cell: "F3" },
        { value: "", cell: "G3" },
        { value: "", cell: "H3" },
        { value: "", cell: "I3" },
        { value: "", cell: "J3" },
        { value: "", cell: "K3" },
        { value: "", cell: "L3" },
        { value: "", cell: "M3" },
        { value: "", cell: "N3" },
        { value: "", cell: "O3" },
        { value: "", cell: "P3" },
        { value: "B-Side", cell: "C5" },
        { value: "A-Side", cell: "D5" },
        { value: "", cell: "C10" },
        { value: "", cell: "F11" },
        { value: "", cell: "F3" },
        { value: "", cell: "G3" },
        { value: "", cell: "H3" },
        { value: "", cell: "I3" },
        { value: "", cell: "J3" },
        { value: "", cell: "K3" },
        { value: "", cell: "L3" },
        { value: "", cell: "M3" },
        { value: "", cell: "N3" },
        { value: "", cell: "O3" },
        { value: "", cell: "P3" },
        { value: "B-Side", cell: "C5" },
        { value: "A-Side", cell: "D5" },
        { value: "", cell: "C10" },
        { value: "", cell: "F11" },
        { value: "", cell: "H11" },
        { value: "", cell: "J11" },
        { value: "", cell: "L11" },
        { value: "", cell: "N11" },
        { value: "", cell: "P11" },
        { value: "", cell: "R11" },
        { value: "", cell: "S2" },
        { value: "", cell: "S3" },
        { value: "Cooling Time", cell: "B10" },
        { value: "Sec", cell: "D10" },
        { value: "Cushion*", cell: "F10" },
        { value: "Transfer*", cell: "H10" },
        { value: "Posn 4", cell: "J10" },
        { value: "Posn 3", cell: "L10" },
        { value: "Posn 2", cell: "N10" },
        { value: "Posn 1", cell: "P10" },
        { value: "Shot Size*", cell: "R10" },
        { value: "Decompression", cell: "R1" },
        { value: "Distance:", cell: "R2" },
        { value: "Speed:", cell: "R3" },
        { value: "Pack & Hold Phase", cell: "A14", backgroundColor: "white" },
        { value: "Pack & Hold Phase", cell: "B14", backgroundColor: "white" },
        { value: "Pack & Hold Phase", cell: "C14", backgroundColor: "white" },
        { value: "Pack & Hold Phase", cell: "D14", backgroundColor: "white" },
        { value: "Pack & Hold Phase", cell: "E14", backgroundColor: "white" },
        { value: "Pack & Hold Phase", cell: "F14", backgroundColor: "white" },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "H14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "I14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "J14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "K14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "L14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "M14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "N14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "O14",
          backgroundColor: "white",
        },
        {
          value: "Injection Phase (Max. Inj. Speed)",
          cell: "P14",
          backgroundColor: "white",
        },
        {
          value: "Barrel Settings Temperatures (image not to scale)",
          cell: "F1",
          backgroundColor: "white",
        },
        { value: "", cell: "A15" },
        { value: "", cell: "B15" },
        { value: "", cell: "C15" },
        { value: "", cell: "D15" },
        { value: "", cell: "E15" },
        { value: "", cell: "F15" },
        { value: "", cell: "A17" },
        { value: "", cell: "B17" },
        { value: "", cell: "C17" },
        { value: "", cell: "D17" },
        { value: "", cell: "E17" },
        { value: "", cell: "F17" },
        { value: "", cell: "A19" },
        { value: "", cell: "B19" },
        { value: "", cell: "C19" },
        { value: "", cell: "D19" },
        { value: "", cell: "E19" },
        { value: "", cell: "F19" },
        { value: "", cell: "H15" },
        { value: "", cell: "J15" },
        { value: "", cell: "L15" },
        { value: "", cell: "N15" },
        { value: "", cell: "P15" },
        { value: "", cell: "H17" },
        { value: "", cell: "J17" },
        { value: "", cell: "L17" },
        { value: "", cell: "N17" },
        { value: "", cell: "P17" },
        { value: "", cell: "R17" },
        { value: "", cell: "R15" },
        { value: "", cell: "S15" },
        { value: "", cell: "R15" },
        { value: "P/H Vel 6", cell: "A16", backgroundColor: "white" },
        { value: "P/H Vel 5", cell: "B16", backgroundColor: "white" },
        { value: "P/H Vel 4", cell: "C16", backgroundColor: "white" },
        { value: "P/H Vel 3", cell: "D16", backgroundColor: "white" },
        { value: "P/H Vel 2", cell: "E16", backgroundColor: "white" },
        { value: "P/H Vel 1", cell: "F16", backgroundColor: "white" },
        { value: "P/H Time 6", cell: "A20", backgroundColor: "white" },
        { value: "P/H Time 5", cell: "B20", backgroundColor: "white" },
        { value: "P/H Time 4", cell: "C20", backgroundColor: "white" },
        { value: "P/H Time 3", cell: "D20", backgroundColor: "white" },
        { value: "P/H Time 2", cell: "E20", backgroundColor: "white" },
        { value: "P/H Time 1", cell: "F20", backgroundColor: "white" },
        { value: "Inj Vel 5", cell: "H16", backgroundColor: "white" },
        { value: "Inj Vel 4", cell: "J16", backgroundColor: "white" },
        { value: "Inj Vel 3", cell: "L16", backgroundColor: "white" },
        { value: "Inj Vel 2", cell: "N16", backgroundColor: "white" },
        { value: "Inj Vel 1", cell: "P16", backgroundColor: "white" },
        { value: "Screw Rot Speed*", cell: "R16" },
        { value: "Inj Press 5", cell: "H18", backgroundColor: "white" },
        { value: "Inj Press 4", cell: "J18", backgroundColor: "white" },
        { value: "Inj Press 3", cell: "L18", backgroundColor: "white" },
        { value: "Inj Press 2", cell: "N18", backgroundColor: "white" },
        { value: "Inj Press 1", cell: "P18", backgroundColor: "white" },
        { value: "Back Pressure*", cell: "R18" },
        { value: "P/H Press 6", cell: "A18", backgroundColor: "white" },
        { value: "P/H Press 5", cell: "B18", backgroundColor: "white" },
        { value: "P/H Press 4", cell: "C18", backgroundColor: "white" },
        { value: "P/H Press 3", cell: "D18", backgroundColor: "white" },
        { value: "P/H Press 2", cell: "E18", backgroundColor: "white" },
        { value: "P/H Press 1", cell: "F18", backgroundColor: "white" },
        { value: "", cell: "G10", backgroundColor: "white" },
        { value: "", cell: "I10", backgroundColor: "white" },
        { value: "", cell: "K10", backgroundColor: "white" },
        { value: "", cell: "M10", backgroundColor: "white" },
        { value: "", cell: "O10", backgroundColor: "white" },
        { value: "", cell: "Q10", backgroundColor: "white" },
        { value: "", cell: "I18", backgroundColor: "white" },
        { value: "", cell: "K18", backgroundColor: "white" },
        { value: "", cell: "M18", backgroundColor: "white" },
        { value: "", cell: "O18", backgroundColor: "white" },
        { value: "Process Inputs", cell: "A6", backgroundColor: "white" },
        {
          value: "The Fields marked with * are mandatory",
          cell: "A7",
          backgroundColor: "white",
        },
      ];

      const CellsWithWhiteBg = [
        { cell: "A5" },
        { cell: "B5" },
        { cell: "A8" },
        { cell: "B8" },
        { cell: "A9" },
        { cell: "A11" },
        { cell: "A10" },
        { cell: "A12" },
        { cell: "B11" },
        { cell: "C11" },
        { cell: "D11" },
        { cell: "E11" },
        { cell: "E10" },
        { cell: "E9" },
        { cell: "E4" },
        { cell: "F4" },
        { cell: "F9" },
        { cell: "G9" },
        { cell: "H9" },
        { cell: "I9" },
        { cell: "J9" },
        { cell: "K9" },
        { cell: "L9" },
        { cell: "M9" },
        { cell: "N9" },
        { cell: "O9" },
        { cell: "P9" },
        { cell: "E1" },
        { cell: "E2" },
        { cell: "E3" },
        { cell: "Q1" },
        { cell: "Q2" },
        { cell: "Q3" },
        { cell: "S5" },
        { cell: "S6" },
        { cell: "S7" },
        { cell: "S8" },
        { cell: "S9" },
        { cell: "B12" },
        { cell: "C12" },
        { cell: "D12" },
        { cell: "E12" },
        { cell: "F12" },
        { cell: "A13" },
        { cell: "B13" },
        { cell: "C13" },
        { cell: "D13" },
        { cell: "E13" },
        { cell: "F13" },
        { cell: "G13" },
        { cell: "H13" },
        { cell: "I13" },
        { cell: "J13" },
        { cell: "K13" },
        { cell: "L13" },
        { cell: "M13" },
        { cell: "N13" },
        { cell: "O13" },
        { cell: "P13" },
        { cell: "Q13" },
        { cell: "R13" },
        { cell: "S13" },
        { cell: "G4" },
        { cell: "S4" },
        { cell: "I5" },
        { cell: "Q9" },
        { cell: "I6" },
        { cell: "I7" },
        { cell: "I8" },
        { cell: "R9" },
        { cell: "A22" },
        { cell: "S10" },
        { cell: "S11" },
        { cell: "S12" },
        { cell: "R12" },
        { cell: "S14" },
        { cell: "Q14" },
        { cell: "Q16" },
        { cell: "Q17" },
        { cell: "Q18" },
        { cell: "Q15" },
        { cell: "I15" },
        { cell: "I16" },
        { cell: "I17" },
        { cell: "I18" },
        { cell: "k15" },
        { cell: "k16" },
        { cell: "k17" },
        { cell: "k18" },
        { cell: "M15" },
        { cell: "M16" },
        { cell: "M17" },
        { cell: "M18" },
        { cell: "O15" },
        { cell: "O16" },
        { cell: "O17" },
        { cell: "O18" },
        { cell: "G14" },
        { cell: "G15" },
        { cell: "G16" },
        { cell: "G17" },
        { cell: "G18" },
        { cell: "G21" },
        { cell: "G19" },
        { cell: "A21" },
      ];

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

        ReadOnlyDataObj.forEach((cell) => {
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

        CellsWithWhiteBg.forEach((cell) => {
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
