import React, { useRef, useEffect } from 'react';
import { SpreadsheetComponent, SheetsDirective, SheetDirective, RowsDirective, RowDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet';
import { registerLicense } from '@syncfusion/ej2-base';
import '@syncfusion/ej2-base/styles/material.css';
import '@syncfusion/ej2-inputs/styles/material.css';
import '@syncfusion/ej2-buttons/styles/material.css';
import '@syncfusion/ej2-splitbuttons/styles/material.css';
import '@syncfusion/ej2-lists/styles/material.css';
import '@syncfusion/ej2-navigations/styles/material.css';
import '@syncfusion/ej2-popups/styles/material.css';
import '@syncfusion/ej2-dropdowns/styles/material.css';
import '@syncfusion/ej2-grids/styles/material.css';
import '@syncfusion/ej2-react-spreadsheet/styles/material.css';

registerLicense('Ngo9BigBOggjHTQxAR8/V1NCaF1cVGhIfEx1RHxQdld5ZFRHallYTnNWUj0eQnxTdEFjUX5acXRRQmJYVEJ1Ww==');

const SpreadsheetExample = ({ pressure, temperature, distance, time, velocity, weight }) => {
  const spreadsheetRef = useRef(null);

  useEffect(() => {
    const onDataBound = () => {
      const spreadsheet = spreadsheetRef.current;

      if (spreadsheet) {
        // Set column widths for A to L
        const columnWidths = Array(12).fill(100); // Adjusted to 100 as per your columns
        spreadsheet.sheets[0].columns.forEach((col, index) => {
          if (index < 12) col.width = columnWidths[index]; // Set width for A to L
        });

        // Set row heights for 1 to 25
        const rowHeights = Array(25).fill(30); // Rows 1 to 25
        spreadsheet.sheets[0].rows.forEach((row, index) => {
          if (index < 25) row.height = rowHeights[index]; // Set height for 1 to 25
        });

        // Hide columns after L (12th column) and rows after 25
        for (let col = 12; col < spreadsheet.sheets[0].columns.length; col++) {
          spreadsheet.sheets[0].columns[col].width = 0; // Set width to 0
        }
        for (let row = 25; row < spreadsheet.sheets[0].rows.length; row++) {
          spreadsheet.sheets[0].rows[row].height = 0; // Set height to 0
        }

        // Adding data and styling to the "HR" sheet
        spreadsheet.updateCell(
          { 
            value: 'Hot Runner Controller Settings', 
            style: { 
              fontWeight: 'bold', 
              textAlign: 'center', 
              verticalAlign: 'middle', 
              backgroundColor: 'rgb(163, 152, 236)', 
              color: 'white',
            } 
          }, 
          'HR!A1'
        );
        spreadsheet.merge('HR!A1:L1');
        spreadsheet.updateCell(
          { 
            value: '', 
            style: { 
              fontWeight: 'bold', 
              textAlign: 'center', 
              verticalAlign: 'middle',  
              backgroundColor: 'white', // Match background color
              border: 'none',           // Remove borders
              color: 'transparent' ,
            } 
          }, 
          'HR!A28'
        );
        spreadsheet.merge('HR!A28:M100');
        

        // Populate header labels
        const headers = ['Zone no.', 'Settings'];
        for (let i = 0; i < 12; i++) {
          const colLabel = i % 2 === 0 ? headers[0] : headers[1];
          spreadsheet.updateCell(
            { 
              value: colLabel, 
              style: { 
                fontWeight: 'bold', 
                textAlign: 'center', 
                verticalAlign: 'middle', 
                backgroundColor: 'rgb(181, 185, 240)', 
                border: '1px solid black' 
              } 
            }, 
            `HR!${String.fromCharCode(65 + i)}2`
          );
        }

        // Populate number ranges
        const numberRanges = [
          { startCol: 'A', startNum: 1 },
          { startCol: 'C', startNum: 25 },
          { startCol: 'E', startNum: 49 },
          { startCol: 'G', startNum: 73 },
          { startCol: 'I', startNum: 97 },
          { startCol: 'K', startNum: 121 }
        ];

        const colors = {
          A: 'rgb(248, 248, 200)', // Light Yellow
          C: 'rgb(248, 248, 200)', // Light Yellow
          E: 'rgb(248, 248, 200)', // Light Yellow
          G: 'rgb(248, 248, 200)', // Light Yellow
          I: 'rgb(248, 248, 200)', // Light Yellow
          K: 'rgb(248, 248, 200)'  // Light Yellow
        };

        // Apply pink background color to the specified ranges
        const pinkColor = '#f7cfd5'; // Pink color

        const applyBackgroundColor = (range) => {
          const [startCell, endCell] = range.split(':');
          const startCol = startCell.charCodeAt(0) - 65; // Convert column letter to index
          const endCol = endCell.charCodeAt(0) - 65;
          const startRow = parseInt(startCell.slice(1)) - 1; // Convert to 0-based index
          const endRow = parseInt(endCell.slice(1)) - 1;

          for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
              spreadsheet.updateCell({
                style: { 
                  backgroundColor: pinkColor 
                }
              }, `HR!${String.fromCharCode(65 + col)}${row + 1}`); // Adjust row index for 1-based
            }
          }
        };

        // Apply pink color to the specified ranges
        applyBackgroundColor('B3:B26');
        applyBackgroundColor('D3:D26');
        applyBackgroundColor('F3:F26');
        applyBackgroundColor('H3:H26');
        applyBackgroundColor('J3:J26');
        applyBackgroundColor('L3:L26');

        // Column M is the 13th column (zero-based index 12)
        const startColumnIndex = 12; // Zero-based index for column M

        // Adjust the starting character code based on the column M
        // Column M is ASCII 77 ('M'), so start from ASCII 77
        // spreadsheet.setColumnsWidth(0, Array.from({ length: 16384 - startColumnIndex }, (_, i) => String.fromCharCode(77 + i)));


        numberRanges.forEach(({ startCol, startNum }) => {
          let number = startNum;
          for (let row = 3; row <= 26; row++) {
            // Update cell value and style
            spreadsheet.updateCell({
              value: number++, 
              style: { 
                textAlign: 'center', 
                verticalAlign: 'middle', 
                backgroundColor: colors[startCol] || 'white' // Apply color based on column
              }
            }, `HR!${startCol}${row}`);
          }
        });

        // Apply borders to cells from A3 to L26
        const startRow = 2; // Row index for 3rd row (0-based index)
        const endRow = 25; // Row index for 26th row (0-based index)
        const startCol = 0; // Column index for A (0-based index)
        const endCol = 11; // Column index for L (0-based index)

        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            spreadsheet.updateCell({
              style: { 
                border: '1px solid black' 
              }
            }, `HR!${String.fromCharCode(65 + col)}${row + 1}`); // Adjust row index for 1-based
          }
        }

        // Add borders to the cells
        const columnsToBorder = ['B', 'D', 'F', 'H', 'J', 'L'];
        columnsToBorder.forEach(col => {
          for (let row = 1; row <= 26; row++) { // Assuming you want to apply the border to the first 25 rows
            spreadsheet.updateCell({ style: { borderRight: '2px solid black' } }, `HR!${col}${row}`);
          }
        });

        for (let col = 0; col < 12; col++) {
          spreadsheet.updateCell({ style: { borderBottom: '2px solid black' } }, `HR!${String.fromCharCode(65 + col)}26`);
        }

        // Adding border to the right side of cell A1 in "HR" sheet
        spreadsheet.updateCell({ style: { border: '2px solid black' } }, 'HR!A1');
      }
    };
    

    // Initialize and bind data to the spreadsheet
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      spreadsheet.dataBound = onDataBound;
    }
  }, []);

  const handlePrint = () => {
    window.print();
  };

  const handleSave = () => {
    alert("Data saved successfully!");
  };

  const handleClose = () => {
    window.close();
  };

  return (
    <div style={{ height: '480px', width: '99%', backgroundColor: '#f0f8ff', padding: '5px' }}> {/* Very light blue background for the page */}
      <SpreadsheetComponent ref={spreadsheetRef} allowScrolling={true} showRibbon={false} showFormulaBar={true} showSheetTabs={false} style={{ height: '150px', width: '830px' }}>
        <SheetsDirective>
          <SheetDirective name="HR" showHeaders={false}>
            <RowsDirective>
              <RowDirective index={0} height={50}/>
              <RowDirective index={1} height={30}/>
              {/* Define additional rows as needed */}
            </RowsDirective>
            <ColumnsDirective>
              <ColumnDirective width={70}/>
              <ColumnDirective width={70}/>
              {/* Define additional columns as needed */}
            </ColumnsDirective>
            {/* Define other directives and cells as needed */}
          </SheetDirective>
        </SheetsDirective>
      </SpreadsheetComponent>
      
      {/* Fixed Bottom Pane */}
      <div style={{
        position: 'fixed',
        bottom: 0,
        width: '96.5vw', // Adjusted width to match container width
        backgroundColor: '#f0f8ff', // Light blue background color
        border: '1.5px solid black',
        padding: '8px',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'left',
        backgroundColor: '#bfd5e9',
        boxShadow: '0px -1px 5px rgba(0,0,0,0.2)' // Added shadow for better visibility
      }}>
        <div style={{ display: 'flex', gap: '15px' }}>
          <span>Pressure: {pressure}</span>
          <span>Temp: {temperature}</span>
          <span>Distance: {distance}</span>
          <span>Time: {time}</span>
          <span>Velocity: {velocity}</span>
          <span>Weight: {weight}</span>
        </div>
        <div style={{ display: 'flex', gap: '10px' }}>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={handlePrint}>Print</button>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={handleSave}>Save</button>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={handleClose}>Close</button>
        </div>
      </div>
    </div>
  );
};

export default SpreadsheetExample;
