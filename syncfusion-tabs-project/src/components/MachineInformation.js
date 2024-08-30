import React, { useState } from 'react';

const MachineInformation = () => {
  const [pressure, setPressure] = useState('');
  const [temperature, setTemperature] = useState('');
  const [distance, setDistance] = useState('');
  const [time, setTime] = useState('');
  const [velocity, setVelocity] = useState('');
  const [weight, setWeight] = useState('');

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
    <div style={styles.container}>
      {/* Machine Information Section */}
      <div style={styles.quadrant}>
        <div style={styles.formGroup}>
          <label style={styles.label}>Processor :</label>
          <input type="text" style={styles.textBox} />
          <label style={styles.labelDate}>Date:</label>
          <input type="date" style={styles.datePicker} />
        </div>
        <div style={styles.formGroup}>
          <label style={styles.label}>Process Sheet ID: *</label>
          <input type="text" style={styles.textBoxWide} />
        </div>
        <div style={styles.formGroup}>
          <label style={styles.label}>Notes :</label>
          <textarea style={styles.rtfBoxWide} />
        </div>
      </div>

      {/* Nozzle Settings Section */}
      <div style={styles.nozzleSettingsContainer}>
        <div style={styles.nozzleSettings}>
          {/* Existing Nozzle Settings */}
          <div style={styles.existingNozzleSettings}>
            <div style={styles.formGroup}>
              <label style={styles.label}>Pressure* :</label>
              <select 
                style={styles.selectBox}
                value={pressure}
                onChange={(e) => setPressure(e.target.value)}
              >
                <option value="psi">psi</option>
                <option value="MPa">MPa</option>
                <option value="bar">bar</option>
              </select>
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Temperature* :</label>
              <select 
                style={styles.selectBox}
                value={temperature}
                onChange={(e) => setTemperature(e.target.value)}
              >
                <option value="degF">deg F</option>
                <option value="degC">deg C</option>
              </select>
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Distances* :</label>
              <select 
                style={styles.selectBox}
                value={distance}
                onChange={(e) => setDistance(e.target.value)}
              >
                <option value="inches">in</option>
                <option value="mm">mm</option>
              </select>
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Time* :</label>
              <select 
                style={styles.selectBox}
                value={time}
                onChange={(e) => setTime(e.target.value)}
              >
                <option value="sec">sec</option>
                <option value="min">min</option>
                <option value="hrs">hrs</option>
              </select>
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Velocity* :</label>
              <select 
                style={styles.selectBox}
                value={velocity}
                onChange={(e) => setVelocity(e.target.value)}
              >
                <option value="inches/sec">inches/sec</option>
                <option value="mm/sec">mm/sec</option>
              </select>
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Weight* :</label>
              <select 
                style={styles.selectBox}
                value={weight}
                onChange={(e) => setWeight(e.target.value)}
              >
                <option value="Gms">Gms</option>
                <option value="Kg">kg</option>
                <option value="oz">oz</option>
                <option value="lbs">lbs</option>
              </select>
            </div>
          </div>

          {/* New Box on the Right */}
          <div style={styles.nozzleSettingsBox}>
            <div style={styles.nozzleSettingsHeader}>Nozzle Settings</div>
            <div style={styles.nozzleSettingsContent}>
              <div style={styles.nozzleSettingsItem}>
                <label style={styles.nozzleSettingsLabel}>Nozzle Type :</label>
                <input type="text" style={styles.smallTextBox} />
              </div>
              <div style={styles.nozzleSettingsItem}>
                <label style={styles.nozzleSettingsLabel}>Nozzle Length :</label>
                <input type="text" style={styles.smallTextBox} />
              </div>
              <div style={styles.nozzleSettingsItem}>
                <label style={styles.nozzleSettingsLabel}>Nozzle Orifice Size :</label>
                <input type="text" style={styles.smallTextBox} />
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Mold Information Section */}
      <div style={styles.quadrant}>
        <h1>3</h1>
      </div>

      {/* Machine Number Section */}
      <div style={styles.quadrant}>
        <h1>4</h1>
      </div>

      {/* Fixed Bottom Pane */}
      <div style={styles.fixedBottomPane}>
        <span>Pressure: {pressure} </span>
        <span>Temp: {temperature} </span>
        <span>Distance: {distance} </span>
        <span>Time: {time} </span>
        <span>Velocity: {velocity} </span>
        <span>Weight: {weight} </span>
        <div style={styles.buttonGroup}>
          <button style={styles.button} onClick={handlePrint}>Print</button>
          <button style={styles.button} onClick={handleSave}>Save</button>
          <button style={styles.button} onClick={handleClose}>Close</button>
        </div>
      </div>
    </div>
  );
};

const styles = {
  container: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr', // Two columns
    gridTemplateRows: '1fr 1fr', // Two rows
    height: '100vh', // Full height of the viewport
    width: '100vw', // Full width of the viewport
    gap: '10px', // Gap between the quadrants
    boxSizing: 'border-box', // Include padding and border in element's total width and height
    backgroundColor: '#bfd5e9', // Very light blue background
    padding: '20px', // Padding around the entire grid
  },
  quadrant: {
    border: '1.5px solid black', // Black border around each quadrant
    padding: '10px', // Padding inside each quadrant
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'flex-start',
    boxSizing: 'border-box', // Include padding and border in element's total width and height
  },
  nozzleSettingsContainer: {
    display: 'flex',
    justifyContent: 'flex-end', // Aligns the nozzle settings to the right
  },
  nozzleSettings: {
    border: '1.5px solid black', // Border around the Nozzle Settings section
    padding: '12px', // Padding inside the border
    display: 'flex',
    gap: '30px', // Space between existing and new nozzle settings
    backgroundColor: '#bfd5e9', // Background color for contrast
    boxSizing: 'border-box', // Include padding and border in element's total width and height
    width: '100%', // Take full width of container
  },
  existingNozzleSettings: {
    display: 'flex',
    flexDirection: 'column',
    gap: '10px',
    flex: 1, // Take remaining space within nozzleSettings
    alignItems: 'flex-end', // Align content to the right
  },
  nozzleSettingsBox: {
    border: '1.5px solid black', // Border around the new Nozzle Settings box
    padding: '10px', // Padding inside the border
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    width: '300px', // Adjust width as needed
    height: '150px',
    backgroundColor: '#bfd5e9', // Ensure background color contrasts with border
    boxSizing: 'border-box', // Include padding and border in element's total width and height
  },
  nozzleSettingsHeader: {
    fontWeight: 'bold',
    marginBottom: '8px', // Adjust margin
    fontSize: '14px', // Adjust font size as needed
  },
  nozzleSettingsBox: {
    border: '1.5px solid black', // Border around the new Nozzle Settings box
    padding: '10px', // Padding inside the border
    display: 'flex',
    flexDirection: 'column', // Align content vertically inside the box
    alignItems: 'flex-start', // Align content to the left
    width: '300px', // Adjust width as needed
    height: 'auto', // Adjust height to fit content
    backgroundColor: '#bfd5e9', // Ensure background color contrasts with border
    boxSizing: 'border-box', // Include padding and border in element's total width and height
  },
  nozzleSettingsContent: {
    display: 'flex',
    flexDirection: 'column',
    gap: '10px', // Space between each row (label + text box)
  },
  nozzleSettingsItem: {
    display: 'flex',
    alignItems: 'center', // Align label and text box in the same row
  },
  nozzleSettingsLabel: {
    width: '140px', // Set a fixed width for labels to align text boxes
    marginRight: '10px', // Space between the label and the text box
  },
  smallTextBox: {
    flex: 1, // Text box will take the remaining space
    width: '120px',  // Adjust width as needed
  },
  formGroup: {
    display: 'flex',
    justifyContent: 'space-between',
    marginBottom: '10px',
  },
  label: {
    marginRight: '10px',
  },
  textBox: {
    width: '100px',
    height: '20px',
  },
  textBoxWide: {
    width: '480px',
    height: '20px',
  },
  rtfBoxWide: {
    width: '480px',
    height: '80px',
  },
  labelDate: {
    marginLeft: '60px',
    marginRight: '10px',
  },
  selectBox: {
    width: '130px',
    height: '25px',
  },
  datePicker: {
    width: '140px',
    height: '20px',
  },
  fixedBottomPane: {
    position: 'fixed',
    bottom: 0,
    left: 0,
    right: 0,
    backgroundColor: '#bfd5e9', // Ensure background color contrasts with text and buttons
    padding: '8px',
    display: 'flex',
    justifyContent: 'space-between',
    borderTop: '1px solid #000', // Add a top border to separate it visually
    gap: '1px', // Reduce the space between spans
  },
  buttonGroup: {
    display: 'flex',
    gap: '10px',
  },
  button: {
    padding: '5px 10px', // Padding inside buttons
    border: 'none',
    cursor: 'pointer',
  },
};

export default MachineInformation;






