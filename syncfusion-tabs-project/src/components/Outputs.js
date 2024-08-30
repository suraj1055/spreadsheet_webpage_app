import React, { useState } from 'react';

const Outputs = () => {
  const [outputData, setOutputData] = useState({
    meltTemperature: '',
    moldTemp: '',
    fillTime: '',
    peakInjectionPressure: '',
    pressAtTransfer: '',
    cushionValue: '',
    screwRotationTime: '',
    cycleTime: '',
    injOnlyShotWeight: '',
    partAndRunnerWeight: '',
  });

  const handleChange = (e) => {
    const { name, value } = e.target;
    setOutputData({
      ...outputData,
      [name]: value,
    });
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    console.log('Output Data:', outputData);
  };

  const formGroupStyle = {
    display: 'flex',
    alignItems: 'center',
    marginBottom: '15px',
  };

  const labelStyle = {
    width: '300px', // Adjust this to control the label width
    textAlign: 'left',
    fontWeight: 'bold',
    color: '#333',
    marginRight: '20px',
  };

  const textInputStyle = {
    width: '200px', // Set a fixed width for all text boxes
    padding: '8px',
    backgroundColor: 'rgb(248, 197, 214)', // Yellowish color for text boxes
    border: '1px solid #ddd',
    borderRadius: '4px',
  };

  const containerStyle = {
    maxWidth: '600px',
    margin: '20px 0',
    backgroundColor: '#e6f7ff', // Light blue background for the container
    padding: '20px',
    borderRadius: '10px',
    boxShadow: '0 0 10px rgba(0,0,0,0.1)',
  };

  const pressure = "";  // Example values, you can pass these as props or retrieve them from state
  const temperature = "";
  const distance = "";
  const time = "";
  const velocity = "";
  const weight = "";

  const handlePrint = () => {
    window.print(); // Example implementation, replace with your actual print logic
  };

  const handleSave = () => {
    alert('Saved!'); // Example implementation, replace with your actual save logic
  };

  const handleClose = () => {
    alert('Closed!'); // Example implementation, replace with your actual close logic
  };

  const fixedBottomPaneStyle = {
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
  };

  const buttonGroupStyle = {
    display: 'flex',
    gap: '10px',
  };

  const buttonStyle = {
    padding: '5px 10px', // Padding inside buttons
    border: 'none',
    cursor: 'pointer',
  };

  return (
    <div>
      <div style={containerStyle}>
        <form onSubmit={handleSubmit}>
          <div style={formGroupStyle}>
            <label style={labelStyle}>1. Melt Temperature :</label>
            <input
              type="text"
              name="meltTemperature"
              value={outputData.meltTemperature}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>2. Mold Temp :</label>
            <input
              type="text"
              name="moldTemp"
              value={outputData.moldTemp}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>3. Fill Time* :</label>
            <input
              type="text"
              name="fillTime"
              value={outputData.fillTime}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>4. Actual Peak Injection Pressure* :</label>
            <input
              type="text"
              name="peakInjectionPressure"
              value={outputData.peakInjectionPressure}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>5. Press At Transfer :</label>
            <input
              type="text"
              name="pressAtTransfer"
              value={outputData.pressAtTransfer}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>6. Cushion Value :</label>
            <input
              type="text"
              name="cushionValue"
              value={outputData.cushionValue}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>7. Screw Rotation Time :</label>
            <input
              type="text"
              name="screwRotationTime"
              value={outputData.screwRotationTime}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>8. Cycle Time :</label>
            <input
              type="text"
              name="cycleTime"
              value={outputData.cycleTime}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>9. Inj Only Shot Weight :</label>
            <input
              type="text"
              name="injOnlyShotWeight"
              value={outputData.injOnlyShotWeight}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
          <div style={formGroupStyle}>
            <label style={labelStyle}>10. Part and Runner Weight :</label>
            <input
              type="text"
              name="partAndRunnerWeight"
              value={outputData.partAndRunnerWeight}
              onChange={handleChange}
              style={textInputStyle}
            />
          </div>
        </form>
      </div>

      {/* Fixed Bottom Pane */}
      <div style={fixedBottomPaneStyle}>
        <span>Pressure: {pressure} </span>
        <span>Temp: {temperature} </span>
        <span>Distance: {distance} </span>
        <span>Time: {time} </span>
        <span>Velocity: {velocity} </span>
        <span>Weight: {weight} </span>
        <div style={buttonGroupStyle}>
          <button style={buttonStyle} onClick={handlePrint}>Print</button>
          <button style={buttonStyle} onClick={handleSave}>Save</button>
          <button style={buttonStyle} onClick={handleClose}>Close</button>
        </div>
      </div>
    </div>
  );
};

export default Outputs;
