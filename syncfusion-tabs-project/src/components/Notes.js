import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';

const Notes = () => {
  const [notes, setNotes] = useState('');
  const navigate = useNavigate();

  const handleNotesChange = (e) => {
    setNotes(e.target.value);
  };

  const handleGoBack = () => {
    navigate('/machine-information'); // Redirect to Machine Information tab or any other desired tab
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

  return (
    <div style={styles.container}>
      <div style={styles.header}>
        <h2 style={styles.headerText}>Notes</h2>
      </div>
      <textarea
        style={styles.textarea}
        value={notes}
        onChange={handleNotesChange}
        placeholder="Type your notes here..."
      />
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
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '80vh',
    backgroundColor: 'antiquewhite',
    padding: '8px',
  },
  header: {
    width: '100%',
    backgroundColor: 'rgb(98, 98, 211)', // Blue strip
    padding: '10px',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
  },
  headerText: {
    color: '#fff',
    fontWeight: 'bold',
    margin: 0,
  },
  textarea: {
    width: '95%',
    height: '350px',
    padding: '10px',
    fontSize: '16px',
    borderRadius: '5px',
    border: '2px solid black',
    marginTop: '20px',
    backgroundColor: '#bfd5e9',
  },
  fixedBottomPane: {
    position: 'fixed',
    bottom: 0,
    left: 0,
    right: 0,
    backgroundColor: '#bfd5e9',
    padding: '8px',
    display: 'flex',
    justifyContent: 'space-between',
    borderTop: '1px solid #000',
    gap: '1px',
  },
  buttonGroup: {
    display: 'flex',
    gap: '10px',
  },
  button: {
    padding: '5px 10px',
    border: 'none',
    cursor: 'pointer',
  },
};

export default Notes;
