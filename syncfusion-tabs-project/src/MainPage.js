import React from 'react';
import { BrowserRouter as Router, Route, Routes, Navigate } from 'react-router-dom';
import { TabComponent, TabItemDirective, TabItemsDirective } from '@syncfusion/ej2-react-navigations';
import MachineInformation from './components/MachineInformation';
import Inputs from './components/Inputs';
import HR from './components/HR';
import Notes from './components/Notes';
import Outputs from './components/Outputs';
import './Tabs.css'; // Import your CSS file

const MainPage = () => {
  return (
    <div>
      <TabComponent>
        <TabItemsDirective>
          <TabItemDirective header={{ text: 'Machine Information' }} content={() => <MachineInformation />} />
          <TabItemDirective header={{ text: 'Inputs' }} content={() => <Inputs />} />
          <TabItemDirective header={{ text: 'Outputs' }} content={() => <Outputs />} />
          <TabItemDirective header={{ text: 'HR' }} content={() => <HR />} />
          <TabItemDirective header={{ text: 'Notes' }} content={() => <Notes />} />
        </TabItemsDirective>
      </TabComponent>
    </div>
  );
};

const App = () => (
  <Router>
    <Routes>
      <Route path="/" element={<Navigate to="/nautilus_js" />} /> {/* Default to the static API route */}
      <Route path="/nautilus_js" element={<MainPage />} /> {/* Render MainPage at static URL */}
    </Routes>
  </Router>
);

export default App;
