import React from "react";
import InputFileUpload from "./component/InputFileUpload";
import "./App.css";

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>מערכת תרגום PDF</h1>
        <p>העלה קובץ PDF לתרגום</p>
        <InputFileUpload />
      </header>
    </div>
  );
}

export default App;