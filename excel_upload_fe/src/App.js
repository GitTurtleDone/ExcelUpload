import logo from "./logo.svg";
import "./App.css";
import React, { useRef, useState } from "react";

function App() {
  const inputRef = useRef(null);
  const [folderPath, setFolderPath] = useState(""); // Path to the interested folder
  const handleButtonClick = () => {
    inputRef.current.click();
  };
  const handleFileSelect = (e) => {
    const selectedFile = e.target.files[0];
    console.log("Selected File: ", selectedFile);
    const filePath = selectedFile
      ? selectedFile.webkitRelativePaht || selectedFile.name
      : "";
    setFolderPath(filePath);
  };
  return (
    <div className="App">
      <label>Path: </label>
      <input type="text" value={folderPath} onChange={handleFileSelect} />
      <button onClick={handleButtonClick}>Browse</button>
      <input
        type="file"
        style={{ display: "none" }}
        ref={inputRef}
        onChange={handleFileSelect}
        directory="true"
        webkitdirectory="true"
      />
    </div>
  );
}

export default App;

/*
import React from 'react';

class FileInput extends React.Component {
  constructor(props) {
    super(props);
    this.inputRef = React.createRef();
  }

  handleBrowseClick = () => {
    this.inputRef.current.click();
  };

  handleFileSelect = (event) => {
    const file = event.target.files[0];
    if (file) {
      this.props.onFileSelect(file.path);
    }
  };

  render() {
    return (
      <div>
        <label>{this.props.label}</label>
        <input type="text" value={this.props.filePath} readOnly />
        <button onClick={this.handleBrowseClick}>Browse</button>
        <input type="file" ref={this.inputRef} style={{ display: 'none' }} onChange={this.handleFileSelect} />
      </div>
    );
  }
}

export default FileInput;


import React from "react";
import FileInput from "./FileInput";

class App extends React.Component {
  state = {
    filePath: "",
  };

  handleFileSelect = (path) => {
    this.setState({ filePath: path });
  };

  render() {
    return (
      <div>
        <FileInput
          label="Select File:"
          filePath={this.state.filePath}
          onFileSelect={this.handleFileSelect}
        />
      </div>
    );
  }
}

export default App;
*/
