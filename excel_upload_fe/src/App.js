import logo from "./logo.svg";
import "./App.css";

function App() {
  return (
    <div className="App">
      {/* {
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <p>
            Edit <code>src/App.js</code> and save to reload.
          </p>
          <a
            className="App-link"
            href="https://reactjs.org"
            target="_blank"
            rel="noopener noreferrer"
          >
            Learn React
          </a>
          <title>Excel Upload</title>
        </header>
      } */}
      <label>Path</label>
      <input type="text" />
      <button>Browse</button>
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
