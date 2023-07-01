import React from "react";

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
        <input
          type="text"
          value={this.props.filePath}
          ref={this.inputRef}
          onChange={this.handleFileSelect}
        />
        <button onClick={this.handleBrowseClick}>Browse</button>
        <input
          type="file"
          ref={this.inputRef}
          //style={{ display: "none" }}
          onChange={this.handleFileSelect}
        />
      </div>
    );
  }
}

export default FileInput;
