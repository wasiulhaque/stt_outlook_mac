import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import { audioBufferToWav, chunkifyAudioBuffer, fileToAudioBuffer } from "../../../modules/chunkify";
import io from "socket.io-client";
import Alert from "@mui/material/Alert";
import Button from "@mui/material/Button";
import FileUploadOutlinedIcon from "@mui/icons-material/FileUploadOutlined";

export default class FileUploaderPage extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.inputRef = React.createRef();
    this.socket = null;
    (this.wavBuffer = null),
      (this.audioChunks = null),
      (this.wavFile = null),
      (this.state = {
        listItems: [],
        uploadedFile: null,
        fileUploaded: false,
        isLoading: false,
        isSocketConnected: false,
        uploadDone: true,
        audioBuffer: true,
        audioFile: null,
      });
  }

  /**
   * Initializes the socket connection
   * "connect" event is reponsible for connecting with the socket
   * "result" event is a listen event to keep listening the responses
   */
  initializeSocket = () => {
    this.socket = io(process.env.SOCKET_ADDRESS, { transports: ["websocket"] });
    this.socket.on("connect", () => {
      console.log("Socket connected");
      this.setState({ isSocketConnected: true });
    });
    this.socket.on("result", (data) => {
      console.log("Received result:", data);
      console.log(data.text);
      this.printInOutlook(data.text);
    });
  };

  componentDidMount() {
    this.setState({});
    if (this.state.isSocketConnected == false && this.state.fileUploaded) {
      this.initializeSocket();
    }
  }

  /**
   * Handles the file upload input option
   */
  handleInputClick = () => {
    if (this.inputRef.current) {
      this.inputRef.current.click();
    }
  };

  /**
   * Handles the Upload button
   */
  handleUploadButton = async () => {
    if (this.state.isSocketConnected == false) {
      this.initializeSocket();
    }
    this.audioBuffer = await fileToAudioBuffer(this.audioFile);
    this.wavFile = await audioBufferToWav(this.audioBuffer);
    this.wavBuffer = await fileToAudioBuffer(this.wavFile);
    this.audioChunks = chunkifyAudioBuffer(this.wavBuffer, process.env.FILE_UPLOAD_CHUNK_SIZE_IN_SECOND);

    const sendChunks = this.audioChunks;
    console.log(sendChunks);

    for (let i = 0; i < sendChunks.length; i++) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64String = window.btoa(event.target.result);
        console.log(base64String);
        this.socket.emit("audio_transmit", {
          file: base64String,
        });
      };
      try {
        reader.readAsBinaryString(sendChunks[i]);
      } catch (error) {
        console.error(error);
      }
    }
    console.log("File uploaded successfully");
  };

  /**
   * Prints the received response from the socket to MS Word
   * Texts are printed from the current cursor position
   * Prints only the first result from the response
   * As the first response is the best prediction
   * @param {string} text
   */
  printInOutlook = async (text) => {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var selectedText = result.value.toString() || ""; // Use an empty string if no text is selected
        selectedText = selectedText.toString().replace("[object Object]", ""); // To remove the object captured everytime from the outlook context
        var insertText = text.toString().split("|");
        var newText = insertText[0] + " ";
        var updatedText = selectedText + newText;

        Office.context.mailbox.item.setSelectedDataAsync(
          updatedText,
          { coercionType: Office.CoercionType.Text },
          function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              var cursorPosition = selectedText.length + newText.length;
              Office.context.mailbox.item.body.getRange("end").select();
              Office.context.mailbox.item.body.getRange("end").moveEnd("character", -cursorPosition);
              Office.context.mailbox.item.body.getRange("end").select();
            } else {
              console.error("Error setting body: " + result.error.message);
            }
          }
        );
      } else {
        console.error("Error getting selected data: " + result.error.message);
      }
    });
  };

  /**
   * Handles the file upload event
   * @param {event} event
   */
  handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      this.setState({ uploadedFile: file });
      this.audioFile = file;
      this.setState({ fileUploaded: true });
      this.handleUploadButton();
    }
  };

  render() {
    const { fileUploaded } = this.state;
    return (
      <div className="ms-welcome">
        <input
          ref={this.inputRef}
          id="file-upload"
          type="file"
          onChange={this.handleFileUpload}
          style={{ display: "none" }}
          accept="audio/*"
        />
        <label htmlFor="file-upload">
          <Button
            color="taskpane_header"
            variant="contained"
            size="large"
            style={{ borderRadius: "0px 10px 10px 0px", height: "50px", width: "112px", fontSize: "13px" }}
            onClick={this.handleInputClick}
          >
            <FileUploadOutlinedIcon fontSize="small" style={{ margin: "1px" }} />
            আপলোড
          </Button>
        </label>
      </div>
    );
  }
}
