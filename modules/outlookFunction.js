/**
   * Prints the received response from the socket to MS Excel
   * Texts are printed from the current cursor position
   * Prints only the first result from the response
   * As the first response is the best prediction
   * @param {string} text
   */
export const printInOutlook = async (text) => {
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