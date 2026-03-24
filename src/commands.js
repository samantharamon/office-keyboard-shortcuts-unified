/* global document, Office */

Office.onReady(() => {
  // Configure the keyboard shortcut to open the task pane.
  Office.actions.associate("ShowTaskpane", () => {
    return Office.addin
      .showAsTaskpane()
      .then(() => {
        return;
      })
      .catch((error) => {
        return error.code;
      });
  });

  // Configure the keyboard shortcut to close the task pane.
  Office.actions.associate("HideTaskpane", () => {
    return Office.addin
      .hide()
      .then(() => {
        return;
      })
      .catch((error) => {
        return error.code;
      });
  });
});

