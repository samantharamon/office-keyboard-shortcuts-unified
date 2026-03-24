/*global document, Office*/

let _count = 0;

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  // Test shared runtime.
  updateCount(); // Update count on first open.
  Office.addin.onVisibilityModeChanged((args) => {
    if (args.visibilityMode === Office.VisibilityMode.taskpane) {
      updateCount(); // Update count on subsequent opens.
    }
  });
});

function updateCount() {
  _count++;
  document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
}