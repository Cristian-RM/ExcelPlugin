(function () {
  "use strict";

  // eslint-disable-next-line no-undef
  Office.onReady().then(function () {
    // TODO1: Assign handler to the OK button.
    // eslint-disable-next-line no-undef
    document.getElementById("ok-button").onclick = sendStringToParentPage;
  });

  // TODO2: Create the OK button handler
  function sendStringToParentPage() {
    // eslint-disable-next-line no-undef
    const userName = document.getElementById("name-box").value;
    // eslint-disable-next-line no-undef
    Office.context.ui.messageParent(userName);
  }
})();
