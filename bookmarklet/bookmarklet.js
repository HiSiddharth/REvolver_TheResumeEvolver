(function() {
  const APPS_SCRIPT_URL = "YOUR_APPSCRIPT_DEPLOYMENT_URL";
  // true: clipboard first, highlighted text fallback; false: highlighted text only
  const USE_CLIPBOARD = false;
  const selectedText = window.getSelection().toString().trim();
  const targetName = "rr-tailoring-" + Date.now();
  const newWindow = window.open("about:blank", targetName);
  if (!newWindow) {
    alert("Resume Revolver: popup blocked. Please allow popups for this site and try again.");
    return;
  }
  const openApp = function(jobDescription) {
    const payload = { jobDescription: jobDescription || "" };
    const base64Data = btoa(unescape(encodeURIComponent(JSON.stringify(payload))));
    newWindow.location.replace(APPS_SCRIPT_URL + "#d=" + encodeURIComponent(base64Data));
  };
  if (USE_CLIPBOARD && navigator.clipboard?.readText) {
    navigator.clipboard.readText()
      .then(function(text) { openApp(text.trim() || selectedText); })
      .catch(function() { openApp(selectedText); });
  } else {
    openApp(selectedText);
  }
})();
