Office.onReady(() => {
  fetchStorageData();
});

function fetchStorageData() {
  const text = document.getElementById("storage-text");
  const label = document.getElementById("label");

  try {
    const mailbox = Office.context.mailbox;
    const userProfile = mailbox.userProfile;

    label.textContent = userProfile.displayName;

    // Try to get quota via mailbox.item or roamingSettings
    // Fall back to showing account info with a manual quota display
    const email = userProfile.emailAddress;

    // Use getCallbackTokenAsync for REST but with isRest: false (EAS token)
    mailbox.getCallbackTokenAsync({ isRest: false }, (tokenResult) => {
      if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
        // Last resort: just show connected status
        text.textContent = `${email} — quota unavailable`;
        return;
      }

      const token = tokenResult.value;

      // Use Outlook REST endpoint for mailbox usage
      const ewsUrl = mailbox.ewsUrl;
      text.textContent = `EWS URL: ${ewsUrl ? ewsUrl.substring(0, 50) : "none"}`;
    });

  } catch (err) {
    text.textContent = "Error: " + err.message;
  }
}
