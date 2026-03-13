Office.onReady(() => {
  fetchStorageData();
});

function fetchStorageData() {
  const text = document.getElementById("storage-text");
  const fill = document.getElementById("bar-fill");

  try {
    const mailbox = Office.context.mailbox;
    const userProfile = mailbox.userProfile;

    // Get mailbox details we can access without EWS
    const displayName = userProfile.displayName;
    const email = userProfile.emailAddress;

    // Try to get diagnostics info
    const diagnostics = mailbox.diagnostics;
    const hostName = diagnostics.hostName;
    const hostVersion = diagnostics.hostVersion;

    text.textContent = `Connected: ${email}`;
    document.getElementById("label").textContent = `${displayName}`;

    // Now try REST-based quota
    fetchQuotaViaRest(mailbox, text, fill);

  } catch (err) {
    text.textContent = "Error: " + err.message;
  }
}

function fetchQuotaViaRest(mailbox, text, fill) {
  try {
    mailbox.getCallbackTokenAsync({ isRest: true }, (tokenResult) => {
      if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
        text.textContent = "Token error: " + tokenResult.error.message;
        return;
      }

      const token = tokenResult.value;
      const restUrl = mailbox.restUrl + "/v2.0/me/MailboxSettings";

      fetch(restUrl, {
        headers: { Authorization: "Bearer " + token }
      })
      .then(res => res.json())
      .then(data => {
        text.textContent = "Data: " + JSON.stringify(data).substring(0, 100);
      })
      .catch(err => {
        text.textContent = "Fetch error: " + err.message;
      });
    });
  } catch (err) {
    text.textContent = "REST error: " + err.message;
  }
}
