Office.onReady(() => {
  fetchStorageData();
});

function fetchStorageData() {
  const text = document.getElementById("storage-text");
  const label = document.getElementById("label");

  try {
    const mailbox = Office.context.mailbox;
    const userProfile = mailbox.userProfile;

    label.textContent = "Mailbox Storage";

    // Try EWS to get actual usage
    const soapRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:ExtendedFieldURI PropertyTag="0x0E08" PropertyType="Long"/>
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>`;

    mailbox.makeEwsRequestAsync(soapRequest, (result) => {
      const totalBytes = 50 * 1e9;

      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        document.getElementById("storage-text").textContent = "EWS fail: " + result.error.message;
        return;
      }

      try {
        const parser = new DOMParser();
        const xml = parser.parseFromString(result.value, "text/xml");
        const sizeNode = xml.querySelector("[PropertyTag='0xe08'], [PropertyTag='0xE08']");

        if (!sizeNode) {
          updateBar(0, totalBytes, true);
          return;
        }

        const usedBytes = parseInt(sizeNode.textContent, 10);
        updateBar(usedBytes, totalBytes, false);
      } catch (err) {
        updateBar(0, totalBytes, true);
      }
    });

  } catch (err) {
    document.getElementById("storage-text").textContent = "Error: " + err.message;
  }
}

function updateBar(usedBytes, totalBytes, unavailable) {
  const fill = document.getElementById("bar-fill");
  const text = document.getElementById("storage-text");

  if (unavailable) {
    text.textContent = "50GB total — usage unavailable";
    fill.style.width = "0%";
    return;
  }

  const usedGB = (usedBytes / 1e9).toFixed(2);
  const totalGB = (totalBytes / 1e9).toFixed(0);
  const pct = Math.min((usedBytes / totalBytes) * 100, 100).toFixed(1);

  fill.style.width = `${pct}%`;
  text.textContent = `${pct}% (${usedGB}GB / ${totalGB}GB)`;

  fill.classList.remove("warn", "danger");
  if (pct >= 90)      fill.classList.add("danger");
  else if (pct >= 75) fill.classList.add("warn");
}
