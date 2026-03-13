const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

Office.onReady(() => {
  fetchStorageData();
  setInterval(fetchStorageData, REFRESH_INTERVAL_MS);
});

function fetchStorageData() {
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

  Office.context.mailbox.makeEwsRequestAsync(soapRequest, (result) => {
    const text = document.getElementById("storage-text");

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      text.textContent = "EWS Error: " + result.error.message;
      return;
    }

    try {
      text.textContent = "Parsing...";
      const parser = new DOMParser();
      const xml = parser.parseFromString(result.value, "text/xml");
      const sizeNode = xml.querySelector("[PropertyTag='0xe08'], [PropertyTag='0xE08']");

      if (!sizeNode) {
        text.textContent = "Error: size node not found. Raw: " + result.value.substring(0, 200);
        return;
      }

      const usedBytes = parseInt(sizeNode.textContent, 10);
      const totalBytes = 50 * 1e9;
      updateBar(usedBytes, totalBytes);
    } catch (err) {
      text.textContent = "Parse error: " + err.message;
    }
  });
}

function updateBar(usedBytes, totalBytes) {
  const usedGB  = (usedBytes  / 1e9).toFixed(2);
  const totalGB = (totalBytes / 1e9).toFixed(0);
  const pct     = Math.min((usedBytes / totalBytes) * 100, 100).toFixed(1);

  const fill = document.getElementById("bar-fill");
  const text = document.getElementById("storage-text");

  fill.style.width = `${pct}%`;
  text.textContent = `${pct}% (${usedGB}GB / ${totalGB}GB)`;

  fill.classList.remove("warn", "danger");
  if (pct >= 90)      fill.classList.add("danger");
  else if (pct >= 75) fill.classList.add("warn");
}
