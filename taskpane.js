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
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("EWS error:", result.error.message);
      document.getElementById("storage-text").textContent = "Unavailable";
      return;
    }

    try {
      const parser = new DOMParser();
      const xml = parser.parseFromString(result.value, "text/xml");
      const sizeNode = xml.querySelector("[PropertyTag='0xe08'], [PropertyTag='0xE08']");

      if (!sizeNode) throw new Error("Size node not found");

      const usedBytes = parseInt(sizeNode.textContent, 10);

      // Exchange doesn't return total quota via EWS easily, so use 50GB as default
      const totalBytes = 50 * 1e9;
      updateBar(usedBytes, totalBytes);
    } catch (err) {
      console.error("Parse error:", err);
      document.getElementById("storage-text").textContent = "Unavailable";
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