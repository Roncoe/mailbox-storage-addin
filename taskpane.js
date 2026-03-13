const MAX_BYTES = 200 * 1e9; // 200GB sanity ceiling
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

Office.onReady(() => {
  fetchStorageData();
  setInterval(fetchStorageData, REFRESH_INTERVAL_MS);
});

async function fetchStorageData() {
  const text = document.getElementById("storage-text");

  try {
    const token = await getToken();

    // Fetch all mail folders, following pagination
    const usedBytes = await getAllFolderBytes(token);

    // Get real quota via EWS
    const totalBytes = await getQuotaBytes();

    // Bounds check
    if (isNaN(usedBytes) || usedBytes < 0 || usedBytes > MAX_BYTES) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Bounds check failed on usedBytes:", usedBytes);
      return;
    }

    updateBar(usedBytes, totalBytes);

  } catch (err) {
    text.textContent = "Unable to retrieve mailbox data.";
    console.error("fetchStorageData error:", err);
  }
}

async function getAllFolderBytes(token) {
  let total = 0;
  let url = "https://graph.microsoft.com/v1.0/me/mailFolders?$select=sizeInBytes&$top=50";

  let pageCount = 0;
  const MAX_PAGES = 20;

  while (url && pageCount < MAX_PAGES) {
    pageCount++;
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + token }
    });

    if (!res.ok) throw new Error("mailFolders error: " + res.status);

    const data = await res.json();
    for (const folder of data.value) {
      total += parseInt(folder.sizeInBytes, 10) || 0;
    }

    url = data["@odata.nextLink"] || null;
  }

  return total;
}

function getQuotaBytes() {
  return new Promise((resolve) => {
    const soap = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:ExtendedFieldURI PropertyTag="0x3FFB" PropertyType="Long"/>
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("EWS quota failed, using 50GB default");
        resolve(50 * 1e9);
        return;
      }

      try {
        const parser = new DOMParser();
        const xml = parser.parseFromString(result.value, "text/xml");

        // Check for parse errors
        const parseError = xml.querySelector("parsererror");
        if (parseError) {
          console.error("EWS XML parse error");
          resolve(50 * 1e9);
          return;
        }

        const quotaNode = xml.querySelector("[PropertyTag='0x3ffb'], [PropertyTag='0x3FFB']");
        if (!quotaNode) {
          console.error("Quota node not found in EWS response");
          resolve(50 * 1e9);
          return;
        }

        // EWS returns quota in KB
        const quotaKB = parseInt(quotaNode.textContent, 10);
        if (isNaN(quotaKB) || quotaKB <= 0) {
          resolve(50 * 1e9);
          return;
        }

        resolve(quotaKB * 1024);
      } catch (err) {
        console.error("EWS quota parse error:", err);
        resolve(50 * 1e9);
      }
    });
  });
}

function getToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Token unavailable"));
      }
    });
  });
}

function updateBar(usedBytes, totalBytes) {
  const fill = document.getElementById("bar-fill");
  const text = document.getElementById("storage-text");

  const usedGB = (usedBytes / 1e9).toFixed(2);
  const totalGB = (totalBytes / 1e9).toFixed(0);
  const pct = Math.min((usedBytes / totalBytes) * 100, 100).toFixed(1);

  fill.style.width = `${pct}%`;
  text.textContent = `${pct}% (${usedGB}GB / ~${totalGB}GB)`;

  fill.classList.remove("warn", "danger");
  if (pct >= 90)      fill.classList.add("danger");
  else if (pct >= 75) fill.classList.add("warn");
}
