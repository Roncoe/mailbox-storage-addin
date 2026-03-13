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
    const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=mailboxSettings", {
      headers: { Authorization: "Bearer " + token }
    });

    if (!res.ok) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Graph error:", res.status, res.statusText);
      return;
    }

    const data = await res.json();

    // MailboxSettings doesn't expose quota directly — use Exchange quota via mailbox usage
    const usageRes = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders?$select=totalItemCount,sizeInBytes&$top=50", {
      headers: { Authorization: "Bearer " + token }
    });

    if (!usageRes.ok) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Usage error:", usageRes.status);
      return;
    }

    const usageData = await usageRes.json();
    const usedBytes = usageData.value.reduce((sum, folder) => sum + (parseInt(folder.sizeInBytes, 10) || 0), 0);
    const totalBytes = 50 * 1e9; // 50GB default, MailboxSettings.Read doesn't expose quota limit

    // Bounds check
    if (isNaN(usedBytes) || usedBytes < 0 || usedBytes > MAX_BYTES) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Bounds check failed:", usageData);
      return;
    }

    updateBar(usedBytes, totalBytes);

  } catch (err) {
    text.textContent = "Unable to retrieve mailbox data.";
    console.error("fetchStorageData error:", err);
  }
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
  text.textContent = `${pct}% (${usedGB}GB / ${totalGB}GB)`;

  fill.classList.remove("warn", "danger");
  if (pct >= 90)      fill.classList.add("danger");
  else if (pct >= 75) fill.classList.add("warn");
}
