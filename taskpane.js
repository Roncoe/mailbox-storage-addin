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
    const res = await fetch("https://graph.microsoft.com/v1.0/me/mailboxSettings", {
      headers: { Authorization: "Bearer " + token }
    });

    if (!res.ok) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Graph error:", res.status, res.statusText);
      return;
    }

    const data = await res.json();

    // MailboxSettings doesn't include quota — fall back to drive quota
    const quotaRes = await fetch("https://graph.microsoft.com/v1.0/me/drive/quota", {
      headers: { Authorization: "Bearer " + token }
    });

    if (!quotaRes.ok) {
      text.textContent = "Unable to retrieve quota data.";
      console.error("Quota error:", quotaRes.status);
      return;
    }

    const quotaData = await quotaRes.json();
    const usedBytes = parseInt(quotaData.used, 10);
    const totalBytes = parseInt(quotaData.total, 10);

    // Bounds check
    if (isNaN(usedBytes) || isNaN(totalBytes) || usedBytes < 0 || totalBytes <= 0 || usedBytes > MAX_BYTES || totalBytes > MAX_BYTES) {
      text.textContent = "Unable to retrieve mailbox data.";
      console.error("Quota bounds check failed:", quotaData);
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
