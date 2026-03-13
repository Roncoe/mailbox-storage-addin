const REFRESH_INTERVAL_MS = 5 * 60 * 1000; // refresh every 5 minutes

Office.onReady(() => {
  fetchStorageData();
  setInterval(fetchStorageData, REFRESH_INTERVAL_MS);
});

async function fetchStorageData() {
  try {
    const token = await getAccessToken();

    const quotaRes = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/quota",
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!quotaRes.ok) throw new Error(`Graph error: ${quotaRes.status}`);

    const quotaData = await quotaRes.json();
    updateBar(quotaData.used, quotaData.total);
  } catch (err) {
    console.error("Storage fetch failed:", err);
    document.getElementById("storage-text").textContent = "Unavailable";
  }
}

function updateBar(usedBytes, totalBytes) {
  const usedGB  = (usedBytes  / 1e9).toFixed(3);
  const totalGB = (totalBytes / 1e9).toFixed(0);
  const pct     = Math.min((usedBytes / totalBytes) * 100, 100).toFixed(2);

  const fill = document.getElementById("bar-fill");
  const text = document.getElementById("storage-text");

  fill.style.width = `${pct}%`;
  text.textContent = `${pct}% (${usedGB}GB / ${totalGB}GB)`;

  fill.classList.remove("warn", "danger");
  if (pct >= 90)      fill.classList.add("danger");
  else if (pct >= 75) fill.classList.add("warn");
}

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken({ allowSignInPrompt: true }, (result) => {
      if (result.status === "succeeded") resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}