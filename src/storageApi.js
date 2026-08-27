// Bu fayl Firebase o'rniga sizning MySQL+PHP API'ingiz bilan gaplashadi.
// API_URL'ni serveringizga yuklagan api.php faylining haqiqiy manziliga almashtiring.

const API_URL = "https://api.buxpiima.uz/api.php";

export async function storageGet(key) {
  const res = await fetch(`${API_URL}?key=${encodeURIComponent(key)}`);
  if (!res.ok) throw new Error("Server xatosi: " + res.status);
  const data = await res.json();
  return data.value ? { value: data.value } : null;
}

export async function storageSet(key, value) {
  const res = await fetch(`${API_URL}?key=${encodeURIComponent(key)}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ value }),
  });
  if (!res.ok) throw new Error("Server xatosi: " + res.status);
  return { value };
}
