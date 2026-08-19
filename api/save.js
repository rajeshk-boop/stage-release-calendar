export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyUIGcy5KGKd_ivD6dKf-tb2TnQ97q6SRWuEtMDqWX_XAQza-AOeZfwZ9B2wcbff_4/exec';
  try {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body),
    });
    const data = await response.json();
    res.status(200).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
