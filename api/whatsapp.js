const { createHmac, timingSafeEqual } = require('node:crypto');

const delivery = new Map();

function equal(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const left = Buffer.from(a);
  const right = Buffer.from(b);
  return left.length === right.length && timingSafeEqual(left, right);
}

function replyFor(message) {
  const text = message?.text?.body?.toLowerCase() ?? '';
  if (/admin|operator|petugas/.test(text)) return 'Ini nomor uji CS GPP. Untuk bicara dengan admin, hubungi kanal CS GPP yang biasa Anda gunakan.';
  if (/rusak|pecah|komplain|keluhan|barang salah/.test(text)) return 'Mohon maaf atas kendalanya. Saat menghubungi admin GPP, siapkan nomor pesanan dan penjelasan singkat agar tim bisa membantu.';
  if (/status|resi|lacak|belum sampai/.test(text)) return 'Untuk cek status pesanan, siapkan nomor pesanan dan hubungi admin GPP. Nomor uji ini belum punya akses ke data pesanan.';
  if (/pesan|order|beli|stok|harga|katalog/.test(text)) return 'Untuk pemesanan, siapkan nama produk dan jumlah yang dibutuhkan. Harga dan stok akan dikonfirmasi oleh admin GPP.';
  return 'Halo, ini nomor uji CS GPP. Coba tulis cara pesan, status pesanan, keluhan barang rusak, atau admin.';
}

async function send(message, phoneId, token, version) {
  const existing = delivery.get(message.id);
  if (existing) return existing;
  const pending = fetch('https://graph.facebook.com/' + version + '/' + phoneId + '/messages', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      messaging_product: 'whatsapp',
      to: message.from,
      context: { message_id: message.id },
      type: 'text',
      text: { preview_url: false, body: replyFor(message) }
    }),
    signal: AbortSignal.timeout(7000)
  }).then(response => {
    if (!response.ok) throw new Error('WhatsApp send failed');
  });
  if (delivery.size >= 512) delivery.delete(delivery.keys().next().value);
  delivery.set(message.id, pending);
  try {
    await pending;
  } catch (error) {
    delivery.delete(message.id);
    throw error;
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  const verifyToken = process.env.WA_VERIFY_TOKEN;
  const secret = process.env.WA_APP_SECRET;
  const phoneId = process.env.WA_PHONE_NUMBER_ID;
  const token = process.env.WA_ACCESS_TOKEN;
  const version = process.env.WA_GRAPH_VERSION || 'v24.0';
  const testers = new Set((process.env.WA_TEST_SENDERS || '').split(',').map(number => number.trim()).filter(number => /^\d{7,15}$/.test(number)));

  if (req.method === 'GET') {
    if (!verifyToken) return res.status(503).send('WhatsApp setup pending');
    const query = new URL(req.url, 'https://webhook.invalid').searchParams;
    const challenge = query.get('hub.challenge');
    if (query.get('hub.mode') !== 'subscribe' || !challenge || !equal(query.get('hub.verify_token'), verifyToken)) {
      return res.status(403).send('Verification failed');
    }
    return res.status(200).send(challenge);
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'GET, POST');
    return res.status(405).send('Method not allowed');
  }
  if (!secret || !token || !/^\d+$/.test(phoneId || '') || !/^v\d+\.\d+$/.test(version) || !testers.size) {
    return res.status(503).send('WhatsApp setup pending');
  }
  const signature = req.headers['x-hub-signature-256'];
  if (typeof signature !== 'string' || !/^sha256=[a-f0-9]{64}$/.test(signature)) {
    return res.status(401).send('Invalid signature');
  }
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 262144) return res.status(413).send('Payload too large');
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks);
  const expected = 'sha256=' + createHmac('sha256', secret).update(raw).digest('hex');
  if (!equal(signature, expected)) return res.status(401).send('Invalid signature');

  let payload;
  try {
    payload = JSON.parse(raw.toString('utf8'));
  } catch {
    return res.status(400).send('Invalid JSON');
  }
  if (payload?.object !== 'whatsapp_business_account') return res.status(200).send('OK');
  try {
    const entries = Array.isArray(payload.entry) ? payload.entry : [];
    const pending = [];
    for (const entry of entries) {
      if (!Array.isArray(entry?.changes)) continue;
      for (const change of entry.changes) {
        const value = change?.value;
        if (change?.field !== 'messages' || value?.metadata?.phone_number_id !== phoneId || !Array.isArray(value.messages)) continue;
        for (const message of value.messages) {
          if (testers.has(message?.from) && typeof message.id === 'string' && message.type === 'text') {
            pending.push(send(message, phoneId, token, version));
          }
        }
      }
    }
    await Promise.all(pending);
    return res.status(200).send('OK');
  } catch {
    return res.status(502).send('WhatsApp reply failed');
  }
};
