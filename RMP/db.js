// RMP — API Client
// Ganti URL di bawah setelah deploy Apps Script sebagai New Deployment
const RMP_API_URL = 'https://script.google.com/macros/s/AKfycbw9RSEEcQVO8q9jI8hvJC6hq9U3g27ZrJjVR3Uaa7k0GROwTOXbnpRAYa5y30VHzP1f/exec';

const DB_RMP = {

  _pin: '',
  setPin(p) { this._pin = String(p).trim(); },

  async _get(params) {
    const res = await fetch(`${RMP_API_URL}?${new URLSearchParams(params)}`, { redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  },

  async _post(payload) {
    const ctrl  = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 15000);
    try {
      const res = await fetch(RMP_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        redirect: 'follow',
        signal: ctrl.signal,
      });
      let json;
      try { json = await res.json(); }
      catch { throw new Error('Response bukan JSON — periksa deployment Apps Script'); }
      return json;
    } catch (err) {
      if (err.name === 'AbortError') throw new Error('Timeout 15 detik');
      throw err;
    } finally { clearTimeout(timer); }
  },

  // Verifikasi PIN
  async verifyPin(pin) {
    try {
      const d = await this._get({ action: 'verifyPin', pin });
      return !!d.ok;
    } catch (e) { console.error('verifyPin:', e); return false; }
  },

  // Load AE hierarchy (tidak butuh PIN)
  async getConfig() {
    try {
      const d = await this._get({ action: 'getConfig' });
      if (d.error) throw new Error(d.error);
      return d.list || [];
    } catch (e) { console.error('getConfig:', e); return null; }
  },

  // Load semua activity
  async getActivities() {
    try {
      const d = await this._get({ action: 'getActivities', pin: this._pin });
      if (d.error) return null;
      return Array.isArray(d) ? d : [];
    } catch (e) { console.error('getActivities:', e); return null; }
  },

  // Submit activity baru
  async submitActivity(data) {
    try {
      return await this._post({ action: 'submitActivity', pin: this._pin, data });
    } catch (e) { console.error('submitActivity:', e); return { error: e.toString() }; }
  },

};

// ── HTML escape ──────────────────────────────────────────────────
function esc(str) {
  return String(str == null ? '' : str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ── Toast ─────────────────────────────────────────────────────────
function showToast(msg, type = 'info') {
  let t = document.getElementById('__toast__');
  if (!t) {
    t = document.createElement('div'); t.id = '__toast__';
    Object.assign(t.style, {
      position:'fixed', bottom:'24px', left:'50%', transform:'translateX(-50%)',
      padding:'12px 24px', borderRadius:'50px', fontWeight:'700', fontSize:'.88rem',
      zIndex:'9999', transition:'opacity .3s', boxShadow:'0 4px 20px rgba(0,0,0,.2)',
      fontFamily:'Segoe UI,system-ui,sans-serif', pointerEvents:'none',
    });
    document.body.appendChild(t);
  }
  const c = { success:{bg:'#dcfce7',color:'#16a34a'}, error:{bg:'#fee2e2',color:'#dc2626'},
              info:{bg:'#dbeafe',color:'#1d4ed8'}, loading:{bg:'#fef3c7',color:'#d97706'} }[type] || {bg:'#dbeafe',color:'#1d4ed8'};
  t.style.background = c.bg; t.style.color = c.color; t.style.opacity = '1';
  t.textContent = msg;
  clearTimeout(t.__t__);
  if (type !== 'loading') t.__t__ = setTimeout(() => { t.style.opacity = '0'; }, 3200);
}

// ── Loading overlay ───────────────────────────────────────────────
function showLoading(show) {
  let ov = document.getElementById('__loading__');
  if (!ov) {
    ov = document.createElement('div'); ov.id = '__loading__';
    Object.assign(ov.style, {
      position:'fixed', inset:'0', background:'rgba(255,255,255,.75)',
      display:'flex', alignItems:'center', justifyContent:'center',
      zIndex:'9998', fontFamily:'Segoe UI,system-ui,sans-serif',
      fontSize:'1rem', fontWeight:'700', color:'#1d4ed8',
    });
    ov.innerHTML = '<div style="text-align:center"><div style="font-size:2rem;margin-bottom:10px">⏳</div>Memuat data...</div>';
    document.body.appendChild(ov);
  }
  ov.style.display = show ? 'flex' : 'none';
}
