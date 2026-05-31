const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

let supabase = null;
if (process.env.SUPABASE_URL && process.env.SUPABASE_KEY) {
  const { createClient } = require('@supabase/supabase-js');
  supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);
}

app.get('/', (req, res) => res.json({ status: 'SAK API v2', supabase: !!supabase }));

async function saveProposal(payload) {
  if (!supabase) return null;
  try {
    const { data } = await supabase.from('proposals').upsert({
      ref: payload.ref, rev: payload.rev, date: payload.date,
      client_name: payload.client_name, client_address: payload.client_address,
      client_attn: payload.client_attn, client_phone: payload.client_phone,
      proj_scope: payload.proj_scope, proj_block: payload.proj_block,
      proj_location: payload.proj_location, contract_type: payload.contract_type,
      duration: payload.duration, duration_unit: payload.duration_unit,
      pay_advance: parseFloat(payload.pay_advance) || 0,
      pay_progress: parseFloat(payload.pay_progress) || 0,
      pay_retention: parseFloat(payload.pay_retention) || 0,
      validity: parseFloat(payload.validity) || 15,
      base_price: payload.base_price || 0, additions: payload.additions || 0,
      omissions: payload.omissions || 0, missed: payload.missed || 0,
      extras_amount: payload.extras_amount || 0,
      grand_total: payload.grand_total || 0,
      status: 'draft', payload
    }, { onConflict: 'ref' }).select('id').single();
    return data?.id || null;
  } catch (e) { console.error('Supabase:', e.message); return null; }
}

app.get('/proposals', async (req, res) => {
  if (!supabase) return res.json({ proposals: [] });
  const { data, error } = await supabase.from('proposals')
    .select('id,ref,rev,date,client_name,proj_block,grand_total,status,created_at')
    .order('created_at', { ascending: false }).limit(100);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ proposals: data || [] });
});

app.get('/proposals/:id', async (req, res) => {
  if (!supabase) return res.status(404).json({ error: 'No DB' });
  const { data, error } = await supabase.from('proposals')
    .select('payload,status').eq('id', req.params.id).single();
  if (error) return res.status(404).json({ error: error.message });
  res.json(data);
});

app.patch('/proposals/:id/status', async (req, res) => {
  if (!supabase) return res.status(404).json({ error: 'No DB' });
  const { status } = req.body;
  if (!['draft','sent','won','lost'].includes(status))
    return res.status(400).json({ error: 'Invalid status' });
  const { error } = await supabase.from('proposals')
    .update({ status }).eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// Generate PDF from DOCX using html-pdf-node as fallback
async function docxToPdf(docxPath, pdfPath) {
  // Try LibreOffice first
  try {
    execSync(`libreoffice --headless --convert-to pdf --outdir "${path.dirname(pdfPath)}" "${docxPath}"`, { timeout: 60000 });
    const generated = docxPath.replace('.docx', '.pdf');
    if (fs.existsSync(generated)) {
      if (generated !== pdfPath) fs.renameSync(generated, pdfPath);
      return true;
    }
  } catch (_) {}

  // Fallback: generate simple PDF via html-pdf-node
  try {
    const htmlPdf = require('html-pdf-node');
    const payload = JSON.parse(fs.readFileSync(docxPath.replace('.docx', '_payload.json'), 'utf8'));
    const html = buildHtmlForPdf(payload);
    const file = { content: html };
    const options = { format: 'A4', margin: { top: '20mm', bottom: '20mm', left: '20mm', right: '20mm' } };
    const pdfBuffer = await htmlPdf.generatePdf(file, options);
    fs.writeFileSync(pdfPath, pdfBuffer);
    return true;
  } catch (e) {
    console.error('PDF fallback error:', e.message);
    return false;
  }
}

function fmt(n) { return Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }

function buildHtmlForPdf(p) {
  const base = (p.price_rows||[]).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
  const grand = p.grand_total || base;
  const rows = (p.price_rows||[]).filter(r=>r.amount>0).map(r=>
    `<tr><td>${r.div}</td><td>${r.desc}</td><td style="text-align:right">SAR ${fmt(r.amount)}</td></tr>`
  ).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body{font-family:Arial,sans-serif;font-size:12px;color:#111;margin:0;padding:20px}
    .header{background:#1a3a6b;color:white;padding:16px 20px;margin-bottom:20px}
    .header h1{font-size:16px;margin:0 0 4px}
    .header p{margin:0;font-size:11px;opacity:.8}
    .ref{float:right;font-size:11px;font-family:monospace}
    table{width:100%;border-collapse:collapse;margin-bottom:16px}
    th{background:#1a3a6b;color:white;padding:7px 10px;text-align:left;font-size:11px}
    td{padding:6px 10px;border-bottom:1px solid #e5e7eb;font-size:11px}
    tr:nth-child(even) td{background:#f9fafb}
    .total-row td{font-weight:bold;background:#1a3a6b;color:white}
    .section{margin-bottom:16px}
    .section-title{font-size:13px;font-weight:bold;color:#1a3a6b;border-bottom:2px solid #1a3a6b;padding-bottom:4px;margin-bottom:10px}
    .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
    .info-item label{font-size:10px;color:#6b7280;display:block}
    .info-item span{font-size:12px;font-weight:500}
    .grand{font-size:15px;font-weight:bold;color:#1a3a6b;text-align:right;padding:10px;background:#e8edf5;border-radius:4px}
    ul{margin:0;padding-left:16px} li{font-size:11px;margin-bottom:3px}
    .footer{margin-top:30px;font-size:10px;color:#9ca3af;text-align:center;border-top:1px solid #e5e7eb;padding-top:10px}
  </style></head><body>
  <div class="header">
    <span class="ref">${p.ref||''}</span>
    <h1>TECHNO-COMMERCIAL PROPOSAL — ${p.rev||'R01'}</h1>
    <p>Soroh Al-Khair General Contracting &nbsp;|&nbsp; ${p.date||''}</p>
  </div>

  <div class="section">
    <div class="section-title">Client & Project Information</div>
    <div class="info-grid">
      <div class="info-item"><label>Client</label><span>${p.client_name||''}</span></div>
      <div class="info-item"><label>Address</label><span>${p.client_address||''}</span></div>
      <div class="info-item"><label>Attention</label><span>${p.client_attn||''}</span></div>
      <div class="info-item"><label>Phone</label><span>${p.client_phone||''}</span></div>
      <div class="info-item"><label>Scope</label><span>${p.proj_scope||''}</span></div>
      <div class="info-item"><label>Block</label><span>${p.proj_block||''}</span></div>
      <div class="info-item"><label>Location</label><span>${p.proj_location||''}</span></div>
      <div class="info-item"><label>Validity</label><span>${p.validity||15} days</span></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Commercial Terms</div>
    <div class="info-grid">
      <div class="info-item"><label>Contract Type</label><span>${p.contract_type||'LUMPSUM'}</span></div>
      <div class="info-item"><label>Duration</label><span>${p.duration||12} ${p.duration_unit||'Months'}</span></div>
      <div class="info-item"><label>Payment</label><span>${p.pay_advance||20}% Advance / ${p.pay_progress||75}% Progress / ${p.pay_retention||5}% Retention</span></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Bill of Quantities — Summary</div>
    <table>
      <thead><tr><th>Division</th><th>Description</th><th>Amount (SAR)</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot>
        ${p.additions>0?`<tr><td colspan="2">Additions</td><td style="text-align:right">SAR ${fmt(p.additions)}</td></tr>`:''}
        ${p.omissions>0?`<tr><td colspan="2">Omissions</td><td style="text-align:right">- SAR ${fmt(p.omissions)}</td></tr>`:''}
        ${p.missed>0?`<tr><td colspan="2">Missed Items</td><td style="text-align:right">SAR ${fmt(p.missed)}</td></tr>`:''}
        <tr class="total-row"><td colspan="2">TOTAL AMOUNT (EXCLUDED VAT)</td><td style="text-align:right">SAR ${fmt(grand)}</td></tr>
      </tfoot>
    </table>
  </div>

  ${p.assumptions?.length?`<div class="section"><div class="section-title">Assumptions</div><ul>${p.assumptions.map(a=>`<li>${a}</li>`).join('')}</ul></div>`:''}
  ${p.exclusions?.length?`<div class="section"><div class="section-title">Exclusions</div><ul>${p.exclusions.map(e=>`<li>${e}</li>`).join('')}</ul></div>`:''}

  <div class="footer">
    Bandariyah Center, Prince Faisal Bin Fahad Street, Al Khobar, KSA &nbsp;|&nbsp; C.R: 2052002095 &nbsp;|&nbsp; www.sakcontracting.com
  </div>
  </body></html>`;
}

app.post('/generate', async (req, res) => {
  const { payload, format = 'both' } = req.body;
  if (!payload?.proj_block) return res.status(400).json({ error: 'Missing payload' });

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sak_'));
  const block = payload.proj_block.replace(/\s+/g, '_');
  const rev = payload.rev || 'R01';
  const docxName = `SAK_Proposal_${block}_${rev}.docx`;
  const pdfName = `SAK_Proposal_${block}_${rev}.pdf`;
  const docxPath = path.join(tmpDir, docxName);
  const pdfPath = path.join(tmpDir, pdfName);
  const payloadJsonPath = path.join(tmpDir, 'payload.json');

  try {
    fs.writeFileSync(payloadJsonPath, JSON.stringify(payload));
    // Save payload for PDF fallback
    fs.writeFileSync(docxPath.replace('.docx','_payload.json'), JSON.stringify(payload));

    const script = path.join(__dirname, 'generate.js');
    execSync(`SAK_OUTDIR="${tmpDir}" node "${script}" '${JSON.stringify(payload).replace(/'/g,"'\\''")}'`,
      { timeout: 90000, shell: true });

    if (!fs.existsSync(docxPath)) throw new Error('DOCX not generated');
    saveProposal(payload);

    if (format === 'docx') {
      res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition',`attachment; filename="${docxName}"`);
      return res.send(fs.readFileSync(docxPath));
    }

    const pdfOk = await docxToPdf(docxPath, pdfPath);

    if (format === 'pdf') {
      if (!pdfOk || !fs.existsSync(pdfPath)) throw new Error('PDF generation failed');
      res.setHeader('Content-Type','application/pdf');
      res.setHeader('Content-Disposition',`attachment; filename="${pdfName}"`);
      return res.send(fs.readFileSync(pdfPath));
    }

    // both
    const docxB64 = fs.readFileSync(docxPath).toString('base64');
    const pdfB64 = pdfOk && fs.existsSync(pdfPath)
      ? fs.readFileSync(pdfPath).toString('base64')
      : null;

    cleanup(tmpDir);
    return res.json({
      success: true,
      docx: { name: docxName, data: docxB64 },
      pdf: pdfB64 ? { name: pdfName, data: pdfB64 } : null
    });

  } catch (err) {
    cleanup(tmpDir);
    return res.status(500).json({ error: err.message });
  }
});

function cleanup(dir) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch(_) {}
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`SAK API v2 on port ${PORT}`));
