const express  = require('express');
const cors     = require('cors');
const { execSync } = require('child_process');
const fs       = require('fs');
const path     = require('path');
const os       = require('os');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ── Supabase client (optional — works without it) ─────────────────────────────
let supabase = null;
if (process.env.SUPABASE_URL && process.env.SUPABASE_KEY) {
  const { createClient } = require('@supabase/supabase-js');
  supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);
  console.log('Supabase connected');
} else {
  console.log('Supabase not configured — running without DB');
}

// ── Health ────────────────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'SAK Proposal API', version: '2.0',
             supabase: !!supabase });
});

// ── Save proposal to Supabase ─────────────────────────────────────────────────
async function saveProposal(payload) {
  if (!supabase) return null;
  try {
    const { data, error } = await supabase
      .from('proposals')
      .upsert({
        ref:           payload.ref,
        rev:           payload.rev,
        date:          payload.date,
        client_name:   payload.client_name,
        client_address:payload.client_address,
        client_attn:   payload.client_attn,
        client_phone:  payload.client_phone,
        proj_scope:    payload.proj_scope,
        proj_block:    payload.proj_block,
        proj_location: payload.proj_location,
        contract_type: payload.contract_type,
        duration:      payload.duration,
        duration_unit: payload.duration_unit,
        pay_advance:   parseFloat(payload.pay_advance)  || 0,
        pay_progress:  parseFloat(payload.pay_progress) || 0,
        pay_retention: parseFloat(payload.pay_retention)|| 0,
        validity:      parseFloat(payload.validity)     || 15,
        base_price:    payload.base_price   || 0,
        additions:     payload.additions    || 0,
        omissions:     payload.omissions    || 0,
        missed:        payload.missed       || 0,
        extras_amount: payload.extras_amount|| 0,
        grand_total:   payload.grand_total  || 0,
        status:        'draft',
        payload:       payload
      }, { onConflict: 'ref' })
      .select('id')
      .single();
    if (error) console.error('Supabase save error:', error.message);
    return data?.id || null;
  } catch (e) {
    console.error('Supabase error:', e.message);
    return null;
  }
}

// ── List proposals ────────────────────────────────────────────────────────────
app.get('/proposals', async (req, res) => {
  if (!supabase) return res.json({ proposals: [] });
  const { data, error } = await supabase
    .from('proposals')
    .select('id,ref,rev,date,client_name,proj_block,grand_total,status,created_at')
    .order('created_at', { ascending: false })
    .limit(100);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ proposals: data || [] });
});

// ── Get single proposal payload ───────────────────────────────────────────────
app.get('/proposals/:id', async (req, res) => {
  if (!supabase) return res.status(404).json({ error: 'No DB' });
  const { data, error } = await supabase
    .from('proposals')
    .select('payload,status')
    .eq('id', req.params.id)
    .single();
  if (error) return res.status(404).json({ error: error.message });
  res.json(data);
});

// ── Update proposal status ────────────────────────────────────────────────────
app.patch('/proposals/:id/status', async (req, res) => {
  if (!supabase) return res.status(404).json({ error: 'No DB' });
  const { status } = req.body;
  const allowed = ['draft','sent','won','lost'];
  if (!allowed.includes(status))
    return res.status(400).json({ error: 'Invalid status' });
  const { error } = await supabase
    .from('proposals')
    .update({ status })
    .eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

// ── Generate DOCX + PDF ───────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  const { payload, format = 'both' } = req.body;
  if (!payload?.proj_block)
    return res.status(400).json({ error: 'Missing payload or proj_block' });

  const tmpDir  = fs.mkdtempSync(path.join(os.tmpdir(), 'sak_'));
  const block   = payload.proj_block.replace(/\s+/g, '_');
  const rev     = payload.rev || 'R01';
  const docxName= `SAK_Proposal_${block}_${rev}.docx`;
  const pdfName = `SAK_Proposal_${block}_${rev}.pdf`;

  try {
    // Write payload JSON to temp file
    const jsonPath = path.join(tmpDir, 'payload.json');
    fs.writeFileSync(jsonPath, JSON.stringify(payload));

    // Run generate.js
    const script = path.join(__dirname, 'generate.js');
    execSync(`SAK_OUTDIR="${tmpDir}" node "${script}" "$(cat "${jsonPath}")"`,
      { timeout: 90000, shell: true });

    const docxPath = path.join(tmpDir, docxName);
    if (!fs.existsSync(docxPath))
      throw new Error('DOCX not found after generation');

    // Save to Supabase (non-blocking)
    saveProposal(payload);

    // Return based on format
    if (format === 'docx') {
      res.setHeader('Content-Type',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename="${docxName}"`);
      const buf = fs.readFileSync(docxPath);
      cleanup(tmpDir); return res.send(buf);
    }

    // Convert to PDF
    execSync(
      `libreoffice --headless --convert-to pdf --outdir "${tmpDir}" "${docxPath}"`,
      { timeout: 90000 });

    const pdfPath = docxPath.replace('.docx', '.pdf');
    if (!fs.existsSync(pdfPath)) throw new Error('PDF not generated');

    if (format === 'pdf') {
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${pdfName}"`);
      const buf = fs.readFileSync(pdfPath);
      cleanup(tmpDir); return res.send(buf);
    }

    // both → base64 JSON
    const docxB64 = fs.readFileSync(docxPath).toString('base64');
    const pdfB64  = fs.readFileSync(pdfPath).toString('base64');
    cleanup(tmpDir);
    return res.json({
      success: true,
      docx: { name: docxName, data: docxB64 },
      pdf:  { name: pdfName,  data: pdfB64  }
    });

  } catch (err) {
    cleanup(tmpDir);
    console.error('Generate error:', err.message);
    return res.status(500).json({ error: err.message });
  }
});

function cleanup(dir) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch(_) {}
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`SAK API v2 on port ${PORT}`));
