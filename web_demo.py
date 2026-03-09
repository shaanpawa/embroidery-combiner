"""
Web demo for testing the Embroidery Combiner logic.
Not production — just a clean test interface.
"""

import json
import os
import struct

from flask import Flask, render_template_string, request, jsonify

from app.core.file_discovery import discover_folder, generate_output_name
from app.core.validator import validate_batch, is_valid_for_combining
from app.core.combiner import combine_designs, save_combined, validate_combined_output
from app.core.converter import check_conversion_capability, batch_convert, cleanup_temp_files

app = Flask(__name__)


HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Embroidery Combiner</title>
<style>
  @font-face {
    font-family: 'Geist';
    src: url('https://cdn.jsdelivr.net/npm/geist@1.3.1/dist/fonts/geist-sans/Geist-Regular.woff2') format('woff2');
    font-weight: 400;
  }
  @font-face {
    font-family: 'Geist';
    src: url('https://cdn.jsdelivr.net/npm/geist@1.3.1/dist/fonts/geist-sans/Geist-Medium.woff2') format('woff2');
    font-weight: 500;
  }
  @font-face {
    font-family: 'Geist Mono';
    src: url('https://cdn.jsdelivr.net/npm/geist@1.3.1/dist/fonts/geist-mono/GeistMono-Regular.woff2') format('woff2');
    font-weight: 400;
  }

  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    font-family: 'Geist', -apple-system, system-ui, sans-serif;
    background: #0a0a0a;
    color: #e0e0e0;
    min-height: 100vh;
    -webkit-font-smoothing: antialiased;
    font-size: 14px;
    line-height: 1.5;
  }

  .container {
    max-width: 640px;
    margin: 0 auto;
    padding: 48px 24px;
  }

  h1 {
    font-size: 18px;
    font-weight: 500;
    letter-spacing: -0.3px;
    margin-bottom: 32px;
  }
  h1 span { color: #444; font-weight: 400; font-size: 13px; margin-left: 8px; }

  /* Folder input */
  .folder-row {
    display: flex;
    gap: 1px;
    background: #1a1a1a;
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 24px;
  }
  .folder-input {
    flex: 1;
    background: #111;
    border: none;
    color: #e0e0e0;
    padding: 12px 14px;
    font-size: 13px;
    font-family: 'Geist Mono', monospace;
    outline: none;
  }
  .folder-input::placeholder { color: #333; }
  .folder-input:focus { background: #141414; }
  .folder-btn {
    background: #161616;
    border: none;
    color: #888;
    padding: 12px 20px;
    font-size: 13px;
    font-family: 'Geist', sans-serif;
    cursor: pointer;
  }
  .folder-btn:hover { background: #1c1c1c; color: #ccc; }

  /* File list */
  .files {
    border: 1px solid #1a1a1a;
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 24px;
    display: none;
  }
  .files.show { display: block; }
  .files-header {
    display: grid;
    grid-template-columns: 50px 1fr 70px 100px;
    padding: 8px 14px;
    font-size: 11px;
    color: #444;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    border-bottom: 1px solid #1a1a1a;
    background: #0d0d0d;
  }
  .files-body { max-height: 360px; overflow-y: auto; }
  .files-body::-webkit-scrollbar { width: 3px; }
  .files-body::-webkit-scrollbar-thumb { background: #222; border-radius: 2px; }

  .file-row {
    display: grid;
    grid-template-columns: 50px 1fr 70px 100px;
    padding: 8px 14px;
    font-size: 13px;
    border-bottom: 1px solid #111;
    align-items: center;
  }
  .file-row:last-child { border-bottom: none; }
  .file-row:hover { background: #0e0e0e; }
  .file-num { font-family: 'Geist Mono', monospace; font-size: 12px; color: #444; text-align: right; padding-right: 12px; }
  .file-name { font-family: 'Geist Mono', monospace; color: #bbb; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file-size { font-family: 'Geist Mono', monospace; font-size: 12px; color: #333; text-align: right; padding-right: 12px; }
  .file-status { font-size: 11px; font-family: 'Geist Mono', monospace; }
  .s-ok { color: #5a5; }
  .s-warn { color: #a85; }
  .s-err { color: #a55; }
  .s-wait { color: #333; }

  .empty { padding: 32px; text-align: center; color: #333; font-size: 13px; }

  /* Warnings */
  .warnings { margin-bottom: 16px; display: none; }
  .warnings.show { display: block; }
  .warn-item { font-size: 12px; color: #666; padding: 2px 0; }
  .warn-item::before { content: ''; display: inline-block; width: 4px; height: 4px; border-radius: 50%; background: #a88530; margin-right: 8px; vertical-align: middle; }

  /* Config */
  .config {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 24px;
    display: none;
  }
  .config.show { display: flex; }
  .config label { font-size: 12px; color: #555; }
  .gap-input {
    width: 52px;
    background: #111;
    border: 1px solid #1a1a1a;
    border-radius: 6px;
    color: #e0e0e0;
    padding: 7px 8px;
    font-size: 13px;
    font-family: 'Geist Mono', monospace;
    text-align: center;
    outline: none;
  }
  .gap-input:focus { border-color: #333; }
  .unit { font-size: 11px; color: #444; font-family: 'Geist Mono', monospace; }
  .sep { width: 1px; height: 20px; background: #1a1a1a; }
  .output-input {
    background: #111;
    border: 1px solid #1a1a1a;
    border-radius: 6px;
    color: #e0e0e0;
    padding: 7px 10px;
    font-size: 13px;
    font-family: 'Geist Mono', monospace;
    outline: none;
    width: 140px;
  }
  .output-input:focus { border-color: #333; }

  /* Combine button */
  .combine-btn {
    background: #e0e0e0;
    color: #0a0a0a;
    border: none;
    padding: 10px 32px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
    font-family: 'Geist', sans-serif;
    cursor: pointer;
    letter-spacing: -0.2px;
    display: none;
  }
  .combine-btn.show { display: inline-block; }
  .combine-btn:hover { opacity: 0.85; }
  .combine-btn:disabled { background: #1a1a1a; color: #333; cursor: not-allowed; opacity: 1; }

  /* Result */
  .result {
    margin-top: 24px;
    border: 1px solid #1a1a1a;
    border-radius: 8px;
    padding: 18px;
    display: none;
  }
  .result.show { display: block; }
  .result.ok { border-color: #1a2a1a; }
  .result.err { border-color: #2a1a1a; }
  .result-title { font-size: 14px; font-weight: 500; margin-bottom: 10px; }
  .result-title.ok { color: #5a5; }
  .result-title.err { color: #a55; }
  .result-stats {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 8px;
  }
  .stat-label { font-size: 10px; color: #444; text-transform: uppercase; letter-spacing: 0.5px; }
  .stat-value { font-size: 14px; font-family: 'Geist Mono', monospace; color: #bbb; }

  .status { font-size: 11px; color: #333; font-family: 'Geist Mono', monospace; margin-top: 12px; }
</style>
</head>
<body>
<div class="container">
  <h1>Embroidery Combiner <span>test</span></h1>

  <div class="folder-row">
    <input class="folder-input" id="folder" placeholder="/path/to/designs" spellcheck="false">
    <button class="folder-btn" onclick="load()">Load</button>
  </div>

  <div class="warnings" id="warnings"></div>

  <div class="files" id="files">
    <div class="files-header">
      <span style="text-align:right;padding-right:12px">#</span>
      <span>File</span>
      <span style="text-align:right;padding-right:12px">Size</span>
      <span>Status</span>
    </div>
    <div class="files-body" id="filesBody"></div>
  </div>

  <div class="config" id="config">
    <label>Gap</label>
    <input class="gap-input" id="gap" type="number" value="3.0" min="0" max="50" step="0.5">
    <span class="unit">mm</span>
    <div class="sep"></div>
    <label>Output</label>
    <input class="output-input" id="output" value="combined.dst" spellcheck="false">
  </div>

  <button class="combine-btn" id="combineBtn" disabled onclick="combine()">Combine</button>

  <div class="result" id="result">
    <div class="result-title" id="resultTitle"></div>
    <div class="result-stats" id="resultStats"></div>
  </div>

  <div class="status" id="status"></div>
</div>

<script>
let files = [];

async function load() {
  const folder = document.getElementById('folder').value.trim();
  if (!folder) return;

  document.getElementById('status').textContent = 'loading...';

  const resp = await fetch('/api/discover', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({folder})
  });
  const data = await resp.json();
  files = data.files || [];

  // Warnings
  const wEl = document.getElementById('warnings');
  if (data.warnings && data.warnings.length) {
    wEl.innerHTML = data.warnings.map(w => '<div class="warn-item">' + esc(w) + '</div>').join('');
    wEl.classList.add('show');
  } else {
    wEl.classList.remove('show');
  }

  // File list
  const fEl = document.getElementById('files');
  const body = document.getElementById('filesBody');

  if (!files.length) {
    fEl.classList.add('show');
    body.innerHTML = '<div class="empty">No embroidery files found</div>';
    hide('config'); hide('combineBtn');
    document.getElementById('status').textContent = '';
    return;
  }

  fEl.classList.add('show');
  body.innerHTML = files.map((f, i) =>
    '<div class="file-row">' +
      '<span class="file-num">' + (f.number !== null ? f.number : '\u2014') + '</span>' +
      '<span class="file-name">' + esc(f.filename) + '</span>' +
      '<span class="file-size">' + esc(f.size_display) + '</span>' +
      '<span class="file-status s-wait" id="st-' + i + '">\u00b7\u00b7\u00b7</span>' +
    '</div>'
  ).join('');

  document.getElementById('output').value = data.output_name || 'combined.dst';
  show('config'); show('combineBtn');

  // Reset result
  document.getElementById('result').className = 'result';

  document.getElementById('status').textContent = files.length + ' files found';
  validate(folder);
}

async function validate(folder) {
  const resp = await fetch('/api/validate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({folder})
  });
  const data = await resp.json();
  let valid = 0;

  (data.results || []).forEach((r, i) => {
    const el = document.getElementById('st-' + i);
    if (!el) return;
    if (r.errors && r.errors.length) {
      el.textContent = r.errors[0];
      el.className = 'file-status s-err';
    } else if (r.warnings && r.warnings.length) {
      el.textContent = r.summary;
      el.className = 'file-status s-warn';
      valid++;
    } else {
      el.textContent = r.summary;
      el.className = 'file-status s-ok';
      valid++;
    }
  });

  document.getElementById('combineBtn').disabled = valid === 0;
  document.getElementById('status').textContent = valid + ' valid, ready to combine';
}

async function combine() {
  const folder = document.getElementById('folder').value.trim();
  const gap = parseFloat(document.getElementById('gap').value) || 3.0;
  const outputName = document.getElementById('output').value.trim() || 'combined.dst';
  const paths = files.filter(f => f.included).map(f => f.path);
  if (!paths.length) return;

  document.getElementById('combineBtn').disabled = true;
  document.getElementById('status').textContent = 'combining...';

  try {
    const resp = await fetch('/api/combine', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({folder, gap, output_name: outputName, files: paths})
    });
    const data = await resp.json();
    const r = document.getElementById('result');

    if (data.success) {
      r.className = 'result show ok';
      document.getElementById('resultTitle').textContent = 'Combined successfully';
      document.getElementById('resultTitle').className = 'result-title ok';
      document.getElementById('resultStats').innerHTML =
        stat('Output', outputName) +
        stat('Stitches', data.info.stitch_count.toLocaleString()) +
        stat('Size', data.info.width_mm.toFixed(1) + ' \u00d7 ' + data.info.height_mm.toFixed(1) + ' mm') +
        stat('Colors', data.info.color_count);
      document.getElementById('status').textContent = 'saved \u2192 ' + outputName;

      // Update file statuses
      files.forEach((f, i) => {
        if (f.included) {
          const el = document.getElementById('st-' + i);
          if (el) { el.textContent = 'done'; el.className = 'file-status s-ok'; }
        }
      });
    } else {
      r.className = 'result show err';
      document.getElementById('resultTitle').textContent = data.error;
      document.getElementById('resultTitle').className = 'result-title err';
      document.getElementById('resultStats').innerHTML = '';
      document.getElementById('status').textContent = 'failed';
    }
  } catch (e) {
    document.getElementById('status').textContent = 'error: ' + e.message;
  }

  document.getElementById('combineBtn').disabled = false;
}

function stat(label, value) {
  return '<div><span class="stat-label">' + label + '</span><div class="stat-value">' + value + '</div></div>';
}
function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
function show(id) { document.getElementById(id).classList.add('show'); }
function hide(id) { document.getElementById(id).classList.remove('show'); }

document.getElementById('folder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/api/discover', methods=['POST'])
def api_discover():
    data = request.json
    folder = data.get('folder', '')

    result = discover_folder(folder)

    files_data = []
    for f in result.files:
        files_data.append({
            'path': f.path,
            'filename': f.filename,
            'extension': f.extension,
            'number': f.number,
            'size_display': f.size_display,
            'included': f.included,
        })

    output_name = generate_output_name(result.files)

    return jsonify({
        'files': files_data,
        'ngs_count': result.ngs_count,
        'dst_count': result.dst_count,
        'warnings': result.warnings,
        'skipped_files': result.skipped_files,
        'output_name': output_name,
    })


@app.route('/api/validate', methods=['POST'])
def api_validate():
    data = request.json
    folder = data.get('folder', '')

    result = discover_folder(folder)
    paths = [f.path for f in result.files]
    val_results = validate_batch(paths)

    results_data = []
    for vr in val_results:
        results_data.append({
            'filename': vr.filename,
            'valid': vr.valid,
            'status': vr.status,
            'summary': vr.summary,
            'errors': vr.errors,
            'warnings': vr.warnings,
            'stitch_count': vr.stitch_count,
            'width_mm': vr.width_mm,
            'height_mm': vr.height_mm,
        })

    return jsonify({'results': results_data})


@app.route('/api/combine', methods=['POST'])
def api_combine():
    import tempfile
    data = request.json
    folder = data.get('folder', '')
    gap = data.get('gap', 3.0)
    output_name = data.get('output_name', 'combined.dst')
    file_paths = data.get('files', [])

    if not file_paths:
        return jsonify({'success': False, 'error': 'No files selected'})

    ngs_paths = [p for p in file_paths if p.lower().endswith('.ngs')]
    dst_paths = [p for p in file_paths if p.lower().endswith('.dst')]

    temp_dir = None
    conversion_results = None

    # If there are NGS files, try to convert them
    if ngs_paths:
        capable, msg = check_conversion_capability()
        if not capable:
            return jsonify({'success': False, 'error': msg})

        try:
            temp_dir = tempfile.mkdtemp(prefix="embroidery_")
            conversion_results = batch_convert(ngs_paths, temp_dir)

            for cr in conversion_results:
                if cr.success and cr.dst_path:
                    dst_paths.append(cr.dst_path)

            failed = [cr for cr in conversion_results if not cr.success]
            if failed:
                names = [os.path.basename(cr.ngs_path) for cr in failed]
                if not dst_paths:
                    return jsonify({
                        'success': False,
                        'error': f'All conversions failed: {", ".join(names)}'
                    })
        except Exception as e:
            return jsonify({'success': False, 'error': f'Conversion error: {e}'})

    if not dst_paths:
        return jsonify({'success': False, 'error': 'No DST files to combine'})

    try:
        if not output_name.lower().endswith('.dst'):
            output_name += '.dst'
        output_path = os.path.join(folder, output_name)

        # Sort by number
        from app.core.file_discovery import extract_number
        dst_paths.sort(key=lambda p: extract_number(os.path.basename(p)) or 0)

        combined = combine_designs(dst_paths, gap_mm=gap)
        save_combined(combined, output_path, overwrite=True)
        info = validate_combined_output(output_path)

        # Cleanup temp conversion files
        if temp_dir and conversion_results:
            cleanup_temp_files(conversion_results, temp_dir)

        return jsonify({'success': True, 'info': info, 'output_path': output_path})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


if __name__ == '__main__':
    app.run(port=5123, debug=False)
