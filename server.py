"""
Fabric DV Image Embedder — Render cloud dashboard.

Architecture (mirrors Procurement / Puma pattern):
  Browser uploads Excel here → Render stores it as a pending session.
  Local dv_watcher.py polls pending sessions, runs Playwright to login to
  SharePoint, downloads images, embeds into Excel, pushes result back.
  All Playwright work happens on the local Windows machine — NOT on Render.

Push endpoints (require X-API-Key: DV_PUSH_API_KEY):
  POST /api/push/dv/status       watcher pushes live status
  POST /api/push/dv/interaction  watcher sets login / MFA prompt
  POST /api/push/dv/result       watcher pushes completed Excel
  GET  /api/dv/pending-sessions  watcher polls for pending work
  POST /api/dv/session-ack       watcher confirms pickup

Browser endpoints:
  GET  /                         dashboard HTML
  POST /api/dv/upload-excel      browser uploads Excel file
  GET  /api/dv/status            browser polls live status
  GET  /api/dv/interaction       browser polls interaction state
  POST /api/dv/interaction       browser submits login / MFA response
  GET  /api/dv/result            browser downloads completed Excel
  GET  /health                   health check
"""
from __future__ import annotations

import base64
import io
import json
import os
import threading
import uuid
from datetime import datetime

from flask import Flask, jsonify, request, send_file

app = Flask(__name__, static_folder='.', static_url_path='')
PORT    = int(os.environ.get('PORT', 10000))
API_KEY = os.environ.get('DV_PUSH_API_KEY', '')

_lock = threading.Lock()


def _idle_interaction() -> dict:
    return {'status': 'idle', 'type': None, 'message': '', 'display_num': '', 'error': '', 'response': None}


_state: dict = {
    'running':         False,
    'status':          'idle',   # idle | queued | processing | done | error
    'stage':           '',
    'message':         '',
    'done':            0,
    'total':           0,
    'ok_count':        0,
    'error_count':     0,
    'pending_sessions': [],      # [{request_id, filename, excel_b64, queued_at}]
    'result_b64':      None,
    'result_filename': None,
    'interaction':     _idle_interaction(),
    'last_push':       None,
}


def _check_key():
    if API_KEY and request.headers.get('X-API-Key') != API_KEY:
        return jsonify({'ok': False, 'error': 'Unauthorized'}), 403
    return None


# ── Static ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return app.send_static_file('image_tool.html')


@app.route('/health')
def health():
    return jsonify({'ok': True, 'time': datetime.now().isoformat()})


# ── Browser: upload Excel ─────────────────────────────────────────────

@app.route('/api/dv/upload-excel', methods=['POST'])
def upload_excel():
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'ok': False, 'error': 'No file provided'}), 400
    content = f.read()
    rid = uuid.uuid4().hex[:12]
    item = {
        'request_id': rid,
        'filename':   os.path.basename(f.filename),
        'excel_b64':  base64.b64encode(content).decode(),
        'queued_at':  datetime.now().isoformat(timespec='seconds'),
    }
    with _lock:
        _state['pending_sessions'].append(item)
        _state['status']          = 'queued'
        _state['stage']           = 'Waiting for watcher to pick up...'
        _state['result_b64']      = None
        _state['result_filename'] = None
        _state['interaction']     = _idle_interaction()
        _state['done']            = 0
        _state['total']           = 0
    return jsonify({'ok': True, 'request_id': rid})


# ── Watcher: poll / ack sessions ─────────────────────────────────────

@app.route('/api/dv/pending-sessions', methods=['GET'])
def pending_sessions():
    err = _check_key()
    if err:
        return err
    with _lock:
        sessions = list(_state['pending_sessions'])
    return jsonify({'ok': True, 'sessions': sessions})


@app.route('/api/dv/session-ack', methods=['POST'])
def session_ack():
    err = _check_key()
    if err:
        return err
    rid = (request.get_json(silent=True) or {}).get('request_id', '')
    with _lock:
        _state['pending_sessions'] = [s for s in _state['pending_sessions'] if s['request_id'] != rid]
    return jsonify({'ok': True})


# ── Watcher: push status ──────────────────────────────────────────────

@app.route('/api/push/dv/status', methods=['POST'])
def push_status():
    err = _check_key()
    if err:
        return err
    d = request.get_json(force=True, silent=True) or {}
    with _lock:
        for key in ('running', 'status', 'stage', 'message', 'done', 'total', 'ok_count', 'error_count'):
            if key in d:
                _state[key] = d[key]
        _state['last_push'] = datetime.now().isoformat(timespec='seconds')
    return jsonify({'ok': True})


# ── Watcher: push interaction (login / MFA prompt) ────────────────────

@app.route('/api/push/dv/interaction', methods=['POST'])
def push_interaction():
    err = _check_key()
    if err:
        return err
    d = request.get_json(force=True, silent=True) or {}
    with _lock:
        current = dict(_state['interaction'])
        # Never overwrite a pending answered response until watcher explicitly clears it
        if current.get('status') != 'answered' or d.get('status') == 'idle':
            _state['interaction'] = {
                'status':      d.get('status', 'idle'),
                'type':        d.get('type'),
                'message':     d.get('message', ''),
                'display_num': d.get('display_num', ''),
                'error':       d.get('error', ''),
                'response':    None,
            }
    return jsonify({'ok': True})


# ── Watcher: push completed Excel ─────────────────────────────────────

@app.route('/api/push/dv/result', methods=['POST'])
def push_result():
    err = _check_key()
    if err:
        return err
    d = request.get_json(force=True, silent=True) or {}
    with _lock:
        _state['result_b64']      = d.get('excel_b64')
        _state['result_filename'] = d.get('filename', 'Output - With Images.xlsx')
        _state['status']          = 'done'
        _state['stage']           = 'Done'
        _state['running']         = False
        _state['interaction']     = _idle_interaction()
    return jsonify({'ok': True})


# ── Browser: poll status ──────────────────────────────────────────────

@app.route('/api/dv/status', methods=['GET'])
def get_status():
    with _lock:
        return jsonify({
            'running':     _state['running'],
            'status':      _state['status'],
            'stage':       _state['stage'],
            'message':     _state['message'],
            'done':        _state['done'],
            'total':       _state['total'],
            'ok_count':    _state['ok_count'],
            'error_count': _state['error_count'],
            'has_result':  bool(_state['result_b64']),
            'interaction': dict(_state['interaction']),
        })


# ── Browser: get / post interaction (login credentials / MFA) ─────────

@app.route('/api/dv/interaction', methods=['GET'])
def get_interaction():
    with _lock:
        return jsonify(dict(_state['interaction']))


@app.route('/api/dv/interaction', methods=['POST'])
def post_interaction():
    d = request.get_json(silent=True) or {}
    with _lock:
        _state['interaction']['response'] = d.get('response')
        _state['interaction']['status']   = 'answered'
    return jsonify({'ok': True})


# ── Browser: download completed Excel ────────────────────────────────

@app.route('/api/dv/result', methods=['GET'])
def get_result():
    with _lock:
        b64      = _state['result_b64']
        filename = _state['result_filename'] or 'Output - With Images.xlsx'
    if not b64:
        return jsonify({'ok': False, 'error': 'No result available'}), 404
    return send_file(
        io.BytesIO(base64.b64decode(b64)),
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
    )


if __name__ == '__main__':
    print(f'\n{"=" * 50}\n  Open in browser: http://localhost:{PORT}\n{"=" * 50}\n')
    app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
