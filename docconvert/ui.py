"""Tiny local web UI: ``docconvert-ui`` opens a page where you drop a file,
pick a format, and download the result. Standard library only."""

from __future__ import annotations

import io
import json
import re
import tempfile
import threading
import webbrowser
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from .core import SUPPORTED, convert

PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>docconvert</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
 body{font-family:-apple-system,system-ui,sans-serif;background:#f5f5f7;margin:0;display:flex;
      min-height:100vh;align-items:center;justify-content:center;color:#1d1d1f}
 main{background:#fff;border-radius:16px;padding:32px;width:min(440px,92vw);box-shadow:0 8px 30px rgba(0,0,0,.08)}
 h1{margin:0 0 4px;font-size:22px} p{margin:0 0 20px;color:#6e6e73;font-size:14px}
 #drop{display:block;border:2px dashed #c7c7cc;border-radius:12px;padding:36px 16px;text-align:center;cursor:pointer;
       transition:.15s} #drop.over{border-color:#0071e3;background:#eef5ff}
 #drop strong{display:block;font-size:15px} #drop span{color:#6e6e73;font-size:13px}
 #name{margin-top:10px;font-size:13px;color:#0071e3;word-break:break-all}
 .row{display:flex;gap:10px;margin-top:18px;align-items:center}
 select,button{font:inherit;font-size:15px;border-radius:10px;padding:10px 14px;border:1px solid #c7c7cc;background:#fff}
 button{flex:1;background:#0071e3;color:#fff;border:0;cursor:pointer;font-weight:600}
 button:disabled{background:#c7c7cc;cursor:default}
 #msg{margin-top:14px;font-size:13px;min-height:18px} .err{color:#d70015}
 input[type=file]{display:none}
</style></head><body><main>
<h1>docconvert</h1><p>Drop a document, choose a format, download.</p>
<label id="drop" for="file"><strong>Drop a file here or click to choose</strong>
<span>FORMATS</span><div id="name"></div></label>
<input type="file" id="file" accept="ACCEPT">
<div class="row">
 <label for="to">Convert to</label>
 <select id="to">OPTIONS</select>
 <button id="go" disabled>Convert</button>
</div>
<div id="msg"></div>
</main>
<script>
const drop=document.getElementById('drop'),input=document.getElementById('file'),
 name=document.getElementById('name'),go=document.getElementById('go'),
 to=document.getElementById('to'),msg=document.getElementById('msg');
let file=null;
function pick(f){ if(!f)return; file=f; name.textContent=f.name; go.disabled=false; msg.textContent='';
 const ext=f.name.split('.').pop().toLowerCase();
 for(const o of to.options){ o.disabled=(o.value===ext); }
 if(to.value===ext){ to.value=(ext==='pdf')?'docx':'pdf'; } }
input.onchange=()=>pick(input.files[0]);
['dragenter','dragover'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',ev=>pick(ev.dataTransfer.files[0]));
go.onclick=async()=>{ if(!file)return; go.disabled=true; msg.className=''; msg.textContent='Converting…';
 const fd=new FormData(); fd.append('file',file); fd.append('to',to.value);
 try{ const r=await fetch('/convert',{method:'POST',body:fd});
  if(!r.ok){ msg.className='err'; msg.textContent=await r.text(); go.disabled=false; return; }
  const blob=await r.blob(); const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=file.name.replace(/\\.[^.]+$/,'')+'.'+to.value; a.click();
  msg.textContent='Done. Saved to your Downloads as '+a.download; }
 catch(e){ msg.className='err'; msg.textContent='Failed: '+e; }
 go.disabled=false; };
</script></body></html>"""

PAGE = (
    PAGE.replace("FORMATS", ", ".join(SUPPORTED))
    .replace("ACCEPT", ",".join("." + f for f in SUPPORTED))
    .replace("OPTIONS", "".join(f'<option value="{f}">{f}</option>' for f in SUPPORTED))
)

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "md": "text/markdown",
    "txt": "text/plain",
    "html": "text/html",
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # keep the terminal quiet
        pass

    def _send(self, code: int, body: bytes, ctype: str, extra: dict | None = None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/":
            self._send(200, PAGE.encode(), "text/html; charset=utf-8")
        else:
            self._send(404, b"not found", "text/plain")

    def do_POST(self):
        if self.path != "/convert":
            return self._send(404, b"not found", "text/plain")
        length = int(self.headers.get("Content-Length", 0))
        raw = b"Content-Type: " + self.headers["Content-Type"].encode() + b"\r\n\r\n" + self.rfile.read(length)
        form = BytesParser(policy=email_policy).parsebytes(raw)
        upload = target = None
        for part in form.iter_parts():
            if part.get_param("name", header="content-disposition") == "file":
                upload = part
            elif part.get_param("name", header="content-disposition") == "to":
                target = part.get_content().strip()
        if upload is None or not target or target not in SUPPORTED:
            return self._send(400, b"missing file or target format", "text/plain")
        filename = Path(upload.get_filename() or "upload").name
        safe = re.sub(r"[^\w.\- ]", "_", filename) or "upload"
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / safe
            src.write_bytes(upload.get_payload(decode=True))
            dst = Path(tmp) / (src.stem + "." + target)
            try:
                convert(src, dst)
            except Exception as exc:  # noqa: BLE001
                return self._send(400, str(exc).encode(), "text/plain; charset=utf-8")
            data = dst.read_bytes()
        self._send(200, data, MIME[target], {"Content-Disposition": f'attachment; filename="{dst.name}"'})


def main(argv: list[str] | None = None) -> int:
    import argparse

    ap = argparse.ArgumentParser(prog="docconvert-ui", description="Open the docconvert web UI.")
    ap.add_argument("-p", "--port", type=int, default=0, help="port (default: random free port)")
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args(argv)

    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"docconvert UI running at {url}  (Ctrl+C to stop)")
    if not args.no_browser:
        threading.Timer(0.3, webbrowser.open, [url]).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
