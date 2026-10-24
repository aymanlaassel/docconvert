import threading
import urllib.request
import uuid
from http.server import ThreadingHTTPServer

from docconvert.ui import Handler


def _multipart(fields: dict, file_name: str, file_bytes: bytes):
    b = uuid.uuid4().hex
    body = b""
    for k, v in fields.items():
        body += f"--{b}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode()
    body += (f"--{b}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{file_name}\"\r\n"
             "Content-Type: application/octet-stream\r\n\r\n").encode() + file_bytes + f"\r\n--{b}--\r\n".encode()
    return body, f"multipart/form-data; boundary={b}"


def test_ui_converts_md_to_docx():
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{server.server_address[1]}"
    try:
        assert b"docconvert" in urllib.request.urlopen(url + "/").read()
        body, ctype = _multipart({"to": "docx"}, "notes.md", b"# Hi\n\nHello there.\n")
        req = urllib.request.Request(url + "/convert", data=body, headers={"Content-Type": ctype})
        resp = urllib.request.urlopen(req)
        assert resp.status == 200
        assert 'filename="notes.docx"' in resp.headers["Content-Disposition"]
        assert resp.read()[:2] == b"PK"
    finally:
        server.shutdown()
