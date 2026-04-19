"""Tests for the PDF to Word converter Flask application."""
import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app import app as flask_app


@pytest.fixture()
def client():
    flask_app.config['TESTING'] = True
    with flask_app.test_client() as c:
        yield c


# ---------------------------------------------------------------------------
# Helper – build a minimal valid single-page PDF in memory
# ---------------------------------------------------------------------------

def _minimal_pdf(text: str = "Hello PDF") -> bytes:
    """Return a minimal but valid PDF byte string."""
    # Build a tiny PDF manually (no external library needed for this fixture)
    content = (
        f"BT /F1 12 Tf 100 700 Td ({text}) Tj ET"
    ).encode()
    stream = b"stream\r\n" + content + b"\r\nendstream"
    objects = []
    offsets = []

    def add(obj_num: int, body: bytes):
        offsets.append(len(b"".join(objects)))
        objects.append(f"{obj_num} 0 obj\n".encode() + body + b"\nendobj\n")

    add(1, b"<< /Type /Catalog /Pages 2 0 R >>")
    add(2, b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    add(3, b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>")
    add(4, f"<< /Length {len(content)} >>\n".encode() + stream)
    add(5, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    body = b"".join(objects)
    xref_offset = len(b"%PDF-1.4\n") + len(body)

    xref_lines = [b"xref\n", f"0 {len(objects) + 1}\n".encode(),
                  b"0000000000 65535 f \n"]
    base = len(b"%PDF-1.4\n")
    for off in offsets:
        xref_lines.append(f"{base + off:010d} 00000 n \n".encode())

    xref = b"".join(xref_lines)
    trailer = (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
        f"startxref\n{len(b'%PDF-1.4\n') + len(body)}\n%%EOF\n"
    ).encode()

    return b"%PDF-1.4\n" + body + xref + trailer


# ---------------------------------------------------------------------------
# Route tests
# ---------------------------------------------------------------------------

class TestIndex:
    def test_get_returns_200(self, client):
        response = client.get('/')
        assert response.status_code == 200

    def test_get_returns_html(self, client):
        response = client.get('/')
        assert b'PDF to Word' in response.data


class TestConvertEndpoint:
    def test_no_file_returns_400(self, client):
        response = client.post('/convert', data={})
        assert response.status_code == 400
        assert b'error' in response.data.lower()

    def test_empty_filename_returns_400(self, client):
        data = {'file': (io.BytesIO(b''), '')}
        response = client.post('/convert', data=data,
                               content_type='multipart/form-data')
        assert response.status_code == 400

    def test_non_pdf_file_returns_400(self, client):
        data = {'file': (io.BytesIO(b'hello world'), 'test.txt')}
        response = client.post('/convert', data=data,
                               content_type='multipart/form-data')
        assert response.status_code == 400
        json_data = response.get_json()
        assert 'error' in json_data

    def test_valid_pdf_returns_docx(self, client):
        pdf_bytes = _minimal_pdf("Test conversion")
        data = {'file': (io.BytesIO(pdf_bytes), 'sample.pdf')}
        response = client.post('/convert', data=data,
                               content_type='multipart/form-data')
        # pdf2docx may warn but should succeed
        assert response.status_code == 200
        assert response.content_type == (
            'application/vnd.openxmlformats-officedocument'
            '.wordprocessingml.document'
        )
        # Check the response contains data
        assert len(response.data) > 0

    def test_valid_pdf_download_name(self, client):
        pdf_bytes = _minimal_pdf()
        data = {'file': (io.BytesIO(pdf_bytes), 'my_document.pdf')}
        response = client.post('/convert', data=data,
                               content_type='multipart/form-data')
        assert response.status_code == 200
        cd = response.headers.get('Content-Disposition', '')
        assert 'my_document.docx' in cd
