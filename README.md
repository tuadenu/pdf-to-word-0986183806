# Batch PDF to Word (Editable Text)

Ung dung Python chuyen doi hang loat PDF sang Word (`.docx`) voi muc tieu xuat ra van ban co the chinh sua.
Giao dien desktop da duoc toi uu theo kieu app thong thuong: co progress bar, bang trang thai tung file, nut mo thu muc ket qua, va nut dung khi dang chay.

## Tinh nang

- Chon thu muc dau vao chua nhieu file PDF.
- Chon thu muc dau ra cho file DOCX.
- Chuyen doi hang loat voi thanh tien do va ước lượng thời gian còn lại.
- Giao dien tieng Viet de su dung de hon.
- Uu tien trich xuat text co san trong PDF.
- OCR fallback cho PDF scan de xuat van ban (khong phai anh).
- Nut Dung de huy qua trinh sau khi xu ly xong tep hien tai.
- Tuy chon bo qua file DOCX da ton tai (skip existing).
- Hien thi thoi gian ước lượng còn lại trong qua trinh chuyen doi.

## Yeu cau

- Python 3.10+
- Tesseract OCR (chi can neu bat OCR fallback)

### Cai Tesseract tren macOS

```bash
brew install tesseract
brew install tesseract-lang
```

Sau do co the dung ngon ngu OCR:
- `eng` (English)
- `vie` (Vietnamese)
- `eng+vie` (ket hop)

## Cai dat

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Chay ung dung

```bash
python pdf_to_word_app.py
```

## Dong goi thanh app macOS (.app)

Neu ban muon mo nhu mot ung dung desktop binh thuong, co the dong goi bang PyInstaller:

```bash
pip install pyinstaller
pyinstaller --windowed --name PDF2WordApp pdf_to_word_app.py
```

Sau khi build xong, app nam trong thu muc:

```bash
dist/PDF2WordApp.app
```

## Ghi chu chat luong

- Neu PDF la file text goc, ket qua thuong tot hon va giu bo cuc tot hon.
- Neu PDF la scan, OCR se nhan dang ky tu de xuat ra van ban co the sua.
- Chat luong OCR phu thuoc do ro net cua file scan va ngon ngu OCR duoc chon.
