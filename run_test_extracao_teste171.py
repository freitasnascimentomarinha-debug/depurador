import tarfile
import io
import re
import sys
from pathlib import Path

TGZ_PATH = Path('arquivos para teste') / 'teste-171.tgz'

CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
PHONE_RE = re.compile(r"(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?\d{4,5}[\s.-]?\d{4}")

VALID_DDDS = {
    '11','12','13','14','15','16','17','18','19',
    '21','22','24',
    '27','28',
    '31','32','33','34','35','37','38',
    '41','42','43','44','45','46',
    '47','48','49',
    '51','53','54','55',
    '61','62','64','63','65','66','67',
    '68','69','71','73','74','75','77','79','81','82','83','84','85','86','87','88','89','91','92','93','94','95','96','97','98','99'
}


def validar_cnpj(cnpj: str) -> bool:
    if not cnpj:
        return False
    dig = re.sub(r"\D", "", cnpj)
    if len(dig) != 14:
        return False
    if dig == dig[0] * 14:
        return False
    def calc(digs, mults):
        s = sum(int(a) * b for a, b in zip(digs, mults))
        r = s % 11
        return 0 if r < 2 else 11 - r
    mult1 = [5,4,3,2,9,8,7,6,5,4,3,2]
    mult2 = [6] + mult1
    d1 = calc(dig[:12], mult1)
    d2 = calc(dig[:12] + str(d1), mult2)
    return dig.endswith(f"{d1}{d2}")


def validar_telefone(tel: str) -> bool:
    if not tel:
        return False
    dig = re.sub(r"\D", "", tel)
    if dig.startswith('55'):
        dig = dig[2:]
    if len(dig) not in (10,11):
        return False
    ddd = dig[:2]
    return ddd in VALID_DDDS


def extract_text_from_pdf_bytes(bts: bytes):
    try:
        import pdfplumber
    except Exception as e:
        return None, f'pdfplumber missing: {e}'
    try:
        with pdfplumber.open(io.BytesIO(bts)) as pdf:
            first = pdf.pages[0].extract_text() or '' if pdf.pages else ''
            last = pdf.pages[-1].extract_text() or '' if pdf.pages else ''
            return (first, last), None
    except Exception as e:
        return None, str(e)


def ocr_image_bytes(bts: bytes):
    try:
        from PIL import Image
        import pytesseract
    except Exception as e:
        return None, f'OCR libs missing: {e}'
    try:
        img = Image.open(io.BytesIO(bts))
        txt = pytesseract.image_to_string(img)
        return txt, None
    except Exception as e:
        return None, str(e)


if not TGZ_PATH.exists():
    print('Arquivo não encontrado:', TGZ_PATH)
    sys.exit(1)

print('Abrindo', TGZ_PATH)
with tarfile.open(TGZ_PATH, mode='r:gz') as tf:
    members = [m for m in tf.getmembers() if m.isfile()]
    print('Arquivos dentro do tgz:', len(members))
    summary = []
    for m in members:
        name = m.name
        lower = name.lower()
        f = tf.extractfile(m)
        b = f.read() if f else b''
        print('\n----', name)
        if lower.endswith('.pdf'):
            (pages, err) = extract_text_from_pdf_bytes(b)
            if err:
                print('PDF error:', err)
                continue
            first, last = pages
            print('>>> trechos extraidos: first len', len(first), 'last len', len(last))
            cnpj_found = CNPJ_RE.findall(first) + CNPJ_RE.findall(last)
            phone_found = PHONE_RE.findall(first) + PHONE_RE.findall(last)
            cnpj_found = list(dict.fromkeys(cnpj_found))
            phone_found = list(dict.fromkeys(phone_found))
            print('CNPJ candidatos:', cnpj_found)
            for c in cnpj_found:
                print('  ', c, 'valid?', validar_cnpj(c))
            print('Telefone candidatos (raw):', phone_found)
            for t in phone_found:
                print('  ', t, 'valid?', validar_telefone(t))
            summary.append((name, cnpj_found, phone_found))
        elif lower.endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')):
            txt, err = ocr_image_bytes(b)
            if err:
                print('OCR error:', err)
                continue
            cnpj_found = CNPJ_RE.findall(txt)
            phone_found = PHONE_RE.findall(txt)
            print('CNPJ candidatos:', cnpj_found)
            for c in cnpj_found:
                print('  ', c, 'valid?', validar_cnpj(c))
            print('Telefone candidatos (raw):', phone_found)
            for t in phone_found:
                print('  ', t, 'valid?', validar_telefone(t))
            summary.append((name, cnpj_found, phone_found))
        else:
            # tentativa de texto
            try:
                txt = b.decode('utf-8', errors='ignore')
            except Exception:
                txt = ''
            cnpj_found = CNPJ_RE.findall(txt)
            phone_found = PHONE_RE.findall(txt)
            print('CNPJ candidatos:', cnpj_found)
            for c in cnpj_found:
                print('  ', c, 'valid?', validar_cnpj(c))
            print('Telefone candidatos (raw):', phone_found)
            for t in phone_found:
                print('  ', t, 'valid?', validar_telefone(t))
            summary.append((name, cnpj_found, phone_found))

print('\n=== Resumo ===')
for name, cs, ps in summary:
    print(name, '->', len(cs), 'CNPJ(s),', len(ps), 'telefone(s)')

print('\nTeste concluído')
