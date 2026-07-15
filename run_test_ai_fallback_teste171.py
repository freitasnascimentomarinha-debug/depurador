import tarfile
import io
import re
import sys
import os
from pathlib import Path

TGZ_PATH = Path('arquivos para teste') / 'teste-171.tgz'
OPENROUTER_KEY = os.environ.get('OPENROUTER_API_KEY') or os.environ.get('OPENROUTER_KEY') or os.environ.get('OPENAI_API_KEY')
MODEL_NAME = os.environ.get('TEST_MODEL') or 'google/gemini-2.5-flash'

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


def call_ai_extract(corpus_text: str, empresa: str, api_key_local: str, model_name: str):
    if not api_key_local:
        return '', ''
    try:
        import requests
        url = 'https://api.openrouter.ai/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {api_key_local}',
            'Content-Type': 'application/json',
        }
        prompt = (
            f"Extraia somente o CNPJ e o telefone da empresa {empresa} a partir do texto fornecido. "
            "Responda estritamente em JSON com chaves 'cnpj' e 'telefone'. Se não encontrar, deixe valor vazio."
            "\n\nTexto:\n" + corpus_text[:6000]
        )
        payload = {
            'model': model_name,
            'messages': [
                {'role': 'system', 'content': 'Você é um assistente que extrai CNPJ e telefone brasileiros.'},
                {'role': 'user', 'content': prompt},
            ],
            'temperature': 0.0,
            'max_tokens': 200,
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        if resp.status_code != 200:
            return '', ''
        data = resp.json()
        text = ''
        try:
            text = data.get('choices', [])[0].get('message', {}).get('content', '')
        except Exception:
            text = data.get('choices', [])[0].get('text', '') if data.get('choices') else ''
        # tenta extrair com regex simples
        cnpj = ''
        tel = ''
        m_c = re.search(r'"cnpj"\s*:\s*"([^"]+)"', text)
        if m_c:
            cnpj = m_c.group(1)
        else:
            m_c2 = CNPJ_RE.search(text)
            if m_c2:
                cnpj = m_c2.group(0)
        m_t = re.search(r'"telefone"\s*:\s*"([^"]+)"', text)
        if m_t:
            tel = m_t.group(1)
        else:
            m_t2 = PHONE_RE.search(text)
            if m_t2:
                tel = m_t2.group(0)
        return cnpj, tel
    except Exception as e:
        return '', ''


if not TGZ_PATH.exists():
    print('Arquivo não encontrado:', TGZ_PATH)
    sys.exit(1)

if not OPENROUTER_KEY:
    print('Nenhuma chave de API encontrada nas variáveis de ambiente (OPENROUTER_API_KEY).')
    print('Defina OPENROUTER_API_KEY e reexecute para usar a IA como fallback.')

print('Usando modelo:', MODEL_NAME)
print('Abrindo', TGZ_PATH)
with tarfile.open(TGZ_PATH, mode='r:gz') as tf:
    members = [m for m in tf.getmembers() if m.isfile()]
    print('Arquivos dentro do tgz:', len(members))
    for m in members:
        name = m.name
        lower = name.lower()
        f = tf.extractfile(m)
        b = f.read() if f else b''
        print('\n----', name)
        texto = ''
        if lower.endswith('.pdf'):
            (pages, err) = extract_text_from_pdf_bytes(b)
            if pages:
                first, last = pages
                texto = (first or '') + '\n' + (last or '')
        elif lower.endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')):
            txt, err = ocr_image_bytes(b)
            if txt:
                texto = txt
        else:
            try:
                texto = b.decode('utf-8', errors='ignore')
            except Exception:
                texto = ''

        # chamada IA (se chave presente)
        ai_cnpj, ai_tel = '', ''
        if OPENROUTER_KEY:
            ai_cnpj, ai_tel = call_ai_extract(texto, name, OPENROUTER_KEY, MODEL_NAME)
        print('IA sugeriu => CNPJ:', ai_cnpj, 'Telefone:', ai_tel)
        print('Validação: CNPJ ok?', validar_cnpj(ai_cnpj), 'Telefone ok?', validar_telefone(ai_tel))

print('\nPassagem IA concluída')
