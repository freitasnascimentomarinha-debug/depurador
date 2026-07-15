import tarfile
import io
import re
import os
from pathlib import Path
import process_db

TGZ = Path('arquivos para teste') / 'teste-171.tgz'
DB_PATH = 'processos_emails.db'

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


def extract_text(conteudo: bytes, nome: str) -> str:
    lower = nome.lower()
    texto = ''
    try:
        if lower.endswith('.pdf'):
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
                    if pdf.pages:
                        first = pdf.pages[0].extract_text() or ''
                        last = pdf.pages[-1].extract_text() or ''
                        texto = (first + '\n' + last)[:8000]
            except Exception:
                texto = ''
        elif lower.endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')):
            try:
                from PIL import Image
                import pytesseract
                img = Image.open(io.BytesIO(conteudo))
                texto = pytesseract.image_to_string(img)
            except Exception:
                texto = ''
        else:
            try:
                texto = conteudo.decode('utf-8', errors='ignore')[:8000]
            except Exception:
                texto = ''
    except Exception:
        texto = ''
    return texto


def find_matched_member(nome_arquivo: str, members: list):
    n = (nome_arquivo or '').lower()
    for m in members:
        if n == m.name.lower() or n.endswith(m.name.lower()) or m.name.lower().endswith(n):
            return m
    # substring match
    for m in members:
        if m.name.lower().find(n) >= 0 or n.find(m.name.lower()) >= 0:
            return m
    return None


if not TGZ.exists():
    print('Arquivo TGZ não encontrado:', TGZ)
    raise SystemExit(1)

conn = process_db.get_connection(DB_PATH)
processos = process_db.listar_processos(conn)
if not processos:
    print('Nenhum processo cadastrado no DB.')
    raise SystemExit(0)

with tarfile.open(TGZ, mode='r:gz') as tf:
    members = [m for m in tf.getmembers() if m.isfile()]

    proposed = []
    for proc in processos:
        pid = proc['id']
        orcs = process_db.listar_orcamentos_do_processo(conn, pid)
        if not orcs:
            continue
        for o in orcs:
            nome_arquivo = (o.get('nome_arquivo') or '').strip()
            if not nome_arquivo:
                continue
            m = find_matched_member(nome_arquivo, members)
            if not m:
                continue
            f = tf.extractfile(m)
            conteudo = f.read() if f else b''
            texto = extract_text(conteudo, m.name)
            cnpj = process_db._extrair_cnpj_do_texto(texto)
            tel = process_db._extrair_telefone_do_texto(texto)
            valid_cnpj = cnpj and validar_cnpj(cnpj)
            valid_tel = tel and validar_telefone(tel)
            # procura fornecedores no processo que casam com nome_arquivo
            fornecedores = process_db.listar_fornecedores_processo(conn, pid)
            for forn in fornecedores:
                nome_f = forn.get('nome') or ''
                # reusar mesma lógica de casamento simples: tokens
                def norm(v: str) -> str:
                    b = unicodedata.normalize('NFKD', (v or '').lower())
                    b = ''.join(ch for ch in b if not unicodedata.combining(ch))
                    b = re.sub(r"[^a-z0-9]+", " ", b)
                    return " ".join(b.split())
                stop = {"ltda","eireli","me","epp","sa","s","a","comercio","servicos","de","da","do","das","dos","empresa","sociedade","limitada"}
                def tokens(v):
                    return {t for t in norm(v).split() if len(t)>=3 and t not in stop}
                tok_ref = tokens(nome_arquivo)
                tok_f = tokens(nome_f)
                inter = len(tok_ref.intersection(tok_f)) if tok_ref and tok_f else 0
                ratio_ref = inter / max(1, len(tok_ref)) if tok_ref else 0
                ratio_f = inter / max(1, len(tok_f)) if tok_f else 0
                casou = (norm(nome_arquivo) in norm(nome_f)) or (norm(nome_f) in norm(nome_arquivo)) or ratio_ref>=0.6 or ratio_f>=0.6 or inter>=2
                if not casou:
                    continue
                # current values
                cur_cnpj = (forn.get('cnpj') or '').strip()
                cur_tel = (forn.get('telefone') or '').strip()
                props = {}
                reason = ''
                if valid_cnpj and not cur_cnpj:
                    props['cnpj'] = cnpj
                    reason += 'heuristica_cnpj '
                if valid_tel and not cur_tel:
                    props['telefone'] = tel
                    reason += 'heuristica_tel '
                if props:
                    proposed.append({
                        'processo_id': pid,
                        'processo_numero': proc.get('numero'),
                        'fornecedor_email': forn.get('email'),
                        'fornecedor_nome': forn.get('nome'),
                        'file': m.name,
                        'proposed': props,
                        'reason': reason.strip(),
                    })

    if not proposed:
        print('Nenhuma atualização proposta pelos heurísticos para os arquivos do TGZ.')
    else:
        print('Atualizações propostas:')
        for p in proposed:
            print('\nProcesso:', p['processo_numero'], 'Fornecedor:', p['fornecedor_email'])
            print(' Arquivo:', p['file'])
            print(' Propostas:', p['proposed'], 'Motivo:', p['reason'])

print('\nDry-run concluído')
