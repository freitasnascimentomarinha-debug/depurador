"""Teste dedicado de OCR para PDF escaneado.

Falha se:
- OCR nao estiver operacional no ambiente.
- O PDF de teste nao acionar tentativa de OCR.
- O texto final permanecer vazio apos OCR.
"""

from __future__ import annotations

import os
import sys

from extract_utils import _extract_text_pages_pdf, ocr_runtime_status

PDF_TESTE = "/workspaces/depurador/arquivos para teste/print site orçamento em pdf.pdf"


def main() -> int:
    ok_ocr, erro = ocr_runtime_status()
    if not ok_ocr:
        print(f"[ERRO] OCR indisponivel: {erro}")
        return 1

    if not os.path.exists(PDF_TESTE):
        print(f"[ERRO] Arquivo de teste nao encontrado: {PDF_TESTE}")
        return 1

    pages = _extract_text_pages_pdf(PDF_TESTE)
    if not pages:
        print("[ERRO] Nenhuma pagina lida do PDF de teste.")
        return 1

    tentou_ocr = sum(1 for p in pages if p.get("tentou_ocr"))
    aplicou_ocr = sum(1 for p in pages if p.get("veio_ocr"))
    texto_total = "\n".join((p.get("texto") or "") for p in pages).strip()

    print(f"Paginas: {len(pages)}")
    print(f"Tentativas de OCR: {tentou_ocr}")
    print(f"Paginas com OCR aplicado: {aplicou_ocr}")
    print(f"Tamanho texto final: {len(texto_total)}")

    if tentou_ocr == 0:
        print("[ERRO] OCR nao foi tentado no PDF-imagem.")
        return 1

    if not texto_total:
        print("[ERRO] OCR foi tentado, mas o texto final ficou vazio.")
        return 1

    print("[OK] OCR para PDF-imagem validado com sucesso.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
