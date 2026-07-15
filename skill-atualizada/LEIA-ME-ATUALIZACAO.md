# Como atualizar a skill pesquisa-precos-navais (v2)

O ambiente de execução desta sessão está sem espaço em disco, então o pacote
`.skill` (zip) não pôde ser gerado automaticamente. Os arquivos atualizados
estão prontos nesta pasta — a atualização leva 2 minutos:

## Passo a passo

1. Localize o arquivo original `pesquisa-precos-navais.skill` (o que você tem/enviou).
   Ele é um ZIP comum — renomeie uma cópia para `.zip` e extraia.
2. Substitua no conteúdo extraído:
   - `SKILL.md`  → pelo `pesquisa-precos-navais/SKILL.md` desta pasta
   - `references/metodologia.md` → pelo desta pasta
3. Adicione o arquivo NOVO:
   - `references/correcoes_matching.md` (memória de correções)
4. NÃO mexa em `scripts/` nem nos demais `references/` — continuam válidos.
5. Compacte a pasta de volta em ZIP (com o SKILL.md na raiz do zip) e renomeie
   a extensão para `.skill`.
6. No Claude: **Settings → Capabilities → Skills**, remova a versão antiga e
   instale o novo `.skill`.

## O que mudou na v2 (resumo)

| Mudança | Onde |
|---|---|
| Matching por código PI/NSN como critério nº 1, com normalização, herança de código e trava "códigos diferentes nunca casam" | SKILL.md passo 5, metodologia §6 |
| Memória de correções consultada antes do matching e alimentada a cada correção do usuário | novo references/correcoes_matching.md |
| Regras explícitas para a zona cinzenta (60–84): UF incompatível bloqueia, mesmo fornecedor bloqueia, na dúvida não casar | SKILL.md passo 5.4 |
| Pré-passagem de hashes de template ANTES da classificação (independe da ordem dos e-mails) | SKILL.md passo 2, metodologia §2.5 |
| Cascata de extração ganhou 6ª camada: leitura visual direta da página quando o OCR falha | SKILL.md passo 4, metodologia §5 |
| Campo `codigo` obrigatório no quotes.json e coluna "Código (PI/NSN)" no mapa | SKILL.md passos 4 e 6 |
| Autoverificação numérica obrigatória do mapa antes do relatório (6 checagens) | SKILL.md passo 7, metodologia §7 |
| Princípio de eficiência: scripts determinísticos primeiro, julgamento só no ambíguo | SKILL.md, início |
