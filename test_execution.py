import sqlite3
from email_classifier import _heuristica
import process_db

# 1. Test _heuristica
print("--- TEST 1: _heuristica ---")
parsed1 = {
    "assunto": "Dúvida sobre o prazo de entrega",
    "corpo": "Poderiam informar o prazo?",
    "remetente_email": "fornecedor@empresa.com"
}
res1 = _heuristica(parsed1)
print(f"Result 1: {res1}")
assert res1 == "duvida", f"Expected 'duvida', got {res1}"

parsed2 = {
    "assunto": "Dúvida sobre a cotação",
    "corpo": "Segue cotação?",
}
res2 = _heuristica(parsed2)
print(f"Result 2: {res2}")
assert res2 == "duvida", f"Expected 'duvida', got {res2}"

print("Test 1 OK")

# 2. Test sqlite behavior with atualizar_participacao
print("--- TEST 2: DB atualizar_participacao ---")
conn = sqlite3.connect(":memory:")
conn.row_factory = sqlite3.Row
process_db._init_db(conn)

processo_id = 1
fornecedor_id = 1

# Insert basic entities or directly insert the record to participacoes to set initial data_pedido_enviado
conn.execute("INSERT INTO processos (id, numero) VALUES (?, ?)", (processo_id, "Proc1"))
conn.execute("INSERT INTO fornecedores (id, nome) VALUES (?, ?)", (fornecedor_id, "Forn1"))
conn.execute(
    "INSERT INTO participacoes (processo_id, fornecedor_id, data_pedido_enviado) VALUES (?, ?, ?)",
    (processo_id, fornecedor_id, "2026-08-16T12:00:00")
)
conn.commit()

# Call atualizar_participacao
process_db.atualizar_participacao(
    conn=conn,
    processo_id=processo_id,
    fornecedor_id=fornecedor_id,
    tipo_email="orcamento_recebido",
    data_envio="2026-08-16T13:00:00",
    data_pedido_inferida="2026-08-15T10:00:00"
)

# Fetch results
row = conn.execute("SELECT * FROM participacoes WHERE processo_id = ? AND fornecedor_id = ?", (processo_id, fornecedor_id)).fetchone()
print(f"Final data_pedido_enviado: {row['data_pedido_enviado']}")
print(f"Final data_primeira_resposta: {row['data_primeira_resposta']}")
print(f"Final enviou_orcamento: {row['enviou_orcamento']}")

assert row["data_pedido_enviado"] == "2026-08-15T10:00:00", f"Expected '2026-08-15T10:00:00', got {row['data_pedido_enviado']}"
print("Test 2 OK")

print("All tests passed successfully!")
