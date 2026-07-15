from streamlit.testing.v1 import AppTest
import sys
import os

try:
    os.environ["OPENROUTER_API_KEY"] = "dummy"
    print("Iniciando AppTest...")
    at = AppTest.from_file("app.py")
    at.run(timeout=120)
    
    # 1. Selecionar 'Pasta local no computador' do st.radio
    print("Atualizando st.radio para 'Pasta local no computador'...")
    # Seleciona 'Pasta local no computador'
    at.radio[0].set_value("Pasta local no computador").run(timeout=120)
    
    # 2. Preencher a pasta '/workspaces/depurador/arquivos para teste' no text_input
    print("Preenchendo caminho da pasta...")
    at.text_input[0].set_value("/workspaces/depurador/arquivos para teste").run(timeout=120)
    
    # 3. Acionar o botão '🚀 Processar orçamentos'
    print("Acionando o botão...")
    for i, b in enumerate(at.button):
         if "Processar orçamentos" in b.label:
             at.button[i].click().run(timeout=120)
             break
             
    # Capturar exceptions se houver algum erro
    if at.exception:
        print("EXCEÇÃO ENCONTRADA:")
        for exc in at.exception:
            print(exc.message)
            print(exc.stack_trace)
    else:
        print("Nenhuma exceção lançada pelo AppTest.")
        print("TEXTOS PRINCIPAIS RENDERIZADOS:")
        for element in at.main:
            if element.type in ["markdown", "text", "subheader", "header", "title", "caption", "code"]:
                print(f"[{element.type}] {element.value}")
except Exception as e:
    print("Erro durante a execução do script de teste:")
    import traceback
    traceback.print_exc()
