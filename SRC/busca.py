import pandas as pd
import os

def iniciar_busca():
    # Acha o caminho do banco mestre
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    caminho_banco = os.path.join(base_dir, "DATA", "output", "Banco_Mestre_Brasul.xlsx")

    # Avisa se o banco não existir
    if not os.path.exists(caminho_banco):
        print("\n❌ BANCO NÃO LOCALIZADO! Rode o 'main.py' primeiro.")
        return

    print("\n📂 Carregando dados da Brasul... segura aí.")
    df = pd.read_excel(caminho_banco).fillna('')

    # Loop infinito pra ficar buscando sem parar
    while True:
        print("\n" + "=" * 90)
        item_alvo = input("🔎 O que você quer achar? (Material ou Código) ou digite 'sair': ").strip().upper()

        if item_alvo.lower() == 'sair': break
        if not item_alvo: continue

        # Procura o que foi digitado no banco
        resultado = df[df['Descricao'].str.contains(item_alvo, na=False) |
                       df['Codigo'].astype(str).str.contains(item_alvo, na=False)]

        if not resultado.empty:
            print(f"\n🎯 Achei {len(resultado)} itens:")
            # Mostra as colunas principais no terminal
            cols_resumo = ['Obra', 'Codigo', 'Descricao', 'UN', 'Qtd_Acum', 'Val_Acum']
            print(resultado[cols_resumo].to_string(index=False))

            # Pergunta se quer salvar esse pedaço em Excel
            salvar = input("\n📄 Quer um Excel só dessa busca? (s/n): ")
            if salvar.lower() == 's':
                nome_arq = f"Relatorio_{item_alvo.replace(' ', '_')}.xlsx"
                resultado.to_excel(os.path.join(base_dir, "DATA", "output", nome_arq), index=False)
                print(f"✔️ Relatório salvo: {nome_arq}")
        else:
            print(f"⚠️ Não achei o item '{item_alvo}'.")

if __name__ == "__main__":
    iniciar_busca()