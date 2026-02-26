import pandas as pd
import os

# Função que faz a busca direto pelo terminal (sem janela)
def iniciar_busca():
    # Acha o caminho subindo uma pasta pra chegar nos dados
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    caminho_banco = os.path.join(base_dir, "DATA", "output", "Banco_Mestre_Brasul.xlsx")

    if not os.path.exists(caminho_banco):
        print("\n❌ BANCO NÃO LOCALIZADO! Rode o 'main.py' v11 primeiro.")
        return

    print("\n📂 Carregando base de dados Brasul v11...")
    df = pd.read_excel(caminho_banco).fillna('')

    # Fica perguntando o que quer buscar até você digitar 'sair'
    while True:
        print("\n" + "=" * 90)
        item_alvo = input("🔎 Digite MATERIAL ou CÓDIGO (ou 'sair'): ").strip().upper()

        if item_alvo.lower() == 'sair': break
        if not item_alvo: continue

        # Procura por descrição ou pelo código do item
        resultado = df[df['Descricao'].str.contains(item_alvo, na=False) |
                       df['Codigo'].astype(str).str.contains(item_alvo, na=False)]

        if not resultado.empty:
            print(f"\n🎯 Encontrado(s) {len(resultado)} item(ns):")
            # Mostra só as 6 colunas principais pra não bagunçar o terminal
            cols_resumo = ['Obra', 'Codigo', 'Descricao', 'UN', 'Qtd_Acum', 'Val_Acum']
            print(resultado[cols_resumo].to_string(index=False))

            # Pergunta se quer salvar esse resultado num Excel
            salvar = input("\n📄 Gerar Excel desta busca? (s/n): ")
            if salvar.lower() == 's':
                nome_arq = f"Relatorio_{item_alvo.replace(' ', '_')}.xlsx"
                resultado.to_excel(os.path.join(base_dir, "DATA", "output", nome_arq), index=False)
                print(f"✔️ Relatório completo salvo como: {nome_arq}")
        else:
            print(f"⚠️ Item '{item_alvo}' não localizado.")

if __name__ == "__main__":
    iniciar_busca()