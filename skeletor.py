import pandas as pd

try:
    print("📂 Lendo o arquivo 'skeletor.xlsx'...")
    df = pd.read_excel('skeletor.xlsx')

    # Renomeia a coluna de preço recomendado para um formato mais limpo (sem quebra de linha)
    if 'RSP \nRecomendado' in df.columns:
        df = df.rename(columns={'RSP \nRecomendado': 'RSP_RECOMENDADO'})
    elif 'RSP Recomenda' in df.columns:
        df = df.rename(columns={df.columns[12]: 'RSP_RECOMENDADO'})

    # Seleciona apenas as 4 colunas solicitadas
    colunas_desejadas = ['CODIGO_MINASSAL', 'EAN', 'PRODUTO', 'RSP_RECOMENDADO']
    
    # Garante que apenas colunas existentes sejam pegas
    colunas_presentes = [c for c in colunas_desejadas if c in df.columns]
    df_filtrado = df[colunas_presentes]

    # Salva o resultado em uma nova planilha limpa
    nome_saida = 'skeletor_resumido.xlsx'
    df_filtrado.to_excel(nome_saida, index=False)

    print(f"\n🎉 Sucesso! Planilha resumida gerada com o nome '{nome_saida}'.")
    print(f"Total de registros processados: {len(df_filtrado)}")

except Exception as e:
    print(f"\n❌ ERRO ENCONTRADO: {e}")

finally:
    print("\n" + "="*50)
    input("Pressione ENTER para encerrar e fechar a janela...")