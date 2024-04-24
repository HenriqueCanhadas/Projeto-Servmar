import pandas as pd

# Ler o arquivo Excel
caminho_excel = r'P:\Vitor\Programacao\teste_Organizar\Resultado_Final_Com_Abas - Copia.xlsx'
excel = pd.ExcelFile(caminho_excel)

# Criar um dicionário para armazenar os DataFrames de cada aba
data_frame_final = {}

#Loop para verificar cada aba no excel
for sheet_name in excel.sheet_names:
    data_frame = pd.read_excel(caminho_excel, sheet_name)
    
    #Pega os valores unicos da coluna 'SAMPLENAME' e 'ANALYTE'
    lista_pm = data_frame['SAMPLENAME'].unique()
    lista_analyte = data_frame['ANALYTE'].unique()

    # Criar um DataFrame vazio com colunas baseadas nos valores únicos da lista_pm
    data_frame_tabelado = pd.DataFrame(index=range(len(lista_pm) + 2), columns=lista_pm)
    
    # Adicionar três colunas Parametro, Cas, Unidade no início do DataFrame
    data_frame_tabelado.insert(0, 'Parametro', '')  # Coluna A vazia
    data_frame_tabelado.insert(1, 'CAS', '')  # Coluna B vazia
    data_frame_tabelado.insert(2, 'Unidade', '')  # Coluna C vazia
    
    # Dicionário para armazenar a correspondência na lista_analyte
    correspondencia_unidades = {}
    correspondencia_cas = {}

    # Iterar sobre os valores de lista_pm
    for coluna in lista_pm:
        # Adicionar valores de lista_analyte, lista_cas e lista_unidades nas colunas A, B e C a partir da terceira linha
        for linha, value in enumerate(lista_analyte):
            # Associar o valor de ANALYTE correspondente a este ANALYTE
            data_frame_tabelado.at[linha + 1, 'Parametro'] = value
            # Associar o valor de UNITS correspondente a este ANALYTE
            data_frame_tabelado.at[linha + 1, 'Unidade'] = correspondencia_unidades.get(value, '')
            # Associar o valor de CASNUMBER correspondente a este ANALYTE
            data_frame_tabelado.at[linha + 1, 'CAS'] = correspondencia_cas.get(value, '')

        # Calcular a quantidade total de valores únicos em ANALYTE para a amostra atual (coluna)
        valores_totais = data_frame.loc[data_frame['SAMPLENAME'] == coluna, 'ANALYTE'].nunique()
        
        # Iterar sobre os valores únicos de ANALYTE até encontrar um valor diferente
        for linha in range(valores_totais):
            # Verificar se já associamos este ANALYTE a um valor de UNITS
            if data_frame['ANALYTE'].unique()[linha] not in correspondencia_unidades:
                # Pegar o valor correspondente em 'UNITS' para o ANALYTE atual
                valor_unidades = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == data_frame['ANALYTE'].unique()[linha]), 'UNITS'].iloc[0]
                # Armazenar a correspondência no dicionário correspondencia_unidades
                correspondencia_unidades[data_frame['ANALYTE'].unique()[linha]] = valor_unidades
        
            # Verificar se já associamos este ANALYTE a um valor de CAS
            if data_frame['ANALYTE'].unique()[linha] not in correspondencia_cas:
                # Pegar o valor correspondente em 'CASNUMBER' para o ANALYTE atual
                valor_cas = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == data_frame['ANALYTE'].unique()[linha]), 'CASNUMBER_x'].iloc[0]
                # Armazenar a correspondência no dicionário correspondencia_cas
                correspondencia_cas[data_frame['ANALYTE'].unique()[linha]] = valor_cas

        # Verificar se coluna existe em SAMPLENAME para inserir os valores de SAMPDATE
        if coluna in data_frame['SAMPLENAME'].values:
            # Pegar o valor da linha correspondente em 'SAMPDATE'
            data_correspondente = data_frame.loc[data_frame['SAMPLENAME'] == coluna, 'SAMPDATE'].iloc[0]
            # Inserir o valor de SAMPDATE no novo DataFrame
            data_frame_tabelado.at[0, coluna] = data_correspondente
            
    # Iterar sobre as amostras (colunas) em lista_pm
    for coluna in lista_pm:
        # Verificar se a amostra (coluna) existe na coluna 'SAMPLENAME' do DataFrame original
        if coluna in data_frame['SAMPLENAME'].values:
            # Iterar sobre os valores únicos em 'ANALYTE'
            for item, valor in enumerate(lista_analyte):
                # Verificar se o valor em 'ANALYTE' existe na coluna 'Parametro' do DataFrame criado
                if valor in data_frame_tabelado['Parametro'].values:
                    # Encontrar a linha correspondente ao valor em 'Parametro'
                    linha_parametro = data_frame_tabelado.index[data_frame_tabelado['Parametro'] == valor][0]
                    # Filtrar o DataFrame original para obter o valor correspondente em 'Result'
                    resultado_correspondente = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == valor), 'Result']
                    # Verificar se há pelo menos um valor antes de tentar acessar o índice
                    if not resultado_correspondente.empty:
                        # Pegar o primeiro valor
                        resultado_correspondente = resultado_correspondente.iloc[0]
                        # Atribuir o valor ao local apropriado no DataFrame criado
                        data_frame_tabelado.at[linha_parametro, coluna] = resultado_correspondente
                    else:
                        # Se não houver valor correspondente, atribuir "n.a" (not applicable)
                        data_frame_tabelado.at[linha_parametro, coluna] = "n.a"
    
    # Adicionar o novo DataFrame ao dicionário
    data_frame_final[sheet_name] = data_frame_tabelado

# Salvar o resultado em um novo arquivo Excel
with pd.ExcelWriter(r'P:\Vitor\Programacao\teste_Organizar\Resultado_Final_Organizado.xlsx') as writer:
    for sheet_name, df in data_frame_final.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
