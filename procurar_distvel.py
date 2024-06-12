import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, Border, Side

def organizar(caminho_arquivo2, global_workbook):
    # Carregar dados do Excel
    df = pd.read_excel(caminho_arquivo2, header=6, sheet_name = "Bin Measure")
    # Converta os valores na coluna 'DAY' para strings
    df['DAY'] = df['DAY'].astype(str)
    # Remova o ponto decimal e zeros à direita ('.0') dos valores na coluna 'DAY'
    df['DAY'] = df['DAY'].str.replace('.0', '')
    # Converta os valores na coluna 'DAY' para inteiros, ignorando os valores NaN
    df['DAY'] = pd.to_numeric(df['DAY'], errors='coerce')
    # Coluna que contém as informações sobre o evento
    coluna_evento = 'Measure'
    # Evento para Distance Traveled
    evento_distance_traveled = 'Mouse 1 Center Distance Traveled (apart 1.000000 second)'
    # Evento para Velocity Average
    evento_velocity_average = 'Mouse 1 Center Velocity Average (apart 1.000000 second)'
    # Colunas desejadas no DataFrame filtrado
    colunas_desejadas = ['DAY', 'ANIMAL', 'Bin1']

    # Cria uma nova planilha no arquivo global para os dados 'TR'
    ws_tr = global_workbook.create_sheet(title='TR')
    # Cria uma nova planilha no arquivo global para os dados 'TT'
    ws_tt = global_workbook.create_sheet(title='TT')

    # Inicializa um DataFrame para armazenar os dados do TR
    df_tr = pd.DataFrame(columns=colunas_desejadas)
    # Inicializa um DataFrame para armazenar os dados do TT
    df_tt = pd.DataFrame(columns=colunas_desejadas)

    #(df['Video Name'].str.contains('TR')) foi retirado, pois o teste pode ter nome de TR
    # Filtra os dados TR com base no evento de distância
    dados_distance_tr = df[(df[coluna_evento] == evento_distance_traveled) & (df['DAY'] == 1)]
    # Filtra os dados TR com base no evento de velocidade
    dados_velocity_tr = df[(df[coluna_evento] == evento_velocity_average) & (df['DAY'] == 1)]

    # Merge dos dados de distância e velocidade do TR com base em 'DAY' e 'ANIMAL'
    df_tr = pd.merge(dados_distance_tr, dados_velocity_tr, on=['DAY', 'ANIMAL'], suffixes=('_dist', '_vel'))

    # Reorganiza as colunas do DataFrame do TR
    df_tr = df_tr.reindex(columns=['DAY', 'ANIMAL', 'Bin1_dist', 'Bin1_vel'])

    # Adiciona uma coluna em branco ao lado de 'Bin1_dist' no DataFrame do TR
    df_tr.insert(df_tr.columns.get_loc('Bin1_dist') + 1, 'Bin1_dist_blank', '')

    # Preenche a coluna em branco do TR com os dados de 'Bin1_dist' divididos por 1000
    df_tr['Bin1_dist_blank'] = df_tr['Bin1_dist'] / 1000

    # Adiciona o DataFrame do TR ao arquivo do Excel
    for row in dataframe_to_rows(df_tr, index=False, header=True):
        ws_tr.append(row)

    #(df['Video Name'].str.contains('TT')) foi retirado
    # Filtra os dados TT com base no evento de distância
    dados_distance_tt = df[(df[coluna_evento] == evento_distance_traveled) & (df['DAY'] == 2)]
    # Filtra os dados TT com base no evento de velocidade
    dados_velocity_tt = df[(df[coluna_evento] == evento_velocity_average) & (df['DAY'] == 2)]

    # Merge dos dados de distância e velocidade do TT com base em 'DAY' e 'ANIMAL'
    df_tt = pd.merge(dados_distance_tt, dados_velocity_tt, on=['DAY', 'ANIMAL'], suffixes=('_dist', '_vel'))

    # Reorganiza as colunas do DataFrame do TT
    df_tt = df_tt.reindex(columns=['DAY', 'ANIMAL', 'Bin1_dist', 'Bin1_vel'])

    # Adiciona uma coluna em branco ao lado de 'Bin1_dist' no DataFrame do TT
    df_tt.insert(df_tt.columns.get_loc('Bin1_dist') + 1, 'Bin1_dist_blank', '')

    # Preenche a coluna em branco do TT com os dados de 'Bin1_dist' divididos por 1000
    df_tt['Bin1_dist_blank'] = df_tt['Bin1_dist'] / 1000

    # Adiciona o DataFrame do TT ao arquivo do Excel
    for row in dataframe_to_rows(df_tt, index=False, header=True):
        ws_tt.append(row)
    
    # Definir tamanhos
    colunas_menores = ['C', 'E']
    coluna_maior = ['D']

    # Define o tamanho desejado para as colunas separadas
    for coluna in colunas_menores:
        ws_tr.column_dimensions[coluna].width = 14
        ws_tt.column_dimensions[coluna].width = 14

    for coluna in coluna_maior:
        ws_tr.column_dimensions[coluna].width = 18
        ws_tt.column_dimensions[coluna].width = 18

    # Renomeia as colunas
    novo_nome_colunas = {
        'DAY': 'Dia',
        'ANIMAL': 'Animal',
        'Bin1_dist': 'Distância',
        'Bin1_dist_blank': 'Distância (metros)',
        'Bin1_vel': 'Velocidade',
    }

    # Itera sobre as células de cabeçalho para renomear as colunas
    for coluna_antiga, novo_nome in novo_nome_colunas.items():
        # Obtém o número da coluna antiga
        for cell in ws_tr[1]:
            if cell.value == coluna_antiga:
                numero_coluna_antiga = cell.column
                break
        # Define o novo nome da coluna
        ws_tr.cell(row=1, column=numero_coluna_antiga, value=novo_nome)

    # Itera sobre as células de cabeçalho para renomear as colunas
    for coluna_antiga, novo_nome in novo_nome_colunas.items():
        # Obtém o número da coluna antiga
        for cell in ws_tt[1]:
            if cell.value == coluna_antiga:
                numero_coluna_antiga = cell.column
                break
        # Define o novo nome da coluna
        ws_tt.cell(row=1, column=numero_coluna_antiga, value=novo_nome)

    # Aplica formatação nas planilhas 'TR' e 'TT'
    for ws in [ws_tr, ws_tt]:
        # Centraliza o conteúdo das células
        for row in ws.iter_rows(min_row=1, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        # Aplica formatação nos títulos das colunas
        for row in ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=ws.max_column):
            for cell in row:
                cell.font = Font(bold=True)
                cell.border = Border(left=Side(style='thin'),
                                     right=Side(style='thin'),
                                     top=Side(style='thin'),
                                     bottom=Side(style='thin'))

    return df_tr, df_tt  # Retorna os DataFrames do TR e TT
