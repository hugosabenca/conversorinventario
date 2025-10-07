import streamlit as st
import pandas as pd
import os
import glob
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import json
import tempfile
import io

# =====================================================================================
# FUNÇÃO PRINCIPAL PARA O FLUXO "BOBINA"
# Esta função agora encapsula a lógica exata dos seus dois scripts.
# =====================================================================================
def processar_fluxo_bobina(pasta_csv_origem, pasta_trabalho_temporaria):
    """
    Executa o processo completo de conversão para "Bobina", imitando os scripts originais.
    Etapa 1: Converte CSVs para Excels formatados.
    Etapa 2: Unifica os Excels formatados em um arquivo final.
    """
    
    # --- ETAPA 1: LÓGICA DO 'conversorcsvparaexcelBOBINAS.py' ---
    st.write("Etapa 1 de 2: Convertendo cada CSV para um Excel formatado...")
    
    arquivos_csv = glob.glob(os.path.join(pasta_csv_origem, '*.csv'))
    if not arquivos_csv:
        st.error("Nenhum arquivo .csv foi encontrado para processar.")
        return None

    for file_path in arquivos_csv:
        dados_processados = []
        try:
            try:
                df_csv = pd.read_csv(file_path, header=None, encoding='utf-8', low_memory=False)
            except UnicodeDecodeError:
                df_csv = pd.read_csv(file_path, header=None, encoding='latin1', low_memory=False)
            mask = df_csv.apply(lambda row: 'date' not in ' '.join(row.astype(str)).lower(), axis=1)
            df_csv = df_csv[mask]
        except Exception as e:
            st.warning(f"Não foi possível ler o arquivo {os.path.basename(file_path)}. Erro: {e}")
            continue

        for index, row in df_csv.iterrows():
            if len(row) < 5: continue
            data_leitura, hora_leitura, _, tipo_codigo, dados_lidos = row[0:5]
            dados_lidos, tipo_codigo = str(dados_lidos).strip(), str(tipo_codigo).strip()
            nova_linha = {"Data da Leitura": data_leitura, "Hora da Leitura": hora_leitura, "Lote": None, "Peso": None}
            
            # (Toda a sua lógica de parsing de códigos de barra é copiada aqui)
            if tipo_codigo == 'Code128':
                if ' ' in dados_lidos: nova_linha.update({"Lote": "erro de leitura", "Peso": "erro de leitura"})
                elif '*' in dados_lidos:
                    try:
                        partes = dados_lidos.split('*')
                        if dados_lidos.startswith('*'):
                            nova_linha.update({"Lote": partes[3].strip(), "Peso": float(partes[2].strip()) / 1000.0})
                        else:
                            nova_linha.update({"Lote": partes[2].strip(), "Peso": float(partes[1].strip()) / 1000.0})
                    except (ValueError, IndexError): nova_linha.update({"Lote": "erro Code128/*", "Peso": "erro Code128/*"})
                else:
                    if dados_lidos.isdigit() and len(dados_lidos) <= 5: nova_linha.update({"Peso": float(dados_lidos) / 1000.0, "Lote": None})
                    else: nova_linha.update({"Lote": dados_lidos, "Peso": None})
            elif tipo_codigo in ['CODE_39', 'CODE_128']:
                nova_linha["Data da Leitura"] = datetime.strptime(str(data_leitura), '%m-%d-%Y').strftime('%d/%m/%Y')
                nova_linha.update({"Lote": dados_lidos, "Peso": None})
            elif tipo_codigo in ['QR_CODE', 'QR']:
                nova_linha["Data da Leitura"] = datetime.strptime(str(data_leitura), '%m-%d-%Y').strftime('%d/%m/%Y')
                if '{' in dados_lidos and '}' in dados_lidos:
                    try:
                        partes = dados_lidos.split('{', 1)
                        identificador = partes[0].strip('"-')
                        dados_json = json.loads('{' + partes[1])
                        nova_linha.update({"Peso": float(dados_json.get('peso', 0)), "Lote": identificador})
                    except (ValueError, IndexError, json.JSONDecodeError): nova_linha.update({"Lote": "erro QR/JSON", "Peso": "erro QR/JSON"})
                else:
                    try:
                        partes = dados_lidos.split('-')
                        nova_linha.update({"Lote": partes[3].strip(), "Peso": float(partes[-1].strip()) / 1000.0})
                    except (ValueError, IndexError): nova_linha.update({"Lote": "erro QR/-", "Peso": "erro QR/-"})
            
            dados_processados.append(nova_linha)

        if not dados_processados: continue

        df_excel = pd.DataFrame(dados_processados)
        if 'Lote' in df_excel.columns: df_excel['Lote'] = df_excel['Lote'].fillna('').astype(str)
        
        output_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}.xlsx"
        output_path = os.path.join(pasta_trabalho_temporaria, output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_excel.to_excel(writer, index=False, sheet_name='Dados')
            worksheet = writer.sheets['Dados']
            for col_cells in worksheet.columns:
                max_len = max(len(str(cell.value)) for cell in col_cells if cell.value is not None)
                worksheet.column_dimensions[get_column_letter(col_cells[0].column)].width = max_len + 2
            
            # Aplica o formato de número na coluna 'Peso' (coluna D)
            col_peso_letra = 'D' 
            for cell in worksheet[col_peso_letra]:
                if cell.row > 1: cell.number_format = '0.000'
    
    st.success("Etapa 1 concluída!")

    # --- ETAPA 2: LÓGICA DO 'unificarBOBINAS.py' ---
    st.write("Etapa 2 de 2: Unificando todos os arquivos Excel em um relatório final...")
    
    arquivo_de_saida_final = os.path.join(pasta_trabalho_temporaria, "Inventario.xlsx")
    arquivos_excel_intermediarios = glob.glob(os.path.join(pasta_trabalho_temporaria, '*.xlsx'))
    
    if not arquivos_excel_intermediarios:
        st.error("Nenhum arquivo Excel intermediário foi gerado na Etapa 1.")
        return None
    
    lista_de_dataframes = []
    for arquivo in arquivos_excel_intermediarios:
        try:
            # Ponto CRÍTICO: ler a coluna 'Lote' como texto para não perder a formatação
            df = pd.read_excel(arquivo, dtype={'Lote': str})
            df['Localização'] = os.path.splitext(os.path.basename(arquivo))[0]
            lista_de_dataframes.append(df)
        except Exception as e:
            st.warning(f"Erro ao ler o arquivo intermediário {os.path.basename(arquivo)}. Erro: {e}")

    if not lista_de_dataframes:
        st.error("Nenhum dado foi lido dos arquivos intermediários. O arquivo final não será gerado.")
        return None

    df_final = pd.concat(lista_de_dataframes, ignore_index=True)

    with pd.ExcelWriter(arquivo_de_saida_final, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Inventario_Unificado')
        worksheet = writer.sheets['Inventario_Unificado']
        for col_cells in worksheet.columns:
            max_len = max(len(str(cell.value)) for cell in col_cells if cell.value is not None)
            worksheet.column_dimensions[get_column_letter(col_cells[0].column)].width = max_len + 2
    
    st.success("Etapa 2 concluída!")
    return arquivo_de_saida_final

# =====================================================================================
# INTERFACE DO STREAMLIT
# =====================================================================================

st.set_page_config(page_title="Conversor de Inventário", layout="wide")
st.title("Conversor de Inventário")

tipo_material = st.selectbox(
    "Tipo de Material:",
    ["Bobina", "Produto Acabado"]
)

grupo_produto = st.text_input("Grupo de Produto:")

uploaded_files = st.file_uploader(
    "Importar arquivos .csv",
    type="csv",
    accept_multiple_files=True
)

if st.button("Converter"):
    if not uploaded_files:
        st.warning("Por favor, carregue pelo menos um arquivo .csv para converter.")
    elif not grupo_produto:
        st.warning("Por favor, preencha o campo 'Grupo de Produto'.")
    else:
        # Foco no fluxo de Bobina
        if tipo_material == "Bobina":
            with st.spinner("Aguarde... A conversão está em andamento."):
                try:
                    # Usar um diretório temporário para gerenciar os arquivos intermediários e o final
                    with tempfile.TemporaryDirectory() as temp_dir:
                        pasta_csv_origem = os.path.join(temp_dir, "csv_original")
                        os.makedirs(pasta_csv_origem)

                        # Salvar arquivos carregados na pasta temporária
                        for uploaded_file in uploaded_files:
                            with open(os.path.join(pasta_csv_origem, uploaded_file.name), "wb") as f:
                                f.write(uploaded_file.getbuffer())

                        # Chamar a função principal do fluxo "Bobina"
                        caminho_arquivo_final = processar_fluxo_bobina(pasta_csv_origem, temp_dir)
                        
                        # --- DOWNLOAD DO ARQUIVO FINAL ---
                        if caminho_arquivo_final and os.path.exists(caminho_arquivo_final):
                            st.success("Conversão finalizada com sucesso!")
                            
                            with open(caminho_arquivo_final, "rb") as file:
                                output_bytes = file.read()
                            
                            nome_arquivo_download = f"Inventario {grupo_produto}.xlsx"
                            
                            st.download_button(
                                label="Clique aqui para baixar o arquivo final",
                                data=output_bytes,
                                file_name=nome_arquivo_download,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("Ocorreu um erro e o arquivo final não pôde ser gerado.")

                except Exception as e:
                    st.error(f"Ocorreu um erro inesperado durante o processo: {e}")
        
        elif tipo_material == "Produto Acabado":
            st.info("A funcionalidade para 'Produto Acabado' será ajustada a seguir. Por favor, selecione 'Bobina' por enquanto.")