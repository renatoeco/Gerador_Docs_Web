import streamlit as st
from docx import Document
from docx.shared import Inches
import pandas as pd
import os
import zipfile
import tempfile
from io import BytesIO
import shutil



####################  functions

# Função do diálogo do resultado final
@st.dialog("Resultado")
def resultado():
    # Validar se todas as informações necessárias estão presentes
    validate()

    # Botão para reiniciar o processo
    if st.button("Reiniciar"):
        st.session_state.clear()
        st.rerun()




# Função para validar se tem todas as informações necessárias.
def validate():
    """
    Verifica se todas as informações necessárias para gerar os documentos
    estão presentes.
    """
    if "sheet_path" in st.session_state and "model_path" in st.session_state:
        funcao_botao()
    else:
        st.session_state.success = "FALTA"    

# Função do botão que vai executar
def funcao_botao():
    """
    Chama a função para gerar os documentos em lote.
    """
    parametro1 = st.session_state.sheet_path
    parametro2 = st.session_state.model_path
    gerar_docs(parametro1, parametro2)


def gerar_docs(caminho_xlsx, caminho_docx):
    # Cria um diretório temporário para armazenar os arquivos gerados
    temp_dir = tempfile.mkdtemp()

    # >>>>>> EXTRAIR <<<<<<
    doc_carregado = Document(caminho_docx)
    df_contratos = pd.read_excel(caminho_xlsx)

    # >>>>>> TRANSFORMAR <<<<<<
    for index, row in df_contratos.iterrows():

        st.session_state.cont += 1

        # Criar um documento novo
        document = Document()
        document.styles['Normal'].font.name = 'Arial'

        # Adiciona o logotipo no cabeçalho
        cabecalho = document.sections[0].header
        paragrafo = cabecalho.paragraphs[0]

        if st.session_state.image_path != "no_image":
            imagem = st.session_state.image_path
            paragrafo.alignment = 1
            paragrafo.add_run().add_picture(imagem, width=Inches(2))

        # Adiciona uma margem abaixo da imagem
        paragrafo_vazio = cabecalho.add_paragraph()
        paragrafo_vazio.space_after = Inches(0.5)

        # Adiciona a imagem no footer
        footer = document.sections[0].footer
        paragrafo_footer = footer.paragraphs[0]

        if st.session_state.image_footer_path != "no_image":
            imagem_footer = st.session_state.image_footer_path
            paragrafo_footer.alignment = 1
            paragrafo_footer.add_run().add_picture(imagem_footer, width=Inches(2))

        # Monta os parágrafos com substituições e preserva a formatação
        for paragrafo in doc_carregado.paragraphs:
            alinhamento = paragrafo.alignment
            novo_paragrafo = document.add_paragraph()
            novo_paragrafo.alignment = alinhamento

            # Itera sobre as "runs" de cada parágrafo para preservar a formatação
            for run in paragrafo.runs:
                texto = run.text
                # Substitui as variáveis no texto
                for coluna in df_contratos.columns:
                    variavel = "{{{" + coluna + "}}}"
                    if variavel in texto:
                        texto = texto.replace(variavel, str(row[coluna]))

                # Adiciona o texto com a formatação
                novo_run = novo_paragrafo.add_run(texto)

                # Copia a formatação do "run" original (negrito, itálico, etc.)
                novo_run.bold = run.bold
                novo_run.italic = run.italic
                novo_run.underline = run.underline
                novo_run.font.size = run.font.size
                novo_run.font.name = run.font.name

        # Cria os nomes dinâmicos para os arquivos finais
        nome_arquivo = os.path.join(temp_dir, f"{row.iloc[0].replace('/', '_')}.docx")
        document.save(nome_arquivo)

    # Compacta os documentos em um arquivo ZIP
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename in os.listdir(temp_dir):
            filepath = os.path.join(temp_dir, filename)
            zip_file.write(filepath, os.path.basename(filepath))

    # Volta para o início do buffer
    zip_buffer.seek(0)


    # SUCESSO <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    st.balloons()
    st.session_state.success = "OK"
    st.session_state.hide_button = True

    # Mensagem de sucesso
    if st.session_state.success == "OK":
        st.success(f'{st.session_state.cont} documentos gerados com sucesso!')
    elif st.session_state.success == "FALTA":
        st.warning("Você precisa escolher a TABELA e o MODELO (etapas 2 e 3).")

    # Oferece o arquivo ZIP para download
    st.download_button(
        label="Baixar ZIP com os arquivos .docx",
        data=zip_buffer,
        file_name="documentos_gerados.zip",
        mime="application/zip"
    )

    # Limpa o diretório temporário
    shutil.rmtree(temp_dir)
    del zip_buffer  # Libera o buffer em memória





# Função principal
def main():
    # Inicializa a contagem de documentos
    if not "cont" in st.session_state: 
        st.session_state.cont = 0

    # Inicializa o marcador de sucesso
    if not "success" in st.session_state: 
        st.session_state.success = ""

    # Inicializa o marcador de esconder o botão
    if not "hide_button" in st.session_state:
        st.session_state.hide_button = False

    # Mostra o logotipo
    logo = "https://avatars.githubusercontent.com/u/85522293?v=4"
    
    with st.columns(3)[1]:
        st.image(logo, output_format="PNG", width=220)

    # Título e subtítulo
    st.markdown("<h1 style='text-align: center;'>Gerador de Documentos</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Para gerar documentos em lote, siga os passos a seguir:</p>", unsafe_allow_html=True)

    st.write('')
    st.write('')
    st.write('')

    
    # Passo 1 - Imagem do cabeçalho ----------------------------------------------
    st.markdown("<h5>1. Escolha uma imagem para o cabeçalho:</h5>", unsafe_allow_html=True)

    # Opção de seleção
    opcao_imagem = st.selectbox('', ('Continuar sem imagem no cabeçalho', 'Carregar imagem'))

    if opcao_imagem == "Carregar imagem":
        arquivo_enviado = st.file_uploader("Escolha uma imagem (PNG, JPG ou JPEG)", type=["png", "jpg", "jpeg"])
        if arquivo_enviado:
            st.session_state.image_path = arquivo_enviado

            # Criando colunas para layout
            col1, col2 = st.columns([1, 3])

            col1.write("Imagem escolhida:")
            col2.image(st.session_state.image_path, width=150)


    elif opcao_imagem == "Continuar sem imagem no cabeçalho":
        st.session_state.image_path = "no_image"
    
    # Passo 2 - Carregar planilha ----------------------------------------------
    st.markdown("<h5 style='padding-top: 40px'>2. Carregue a tabela com as informações:</h5>", unsafe_allow_html=True)
    st.markdown("<span style='color:gray;'>Apenas são aceitas tabelas no formato .xlsx.</span>", unsafe_allow_html=True)
    st.markdown("<span style='color:gray;'>A <strong>primeira coluna da tabela</strong> será usada no <strong>nome dos arquivos</strong> gerados.</span>", unsafe_allow_html=True)

    # Upload do arquivo diretamente (sem botão separado)
    arquivo_enviado = st.file_uploader("", type=["xlsx"])

    # Salva no session_state apenas se o usuário enviar um arquivo
    if arquivo_enviado is not None:
        st.session_state.sheet_path = arquivo_enviado

    # Passo 3 - Carregar modelo ----------------------------------------------
    st.markdown("<h5 style='padding-top: 40px'>3. Carregue o modelo do documento:</h5>", unsafe_allow_html=True)
    st.markdown("<span style='color:gray;'>Apenas são aceitos documentos no formato .docx.</span>", unsafe_allow_html=True)
    st.markdown("<span style='color:gray;'>As variárveis no modelo devem ter o mesmo nome das colunas da planilha, embaladas com chaves triplas {{{ }}}.</span>", unsafe_allow_html=True)

    # Upload do modelo diretamente (sem botão separado)
    arquivo_enviado = st.file_uploader("Escolha o modelo (.docx)", type=["docx"])

    # Salva no session_state apenas se o usuário enviar um arquivo
    if arquivo_enviado is not None:
        st.session_state.model_path = arquivo_enviado

    # Passo 4 - Imagem do rodapé ----------------------------------------------
    st.markdown("<h5 style='padding-top: 40px'>4. Escolha uma imagem para o rodapé:</h5>", unsafe_allow_html=True)
    
    opcao_imagem_rodape = st.selectbox('', ('Continuar sem imagem no rodapé', 'Carregar imagem'))

    if opcao_imagem_rodape == "Carregar imagem":
        arquivo_enviado = st.file_uploader("Escolha uma imagem para o rodapé", type=["png", "jpg", "jpeg"])
        if arquivo_enviado:
            st.session_state.image_footer_path = arquivo_enviado

    elif opcao_imagem_rodape == "Continuar sem imagem no rodapé":
        st.session_state.image_footer_path = "no_image"

    # Verificando se as informações estão presentes no session_state
    if "image_footer_path" in st.session_state:
        if st.session_state.image_footer_path == "no_image":
            st.write("Não usar imagem no rodapé.")
        else:
            col1, col2 = st.columns([1, 3])
            col1.write("Imagem escolhida:")
            col2.image(st.session_state.image_footer_path, width=150)

    # Passo 5 - Escolha o nome dos arquivos ----------------------------------------------
    st.markdown("<h5 style='padding-top: 40px'>5. Qual nome deseja dar aos documentos?</h5>", unsafe_allow_html=True)
    if "docs_name" not in st.session_state:
        nome_doc_input = st.text_input("Digite e APERTE ENTER")
        if nome_doc_input != "":
            st.session_state.docs_name = nome_doc_input

    # Passo 6 - Gerar documentos ----------------------------------------------
    st.markdown("<h5 style='padding-top: 40px'>6. Clique no botão abaixo para gerar os documentos:</h5>", unsafe_allow_html=True)
    st.button("Gerar documentos!", type="primary", disabled=st.session_state.hide_button, on_click=resultado)



if __name__ == "__main__":
    main()
