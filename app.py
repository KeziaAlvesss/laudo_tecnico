import streamlit as st
from xhtml2pdf import pisa
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configurações de e-mail
EMAIL_REMETENTE = "keziaalves318@gmail.com"
EMAIL_SENHA = "gkblniuaemhffiwq"  # App Password do Gmail
EMAIL_DESTINO = "qualidade@bonsono.com.br"

def gerar_html_laudo(dados):
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20mm;
                line-height: 1.6;
            }}
            .header {{
                text-align: center;
                margin-bottom: 20px;
            }}
            h2 {{
                text-align: center;
                margin-top: 10px;
                margin-bottom: 15px;
                font-size: 18px;
                font-weight: bold;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                font-size: 14px;
            }}
            th, td {{
                border: 1px solid #000;
                padding: 6px;
                text-align: left;
            }}
            th {{
                background-color: #f2f2f2;
                font-weight: bold;
            }}
            .section-title {{
                font-weight: bold;
                margin-top: 20px;
                margin-bottom: 5px;
                font-size: 14px;
            }}
            .signature-box {{
                margin-top: 40px;
                display: flex;
                justify-content: space-between;
                gap: 20px;
            }}
            .signature-item {{
                width: 45%;
                text-align: center;
                padding: 5px;
            }}
            .observation-box {{
                border: 1px solid #000;
                padding: 10px;
                min-height: 80px;
                margin-top: 10px;
            }}
            hr {{
                border: none;
                border-bottom: 1px solid #000;
                width: 80%;
                margin: 5px auto;
            }}
        </style>
    </head>
    <body>

        <!-- CABEÇALHO -->
        <div class="header">
            <h2>LAUDO TÉCNICO – BLOCO LAMINADO</h2>
        </div>

        <!-- SEÇÃO 1: IDENTIFICAÇÃO DO PEDIDO -->
        <div class="section-title">1. Identificação do Pedido</div>
        <table>
            <tr>
                <th>Número do Pedido:</th>
                <td>{dados['pedido']}</td>
                <th>Cliente:</th>
                <td>{dados['cliente']}</td>
            </tr>
        </table>

        <!-- SEÇÃO 2: IDENTIFICAÇÃO DO BLOCO -->
        <div class="section-title">2. Identificação do Bloco</div>
        <table>
            <tr>
                <th>Densidade do Bloco (kg/m³):</th>
                <td>{dados['densidade_bloco']}</td>
            </tr>
            <tr>
                <th>Medidas do Bloco (Largura x Comprimento x Altura) (cm):</th>
                <td>{dados['bloco_largura_cm']} × {dados['bloco_comprimento_cm']} × {dados['bloco_altura_cm']}</td>
            </tr>
            <tr>
                <th>Número do Bloco:</th>
                <td>{dados['num_bloco']}</td>
            </tr>
            <tr>
                <th>Lote do Bloco:</th>
                <td>{dados['lote_bloco']}</td>
            </tr>
        </table>

        <!-- SEÇÃO 3: CÁLCULO DA DENSIDADE DA LÂMINA -->
        <div class="section-title">3. Cálculo da Densidade da Lâmina</div>
        <table>
            <tr>
                <th>Lâmina de Topo, Meio ou Fundo</th>
                <td>{dados['posicao_lamina']}</td>
            </tr>
            <tr>
                <th>Peso da Lâmina (kg):</th>
                <td>{dados['peso_kg']:.2f} kg</td>
            </tr>
            <tr>
                <th>Largura da Lâmina (m):</th>
                <td>{dados['largura_lamina_m']:.2f} m</td>
            </tr>
            <tr>
                <th>Comprimento da Lâmina (m):</th>
                <td>{dados['comprimento_lamina_m']:.2f} m</td>
            </tr>
            <tr>
                <th>Altura da Lâmina (m):</th>
                <td>{dados['altura_lamina_m']:.2f} m</td>
            </tr>
            <tr>
                <th>Volume da Lâmina (m³):</th>
                <td>{dados['volume_m3']:.3f} m³</td>
            </tr>
            <tr>
                <th>Densidade Obtida (kg/m³):<br>(Peso/Volume)</th>
                <td>{dados['densidade_obtida']:.2f} kg/m³</td>
            </tr>
        </table>

        <!-- SEÇÃO 4: OBSERVAÇÕES -->
        <div class="section-title">4. Observações</div>
        <div class="observation-box">
            {dados['observacoes'].replace(chr(10), '<br>')}
        </div>

        <!-- ASSINATURA E DATA -->
        <div class="signature-box">
            <div class="signature-item">
                Data: {dados['data']}<br>
                <hr>
            </div>
            <div class="signature-item">
                Responsável: {dados['responsavel']}<br>
                <hr>
            </div>
        </div>

    </body>
    </html>
    """

def gerar_pdf(html_content):
    try:
        pdf_buffer = BytesIO()
        result = pisa.CreatePDF(
            html_content,
            dest=pdf_buffer,
            encoding='utf-8'
        )
        if not result.err:
            return pdf_buffer.getvalue()
        else:
            st.error("Erro ao gerar PDF")
            return None
    except Exception as e:
        st.error(f"Exceção ao gerar PDF: {e}")
        return None

def enviar_email_com_pdf(pdf_bytes, assunto="Laudo Técnico - BonSono"):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    msg['Subject'] = assunto

    corpo = "Prezado(a),\n\nEm anexo, segue o laudo técnico do bloco laminado."
    msg.attach(MIMEText(corpo, 'plain'))

    parte = MIMEBase('application', 'octet-stream')
    parte.set_payload(pdf_bytes)
    encoders.encode_base64(parte)
    parte.add_header('Content-Disposition', 'attachment; filename=laudo_tecnico.pdf')
    msg.attach(parte)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, EMAIL_SENHA)
        texto = msg.as_string()
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINO, texto)
        server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e:
        return False, str(e)

# === Interface Streamlit ===
st.set_page_config(page_title="Laudo Técnico - BonSono", layout="wide")
st.title("📄 Laudo Técnico – Bloco Laminado")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Identificação do Pedido")
    pedido = st.text_input("Número do Pedido", key="pedido")
    cliente = st.text_input("Cliente", key="cliente")

    st.subheader("2. Identificação do Bloco")
    densidade_bloco = st.text_input("Densidade do Bloco (kg/m³)", key="densidade_bloco")
    num_bloco = st.text_input("Número do Bloco", key="num_bloco")
    lote_bloco = st.text_input("Lote do Bloco", key="lote_bloco")

    c1, c2, c3 = st.columns(3)
    with c1:
        bloco_largura_cm = st.number_input("Largura (cm)", format="%.2f", key="bloco_largura_cm")
    with c2:
        bloco_comprimento_cm = st.number_input("Comprimento (cm)", format="%.2f", key="bloco_comprimento_cm")
    with c3:
        bloco_altura_cm = st.number_input("Altura (cm)", format="%.2f", key="bloco_altura_cm")

with col2:
    st.subheader("3. Cálculo da Densidade da Lâmina")
    posicao_lamina = st.selectbox("Posição da Lâmina", ["Topo", "Meio", "Fundo"], key="posicao_lamina")

    c_larg, c_comp, c_alt = st.columns(3)
    with c_larg:
        largura_lamina_m = st.number_input("Largura da Lâmina (m)", format="%.4f", value=0.0, key="largura_lamina_m")
    with c_comp:
        comprimento_lamina_m = st.number_input("Comprimento da Lâmina (m)", format="%.4f", value=0.0, key="comprimento_lamina_m")
    with c_alt:
        altura_lamina_m = st.number_input("Altura da Lâmina (m)", format="%.4f", value=0.0, key="altura_lamina_m")

    peso_kg = st.number_input("Peso da Lâmina (kg)", format="%.4f", key="peso_kg")

    volume_m3 = largura_lamina_m * comprimento_lamina_m * altura_lamina_m
    densidade_obtida = peso_kg / volume_m3 if volume_m3 > 0 else 0

    st.text(f"Volume da Lâmina (m³): {volume_m3:.3f}".replace('.', ','))
    st.text(f"Densidade Obtida (kg/m³): {densidade_obtida:.2f}".replace('.', ','))

st.subheader("4. Observações")
observacoes = st.text_area("Observações técnicas", key="observacoes")

c_data, c_resp = st.columns(2)
with c_data:
    data_input = st.date_input("Data", key="data")
    data = data_input.strftime("%d/%m/%Y") if data_input else ""
with c_resp:
    responsavel = st.text_input("Assinatura do Responsável", key="responsavel")

# Botões de ação
st.markdown("---")
col_btn1, col_btn2, col_btn3 = st.columns(3)

with col_btn1:
    if st.button("📥 Baixar PDF"):
        dados = {
            'pedido': st.session_state.get("pedido", ""),
            'cliente': st.session_state.get("cliente", ""),
            'densidade_bloco': st.session_state.get("densidade_bloco", ""),
            'num_bloco': st.session_state.get("num_bloco", ""),
            'lote_bloco': st.session_state.get("lote_bloco", ""),
            'bloco_largura_cm': st.session_state.get("bloco_largura_cm", 0.0),
            'bloco_comprimento_cm': st.session_state.get("bloco_comprimento_cm", 0.0),
            'bloco_altura_cm': st.session_state.get("bloco_altura_cm", 0.0),
            'posicao_lamina': st.session_state.get("posicao_lamina", "Topo"),
            'peso_kg': st.session_state.get("peso_kg", 0.0),
            'largura_lamina_m': st.session_state.get("largura_lamina_m", 0.0),
            'comprimento_lamina_m': st.session_state.get("comprimento_lamina_m", 0.0),
            'altura_lamina_m': st.session_state.get("altura_lamina_m", 0.0),
            'volume_m3': volume_m3,
            'densidade_obtida': densidade_obtida,
            'observacoes': st.session_state.get("observacoes", ""),
            'data': data,
            'responsavel': st.session_state.get("responsavel", ""),
        }
        html = gerar_html_laudo(dados)
        pdf = gerar_pdf(html)
        if pdf:
            st.download_button(
                "⬇️ Clique para baixar",
                data=pdf,
                file_name="laudo_tecnico.pdf",
                mime="application/pdf"
            )
        else:
            st.error("Falha ao gerar PDF.")

with col_btn2:
    if st.button("📤 Enviar por E-mail"):
        pedido_val = st.session_state.get("pedido", "").strip()
        cliente_val = st.session_state.get("cliente", "").strip()
        if not pedido_val or not cliente_val:
            st.error("Preencha pelo menos o Pedido e o Cliente.")
        else:
            dados = { ... }  # (mesmo que acima)
            html = gerar_html_laudo(dados)
            pdf = gerar_pdf(html)
            if pdf:
                sucesso, msg = enviar_email_com_pdf(pdf)
                if sucesso:
                    st.success(msg)
                else:
                    st.error(f"Erro: {msg}")
            else:
                st.error("Falha ao gerar PDF.")

with col_btn3:
    if st.button("🗑️ Limpar"):
        keys_to_clear = [
            "pedido", "cliente", "densidade_bloco", "num_bloco", "lote_bloco",
            "bloco_largura_cm", "bloco_comprimento_cm", "bloco_altura_cm",
            "posicao_lamina", "largura_lamina_m", "comprimento_lamina_m", "altura_lamina_m",
            "peso_kg", "observacoes", "responsavel", "data"
        ]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()