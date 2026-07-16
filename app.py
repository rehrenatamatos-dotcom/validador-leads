# Validador de Leads — Soluções Industriais
# Web app gratuito (Streamlit + Groq) para validar leads contra o perfil do cliente.

import csv
import io
import json
import re
import time

import requests
import streamlit as st

MODELO = "llama-3.3-70b-versatile"
URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"
TAMANHO_LOTE = 15
MAX_TENTATIVAS = 4
STATUS_VALIDOS = {"Dentro do foco", "Fora do foco", "Aberto"}

AZUL = "#185FA5"
AZUL_ESCURO = "#0C447C"
AZUL_CLARO = "#E6F1FB"

PROMPT_SISTEMA = """Você é um analista de qualidade de leads de uma plataforma de geração de leads B2B.
Sua tarefa: para cada lead abaixo, decidir se ele está DENTRO ou FORA do foco do cliente descrito no perfil, ou se está ABERTO (mensagem sem informação suficiente para avaliar).

Critérios (avalie em conjunto, nenhum sozinho decide):
1. Material/produto compatível com o portfólio do cliente. Se o cliente trabalha metal e o lead pede madeira/tecido/PVC, é forte sinal de "Fora do foco", mesmo que o serviço seja o mesmo.
2. B2B vs. uso pessoal. Pedidos claramente domésticos/pontuais de pessoa física pesam para "Fora do foco" quando o cliente atende indústria/B2B.
3. Especificidade técnica. Medidas, normas, quantidade definida, nome de empresa/CNPJ pesam para "Dentro do foco". Mensagens vagas sem material/aplicação tendem a "Aberto".
4. Sinais de ruído. Teste interno (QA, e-mails de qualidade), spam, concorrente se oferecendo, ou lead avisando que já comprou em outro lugar = "Fora do foco" independente do produto.

Regras de saída:
- STATUS deve ser EXATAMENTE um destes: "Dentro do foco", "Fora do foco", "Aberto".
- MOTIVO: uma frase objetiva em português citando a evidência da própria mensagem.
- Duplicatas (mesma mensagem) recebem a mesma classificação.
- Responda SOMENTE com um objeto JSON: {"resultados": [{"id": "...", "status": "...", "motivo": "..."}]} — um item por lead, na mesma ordem."""


# ---------- Funções ----------

def obter_chave():
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        import os
        return os.environ.get("GROQ_API_KEY", "")


def buscar_site(url, limite=8000):
    try:
        r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        html = r.text
    except Exception:
        return ""
    html = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", html)
    texto = re.sub(r"(?s)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", texto).strip()[:limite]


def chamar_groq(api_key, perfil, lote):
    leads_texto = "\n\n".join(
        f"LEAD id={l['id']}\nMensagem: {l['mensagem']}\nContexto extra: {l['extra']}"
        for l in lote
    )
    corpo = {
        "model": MODELO,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": PROMPT_SISTEMA},
            {"role": "user", "content": f"PERFIL DO CLIENTE:\n{perfil}\n\nLEADS A CLASSIFICAR:\n{leads_texto}"},
        ],
    }
    ultima = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = requests.post(
                URL_GROQ,
                json=corpo,
                headers={"Authorization": f"Bearer {api_key}"},
                timeout=120,
            )
            if r.status_code == 429 or r.status_code >= 500:
                time.sleep(min(30, 4 * tentativa))
                continue
            r.raise_for_status()
            texto = r.json()["choices"][0]["message"]["content"]
            parsed = json.loads(texto)
            if isinstance(parsed, list):
                return parsed
            for v in parsed.values():
                if isinstance(v, list):
                    return v
            raise ValueError("Resposta sem lista de resultados.")
        except Exception as e:
            ultima = e
            time.sleep(2)
    raise ultima


# ---------- Interface ----------

st.set_page_config(page_title="Validador de Leads", page_icon="✅", layout="centered")

st.markdown(f"""
<style>
  .header-si {{
    background: {AZUL}; border-radius: 10px; padding: 18px 24px; margin-bottom: 24px;
  }}
  .header-si h1 {{ color: #fff; font-size: 1.4rem; margin: 0; }}
  .header-si p {{ color: {AZUL_CLARO}; margin: 2px 0 0; font-size: 0.85rem; }}
  .stButton > button[kind="primary"] {{ background: {AZUL}; border: none; }}
  div[data-testid="stMetric"] {{
    background: {AZUL_CLARO}; border-radius: 10px; padding: 12px 16px;
  }}
  div[data-testid="stMetric"] label {{ color: {AZUL}; }}
  div[data-testid="stMetricValue"] {{ color: {AZUL_ESCURO}; }}
</style>
<div class="header-si">
  <h1>Validador de Leads</h1>
  <p>Soluções Industriais — classifica cada lead como Dentro do foco, Fora do foco ou Aberto</p>
</div>
""", unsafe_allow_html=True)

st.markdown(f"<p style='color:{AZUL_ESCURO}; font-weight:600;'>1. Arquivos do cliente (briefing obrigatório; contexto e nutrições melhoram a precisão)</p>", unsafe_allow_html=True)
arquivos_txt = st.file_uploader(
    "Envie um ou mais arquivos .txt / .md", type=["txt", "md"],
    accept_multiple_files=True, label_visibility="collapsed",
)

st.markdown(f"<p style='color:{AZUL_ESCURO}; font-weight:600;'>2. Site do cliente (opcional)</p>", unsafe_allow_html=True)
site = st.text_input("URL do site", placeholder="https://www.sitedocliente.com.br", label_visibility="collapsed")

st.markdown(f"<p style='color:{AZUL_ESCURO}; font-weight:600;'>3. CSV de leads</p>", unsafe_allow_html=True)
arquivo_csv = st.file_uploader("Envie o CSV", type=["csv"], label_visibility="collapsed")

with st.expander("Configurações das colunas"):
    col_id = st.text_input("Coluna de ID", value="ID do Orçamento")
    col_msg = st.text_input("Coluna da mensagem", value="Mensagem do Cliente")

if st.button("Validar leads", type="primary", use_container_width=True):
    api_key = obter_chave()
    if not api_key:
        st.error("Chave do Groq não configurada. O administrador precisa adicionar o segredo GROQ_API_KEY nas configurações do Space.")
        st.stop()
    if not arquivos_txt:
        st.error("Envie ao menos um arquivo .txt com o briefing do cliente.")
        st.stop()
    if not arquivo_csv:
        st.error("Envie o CSV de leads.")
        st.stop()

    # Perfil do cliente
    partes = []
    for f in arquivos_txt:
        conteudo = f.read().decode("utf-8", errors="replace").strip()
        if conteudo:
            partes.append(f"===== {f.name} =====\n{conteudo}")
    perfil = "\n\n".join(partes)
    if site.strip():
        with st.spinner("Lendo o site do cliente..."):
            texto_site = buscar_site(site.strip())
        if texto_site:
            perfil += f"\n\n===== SITE DO CLIENTE ({site.strip()}) =====\n{texto_site}"
        else:
            st.warning("Não consegui ler o site — prosseguindo sem ele.")

    # CSV
    try:
        conteudo_csv = arquivo_csv.read().decode("utf-8-sig", errors="replace")
        linhas = list(csv.reader(io.StringIO(conteudo_csv)))
    except Exception as e:
        st.error(f"Não consegui ler o CSV: {e}")
        st.stop()
    if len(linhas) < 2:
        st.error("O CSV está vazio ou só tem o cabeçalho.")
        st.stop()

    cabecalho = [c.strip() for c in linhas[0]]
    if col_msg not in cabecalho:
        st.error(f'Coluna "{col_msg}" não encontrada. Colunas do arquivo: {", ".join(cabecalho)}')
        st.stop()
    idx_id = cabecalho.index(col_id) if col_id in cabecalho else -1
    idx_msg = cabecalho.index(col_msg)

    registros = [r for r in linhas[1:] if any(c.strip() for c in r)]
    leads = []
    for i, r in enumerate(registros):
        extra = "; ".join(
            f"{cabecalho[j]}: {r[j]}" for j in range(len(cabecalho))
            if j not in (idx_id, idx_msg) and j < len(r) and r[j].strip()
        )[:500]
        leads.append({
            "id": f"L{i+2}",
            "mensagem": r[idx_msg] if idx_msg < len(r) else "",
            "extra": extra,
        })

    # Classificação
    classificacoes = {}
    total_lotes = (len(leads) + TAMANHO_LOTE - 1) // TAMANHO_LOTE
    progresso = st.progress(0, text=f"Classificando {len(leads)} leads...")
    falhas = 0
    for n in range(total_lotes):
        lote = leads[n * TAMANHO_LOTE:(n + 1) * TAMANHO_LOTE]
        try:
            resultado = chamar_groq(api_key, perfil, lote)
        except Exception:
            falhas += len(lote)
            resultado = []
        for item in resultado:
            status = str(item.get("status", "")).strip()
            if status not in STATUS_VALIDOS:
                status = "Aberto"
            classificacoes[str(item.get("id", "")).strip()] = {
                "status": status,
                "motivo": str(item.get("motivo", "")).strip(),
            }
        progresso.progress((n + 1) / total_lotes, text=f"Lote {n+1} de {total_lotes} concluído")
        if n + 1 < total_lotes:
            time.sleep(3)
    progresso.empty()

    # CSV de saída
    saida = io.StringIO()
    w = csv.writer(saida)
    w.writerow(cabecalho + ["STATUS", "MOTIVO"])
    contagem = {"Dentro do foco": 0, "Fora do foco": 0, "Aberto": 0}
    for i, r in enumerate(registros):
        c = classificacoes.get(leads[i]["id"], {
            "status": "Aberto", "motivo": "Não classificado pela IA — revisar manualmente.",
        })
        contagem[c["status"]] += 1
        w.writerow(r + [c["status"], c["motivo"]])

    total = len(registros)
    st.success(f"Validação concluída — {total} leads processados!")
    c1, c2, c3 = st.columns(3)
    c1.metric("Dentro do foco", f"{contagem['Dentro do foco']/total:.0%}", f"{contagem['Dentro do foco']} leads", delta_color="off")
    c2.metric("Fora do foco", f"{contagem['Fora do foco']/total:.0%}", f"{contagem['Fora do foco']} leads", delta_color="off")
    c3.metric("Aberto", f"{contagem['Aberto']/total:.0%}", f"{contagem['Aberto']} leads", delta_color="off")
    if falhas:
        st.warning(f"{falhas} lead(s) ficaram sem resposta da IA e foram marcados como Aberto — rode de novo se quiser reprocessar.")

    nome_saida = re.sub(r"\.csv$", "", arquivo_csv.name, flags=re.I) + " - Validado.csv"
    st.download_button(
        "Baixar CSV validado", data=saida.getvalue().encode("utf-8-sig"),
        file_name=nome_saida, mime="text/csv", use_container_width=True,
    )

st.markdown(
    f"<p style='text-align:center; color:#8a8a8a; font-size:0.75rem; margin-top:32px;'>"
    f"Validador de Leads · Soluções Industriais</p>",
    unsafe_allow_html=True,
)
