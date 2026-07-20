# Validador de Leads v3 — Soluções Industriais
# v2 (busca automática no Metabase) + dashboard HTML de resultados para download.

import csv
import io
import json
import os
import re
import time
from datetime import date, datetime, timedelta

import requests
import streamlit as st

MODELO = "meta-llama/llama-4-scout-17b-16e-instruct"
MODELO_RESERVA = "llama-3.3-70b-versatile"
URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"
TAMANHO_LOTE = 20
MAX_TENTATIVAS = 5
LIMITE_FONTE = 4000
LIMITE_PERFIL = 9000
STATUS_VALIDOS = {"Dentro do foco", "Fora do foco", "Aberto"}

CARD_ORCAMENTOS = 47   # question do Metabase: base de orçamentos por chave única
CARD_BRIEFING = 286    # question do Metabase: briefing do cliente por chave única
CARD_ANUNCIOS = 185    # question do Metabase: anúncios por chave única

AZUL = "#185FA5"
AZUL_ESCURO = "#0C447C"
AZUL_CLARO = "#E6F1FB"

ARQUIVO_HISTORICO = "historico.json"
PASTA_RESULTADOS = "resultados"

PROMPT_SISTEMA = """Você é um analista de qualidade de leads de uma plataforma de geração de leads B2B.
Sua tarefa: para cada lead abaixo, decidir se ele está DENTRO ou FORA do foco do cliente descrito no perfil, ou se está ABERTO (mensagem sem informação suficiente para avaliar).

ANTES DE CLASSIFICAR: identifique no perfil O QUE o cliente vende — um PRODUTO (ex. máquinas, equipamentos) ou um SERVIÇO (ex. corte sob medida, usinagem para terceiros). Essa distinção é o critério mais importante.

Critérios (avalie em conjunto, nenhum sozinho decide):
1. Produto vs. serviço. Se o cliente VENDE MÁQUINAS e o lead quer CONTRATAR o serviço (ex. "preciso cortar 50 chapas", "orçamento para corte de peças"), é "Fora do foco" — mesmo que a mensagem seja tecnicamente detalhada (material, medidas, CNPJ). Especificidade técnica NÃO transforma um pedido de serviço em lead de máquina. O inverso também vale (cliente presta serviço e lead quer comprar máquina). Salvo se o perfil disser que o cliente atende ambos.
2. Modalidade compatível. Pedidos de assistência técnica/manutenção, aluguel de máquina, ou peças/componentes avulsos são "Fora do foco" quando o cliente vende equipamentos novos — salvo indicação contrária no perfil ou nas regras específicas.
3. Material/produto compatível com o portfólio do cliente. Se o cliente trabalha metal e o lead pede madeira/tecido/PVC, é forte sinal de "Fora do foco", mesmo que o serviço seja o mesmo.
4. B2B vs. uso pessoal. Pedidos claramente domésticos/pontuais de pessoa física pesam para "Fora do foco" quando o cliente atende indústria/B2B.
5. Especificidade técnica. Medidas, normas, quantidade definida, nome de empresa/CNPJ pesam para "Dentro do foco" — mas SOMENTE quando o pedido é da modalidade certa (ver critério 1).
6. Sinais de ruído. Teste interno (QA, e-mails de qualidade), spam, concorrente se oferecendo, marca/modelo que o cliente não vende, ou lead avisando que já comprou em outro lugar = "Fora do foco" independente do produto.
7. Regras específicas do cliente (se fornecidas no perfil) têm prioridade sobre os critérios gerais.

Regras de saída:
- STATUS deve ser EXATAMENTE um destes: "Dentro do foco", "Fora do foco", "Aberto".
- MOTIVO: uma frase objetiva em português citando a evidência da própria mensagem. O motivo deve justificar o STATUS escolhido, não outro.
- Mensagens vagas demais para julgar (ex. apenas "aço inox", apenas "me manda o e-mail", uma palavra solta sem contexto de compra) = "Aberto". NUNCA marque "Dentro do foco" sem evidência de interesse na modalidade certa (compra do produto que o cliente vende).
- Peças, componentes e insumos avulsos (ex. fonte, tubo de laser, lentes) = "Fora do foco" quando o cliente vende máquinas completas, salvo indicação contrária no perfil.
- Mensagens idênticas ou quase idênticas (mesmo texto em vários leads) DEVEM receber exatamente a mesma classificação e o mesmo motivo — revise antes de responder.
- Responda SOMENTE com um objeto JSON: {"resultados": [{"id": "...", "status": "...", "motivo": "..."}]} — um item por lead, na mesma ordem."""


# ---------- Histórico ----------

def carregar_historico():
    try:
        with open(ARQUIVO_HISTORICO, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def salvar_no_historico(registro, csv_bytes=None, dash_bytes=None):
    historico = carregar_historico()
    historico.insert(0, registro)
    historico = historico[:200]
    try:
        with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False)
        os.makedirs(PASTA_RESULTADOS, exist_ok=True)
        rid = registro.get("id", "")
        if rid and csv_bytes:
            with open(os.path.join(PASTA_RESULTADOS, f"{rid}.csv"), "wb") as f:
                f.write(csv_bytes)
        if rid and dash_bytes:
            with open(os.path.join(PASTA_RESULTADOS, f"{rid}.html"), "wb") as f:
                f.write(dash_bytes)
    except Exception:
        pass


def excluir_do_historico(rid):
    historico = [h for h in carregar_historico() if h.get("id") != rid]
    try:
        with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False)
        for ext in (".csv", ".html"):
            caminho = os.path.join(PASTA_RESULTADOS, f"{rid}{ext}")
            if os.path.exists(caminho):
                os.remove(caminho)
    except Exception:
        pass


def ler_resultado_salvo(rid, ext):
    try:
        with open(os.path.join(PASTA_RESULTADOS, f"{rid}{ext}"), "rb") as f:
            return f.read()
    except Exception:
        return None


# ---------- Metabase ----------

def secret(nome, padrao=""):
    try:
        return str(st.secrets.get(nome, padrao)).strip()
    except Exception:
        return os.environ.get(nome, padrao).strip()


def cabecalhos_metabase():
    """Autentica no Metabase por API key ou por login/senha (sessão)."""
    api_key = secret("METABASE_API_KEY")
    if api_key:
        return {"X-API-KEY": api_key}
    if "mb_sessao" in st.session_state:
        return {"X-Metabase-Session": st.session_state["mb_sessao"]}
    usuario, senha = secret("METABASE_USER"), secret("METABASE_PASSWORD")
    if not (usuario and senha):
        return None
    r = requests.post(
        secret("METABASE_URL").rstrip("/") + "/api/session",
        json={"username": usuario, "password": senha},
        timeout=30,
    )
    r.raise_for_status()
    st.session_state["mb_sessao"] = r.json()["id"]
    return {"X-Metabase-Session": st.session_state["mb_sessao"]}


def consultar_question(card_id, parametros):
    """Baixa o resultado CSV de uma question do Metabase com parâmetros."""
    url = secret("METABASE_URL").rstrip("/") + f"/api/card/{card_id}/query/csv"
    headers = cabecalhos_metabase()
    if headers is None:
        raise RuntimeError(
            "Credenciais do Metabase não configuradas. Adicione nos Secrets: "
            "METABASE_URL e (METABASE_API_KEY ou METABASE_USER + METABASE_PASSWORD)."
        )
    params_mb = [
        {"type": tipo, "target": ["variable", ["template-tag", nome]], "value": valor}
        for nome, valor, tipo in parametros
    ]
    r = requests.post(url, headers=headers, data={"parameters": json.dumps(params_mb)}, timeout=120)
    if r.status_code == 401 and "mb_sessao" in st.session_state:
        del st.session_state["mb_sessao"]           # sessão expirada → renova e tenta de novo
        headers = cabecalhos_metabase()
        r = requests.post(url, headers=headers, data={"parameters": json.dumps(params_mb)}, timeout=120)
    r.raise_for_status()
    return r.content.decode("utf-8-sig", errors="replace")


# ---------- Fontes auxiliares ----------

def csv_anuncios_para_texto(conteudo_csv, limite=LIMITE_FONTE):
    """Transforma o CSV de anúncios em uma lista compacta de nomes."""
    linhas = list(csv.reader(io.StringIO(conteudo_csv)))
    if len(linhas) < 2:
        return ""
    itens = []
    for r in linhas[1:]:
        valores = [c.strip() for c in r if c.strip()]
        if valores:
            itens.append(" | ".join(valores))
    return "\n".join(itens)[:limite]


def csv_briefing_para_texto(conteudo_csv, limite=LIMITE_FONTE):
    """Converte o CSV do briefing (colunas longas) em texto legível para a IA."""
    linhas = list(csv.reader(io.StringIO(conteudo_csv)))
    if len(linhas) < 2:
        return ""
    h = linhas[0]
    blocos = []
    for r in linhas[1:]:
        campos = [f"{h[j]}: {r[j].strip()}" for j in range(min(len(h), len(r))) if r[j].strip()]
        blocos.append("\n".join(campos))
    return ("\n\n--- produto/linha seguinte ---\n\n".join(blocos))[:limite]


# ---------- Dashboard HTML ----------

MODELO_DASH = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relatório de Validação — __EMPRESA__</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', Arial, sans-serif; color: #EAF3FC; padding: 24px; min-height: 100vh;
    background:
      linear-gradient(rgba(12, 68, 124, 0.86), rgba(10, 58, 107, 0.90)),
      url('https://images.unsplash.com/photo-1513828583688-c52646db42da?w=1600&q=60') center / cover fixed no-repeat;
  }
  .wrap { max-width: 900px; margin: 0 auto; }
  .vidro {
    background: rgba(255, 255, 255, 0.13);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.32);
    border-radius: 16px;
  }
  .topo { padding: 22px 28px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
  .topo h1 { color: #fff; font-size: 20px; font-weight: 600; }
  .topo p { color: #B5D4F4; font-size: 13px; margin-top: 4px; }
  .topo .marca { color: #B5D4F4; font-size: 12px; text-align: right; }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin: 14px 0; }
  .card { padding: 16px; text-align: center; }
  .card .rotulo { font-size: 12px; color: #B5D4F4; margin-bottom: 4px; }
  .card .valor { font-size: 30px; font-weight: 600; }
  .card .sub { font-size: 12px; color: rgba(234,243,252,0.7); margin-top: 2px; }
  .azul { color: #ffffff; } .verde { color: #9FE1A5; } .vermelho { color: #F5A9A9; } .ambar { color: #FAD98F; }
  .paineis { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .painel { padding: 18px; }
  .painel h2 { font-size: 14px; color: #ffffff; margin-bottom: 12px; }
  .rodape { text-align: center; color: rgba(234,243,252,0.6); font-size: 11px; margin-top: 18px; }
  @media (max-width: 700px) { .paineis { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="wrap">
  <div class="topo vidro">
    <div>
      <h1>Relatório de Validação de Leads</h1>
      <p>__EMPRESA__ &nbsp;·&nbsp; chave __CHAVE__ &nbsp;·&nbsp; __PERIODO__</p>
    </div>
    <div class="marca">Soluções Industriais<br>gerado em __GERADO__</div>
  </div>

  <div class="cards">
    <div class="card vidro"><div class="rotulo">Total de leads</div><div class="valor azul">__TOTAL__</div><div class="sub">no período</div></div>
    <div class="card vidro"><div class="rotulo">Dentro do foco</div><div class="valor verde">__PCT_DENTRO__%</div><div class="sub">__N_DENTRO__ leads</div></div>
    <div class="card vidro"><div class="rotulo">Fora do foco</div><div class="valor vermelho">__PCT_FORA__%</div><div class="sub">__N_FORA__ leads</div></div>
    <div class="card vidro"><div class="rotulo">Aberto</div><div class="valor ambar">__PCT_ABERTO__%</div><div class="sub">__N_ABERTO__ leads</div></div>
  </div>

  <div class="paineis">
    <div class="painel vidro">
      <h2>Distribuição por status</h2>
      <canvas id="rosca" height="220"></canvas>
    </div>
    <div class="painel vidro">
      <h2>Leads por status</h2>
      <canvas id="barras" height="220"></canvas>
    </div>
  </div>

  <p class="rodape">Validador de Leads · Soluções Industriais · uso interno</p>
</div>
<script>
const CORES = ["#639922", "#E24B4A", "#EF9F27"];
const ROTULOS = ["Dentro do foco", "Fora do foco", "Aberto"];
const DADOS = [__N_DENTRO__, __N_FORA__, __N_ABERTO__];
const CLARO = "#EAF3FC";
new Chart(document.getElementById("rosca"), {
  type: "doughnut",
  data: { labels: ROTULOS, datasets: [{ data: DADOS, backgroundColor: CORES, borderWidth: 2, borderColor: "rgba(255,255,255,0.4)" }] },
  options: { plugins: { legend: { position: "bottom", labels: { color: CLARO } } }, cutout: "58%" }
});
new Chart(document.getElementById("barras"), {
  type: "bar",
  data: { labels: ROTULOS, datasets: [{ data: DADOS, backgroundColor: CORES, borderRadius: 6 }] },
  options: {
    plugins: { legend: { display: false } },
    scales: {
      y: { beginAtZero: true, ticks: { precision: 0, color: CLARO }, grid: { color: "rgba(255,255,255,0.15)" } },
      x: { ticks: { color: CLARO }, grid: { display: false } }
    }
  }
});
</script>
</body>
</html>"""


def gerar_dashboard_html(empresa, chave, periodo, total, contagem):
    def pct(n):
        return str(round(100 * n / total)) if total else "0"
    html = MODELO_DASH
    trocas = {
        "__EMPRESA__": empresa,
        "__CHAVE__": chave,
        "__PERIODO__": periodo,
        "__GERADO__": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "__TOTAL__": str(total),
        "__PCT_DENTRO__": pct(contagem["Dentro do foco"]),
        "__PCT_FORA__": pct(contagem["Fora do foco"]),
        "__PCT_ABERTO__": pct(contagem["Aberto"]),
        "__N_DENTRO__": str(contagem["Dentro do foco"]),
        "__N_FORA__": str(contagem["Fora do foco"]),
        "__N_ABERTO__": str(contagem["Aberto"]),
    }
    for k, v in trocas.items():
        html = html.replace(k, v)
    return html


# ---------- IA (Groq) ----------

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
            r = requests.post(URL_GROQ, json=corpo,
                              headers={"Authorization": f"Bearer {api_key}"}, timeout=120)
            if r.status_code == 404:
                corpo["model"] = MODELO_RESERVA
                continue
            if r.status_code == 429 or r.status_code >= 500:
                try:
                    espera = float(r.headers.get("retry-after", 0))
                except (TypeError, ValueError):
                    espera = 0
                if espera > 300:
                    raise RuntimeError(
                        f"Cota diária da IA esgotada — o Groq pediu {espera/3600:.1f}h de espera. "
                        "Renova às 21h (horário de Brasília), ou troque a GROQ_API_KEY nos Secrets "
                        "por uma chave de outra conta para continuar agora."
                    )
                time.sleep(min(60, espera + 1) if espera else min(30, 5 * tentativa))
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
        except RuntimeError:
            raise                      # cota diária esgotada: não adianta re-tentar
        except Exception as e:
            ultima = e
            time.sleep(2)
    raise ultima


# ---------- Interface ----------

st.set_page_config(page_title="Validador de Leads v3", page_icon="✅", layout="centered")

IMG_FUNDO = "https://images.unsplash.com/photo-1513828583688-c52646db42da?w=1600&q=60"

st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"] {{
    background:
      linear-gradient(rgba(12, 68, 124, 0.86), rgba(10, 58, 107, 0.90)),
      url('{IMG_FUNDO}') center / cover fixed no-repeat;
  }}
  [data-testid="stHeader"] {{ background: transparent; }}

  .block-container {{
    background: rgba(255, 255, 255, 0.13);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.32);
    border-radius: 18px;
    padding: 2.2rem 2.4rem 2.6rem;
    margin-top: 2rem;
    max-width: 780px;
  }}

  .block-container label, .block-container p, .block-container .stMarkdown,
  .block-container [data-testid="stWidgetLabel"] p {{ color: #EAF3FC !important; }}

  .header-si h1 {{ color: #ffffff; font-size: 1.5rem; margin: 0 0 2px; }}
  .header-si p {{ color: {AZUL_CLARO}; margin: 0 0 18px; font-size: 0.9rem; }}

  .stButton > button[kind="primary"] {{
    background: {AZUL}; border: 1px solid rgba(255,255,255,0.35); font-weight: 600;
  }}
  .stButton > button[kind="primary"] p {{ color: #ffffff !important; }}
  .stButton > button[kind="primary"]:hover {{ background: {AZUL_ESCURO}; }}
  .stButton > button[kind="secondary"] {{
    background: rgba(255, 255, 255, 0.14);
    border: 1px solid rgba(255, 255, 255, 0.4);
  }}
  .stButton > button[kind="secondary"] p {{ color: #ffffff !important; }}
  .stDownloadButton > button {{
    background: {AZUL}; border: 1px solid rgba(255,255,255,0.35); font-weight: 600;
  }}
  .stDownloadButton > button p {{ color: #ffffff !important; }}
  .stDownloadButton > button:hover {{ background: {AZUL_ESCURO}; }}

  div[data-testid="stMetric"] {{
    background: rgba(255, 255, 255, 0.15);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border: 1px solid rgba(255, 255, 255, 0.3);
    border-radius: 14px; padding: 12px 16px;
  }}
  div[data-testid="stMetric"] label {{ color: {AZUL_CLARO} !important; }}
  div[data-testid="stMetricValue"] {{ color: #ffffff; }}
  div[data-testid="stMetricDelta"] {{ color: {AZUL_CLARO}; }}
</style>
<div class="header-si">
  <h1>Validador de Leads <span style="font-size:0.8rem; background:rgba(255,255,255,0.25); border:1px solid rgba(255,255,255,0.4); color:#fff; padding:2px 10px; border-radius:20px; vertical-align:middle;">v3</span></h1>
  <p>Soluções Industriais</p>
</div>
""", unsafe_allow_html=True)

CAMPOS_FORM = ("f_chave", "f_obs")
if st.session_state.pop("limpar_form", False):
    for k in CAMPOS_FORM:
        st.session_state.pop(k, None)
    st.session_state.pop("resultado", None)

col_a, col_c, col_d = st.columns([1.2, 1, 1])
with col_a:
    chave_unica = st.text_input("Chave única do cliente", placeholder="Ex.: 12-34567-1", key="f_chave")
with col_c:
    data_inicio = st.date_input("Data início", value=date.today() - timedelta(days=90), format="DD/MM/YYYY")
with col_d:
    data_fim = st.date_input("Data fim", value=date.today(), format="DD/MM/YYYY")

obs = st.text_area(
    "Outras observações (opcional)",
    placeholder="Ex.: cliente só vende máquinas (serviço, assistência, aluguel e peças = fora do foco); lote mínimo 500 peças; atende só Sul e Sudeste.",
    height=90, key="f_obs",
)


def montar_regras():
    if obs.strip():
        return f"- Observações do projeto (prioridade máxima): {obs.strip()}"
    return ""


col_btn1, col_btn2 = st.columns([3, 1])
with col_btn2:
    if st.button("Limpar campos", use_container_width=True):
        st.session_state["limpar_form"] = True
        st.rerun()
with col_btn1:
    validar = st.button("Enviar", type="primary", use_container_width=True)

if validar:
    api_key = secret("GROQ_API_KEY")
    if not api_key:
        st.error("Segredo GROQ_API_KEY não configurado.")
        st.stop()
    if not secret("METABASE_URL"):
        st.error("Segredo METABASE_URL não configurado (ex.: https://metabase.ferramentademarketing.com.br).")
        st.stop()
    if not chave_unica.strip():
        st.error("Preencha a chave única do cliente.")
        st.stop()

    # 1. Briefing (question 286 — por chave única)
    with st.spinner("Buscando briefing no Metabase..."):
        try:
            csv_briefing = consultar_question(CARD_BRIEFING, [
                ("chave_unica", chave_unica.strip(), "category"),
            ])
        except Exception as e:
            st.error(f"Erro ao buscar o briefing (question {CARD_BRIEFING}): {e}")
            st.stop()
    texto_briefing = csv_briefing_para_texto(csv_briefing)
    if not texto_briefing:
        st.error(f'Briefing vazio para a chave "{chave_unica}". Confira se a chave está correta.')
        st.stop()

    # Nome da empresa extraído do próprio briefing (para arquivo e histórico)
    nome_empresa = chave_unica.strip()
    try:
        _linhas_b = list(csv.reader(io.StringIO(csv_briefing)))
        _h = [c.strip().lower() for c in _linhas_b[0]]
        for _cand in ("nome da empresa", "nome fantasia", "empresa"):
            if _cand in _h:
                _v = _linhas_b[1][_h.index(_cand)].strip()
                if _v:
                    nome_empresa = _v
                break
    except Exception:
        pass
    st.caption(f"Cliente identificado: {nome_empresa}")

    # 2. Orçamentos (question 47)
    with st.spinner("Buscando orçamentos no Metabase..."):
        try:
            csv_orcamentos = consultar_question(CARD_ORCAMENTOS, [
                ("chave_unica", chave_unica.strip(), "category"),
                ("data_inicio", data_inicio.isoformat(), "date/single"),
                ("data_fim", data_fim.isoformat(), "date/single"),
            ])
        except Exception as e:
            st.error(f"Erro ao buscar os orçamentos (question {CARD_ORCAMENTOS}): {e}")
            st.stop()
    linhas = list(csv.reader(io.StringIO(csv_orcamentos)))
    if len(linhas) < 2:
        st.error("Nenhum orçamento encontrado para essa chave única nesse período.")
        st.stop()

    # 3. Anúncios ativos (question 185 — opcional, não bloqueia se falhar)
    texto_anuncios = ""
    with st.spinner("Buscando anúncios do cliente no Metabase..."):
        try:
            csv_anuncios = consultar_question(CARD_ANUNCIOS, [
                ("chave_unica", chave_unica.strip(), "category"),
            ])
            texto_anuncios = csv_anuncios_para_texto(csv_anuncios)
        except Exception:
            st.warning("Não consegui buscar os anúncios (question 185) — prosseguindo sem eles.")

    # 4. Perfil do cliente
    perfil = f"===== BRIEFING DO CLIENTE (Metabase) =====\n{texto_briefing}"
    regras_projeto = montar_regras()
    if regras_projeto:
        perfil = f"===== OBSERVAÇÕES DO PROJETO (prioridade máxima) =====\n{regras_projeto}\n\n" + perfil
    if texto_anuncios:
        perfil += f"\n\n===== ANÚNCIOS ATIVOS DO CLIENTE (termos anunciados) =====\n{texto_anuncios}"
    if len(perfil) > LIMITE_PERFIL:
        perfil = perfil[:LIMITE_PERFIL] + "\n[... perfil truncado para caber no limite da IA gratuita ...]"

    # 5. Montagem dos leads
    cabecalho = [c.strip() for c in linhas[0]]
    col_msg = "Mensagem do Cliente"
    if col_msg not in cabecalho:
        st.error(f'Coluna "{col_msg}" não encontrada no retorno do Metabase. Colunas: {", ".join(cabecalho)}')
        st.stop()
    idx_msg = cabecalho.index(col_msg)
    registros = [r for r in linhas[1:] if any(c.strip() for c in r)]
    st.info(f"{len(registros)} leads encontrados de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}.")

    leads = []
    for i, r in enumerate(registros):
        extra = "; ".join(
            f"{cabecalho[j]}: {r[j]}" for j in range(len(cabecalho))
            if j != idx_msg and j < len(r) and r[j].strip()
        )[:500]
        leads.append({"id": f"L{i+2}", "mensagem": r[idx_msg] if idx_msg < len(r) else "", "extra": extra})

    # 6. Classificação com re-tentativas
    classificacoes = {}
    erros_ia = []

    def processar(lista, tamanho_lote, rotulo):
        total_lotes = (len(lista) + tamanho_lote - 1) // tamanho_lote
        progresso = st.progress(0, text=f"{rotulo}: {len(lista)} leads...")
        for n in range(total_lotes):
            lote = lista[n * tamanho_lote:(n + 1) * tamanho_lote]
            try:
                resultado = chamar_groq(api_key, perfil, lote)
            except RuntimeError as e:
                progresso.empty()
                st.error(str(e))       # cota diária esgotada: para tudo na hora
                st.stop()
            except Exception as e:
                erros_ia.append(f"{type(e).__name__}: {e}")
                resultado = []
            for item in resultado:
                status = str(item.get("status", "")).strip()
                if status not in STATUS_VALIDOS:
                    status = "Aberto"
                classificacoes[str(item.get("id", "")).strip()] = {
                    "status": status, "motivo": str(item.get("motivo", "")).strip(),
                }
            progresso.progress((n + 1) / total_lotes, text=f"{rotulo}: lote {n+1} de {total_lotes}")
            if n + 1 < total_lotes:
                time.sleep(1)
        progresso.empty()

    processar(leads, TAMANHO_LOTE, "Classificando")
    pendentes = [l for l in leads if l["id"] not in classificacoes]
    if pendentes:
        st.info(f"{len(pendentes)} lead(s) sem resposta — reprocessando em lotes menores...")
        time.sleep(5)
        processar(pendentes, 5, "Reprocessando")
    pendentes = [l for l in leads if l["id"] not in classificacoes]
    if pendentes:
        time.sleep(5)
        processar(pendentes, 1, "Última passada")
    falhas = sum(1 for l in leads if l["id"] not in classificacoes)

    # 7. CSV final + resumo
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
    periodo_txt = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
    base_nome = f"{nome_empresa} - {data_inicio.isoformat()} a {data_fim.isoformat()}"
    dash_html = gerar_dashboard_html(nome_empresa, chave_unica.strip(), periodo_txt, total, contagem)

    # Resultado fica guardado na sessão: os downloads não somem ao clicar
    st.session_state["resultado"] = {
        "empresa": nome_empresa,
        "total": total,
        "contagem": contagem,
        "falhas": falhas,
        "erro_ia": erros_ia[-1] if erros_ia else "",
        "csv_bytes": saida.getvalue().encode("utf-8-sig"),
        "csv_nome": f"{base_nome} - Validado.csv",
        "dash_bytes": dash_html.encode("utf-8"),
        "dash_nome": f"{base_nome} - Dashboard.html",
    }

    salvar_no_historico({
        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
        "Data da solicitação": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Empresa": nome_empresa,
        "Chave única": chave_unica.strip(),
        "Período": periodo_txt,
        "Leads": total,
        "Dentro do foco": contagem["Dentro do foco"],
        "Fora do foco": contagem["Fora do foco"],
        "Aberto": contagem["Aberto"],
        "csv_nome": f"{base_nome} - Validado.csv",
        "dash_nome": f"{base_nome} - Dashboard.html",
    }, csv_bytes=st.session_state["resultado"]["csv_bytes"],
       dash_bytes=st.session_state["resultado"]["dash_bytes"])

res = st.session_state.get("resultado")
if res:
    total, contagem = res["total"], res["contagem"]
    st.success(f"Validação concluída — {total} leads processados! ({res['empresa']})")
    c1, c2, c3 = st.columns(3)
    c1.metric("Dentro do foco", f"{contagem['Dentro do foco']/total:.0%}", f"{contagem['Dentro do foco']} leads", delta_color="off")
    c2.metric("Fora do foco", f"{contagem['Fora do foco']/total:.0%}", f"{contagem['Fora do foco']} leads", delta_color="off")
    c3.metric("Aberto", f"{contagem['Aberto']/total:.0%}", f"{contagem['Aberto']} leads", delta_color="off")
    if res["falhas"]:
        st.warning(f"{res['falhas']} lead(s) sem resposta da IA (marcados como Aberto) — rode de novo para reprocessar.")
        if res["erro_ia"]:
            st.error(f"Motivo técnico da falha na IA: {res['erro_ia'][:400]}")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button("Baixar CSV validado", data=res["csv_bytes"],
                           file_name=res["csv_nome"], mime="text/csv",
                           use_container_width=True, key="dl_csv")
    with col_dl2:
        st.download_button("Baixar dashboard (HTML)", data=res["dash_bytes"],
                           file_name=res["dash_nome"], mime="text/html",
                           use_container_width=True, key="dl_dash")

@st.dialog("Excluir pesquisa")
def dialogo_excluir(rid, rotulo):
    st.write(f"Tem certeza que deseja excluir a pesquisa **{rotulo}**?")
    st.caption("O registro e os arquivos (CSV e dashboard) dela serão apagados. Essa ação não tem volta.")
    cd1, cd2 = st.columns(2)
    if cd1.button("Sim, excluir", type="primary", use_container_width=True, key=f"conf_{rid}"):
        excluir_do_historico(rid)
        st.rerun()
    if cd2.button("Cancelar", use_container_width=True, key=f"canc_{rid}"):
        st.rerun()


st.markdown("<p style='font-weight:600; margin-top:32px;'>Histórico de validações</p>", unsafe_allow_html=True)
historico = carregar_historico()
if historico:
    filtro_chave = st.text_input(
        "Buscar por chave única",
        placeholder="Digite a chave para filtrar o histórico (ex.: 12-34567-1)",
        key="f_filtro_hist",
    )
    if filtro_chave.strip():
        alvo = filtro_chave.strip().lower()
        historico = [h for h in historico if alvo in str(h.get("Chave única", "")).lower()]
        if not historico:
            st.caption("Nenhuma validação encontrada para essa chave.")
    cab = st.columns([1.5, 1.6, 1.3, 1.9, 1.1, 0.4])
    for col, titulo in zip(cab, ("Data", "Empresa", "Chave", "Período", "Leads (D/F/A)", "")):
        col.markdown(f"<span style='font-size:0.72rem; color:#B5D4F4; font-weight:600;'>{titulo}</span>", unsafe_allow_html=True)

    for h in historico[:15]:
        rid = h.get("id", "")
        c = st.columns([1.5, 1.6, 1.3, 1.9, 1.1, 0.4])
        c[0].markdown(f"<span style='font-size:0.78rem;'>{h.get('Data da solicitação', '')}</span>", unsafe_allow_html=True)
        c[1].markdown(f"<span style='font-size:0.78rem;'>{h.get('Empresa', '')}</span>", unsafe_allow_html=True)
        c[2].markdown(f"<span style='font-size:0.78rem;'>{h.get('Chave única', '')}</span>", unsafe_allow_html=True)
        c[3].markdown(f"<span style='font-size:0.78rem;'>{h.get('Período', '')}</span>", unsafe_allow_html=True)
        c[4].markdown(
            f"<span style='font-size:0.78rem;'>{h.get('Leads', '')} "
            f"(<span style='color:#9FE1A5;'>{h.get('Dentro do foco', '')}</span>/"
            f"<span style='color:#F5A9A9;'>{h.get('Fora do foco', '')}</span>/"
            f"<span style='color:#FAD98F;'>{h.get('Aberto', '')}</span>)</span>",
            unsafe_allow_html=True,
        )
        if rid and c[5].button("✕", key=f"x_{rid}", help="Excluir esta pesquisa"):
            dialogo_excluir(rid, f"{h.get('Empresa', '')} · {h.get('Data da solicitação', '')}")
    com_id = [h for h in historico if h.get("id")]
    if com_id:
        st.markdown("<p style='font-weight:600; margin-top:16px;'>Baixar arquivos de uma validação</p>", unsafe_allow_html=True)
        rotulos = {
            f"{h['Data da solicitação']} · {h.get('Empresa', '?')} · {h.get('Leads', '?')} leads": h
            for h in com_id[:15]
        }
        escolha = st.selectbox("Selecione a validação", list(rotulos.keys()), label_visibility="collapsed")
        sel = rotulos[escolha]
        rid = sel["id"]
        csv_salvo = ler_resultado_salvo(rid, ".csv")
        dash_salvo = ler_resultado_salvo(rid, ".html")
        cg1, cg2 = st.columns(2)
        with cg1:
            if csv_salvo:
                st.download_button("Baixar CSV", data=csv_salvo,
                                   file_name=sel.get("csv_nome", f"{rid}.csv"),
                                   mime="text/csv", use_container_width=True, key=f"hcsv_{rid}")
            else:
                st.caption("CSV não disponível (app reiniciou)")
        with cg2:
            if dash_salvo:
                st.download_button("Baixar dashboard", data=dash_salvo,
                                   file_name=sel.get("dash_nome", f"{rid}.html"),
                                   mime="text/html", use_container_width=True, key=f"hdash_{rid}")
            else:
                st.caption("Dashboard não disponível (app reiniciou)")
else:
    st.caption("Nenhuma validação registrada ainda. As próximas aparecerão aqui com data, empresa e resultado.")

st.markdown(
    "<p style='text-align:center; color:#B5D4F4; font-size:0.75rem; margin-top:32px;'>"
    "Validador de Leads v3 · Soluções Industriais</p>",
    unsafe_allow_html=True,
)
