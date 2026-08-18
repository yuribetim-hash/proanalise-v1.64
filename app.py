import streamlit as st
import os
import json
import random
from io import BytesIO
from datetime import datetime, timedelta, timezone
from docxtpl import DocxTemplate, RichText
import hashlib
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import time
import re

# ============================================
# CONFIGURAÇÃO DE FUSO HORÁRIO (BRASÍLIA UTC-3)
# ============================================
BRASILIA_TZ = timezone(timedelta(hours=-3))

def agora_brasilia():
    return datetime.now(BRASILIA_TZ)

st.set_page_config(
    page_title="Proanalise v1.641",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# INICIALIZAÇÃO DAS VARIÁVEIS DE SESSÃO
# ============================================
if "tema_mode" not in st.session_state:
    st.session_state["tema_mode"] = "claro"
if "ultimo_backup" not in st.session_state:
    st.session_state["ultimo_backup"] = agora_brasilia()
if "marcadas_revisao" not in st.session_state:
    st.session_state["marcadas_revisao"] = set()
if "anotacoes_pessoais" not in st.session_state:
    st.session_state["anotacoes_pessoais"] = {}
if "analise_ativa" not in st.session_state:
    st.session_state["analise_ativa"] = False
if "analise_concluida" not in st.session_state:
    st.session_state["analise_concluida"] = False
if "tempo_inicio" not in st.session_state:
    st.session_state["tempo_inicio"] = None
if "tempo_fim" not in st.session_state:
    st.session_state["tempo_fim"] = None
if "tipo_analise" not in st.session_state:
    st.session_state["tipo_analise"] = "Aceite urbanístico"
if "perguntas_ativas" not in st.session_state:
    st.session_state["perguntas_ativas"] = []
if "botao_flutuante" not in st.session_state:
    st.session_state["botao_flutuante"] = False
if "dados_antigos" not in st.session_state:
    st.session_state["dados_antigos"] = None
if "etapa" not in st.session_state:
    st.session_state["etapa"] = "1. Protocolo"
if "protocolo" not in st.session_state:
    st.session_state["protocolo"] = ""
if "protocolo_anterior" not in st.session_state:
    st.session_state["protocolo_anterior"] = ""
if "tipo" not in st.session_state:
    st.session_state["tipo"] = "Loteamento"
if "interessado" not in st.session_state:
    st.session_state["interessado"] = ""
if "n_lotes" not in st.session_state:
    st.session_state["n_lotes"] = 1
if "matriculas" not in st.session_state:
    st.session_state["matriculas"] = ""
if "analista" not in st.session_state:
    st.session_state["analista"] = ""
if "matricula_analista" not in st.session_state:
    st.session_state["matricula_analista"] = ""
if "setor" not in st.session_state:
    st.session_state["setor"] = ""
if "n_analise" not in st.session_state:
    st.session_state["n_analise"] = ""
if "pendencias_manuais" not in st.session_state:
    st.session_state["pendencias_manuais"] = {}
if "respostas_analise" not in st.session_state:
    st.session_state["respostas_analise"] = {}
if "observacoes_analise" not in st.session_state:
    st.session_state["observacoes_analise"] = {}
if "pendencias_analise" not in st.session_state:
    st.session_state["pendencias_analise"] = {}
if "analise_estado" not in st.session_state:
    st.session_state["analise_estado"] = "em_andamento"
if "analista_responsavel" not in st.session_state:
    st.session_state["analista_responsavel"] = ""
if "analista_revisor" not in st.session_state:
    st.session_state["analista_revisor"] = ""
if "modo_revisao" not in st.session_state:
    st.session_state["modo_revisao"] = False

HASH_SALT = "proanalise_salt_2024"

# ============================================
# FUNÇÕES AUXILIARES
# ============================================
def serialize_datetime(dt):
    if dt is None:
        return None
    if isinstance(dt, datetime):
        return dt.isoformat()
    return dt

def deserialize_datetime(dt_str):
    if dt_str is None:
        return None
    try:
        dt = datetime.fromisoformat(dt_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=BRASILIA_TZ)
        return dt
    except:
        return None

def dados_analise_persistente():
    """Retorna os dados da análise que devem sobreviver ao rerun/sessão."""
    return {
        "protocolo": st.session_state.get("protocolo", ""),
        "tipo": st.session_state.get("tipo", ""),
        "tipo_analise": st.session_state.get("tipo_analise", ""),
        "interessado": st.session_state.get("interessado", ""),
        "n_lotes": st.session_state.get("n_lotes", 1),
        "matriculas": st.session_state.get("matriculas", ""),
        "analista": st.session_state.get("analista", ""),
        "matricula_analista": st.session_state.get("matricula_analista", ""),
        "setor": st.session_state.get("setor", ""),
        "n_analise": st.session_state.get("n_analise", ""),
        "respostas_analise": dict(st.session_state.get("respostas_analise", {})),
        "observacoes_analise": dict(st.session_state.get("observacoes_analise", {})),
        "pendencias_analise": dict(st.session_state.get("pendencias_analise", {})),
        "analise_ativa": st.session_state.get("analise_ativa", False),
        "analise_concluida": st.session_state.get("analise_concluida", False),
        "analise_estado": st.session_state.get("analise_estado", "em_andamento"),
        "analista_responsavel": st.session_state.get("analista_responsavel", ""),
        "analista_revisor": st.session_state.get("analista_revisor", ""),
        "tempo_inicio": serialize_datetime(st.session_state.get("tempo_inicio")),
        "tempo_fim": serialize_datetime(st.session_state.get("tempo_fim")),
        "etapa": st.session_state.get("etapa", ""),
        "marcadas_revisao": list(st.session_state.get("marcadas_revisao", set())),
        "anotacoes_pessoais": st.session_state.get("anotacoes_pessoais", {}),
        "data_atualizacao": agora_brasilia().isoformat()
    }

def salvar_estado_analise_persistente():
    """Grava o estado atual em arquivo próprio do protocolo."""
    protocolo = st.session_state.get("protocolo", "").strip()
    if not protocolo:
        return None
    pasta = get_pasta_protocolo(protocolo)
    os.makedirs(pasta, exist_ok=True)
    arquivo = os.path.join(pasta, "estado_analise.json")
    dados = dados_analise_persistente()
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)
    return arquivo

def carregar_estado_analise_persistente(protocolo):
    protocolo = str(protocolo or "").strip()
    if not protocolo:
        return None
    arquivo = os.path.join(get_pasta_protocolo(protocolo), "estado_analise.json")
    if not os.path.exists(arquivo):
        return None
    try:
        with open(arquivo, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None

def aplicar_estado_analise(dados, etapa=None):
    """Carrega uma análise persistida para a sessão atual."""
    if not dados:
        return False
    chaves = [
        "protocolo", "tipo", "tipo_analise", "interessado", "n_lotes",
        "matriculas", "analista", "matricula_analista", "setor", "n_analise",
        "respostas_analise", "observacoes_analise", "pendencias_analise",
        "analise_ativa", "analise_concluida", "analise_estado",
        "analista_responsavel", "analista_revisor", "marcadas_revisao",
        "anotacoes_pessoais"
    ]
    for chave in chaves:
        if chave in dados:
            st.session_state[chave] = dados[chave]
    st.session_state["tempo_inicio"] = deserialize_datetime(dados.get("tempo_inicio"))
    st.session_state["tempo_fim"] = deserialize_datetime(dados.get("tempo_fim"))
    st.session_state["marcadas_revisao"] = set(dados.get("marcadas_revisao", []))
    if etapa:
        st.session_state["etapa"] = etapa
    return True

def listar_protocolos_aguardando_revisao():
    """Lista protocolos persistidos como aguardando revisão.

    O arquivo estado_analise.json é a fonte principal. Backups antigos são
    mantidos como fallback para não perder protocolos já existentes.
    """
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return []
    encontrados = {}

    for nome in os.listdir(pasta_dados):
        pasta = os.path.join(pasta_dados, nome)
        if not os.path.isdir(pasta):
            continue
        arquivo = os.path.join(pasta, "estado_analise.json")
        if not os.path.exists(arquivo):
            continue
        try:
            with open(arquivo, "r", encoding="utf-8") as f:
                dados = json.load(f)
            if dados.get("analise_estado") == "aguardando_revisao":
                dt = deserialize_datetime(dados.get("data_atualizacao"))
                data_formatada = dt.strftime("%d/%m/%Y %H:%M") if dt else "Data desconhecida"
                protocolo = dados.get("protocolo", nome.replace("-", "/"))
                encontrados[protocolo] = {
                    "protocolo": protocolo,
                    "analista_responsavel": dados.get("analista_responsavel") or dados.get("analista", "Não informado"),
                    "data_backup": data_formatada,
                    "caminho_backup": arquivo,
                    "backup_data": dados
                }
        except (OSError, json.JSONDecodeError):
            continue

    # Fallback para backups antigos que ainda não possuem estado_analise.json.
    pasta_backup = os.path.join(pasta_dados, "backups")
    if os.path.exists(pasta_backup):
        for arquivo in os.listdir(pasta_backup):
            if not arquivo.startswith("backup_") or not arquivo.endswith(".json"):
                continue
            caminho = os.path.join(pasta_backup, arquivo)
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                if dados.get("analise_estado") != "aguardando_revisao":
                    continue
                protocolo = dados.get("protocolo", "Desconhecido")
                if protocolo in encontrados:
                    continue
                data_str = arquivo.replace("backup_", "").replace(".json", "")
                try:
                    data_obj = datetime.strptime(data_str, "%Y%m%d_%H%M%S")
                    data_formatada = data_obj.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    data_formatada = "Data desconhecida"
                encontrados[protocolo] = {
                    "protocolo": protocolo,
                    "analista_responsavel": dados.get("analista_responsavel") or dados.get("analista", "Não informado"),
                    "data_backup": data_formatada,
                    "caminho_backup": caminho,
                    "backup_data": dados
                }
            except (OSError, json.JSONDecodeError):
                continue

    return sorted(encontrados.values(), key=lambda x: x["data_backup"], reverse=True)

# ============================================
# FUNÇÕES DE BACKUP
# ============================================
def fazer_backup_automatico():
    if not st.session_state.get("analise_ativa", False):
        return False
    agora = agora_brasilia()
    diff = (agora - st.session_state["ultimo_backup"]).total_seconds()
    if diff >= 300:
        pasta_backup = os.path.join("dados", "backups")
        os.makedirs(pasta_backup, exist_ok=True)
        timestamp = agora.strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(pasta_backup, f"backup_{timestamp}.json")
        dados_backup = {
            "protocolo": st.session_state.get("protocolo", ""),
            "tipo": st.session_state.get("tipo", ""),
            "tipo_analise": st.session_state.get("tipo_analise", ""),
            "interessado": st.session_state.get("interessado", ""),
            "n_lotes": st.session_state.get("n_lotes", 1),
            "matriculas": st.session_state.get("matriculas", ""),
            "analista": st.session_state.get("analista", ""),
            "matricula_analista": st.session_state.get("matricula_analista", ""),
            "setor": st.session_state.get("setor", ""),
            "n_analise": st.session_state.get("n_analise", ""),
            "respostas_analise": st.session_state.get("respostas_analise", {}),
            "observacoes_analise": st.session_state.get("observacoes_analise", {}),
            "pendencias_analise": st.session_state.get("pendencias_analise", {}),
            "analise_ativa": st.session_state.get("analise_ativa", False),
            "analise_concluida": st.session_state.get("analise_concluida", False),
            "analise_estado": st.session_state.get("analise_estado", "em_andamento"),
            "analista_responsavel": st.session_state.get("analista_responsavel", ""),
            "analista_revisor": st.session_state.get("analista_revisor", ""),
            "tempo_inicio": serialize_datetime(st.session_state.get("tempo_inicio")),
            "tempo_fim": serialize_datetime(st.session_state.get("tempo_fim")),
            "etapa": st.session_state.get("etapa", ""),
            "marcadas_revisao": list(st.session_state.get("marcadas_revisao", set())),
            "anotacoes_pessoais": st.session_state.get("anotacoes_pessoais", {}),
            "data_backup": timestamp
        }
        with open(backup_file, "w", encoding="utf-8") as f:
            json.dump(dados_backup, f, indent=4, ensure_ascii=False)
        backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")])
        if len(backups) > 10:
            for old_backup in backups[:-10]:
                os.remove(os.path.join(pasta_backup, old_backup))
        st.session_state["ultimo_backup"] = agora
        return True
    return False

def salvar_backup_manual():
    pasta_backup = os.path.join("dados", "backups")
    os.makedirs(pasta_backup, exist_ok=True)
    agora = agora_brasilia()
    timestamp = agora.strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(pasta_backup, f"backup_{timestamp}.json")
    dados_backup = {
        "protocolo": st.session_state.get("protocolo", ""),
        "tipo": st.session_state.get("tipo", ""),
        "tipo_analise": st.session_state.get("tipo_analise", ""),
        "interessado": st.session_state.get("interessado", ""),
        "n_lotes": st.session_state.get("n_lotes", 1),
        "matriculas": st.session_state.get("matriculas", ""),
        "analista": st.session_state.get("analista", ""),
        "matricula_analista": st.session_state.get("matricula_analista", ""),
        "setor": st.session_state.get("setor", ""),
        "n_analise": st.session_state.get("n_analise", ""),
        "respostas_analise": st.session_state.get("respostas_analise", {}),
        "observacoes_analise": st.session_state.get("observacoes_analise", {}),
        "pendencias_analise": st.session_state.get("pendencias_analise", {}),
        "analise_ativa": st.session_state.get("analise_ativa", False),
        "analise_concluida": st.session_state.get("analise_concluida", False),
        "analise_estado": st.session_state.get("analise_estado", "em_andamento"),
        "analista_responsavel": st.session_state.get("analista_responsavel", ""),
        "analista_revisor": st.session_state.get("analista_revisor", ""),
        "tempo_inicio": serialize_datetime(st.session_state.get("tempo_inicio")),
        "tempo_fim": serialize_datetime(st.session_state.get("tempo_fim")),
        "etapa": st.session_state.get("etapa", ""),
        "marcadas_revisao": list(st.session_state.get("marcadas_revisao", set())),
        "anotacoes_pessoais": st.session_state.get("anotacoes_pessoais", {}),
        "data_backup": timestamp
    }
    with open(backup_file, "w", encoding="utf-8") as f:
        json.dump(dados_backup, f, indent=4, ensure_ascii=False)
    backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")])
    if len(backups) > 10:
        for old_backup in backups[:-10]:
            os.remove(os.path.join(pasta_backup, old_backup))
    return backup_file

def restaurar_analise():
    pasta_backup = os.path.join("dados", "backups")
    if not os.path.exists(pasta_backup):
        return None
    backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")], reverse=True)
    if not backups:
        return None
    opcoes = {}
    for b in backups[:10]:
        data_str = b.replace("backup_", "").replace(".json", "")
        data_obj = datetime.strptime(data_str, "%Y%m%d_%H%M%S")
        opcoes[b] = data_obj.strftime("%d/%m/%Y %H:%M:%S")
    return opcoes

# ============================================
# FUNÇÕES DE PERMISSÕES
# ============================================
def carregar_usuarios(caminho="usuarios.txt"):
    if not os.path.exists(caminho):
        st.error("Arquivo usuarios.txt não encontrado.")
        st.stop()
    usuarios = {}
    with open(caminho, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            partes = linha.split(";")
            if len(partes) >= 4:
                usuario, senha, papel, nivel = partes[0], partes[1], partes[2], partes[3]
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": papel.strip(),
                    "nivel": int(nivel.strip())
                }
            elif len(partes) == 3:
                usuario, senha, papel = partes
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": papel.strip(),
                    "nivel": 2 if "Sênior" in papel else 1
                }
            else:
                usuario, senha = partes[0], partes[1]
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": "Analista",
                    "nivel": 2
                }
    return usuarios

def tem_permissao(nivel_necessario):
    if "usuario_info" not in st.session_state:
        return False
    return st.session_state["usuario_info"]["nivel"] >= nivel_necessario

def pode_ver_menu(menu_item):
    niveis_necessarios = {
        "1. Protocolo": 1,
        "2. Analista": 1,
        "3. Análise": 1,
        "4. Revisão": 1,
        "5. Gerar parecer": 1,
        "6. Dashboard": 2,
        "7. Comparador": 3
    }
    return tem_permissao(niveis_necessarios.get(menu_item, 1))

# ============================================
# CARREGAMENTO DE PERGUNTAS
# ============================================
def carregar_perguntas_por_tipo(tipo_empreendimento, tipo_analise):
    if tipo_empreendimento == "Desmembramentos":
        mapa_arquivos = {
            "Retificação": "retificacao.txt",
            "Desdobro de áreas": "desdobro.txt",
            "Unificação de áreas": "unificacao.txt",
            "Alteração de divisas": "alteracao_divisas.txt",
            "Estremação": "estremacao.txt",
            "Usucapião": "usucapiao.txt",
            "Retificação com desdobro": "retificacao_desdobro.txt",
            "Retificação com alteração de divisas": "retificacao_alteracao_divisas.txt",
            "Retificação com unificação": "retificacao_unificacao.txt",
            "Unificação com desdobro": "unificacao_desdobro.txt",
            "Desdobro para anexação": "desdobro_anexacao.txt",
            "Anuência rural": "anuencia_rural.txt"
        }
        nome_arquivo = mapa_arquivos.get(tipo_analise, "retificacao.txt")
        for caminho in [os.path.join("perguntas", nome_arquivo), nome_arquivo]:
            if os.path.exists(caminho):
                return carregar_perguntas_txt(caminho)
        if os.path.exists("perguntas/desmembramentos.txt"):
            return carregar_perguntas_txt("perguntas/desmembramentos.txt")
        elif os.path.exists("desmembramentos.txt"):
            return carregar_perguntas_txt("desmembramentos.txt")
        if os.path.exists("perguntas.txt"):
            return carregar_perguntas_txt("perguntas.txt")
        st.error(f"Arquivo de perguntas para '{tipo_analise}' não encontrado.")
        return []
    if tipo_analise == "Aceite urbanístico":
        if tipo_empreendimento == "Loteamento":
            arquivo = "perguntas/loteamento_aceite.txt"
        else:
            arquivo = "perguntas/condominio_aceite.txt"
    else:  # Alvará
        if tipo_empreendimento == "Loteamento":
            arquivo = "perguntas/loteamento_alvara.txt"
        else:
            arquivo = "perguntas/condominio_alvara.txt"
    if os.path.exists(arquivo):
        return carregar_perguntas_txt(arquivo)
    if os.path.exists("perguntas.txt"):
        return carregar_perguntas_txt("perguntas.txt")
    st.error(f"Arquivo de perguntas não encontrado: {arquivo}")
    return []

def carregar_perguntas_txt(caminho):
    perguntas = []
    bloco = {}
    if not os.path.exists(caminho):
        return []
    with open(caminho, "r", encoding="utf-8") as f:
        linhas = f.readlines()
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            if bloco:
                perguntas.append(bloco)
                bloco = {}
            continue
        if linha.startswith("GRUPO:"):
            bloco["grupo"] = linha.replace("GRUPO:", "").strip()
        elif linha.startswith("ID:"):
            bloco["id"] = linha.replace("ID:", "").strip()
        elif linha.startswith("PERGUNTA:"):
            bloco["pergunta"] = linha.replace("PERGUNTA:", "").strip()
        elif linha.startswith("OPCOES:"):
            bloco["opcoes"] = [op.strip() for op in linha.replace("OPCOES:", "").strip().split(";")]
        elif linha.startswith("CONFORMES:"):
            bloco["conformes"] = [op.strip() for op in linha.replace("CONFORMES:", "").strip().split(";")]
        elif linha.startswith("REGRA_"):
            chave, valor = linha.split(":", 1)
            resposta = chave.replace("REGRA_", "").strip()
            bloco.setdefault("regras", {})[resposta] = {"texto": valor.strip()}
    if bloco:
        perguntas.append(bloco)
    return perguntas

# ============================================
# FUNÇÕES DE MULTIPLOS ANALISTAS
# ============================================
def get_pasta_protocolo(protocolo):
    protocolo_limpo = protocolo.replace("/", "-").strip()
    return os.path.join("dados", protocolo_limpo)

def salvar_analise_analista(protocolo, analista, papel, respostas, observacoes, pendencias):
    pasta = get_pasta_protocolo(protocolo)
    os.makedirs(pasta, exist_ok=True)
    analista_hash = hashlib.md5(f"{analista}_{papel}_{HASH_SALT}".encode()).hexdigest()[:8]
    arquivo = os.path.join(pasta, f"analise_{analista_hash}_{papel}.json")
    registro = {
        "protocolo": protocolo,
        "analista": analista,
        "papel": papel,
        "data": agora_brasilia().strftime("%d/%m/%Y %H:%M:%S"),
        "respostas": respostas,
        "observacoes": observacoes,
        "pendencias": pendencias,
        "hash": analista_hash
    }
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(registro, f, indent=4, ensure_ascii=False)
    return analista_hash

def carregar_analises_analistas(protocolo):
    pasta = get_pasta_protocolo(protocolo)
    if not os.path.exists(pasta):
        return []
    analises = []
    for arquivo in os.listdir(pasta):
        if arquivo.startswith("analise_") and arquivo.endswith(".json"):
            with open(os.path.join(pasta, arquivo), "r", encoding="utf-8") as f:
                analises.append(json.load(f))
    return analises

def comparar_analises(analises):
    if len(analises) < 2:
        return None
    resultado = {
        "total_diferencas": 0,
        "diferencas_por_pergunta": {},
        "analistas": [a["analista"] for a in analises],
        "papeis": [a["papel"] for a in analises],
        "data_analises": [a["data"] for a in analises]
    }
    todas_perguntas = set()
    for analise in analises:
        todas_perguntas.update(analise["respostas"].keys())
    for pergunta in todas_perguntas:
        respostas_analistas = {}
        for analise in analises:
            resp = analise["respostas"].get(pergunta, "Não respondida")
            respostas_analistas[analise["analista"]] = resp
        valores_unicos = set(respostas_analistas.values())
        if len(valores_unicos) > 1:
            resultado["diferencas_por_pergunta"][pergunta] = respostas_analistas
            resultado["total_diferencas"] += 1
    return resultado

def render_comparador_analises(protocolo_atual):
    if not tem_permissao(3):
        st.error("❌ Acesso negado! Apenas Analistas Responsáveis podem acessar o Comparador.")
        return
    st.subheader("🔍 Comparador de Análises")
    analises = carregar_analises_analistas(protocolo_atual)
    if not analises:
        st.info("Nenhuma análise de outro analista encontrada para este protocolo.")
        return
    st.write(f"**Total de análises encontradas:** {len(analises)}")
    for a in analises:
        st.write(f"- {a['analista']} ({a['papel']}) - {a['data']}")
    if st.button("Comparar Análises", use_container_width=True):
        comparacao = comparar_analises(analises)
        if comparacao and comparacao["total_diferencas"] > 0:
            st.warning(f"⚠️ **{comparacao['total_diferencas']} divergências encontradas**")
            for pergunta, respostas in comparacao["diferencas_por_pergunta"].items():
                with st.expander(f"📌 Pergunta ID: {pergunta}"):
                    for analista, resposta in respostas.items():
                        st.write(f"**{analista}:** {resposta}")
        else:
            st.success("✅ Todas as análises estão consistentes!")

# ============================================
# FUNÇÕES DE MÉTRICAS E BUSCA
# ============================================
def calcular_tempo_medio_analise():
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return None
    tempos = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
            if len(analises) >= 2:
                analises.sort()
                primeira = analises[0]
                ultima = analises[-1]
                with open(os.path.join(protocolo_path, primeira), "r", encoding="utf-8") as f:
                    data_primeira = datetime.strptime(json.load(f)["data"], "%d/%m/%Y")
                with open(os.path.join(protocolo_path, ultima), "r", encoding="utf-8") as f:
                    data_ultima = datetime.strptime(json.load(f)["data"], "%d/%m/%Y")
                tempo = (data_ultima - data_primeira).days
                tempos.append(tempo)
    if tempos:
        return sum(tempos) / len(tempos)
    return None

def gerar_grafico_inconformidades(respostas, grupos_inconformes):
    if not grupos_inconformes:
        return None
    dados = []
    for grupo, itens in grupos_inconformes.items():
        dados.append({"Grupo": grupo, "Inconformidades": len(itens)})
    df = pd.DataFrame(dados)
    fig = px.bar(df, x="Grupo", y="Inconformidades", 
                 title="Inconformidades por Grupo",
                 color="Inconformidades",
                 color_continuous_scale="Reds")
    fig.update_layout(
        xaxis_title="Grupo",
        yaxis_title="Número de Inconformidades",
        showlegend=False,
        height=400
    )
    return fig

def gerar_grafico_tempo_analises():
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return None
    dados_tempo = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
            if analises:
                with open(os.path.join(protocolo_path, analises[0]), "r", encoding="utf-8") as f:
                    primeira = json.load(f)
                with open(os.path.join(protocolo_path, analises[-1]), "r", encoding="utf-8") as f:
                    ultima = json.load(f)
                data_inicio = datetime.strptime(primeira["data"], "%d/%m/%Y")
                data_fim = datetime.strptime(ultima["data"], "%d/%m/%Y")
                dias = (data_fim - data_inicio).days
                dados_tempo.append({
                    "Protocolo": protocolo_dir.replace("-", "/"),
                    "Dias de Análise": dias,
                    "Conclusão": ultima.get("conclusao", "Em análise")
                })
    if dados_tempo:
        df = pd.DataFrame(dados_tempo)
        fig = px.bar(df, x="Protocolo", y="Dias de Análise", 
                     title="Tempo de Análise por Protocolo",
                     color="Conclusão",
                     color_discrete_map={"FAVORÁVEL": "green", "DESFAVORÁVEL": "red"})
        fig.update_layout(xaxis_tickangle=-45, height=400)
        return fig
    return None

def buscar_protocolos(termo_busca, filtro_status=None, filtro_analista=None, data_inicio=None, data_fim=None):
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return []
    resultados = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            if termo_busca.lower() in protocolo_dir.lower():
                status = "Em análise"
                analista = "Não informado"
                data_analise = datetime.fromtimestamp(os.path.getmtime(protocolo_path))
                analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
                if analises:
                    ultima = sorted(analises)[-1]
                    with open(os.path.join(protocolo_path, ultima), "r", encoding="utf-8") as f:
                        dados = json.load(f)
                        status = dados.get("conclusao", "Em análise")
                        analista = dados.get("analista", "Não informado")
                if filtro_status and filtro_status != "Todos":
                    if status != filtro_status:
                        continue
                if filtro_analista and filtro_analista != "Todos":
                    if analista != filtro_analista:
                        continue
                if data_inicio:
                    if data_analise.date() < data_inicio:
                        continue
                if data_fim:
                    if data_analise.date() > data_fim:
                        continue
                resultados.append({
                    "protocolo": protocolo_dir.replace("-", "/"),
                    "status": status,
                    "analista": analista,
                    "data_ultima": data_analise.strftime("%d/%m/%Y"),
                    "qtd_analises": len(analises)
                })
    return resultados

def proxima_pergunta_nao_respondida(respostas, perguntas):
    for idx, p in enumerate(perguntas):
        resposta = respostas.get(p["id"])
        if resposta in ("", None, "Selecione..."):
            return p["id"], idx
    return None, None

def resposta_preenchida(valor):
    return valor not in ("", None, "Selecione...")

# ============================================
# TEMA E CSS
# ============================================
def carregar_tema():
    if "tema_mode" not in st.session_state:
        st.session_state["tema_mode"] = "claro"
    tema_claro = {
        "cores": {
            "primaria": "#0a2a3a",
            "primaria_clara": "#1a5276",
            "primaria_muito_escura": "#051a24",
            "secundaria": "#2c6b96",
            "sucesso": "#0d6e2e",
            "erro": "#b42318",
            "alerta": "#b54708",
            "info": "#175cd3",
            "fundo_claro": "#f8fafd",
            "fundo_branco": "#ffffff",
            "texto_principal": "#1a1a1a",
            "texto_secundario": "#2c3e50",
            "texto_caption": "#666666",
            "texto_titulo": "#051a24",
            "fundo_app": "linear-gradient(135deg, #e8f0fe 0%, #d4e4fc 100%)",
            "sidebar_fundo": "linear-gradient(180deg, #0a2a3a 0%, #051a24 100%)"
        },
        "botoes": {
            "primario_fundo": "#0a5c2a",
            "primario_fundo_hover": "#0d6e2e",
            "secundario_fundo": "#0a2a3a",
            "secundario_fundo_hover": "#1a5276",
            "texto": "#ffffff"
        },
        "status": {
            "conforme": {"fundo": "#ecfdf3", "borda": "#067647", "texto": "#067647", "icone": "✅"},
            "inconforme": {"fundo": "#fef3f2", "borda": "#b42318", "texto": "#7a271a", "icone": "⛔"},
            "pendente": {"fundo": "#fffaeb", "borda": "#b54708", "texto": "#7a4a0a", "icone": "⏳"},
            "nao_se_enquadra": {"fundo": "#eff6ff", "borda": "#175cd3", "texto": "#0e4a8a", "icone": "ℹ️"}
        }
    }
    tema_escuro = {
        "cores": {
            "primaria": "#e8f0fe",
            "primaria_clara": "#d4e4fc",
            "primaria_muito_escura": "#c5d5e6",
            "secundaria": "#2c6b96",
            "sucesso": "#0f8a3a",
            "erro": "#f5c2c7",
            "alerta": "#fedf89",
            "info": "#c7d7fe",
            "fundo_claro": "#1e1e2e",
            "fundo_branco": "#2a2a3e",
            "texto_principal": "#e0e0e0",
            "texto_secundario": "#b0b0b0",
            "texto_caption": "#888888",
            "texto_titulo": "#ffffff",
            "fundo_app": "linear-gradient(135deg, #1a1a2e 0%, #0a0a15 100%)",
            "sidebar_fundo": "linear-gradient(180deg, #0a0a15 0%, #05050a 100%)"
        },
        "botoes": {
            "primario_fundo": "#0f8a3a",
            "primario_fundo_hover": "#0d6e2e",
            "secundario_fundo": "#2c6b96",
            "secundario_fundo_hover": "#1a5276",
            "texto": "#ffffff"
        },
        "status": {
            "conforme": {"fundo": "#1a3a2a", "borda": "#0f8a3a", "texto": "#90ee90", "icone": "✅"},
            "inconforme": {"fundo": "#3a1a1a", "borda": "#f5c2c7", "texto": "#f5c2c7", "icone": "⛔"},
            "pendente": {"fundo": "#3a2a1a", "borda": "#fedf89", "texto": "#fedf89", "icone": "⏳"},
            "nao_se_enquadra": {"fundo": "#1a2a3a", "borda": "#c7d7fe", "texto": "#c7d7fe", "icone": "ℹ️"}
        }
    }
    if st.session_state["tema_mode"] == "escuro":
        return tema_escuro
    return tema_claro

tema = carregar_tema()
css_tema = f"""
<style>
    .stApp {{ background: {tema["cores"]["fundo_app"]}; }}
    .main > div {{ background-color: {tema["cores"]["fundo_branco"]}; border-radius: 12px; padding: 1rem; }}
    [data-testid="stSidebar"] {{ background: {tema["cores"]["sidebar_fundo"]}; }}
    [data-testid="stSidebar"] * {{ color: {tema["botoes"]["texto"]} !important; }}
    h1 {{ color: {tema["cores"]["texto_titulo"]} !important; font-weight: 700 !important; font-size: 24px !important; }}
    h2, h3, h4 {{ color: {tema["cores"]["primaria"]} !important; font-weight: 600 !important; }}
    .stCaption {{ color: {tema["cores"]["texto_caption"]} !important; font-size: 14px !important; }}
    .stTextInput input, .stTextArea textarea, .stNumberInput input {{
        background-color: {tema["cores"]["fundo_branco"]} !important;
        color: {tema["cores"]["texto_principal"]} !important;
        border: 1px solid {tema["cores"]["primaria_clara"]} !important;
        border-radius: 6px !important;
        caret-color: {tema["cores"]["primaria"]} !important;
    }}
    .stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus {{
        border-color: {tema["cores"]["primaria"]} !important;
        box-shadow: 0 0 0 2px rgba(26, 82, 118, 0.2) !important;
        outline: none !important;
    }}
    .stTextInput label, .stSelectbox label, .stTextArea label, .stNumberInput label {{
        color: {tema["cores"]["primaria_clara"]} !important;
        font-weight: 600 !important;
    }}
    .stSelectbox select {{
        background-color: {tema["cores"]["fundo_branco"]} !important;
        color: {tema["cores"]["texto_principal"]} !important;
        border: 1px solid {tema["cores"]["primaria_clara"]} !important;
        border-radius: 6px !important;
        font-size: 14px !important;
    }}
    div[data-baseweb="menu"] {{
        background-color: #1a1a2e !important;
        border: 1px solid #2c6b96 !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
    }}
    div[data-baseweb="menu"] div {{
        background-color: #1a1a2e !important;
        color: white !important;
        padding: 8px 16px !important;
        font-size: 14px !important;
        cursor: pointer !important;
    }}
    div[data-baseweb="menu"] div:hover {{
        background-color: #2c6b96 !important;
        color: white !important;
    }}
    div[data-baseweb="menu"] div[aria-selected="true"] {{
        background-color: #0d6e2e !important;
        color: white !important;
    }}
    .stButton > button {{
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }}
    .stButton > button[kind="primary"] {{
        background-color: {tema["botoes"]["primario_fundo"]} !important;
        color: white !important;
    }}
    .stButton > button[kind="primary"]:hover {{
        background-color: {tema["botoes"]["primario_fundo_hover"]} !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        color: white !important;
    }}
    .stButton > button:not([kind="primary"]) {{
        background-color: {tema["botoes"]["secundario_fundo"]} !important;
        color: white !important;
    }}
    .stButton > button:not([kind="primary"]):hover {{
        background-color: {tema["botoes"]["secundario_fundo_hover"]} !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        color: white !important;
    }}
    .stButton > button:disabled {{
        background-color: #cccccc !important;
        color: #666666 !important;
    }}
    .status-badge-conforme {{
        background-color: {tema["status"]["conforme"]["fundo"]};
        border-left: 4px solid {tema["status"]["conforme"]["borda"]};
        padding: 8px 12px;
        border-radius: 6px;
        margin-top: 28px;
        color: {tema["status"]["conforme"]["texto"]};
        font-weight: 600;
    }}
    .status-badge-inconforme {{
        background-color: {tema["status"]["inconforme"]["fundo"]};
        border-left: 4px solid {tema["status"]["inconforme"]["borda"]};
        padding: 8px 12px;
        border-radius: 6px;
        margin-top: 28px;
        color: {tema["status"]["inconforme"]["texto"]};
        font-weight: 600;
    }}
    .status-badge-pendente {{
        background-color: {tema["status"]["pendente"]["fundo"]};
        border-left: 4px solid {tema["status"]["pendente"]["borda"]};
        padding: 8px 12px;
        border-radius: 6px;
        margin-top: 28px;
        color: {tema["status"]["pendente"]["texto"]};
        font-weight: 600;
    }}
    .status-badge-na {{
        background-color: {tema["status"]["nao_se_enquadra"]["fundo"]};
        border-left: 4px solid {tema["status"]["nao_se_enquadra"]["borda"]};
        padding: 8px 12px;
        border-radius: 6px;
        margin-top: 28px;
        color: {tema["status"]["nao_se_enquadra"]["texto"]};
        font-weight: 600;
    }}
    .progress-wrap {{
        width: 100%;
        background: #e9ecef;
        border-radius: 999px;
        height: 14px;
        overflow: hidden;
        margin: 8px 0;
    }}
    .progress-bar {{
        height: 14px;
        border-radius: 999px;
        transition: width 0.3s ease;
    }}
    .card {{
        padding: 12px 16px;
        border: 1px solid #c5d5e6;
        border-radius: 10px;
        background: {tema["cores"]["fundo_claro"]};
        margin-bottom: 0.6rem;
        color: {tema["cores"]["texto_principal"]};
    }}
    [data-testid="stMetric"] {{
        background-color: {tema["cores"]["fundo_claro"]};
        border-radius: 8px;
        padding: 10px;
        border: 1px solid #c5d5e6;
    }}
    [data-testid="stMetric"] label {{
        color: {tema["cores"]["primaria"]} !important;
        font-weight: 600 !important;
    }}
    [data-testid="stMetric"] .stMetricValue {{
        color: {tema["cores"]["texto_principal"]} !important;
        font-weight: 700 !important;
    }}
    .streamlit-expanderHeader {{
        background-color: {tema["cores"]["fundo_claro"]} !important;
        color: {tema["cores"]["primaria"]} !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }}
    .streamlit-expanderContent {{
        background-color: {tema["cores"]["fundo_branco"]} !important;
        border-radius: 0 0 8px 8px !important;
    }}
    .card-revisao {{
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 8px 12px;
        border-radius: 6px;
        margin: 5px 0;
    }}
    .floating-btn {{
        position: fixed;
        bottom: 20px;
        right: 20px;
        z-index: 999;
        border-radius: 50px !important;
        padding: 10px 20px !important;
        background-color: {tema["botoes"]["primario_fundo"]} !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
        animation: pulse 2s infinite;
    }}
    @keyframes pulse {{
        0% {{ transform: scale(1); }}
        50% {{ transform: scale(1.05); }}
        100% {{ transform: scale(1); }}
    }}
    .floating-btn:hover {{
        background-color: {tema["botoes"]["primario_fundo_hover"]} !important;
        transform: translateY(-2px);
    }}
    p, li, .stMarkdown, .stText {{
        color: {tema["cores"]["texto_secundario"]};
    }}
    hr {{
        border-color: {tema["cores"]["primaria_clara"]};
    }}
    .stInfo, .stSuccess, .stWarning, .stError {{
        border-radius: 8px !important;
    }}
</style>
"""
st.markdown(css_tema, unsafe_allow_html=True)

st.markdown("""
<script>
document.addEventListener('DOMContentLoaded', function() {
    function styleDropdowns() {
        const dropdowns = document.querySelectorAll('[data-baseweb="menu"]');
        dropdowns.forEach(dropdown => {
            dropdown.style.backgroundColor = '#1a1a2e';
            dropdown.style.border = '1px solid #2c6b96';
            dropdown.style.borderRadius = '8px';
            const options = dropdown.querySelectorAll('[role="option"]');
            options.forEach(option => {
                option.style.backgroundColor = '#1a1a2e';
                option.style.color = 'white';
                option.style.padding = '8px 16px';
                if (option.getAttribute('aria-selected') === 'true') {
                    option.style.backgroundColor = '#0d6e2e';
                }
                option.addEventListener('mouseenter', function() {
                    this.style.backgroundColor = '#2c6b96';
                });
                option.addEventListener('mouseleave', function() {
                    if (this.getAttribute('aria-selected') === 'true') {
                        this.style.backgroundColor = '#0d6e2e';
                    } else {
                        this.style.backgroundColor = '#1a1a2e';
                    }
                });
            });
        });
    }
    const observer = new MutationObserver(function() { styleDropdowns(); });
    observer.observe(document.body, { childList: true, subtree: true });
    styleDropdowns();
});
</script>
""", unsafe_allow_html=True)

# ============================================
# LOGIN
# ============================================
def tela_login():
    col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
    with col_logo2:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=200)
    st.title("📐 Proanalise v1.622")
    st.caption("Sistema de análise urbanística padronizada com geração de parecer técnico")
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("### Acesso ao sistema")
        usuarios = carregar_usuarios()
        user = st.text_input("Usuário", key="login_user")
        senha = st.text_input("Senha", type="password", key="login_senha")
        if st.button("Entrar", use_container_width=True, key="btn_login", type="primary"):
            if user in usuarios and usuarios[user]["senha"] == senha:
                st.session_state["logado"] = True
                st.session_state["usuario"] = user
                st.session_state["usuario_info"] = usuarios[user]
                st.session_state["papel"] = usuarios[user]["papel"]
                st.session_state["nivel"] = usuarios[user]["nivel"]
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

if "logado" not in st.session_state:
    st.session_state["logado"] = False
if not st.session_state["logado"]:
    tela_login()
    st.stop()

# ============================================
# SIDEBAR
# ============================================
st.sidebar.title("📐 Proanalise v1.622")
st.sidebar.write(f"👤 {st.session_state['usuario']} - {st.session_state.get('papel', 'Analista')}")
st.sidebar.write(f"🔒 Nível: {st.session_state.get('nivel', 1)}")

if st.session_state.get("analise_ativa", False):
    estado = st.session_state.get("analise_estado", "em_andamento")
    if estado == "em_andamento":
        st.sidebar.success("🚀 Análise em andamento")
    elif estado == "aguardando_revisao":
        st.sidebar.warning("⏳ Aguardando revisão")
    elif estado == "revisado":
        st.sidebar.info("✅ Análise revisada")
    if st.session_state.get("tempo_inicio"):
        tempo_decorrido = agora_brasilia() - st.session_state["tempo_inicio"]
        horas = tempo_decorrido.seconds // 3600
        minutos = (tempo_decorrido.seconds % 3600) // 60
        st.sidebar.caption(f"⏱️ Tempo: {horas}h {minutos}min")
else:
    if st.session_state.get("analise_concluida", False):
        st.sidebar.info("✅ Análise concluída")
    else:
        st.sidebar.warning("⏸️ Nenhuma análise ativa")

st.sidebar.markdown("---")
with st.sidebar.expander("📋 Resumo do Protocolo", expanded=True):
    col1, col2 = st.sidebar.columns(2)
    with col1:
        st.caption("Protocolo")
        st.write(f"**{st.session_state.get('protocolo', '—')}**")
        st.caption("Tipo")
        st.write(f"**{st.session_state.get('tipo', '—')}**")
        st.caption("Análise")
        st.write(f"**{st.session_state.get('tipo_analise', '—')}**")
    with col2:
        st.caption("Lotes")
        st.write(f"**{st.session_state.get('n_lotes', '—')}**")
        st.caption("Requerente")
        st.write(f"**{st.session_state.get('interessado', '—')[:20]}**")
    st.caption("Matrícula(s)")
    st.write(f"**{st.session_state.get('matriculas', '—')[:30]}**")
    st.caption("Estado")
    st.write(f"**{st.session_state.get('analise_estado', 'em_andamento').replace('_', ' ').title()}**")

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Busca Rápida")
termo_busca = st.sidebar.text_input("N° Protocolo", placeholder="Digite o protocolo...", key="busca_rapida")
if termo_busca:
    dados_busca = carregar_estado_analise_persistente(termo_busca)
    if dados_busca:
        estado_busca = dados_busca.get("analise_estado", "em_andamento")
        responsavel_busca = dados_busca.get("analista_responsavel") or dados_busca.get("analista", "Não informado")
        st.sidebar.write(f"📋 {dados_busca.get('protocolo', termo_busca)} - {estado_busca.replace('_', ' ').title()}")
        if estado_busca == "aguardando_revisao":
            usuario_atual_busca = st.session_state.get("usuario", "")
            if usuario_atual_busca and usuario_atual_busca == responsavel_busca:
                st.sidebar.warning("⚠️ O Analista 1 não pode revisar a própria análise.")
            elif st.sidebar.button("🔍 Abrir para revisão", key="abrir_revisao_busca", use_container_width=True):
                aplicar_estado_analise(dados_busca, "4. Revisão")
                st.session_state["modo_revisao"] = True
                st.rerun()
        elif estado_busca == "revisado":
            st.sidebar.success(f"✅ Revisado por: {dados_busca.get('analista_revisor', 'Não informado')}")
    else:
        resultados = buscar_protocolos(termo_busca)
        if resultados:
            for r in resultados[:5]:
                st.sidebar.write(f"📋 {r['protocolo']} - {r['status']}")
        else:
            st.sidebar.info("Nenhum protocolo encontrado.")

st.sidebar.markdown("---")
st.sidebar.subheader("📋 Protocolos para revisão")
protocolos_revisao = listar_protocolos_aguardando_revisao()
if protocolos_revisao:
    for item in protocolos_revisao:
        with st.sidebar.expander(f"📌 {item['protocolo']}", expanded=False):
            st.caption(f"Analista: {item['analista_responsavel']}")
            st.caption(f"Backup: {item['data_backup']}")
            if st.button(f"Ir para revisão", key=f"btn_revisar_{item['protocolo']}", use_container_width=True):
                dados = item["backup_data"]
                responsavel = dados.get("analista_responsavel") or dados.get("analista", "")
                if st.session_state.get("usuario", "") == responsavel:
                    st.error("⚠️ O analista responsável não pode revisar a própria análise.")
                else:
                    aplicar_estado_analise(dados, "4. Revisão")
                    st.session_state["modo_revisao"] = True
                    st.rerun()
else:
    st.sidebar.info("Nenhum protocolo aguardando revisão no momento.")

st.sidebar.markdown("---")
menus_base = ["1. Protocolo", "2. Analista", "3. Análise", "4. Revisão", "5. Gerar parecer"]
menus_nivel2 = ["6. Dashboard"]
menus_nivel3 = ["7. Comparador"]
menus_disponiveis = []
for menu in menus_base:
    if pode_ver_menu(menu):
        menus_disponiveis.append(menu)
if tem_permissao(2):
    menus_disponiveis.extend(menus_nivel2)
if tem_permissao(3):
    menus_disponiveis.extend(menus_nivel3)
if st.session_state["etapa"] not in menus_disponiveis:
    st.session_state["etapa"] = menus_disponiveis[0]
etapa_atual = st.sidebar.radio("📋 Etapas", menus_disponiveis, index=menus_disponiveis.index(st.session_state["etapa"]))
if etapa_atual != st.session_state["etapa"]:
    st.session_state["etapa"] = etapa_atual
    st.rerun()

st.sidebar.markdown("---")
col_salvar, col_restaurar = st.sidebar.columns(2)
with col_salvar:
    if st.button("💾 Salvar manual", use_container_width=True):
        salvar_backup_manual()
        analista = st.session_state.get("analista", "Não informado")
        n_analise = st.session_state.get("n_analise", "Não informado")
        protocolo = st.session_state.get("protocolo", "Não informado")
        horario = agora_brasilia().strftime("%d/%m/%Y %H:%M:%S")
        st.sidebar.success(
            f"Análise salva com sucesso!\n\n"
            f"Protocolo: {protocolo}\n"
            f"Analista: {analista}\n"
            f"N° Análise: {n_analise}\n"
            f"Horário: {horario}"
        )
with col_restaurar:
    backups_disp = restaurar_analise()
    if backups_disp:
        backup_selecionado = st.sidebar.selectbox("", list(backups_disp.keys()), 
                                                   format_func=lambda x: backups_disp[x],
                                                   label_visibility="collapsed")
        if st.button("🔄 Restaura análise", use_container_width=True):
            caminho_backup = os.path.join("dados", "backups", backup_selecionado)
            with open(caminho_backup, "r", encoding="utf-8") as f:
                dados_restaurados = json.load(f)
                for key, value in dados_restaurados.items():
                    if key in st.session_state:
                        st.session_state[key] = value
                st.session_state["respostas_analise"] = dados_restaurados.get("respostas_analise", {})
                st.session_state["observacoes_analise"] = dados_restaurados.get("observacoes_analise", {})
                st.session_state["pendencias_analise"] = dados_restaurados.get("pendencias_analise", {})
                st.session_state["tempo_inicio"] = deserialize_datetime(dados_restaurados.get("tempo_inicio"))
                st.session_state["tempo_fim"] = deserialize_datetime(dados_restaurados.get("tempo_fim"))
                st.session_state["marcadas_revisao"] = set(dados_restaurados.get("marcadas_revisao", []))
                st.session_state["analise_estado"] = dados_restaurados.get("analise_estado", "em_andamento")
                st.session_state["analista_responsavel"] = dados_restaurados.get("analista_responsavel", "")
                st.session_state["analista_revisor"] = dados_restaurados.get("analista_revisor", "")
                st.success("✅ Análise restaurada com sucesso!")
            st.rerun()
    else:
        st.sidebar.info("Nenhum backup disponível. Salve uma análise primeiro.")

st.sidebar.markdown("---")
with st.sidebar.expander("📓 Anotações Pessoais", expanded=False):
    anotacao_atual = st.session_state["anotacoes_pessoais"].get(st.session_state.get("protocolo", ""), "")
    nova_anotacao = st.text_area("Não vão para o parecer", value=anotacao_atual, height=150)
    if nova_anotacao != anotacao_atual:
        st.session_state["anotacoes_pessoais"][st.session_state.get("protocolo", "")] = nova_anotacao

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Aparência")
col_tema1, col_tema2 = st.sidebar.columns(2)
with col_tema1:
    if st.button("☀️ Claro", use_container_width=True):
        st.session_state["tema_mode"] = "claro"
        st.rerun()
with col_tema2:
    if st.button("🌙 Escuro", use_container_width=True):
        st.session_state["tema_mode"] = "escuro"
        st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Sair", use_container_width=True, key="btn_sair"):
    st.session_state["logado"] = False
    st.session_state.pop("dados_antigos", None)
    st.rerun()

# ============================================
# FUNÇÕES DE RENDERIZAÇÃO
# ============================================
def render_status_badge(status):
    if status == "conforme":
        st.markdown(f"<div class='status-badge-conforme'>{tema['status']['conforme']['icone']} <strong>CONFORME</strong></div>", unsafe_allow_html=True)
    elif status == "inconforme":
        st.markdown(f"<div class='status-badge-inconforme'>{tema['status']['inconforme']['icone']} <strong>INCONFORME</strong></div>", unsafe_allow_html=True)
    elif status == "pendente":
        st.markdown(f"<div class='status-badge-pendente'>{tema['status']['pendente']['icone']} <strong>PENDENTE</strong></div>", unsafe_allow_html=True)
    elif status == "na":
        st.markdown(f"<div class='status-badge-na'>{tema['status']['nao_se_enquadra']['icone']} <strong>NÃO SE ENQUADRA</strong></div>", unsafe_allow_html=True)

def cor_progresso(pct):
    r = int(255 * (1 - pct))
    g = int(180 * pct + 60)
    b = 60
    return f"rgb({r},{g},{b})"

def render_progresso(preenchidas, total, pct, destino):
    cor = cor_progresso(pct)
    html = f"""
    <div><b>{preenchidas}/{total}</b> respostas preenchidas ({int(pct*100)}%)</div>
    <div class="progress-wrap">
        <div class="progress-bar" style="width:{pct*100:.1f}%; background:{cor};"></div>
    </div>
    """
    destino.markdown(html, unsafe_allow_html=True)

# ============================================
# CARREGAR PERGUNTAS
# ============================================
if st.session_state.get("tipo") and st.session_state.get("tipo_analise"):
    perguntas = carregar_perguntas_por_tipo(st.session_state["tipo"], st.session_state["tipo_analise"])
else:
    perguntas = []

col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=100)
with col_titulo:
    st.title("📐 Proanalise v1.622")
    st.caption("Sistema de análise urbanística padronizada com geração de parecer técnico")

tema = carregar_tema()

# ============================================
# FUNÇÕES PRINCIPAIS DEPENDENTES DE PERGUNTAS
# ============================================
def definir_conclusao(respostas, pendencias_manuais=None):
    for p in perguntas:
        resposta = respostas.get(p["id"])
        if not resposta_preenchida(resposta):
            return "PENDENTE"
    for p in perguntas:
        resposta = respostas.get(p["id"])
        if resposta_preenchida(resposta):
            conformes = p.get("conformes", ["Sim", "Não se enquadra"])
            if resposta not in conformes and resposta in p.get("regras", {}):
                return "DESFAVORÁVEL"
    if pendencias_manuais:
        for grupo, pendencias in pendencias_manuais.items():
            if isinstance(pendencias, list):
                for pendencia in pendencias:
                    if pendencia and pendencia.strip():
                        return "DESFAVORÁVEL"
            elif pendencias and pendencias.strip():
                return "DESFAVORÁVEL"
    return "FAVORÁVEL"

def montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais=None):
    grupos = {}
    for p in perguntas:
        pid = p["id"]
        resp = respostas.get(pid)
        if not resposta_preenchida(resp):
            continue
        conformes = p.get("conformes", ["Sim", "Não se enquadra"])
        if resp not in conformes and resp in p.get("regras", {}):
            grupo = p["grupo"]
            texto = p["regras"][resp]["texto"]
            obs = observacoes.get(pid, "").strip()
            if obs:
                texto += f"\nObservação: {obs}"
            grupos.setdefault(grupo, []).append(texto)
    if pendencias_manuais:
        for grupo, pendencias in pendencias_manuais.items():
            if isinstance(pendencias, list):
                for pendencia in pendencias:
                    if pendencia and pendencia.strip():
                        grupos.setdefault(grupo, []).append(pendencia)
            elif pendencias and pendencias.strip():
                grupos.setdefault(grupo, []).append(pendencias)
    return grupos

def montar_inconformidades_rt(respostas, observacoes, pendencias_manuais=None):
    grupos = montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais)
    rt = RichText()
    contador = 1
    if grupos:
        for grupo, itens in grupos.items():
            rt.add(grupo.upper(), bold=True)
            rt.add("\n\n")
            for item in itens:
                rt.add(f"{contador}. {item}")
                rt.add("\n\n")
                contador += 1
    else:
        rt.add("Não foram identificadas inconformidades.")
    return rt

def gerar_docx(dados, respostas, observacoes, conclusao, analista, matricula, setor, n_analise, pendencias_manuais=None):
    if not os.path.exists("modelo_parecer.docx"):
        st.error("Arquivo modelo_parecer.docx não encontrado.")
        st.stop()
    doc = DocxTemplate("modelo_parecer.docx")
    inconformidades_rt = montar_inconformidades_rt(respostas, observacoes, pendencias_manuais)
    matriculas_str = dados.get("matriculas", "")
    if isinstance(matriculas_str, list):
        matriculas_str = ", ".join(matriculas_str)
    context = {
        "protocolo": dados["protocolo"],
        "tipo": dados["tipo"],
        "interessado": dados["interessado"],
        "n_lotes": dados["n_lotes"],
        "matriculas": matriculas_str,
        "inconformidades": inconformidades_rt,
        "conclusao": conclusao,
        "data": f"Data: {agora_brasilia().strftime('%d/%m/%Y')}",
        "analista": f"Analista: {analista}",
        "matricula": matricula,
        "setor": setor,
        "n_analise": n_analise
    }
    doc.render(context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def visualizar_parecer_html(dados, respostas, observacoes, conclusao, analista, n_analise, pendencias_manuais=None):
    grupos_inconformes = montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais)
    html = f"""
    <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ccc; border-radius: 10px; background-color: white;">
        <h2 style="color: #1a5276;">PARECER DE ANÁLISE URBANÍSTICA</h2>
        <hr>
        <p><strong>N° Protocolo:</strong> {dados['protocolo']}</p>
        <p><strong>Análise n°:</strong> {n_analise}</p>
        <p><strong>Tipo do Empreendimento:</strong> {dados['tipo']}</p>
        <p><strong>Tipo de Análise:</strong> {st.session_state.get('tipo_analise', '—')}</p>
        <p><strong>Requerente:</strong> {dados['interessado']}</p>
        <p><strong>Número de Lotes:</strong> {dados['n_lotes']}</p>
        <p><strong>Matrícula(s):</strong> {dados['matriculas']}</p>
        <p><strong>Analista:</strong> {analista}</p>
        <p><strong>Data:</strong> {agora_brasilia().strftime('%d/%m/%Y')}</p>
        <hr>
        <h3 style="color: #1a5276;">INCONFORMIDADES IDENTIFICADAS:</h3>
    """
    if grupos_inconformes:
        contador = 1
        for grupo, itens in grupos_inconformes.items():
            html += f"<h4>{grupo.upper()}</h4>"
            for item in itens:
                html += f"<p><strong>{contador}.</strong> {item.replace(chr(10), '<br>')}</p>"
                contador += 1
    else:
        html += "<p>Não foram identificadas inconformidades.</p>"
    html += f"""
        <hr>
        <h3 style="color: #1a5276;">CONCLUSÃO</h3>
        <p>Diante do exposto, o parecer é <strong>{conclusao}</strong>.</p>
    </div>
    """
    return html

def progresso_percentual(respostas):
    total = len(perguntas)
    preenchidas = sum(1 for v in respostas.values() if resposta_preenchida(v))
    if total == 0:
        return 0, 0, 0.0
    pct = preenchidas / total
    return preenchidas, total, pct

def resumo_status_pergunta(p, resposta):
    if not resposta_preenchida(resposta):
        return "pendente"
    conformes = p.get("conformes", ["Sim", "Não se enquadra"])
    if resposta == "Não se enquadra":
        return "na"
    if resposta in conformes:
        return "conforme"
    if resposta in p.get("regras", {}):
        return "inconforme"
    return "neutro"

# ============================================
# ETAPA 1 - PROTOCOLO
# ============================================
if st.session_state["etapa"] == "1. Protocolo":
    st.header("📋 Dados do protocolo")
    protocolo = st.text_input("N° Protocolo", value=st.session_state["protocolo"], key="protocolo_input")
    if protocolo != st.session_state.get("protocolo_anterior", ""):
        st.session_state["protocolo_anterior"] = protocolo
        if protocolo:
            pasta_backup = os.path.join("dados", "backups")
            if os.path.exists(pasta_backup):
                backups = [f for f in os.listdir(pasta_backup) if f.startswith("backup_") and f.endswith(".json")]
                for backup in sorted(backups, reverse=True):
                    caminho = os.path.join(pasta_backup, backup)
                    try:
                        with open(caminho, "r", encoding="utf-8") as f:
                            dados = json.load(f)
                        if dados.get("protocolo") == protocolo:
                            for key, value in dados.items():
                                if key in st.session_state:
                                    st.session_state[key] = value
                            st.session_state["respostas_analise"] = dados.get("respostas_analise", {})
                            st.session_state["observacoes_analise"] = dados.get("observacoes_analise", {})
                            st.session_state["pendencias_analise"] = dados.get("pendencias_analise", {})
                            st.session_state["tempo_inicio"] = deserialize_datetime(dados.get("tempo_inicio"))
                            st.session_state["tempo_fim"] = deserialize_datetime(dados.get("tempo_fim"))
                            st.session_state["marcadas_revisao"] = set(dados.get("marcadas_revisao", []))
                            st.session_state["analise_estado"] = dados.get("analise_estado", "em_andamento")
                            st.session_state["analista_responsavel"] = dados.get("analista_responsavel", "")
                            st.session_state["analista_revisor"] = dados.get("analista_revisor", "")
                            st.session_state["analise_ativa"] = dados.get("analise_ativa", False)
                            st.session_state["analise_concluida"] = dados.get("analise_concluida", False)
                            perguntas = carregar_perguntas_por_tipo(st.session_state["tipo"], st.session_state["tipo_analise"])
                            st.rerun()
                            break
                    except:
                        continue
    if protocolo != st.session_state["protocolo"]:
        st.session_state["protocolo"] = protocolo
        if st.session_state.get("analise_ativa"):
            st.session_state["analise_ativa"] = False

    col1, col2 = st.columns(2)
    with col1:
        opcoes_tipo = ["Loteamento", "Condomínio fechado de lotes", "Desmembramentos"]
        indice_tipo = opcoes_tipo.index(st.session_state["tipo"]) if st.session_state["tipo"] in opcoes_tipo else 0
        tipo = st.selectbox("Tipo do Empreendimento", opcoes_tipo, index=indice_tipo, key="tipo_select")
        if tipo != st.session_state["tipo"]:
            st.session_state["tipo"] = tipo
            perguntas = carregar_perguntas_por_tipo(tipo, st.session_state["tipo_analise"])
    with col2:
        if tipo == "Desmembramentos":
            opcoes_analise = [
                "Retificação", "Desdobro de áreas", "Unificação de áreas",
                "Alteração de divisas", "Estremação", "Usucapião",
                "Retificação com desdobro", "Retificação com alteração de divisas",
                "Retificação com unificação", "Unificação com desdobro",
                "Desdobro para anexação", "Anuência rural"
            ]
        else:
            opcoes_analise = ["Aceite urbanístico", "Alvará"]
        if st.session_state["tipo_analise"] not in opcoes_analise:
            st.session_state["tipo_analise"] = opcoes_analise[0]
        indice_analise = opcoes_analise.index(st.session_state["tipo_analise"])
        tipo_analise = st.selectbox("Tipo de Análise", opcoes_analise, index=indice_analise, key="tipo_analise_select")
        if tipo_analise != st.session_state["tipo_analise"]:
            st.session_state["tipo_analise"] = tipo_analise
            perguntas = carregar_perguntas_por_tipo(tipo, tipo_analise)

    interessado = st.text_input("Requerente", value=st.session_state["interessado"], key="interessado_input")
    st.session_state["interessado"] = interessado
    n_lotes = st.number_input("Número de Lotes", min_value=1, value=st.session_state["n_lotes"], key="n_lotes_input")
    st.session_state["n_lotes"] = n_lotes
    matriculas = st.text_area("Matrícula(s) do Empreendimento", value=st.session_state["matriculas"], 
                               key="matriculas_input", placeholder="Digite a(s) matrícula(s) separadas por vírgula")
    st.session_state["matriculas"] = matriculas
    st.markdown("---")

    estado = st.session_state.get("analise_estado", "em_andamento")
    if estado == "em_andamento":
        if st.session_state.get("analise_ativa"):
            st.success("🚀 Análise em andamento")
            if st.button("⏩ Continuar Análise", use_container_width=True, type="primary"):
                st.session_state["etapa"] = "2. Analista"
                st.rerun()
        else:
            if st.button("🚀 Iniciar Análise", use_container_width=True, type="primary"):
                if st.session_state["protocolo"] and st.session_state["interessado"]:
                    st.session_state["analise_ativa"] = True
                    st.session_state["analise_concluida"] = False
                    st.session_state["analise_estado"] = "em_andamento"
                    st.session_state["analista_responsavel"] = st.session_state["usuario"]
                    st.session_state["tempo_inicio"] = agora_brasilia()
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
                else:
                    st.error("⚠️ Informe o protocolo e o requerente")
    elif estado == "aguardando_revisao":
        st.warning("⏳ Análise aguardando revisão por outro analista")
        responsavel = st.session_state.get("analista_responsavel", "")
        if responsavel:
            st.info(f"Analista responsável: {responsavel}")
        if st.button("🔍 Ir para Revisão", use_container_width=True):
            st.session_state["etapa"] = "4. Revisão"
            st.rerun()
    elif estado == "revisado":
        st.success("✅ Análise revisada e aprovada")
        revisor = st.session_state.get("analista_revisor", "")
        if revisor:
            st.info(f"Revisado por: {revisor}")
        if st.button("📄 Gerar Parecer", use_container_width=True, type="primary"):
            st.session_state["etapa"] = "5. Gerar parecer"
            st.rerun()

# ============================================
# ETAPA 2 - ANALISTA
# ============================================
elif st.session_state["etapa"] == "2. Analista":
    st.header("👤 Dados do analista")
    st.info(f"📌 Protocolo: **{st.session_state['protocolo']}**")
    analista = st.text_input("Nome do Analista", value=st.session_state["analista"], key="analista_input")
    st.session_state["analista"] = analista
    matricula_analista = st.text_input("Matrícula do Analista", value=st.session_state["matricula_analista"], key="matricula_analista_input")
    st.session_state["matricula_analista"] = matricula_analista
    setor = st.text_input("Setor", value=st.session_state["setor"], key="setor_input")
    st.session_state["setor"] = setor
    n_analise_sugerida = "1"
    if st.session_state["protocolo"]:
        pasta = get_pasta_protocolo(st.session_state["protocolo"])
        if os.path.exists(pasta):
            analises = [f for f in os.listdir(pasta) if f.startswith("AN") and f.endswith(".json")]
            if analises:
                ultimo_num = max([int(a.replace("AN", "").replace(".json", "")) for a in analises])
                n_analise_sugerida = str(ultimo_num + 1)
    if not st.session_state["n_analise"]:
        st.session_state["n_analise"] = n_analise_sugerida
    n_analise = st.text_input("Nº da Análise", value=st.session_state["n_analise"], key="n_analise_input")
    st.session_state["n_analise"] = n_analise
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "1. Protocolo"
            st.rerun()
    with col2:
        if st.button("Prosseguir →", use_container_width=True, type="primary"):
            if st.session_state["analista"] and st.session_state["n_analise"]:
                st.session_state["etapa"] = "3. Análise"
                st.rerun()
            else:
                st.error("⚠️ Preencha todos os campos")

# ============================================
# ETAPA 3 - ANÁLISE
# ============================================
elif st.session_state["etapa"] == "3. Análise":
    estado = st.session_state.get("analise_estado", "em_andamento")
    usuario_atual = st.session_state.get("usuario", "")
    responsavel = st.session_state.get("analista_responsavel", "")
    modo_leitura = False
    if estado == "aguardando_revisao" and usuario_atual != responsavel:
        modo_leitura = True
        st.info("🔍 Modo de visualização – você pode ver as respostas, mas não editá-las.")
    elif estado == "revisado":
        st.warning("⚠️ Esta análise já foi revisada e aprovada. Não é mais possível editar.")
        modo_leitura = True

    if not st.session_state.get("analise_ativa", False) and not modo_leitura:
        if st.session_state.get("respostas_analise") and not st.session_state.get("analise_concluida", False):
            st.session_state["analise_ativa"] = True
        else:
            st.error("❌ Nenhuma análise ativa. Volte à Etapa 1 e clique em 'Iniciar Análise'")
            if st.button("← Voltar à Etapa 1"):
                st.session_state["etapa"] = "1. Protocolo"
                st.rerun()
            st.stop()

    st.header("🔍 Análise técnica" + (" (somente leitura)" if modo_leitura else ""))
    st.info(f"📌 Protocolo: **{st.session_state['protocolo']}** | Analista: **{st.session_state['analista']}**")
    fazer_backup_automatico()

    respostas = st.session_state.get("respostas_analise", {})
    observacoes = st.session_state.get("observacoes_analise", {})
    pendencias_manuais = st.session_state.get("pendencias_analise", {})

    if not modo_leitura:
        proximo_id, _ = proxima_pergunta_nao_respondida(respostas, perguntas)
        if proximo_id:
            st.session_state["botao_flutuante"] = True
        else:
            st.session_state["botao_flutuante"] = False

    # Botão preencher aleatoriamente (nível 3)
    if tem_permissao(3) and not modo_leitura:
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("🎲 Preencher respostas aleatoriamente", use_container_width=True, help="Preenche todas as perguntas não respondidas com opções aleatórias (para testes)"):
                for p in perguntas:
                    pid = p["id"]
                    if not resposta_preenchida(respostas.get(pid)):
                        opcoes = p["opcoes"]
                        if opcoes:
                            opcoes_validas = [op for op in opcoes if op != "Selecione..."]
                            if opcoes_validas:
                                # IMPORTANTE: além de salvar no dicionário de respostas,
                                # atualiza o estado do próprio widget st.selectbox.
                                # Sem isso, no st.rerun() o selectbox pode recuperar o
                                # valor antigo ("Selecione...") e sobrescrever a resposta
                                # aleatória recém-gerada.
                                resposta_aleatoria = random.choice(opcoes_validas)
                                respostas[pid] = resposta_aleatoria
                                st.session_state[f"resp_{pid}"] = resposta_aleatoria
                st.session_state["respostas_analise"] = dict(respostas)
                st.session_state["observacoes_analise"] = observacoes
                st.session_state["pendencias_analise"] = pendencias_manuais
                st.rerun()

    grupos_ordenados = []
    for p in perguntas:
        grupo = p["grupo"]
        if grupo not in grupos_ordenados:
            grupos_ordenados.append(grupo)

    inconformes_sidebar = []
    for grupo in grupos_ordenados:
        perguntas_grupo = [p for p in perguntas if p["grupo"] == grupo]
        with st.expander(f"📁 {grupo}", expanded=False):
            for idx, p in enumerate(perguntas_grupo):
                pid = p["id"]
                valor_salvo = respostas.get(pid, "Selecione...")
                obs_salva = observacoes.get(pid, "")
                if st.session_state["dados_antigos"] and pid not in respostas:
                    valor_salvo = st.session_state["dados_antigos"]["respostas"].get(pid, "Selecione...")
                    obs_salva = st.session_state["dados_antigos"]["observacoes"].get(pid, "")
                opcoes = ["Selecione..."] + p["opcoes"]
                idx_padrao = opcoes.index(valor_salvo) if valor_salvo in opcoes else 0

                col_pergunta, col_status, col_marcar = st.columns([3, 1, 0.5])
                with col_pergunta:
                    if modo_leitura:
                        st.markdown(f"**{p['pergunta']}**")
                        st.write(f"**Resposta:** {valor_salvo}")
                    else:
                        resposta = st.selectbox(p["pergunta"], opcoes, index=idx_padrao, key=f"resp_{pid}", help=f"ID: {pid}")
                        respostas[pid] = resposta

                with col_status:
                    status = resumo_status_pergunta(p, valor_salvo if modo_leitura else respostas.get(pid, "Selecione..."))
                    render_status_badge(status)
                    if status == "inconforme":
                        inconformes_sidebar.append(p["pergunta"])

                if not modo_leitura:
                    with col_marcar:
                        marcada = pid in st.session_state["marcadas_revisao"]
                        if st.button("🔖", key=f"marcar_{pid}", help="Marcar para revisão"):
                            if marcada:
                                st.session_state["marcadas_revisao"].discard(pid)
                            else:
                                st.session_state["marcadas_revisao"].add(pid)
                            st.rerun()
                    if pid in st.session_state["marcadas_revisao"]:
                        st.markdown("<div class='card-revisao'>🔖 Marcada para revisão posterior</div>", unsafe_allow_html=True)
                    obs = st.text_area("📝 Observação (opcional)", value=obs_salva, key=f"obs_{pid}", height=68)
                    observacoes[pid] = obs
                else:
                    if obs_salva:
                        st.markdown(f"**Observação:** {obs_salva}")
                st.markdown("---")

            if not modo_leitura:
                st.markdown("### 📝 Inconformidades Diversas")
                if grupo not in pendencias_manuais:
                    pendencias_manuais[grupo] = []
                if pendencias_manuais[grupo]:
                    for i, pendencia in enumerate(pendencias_manuais[grupo]):
                        if pendencia:
                            col_p, col_b = st.columns([10, 1])
                            with col_p:
                                st.markdown(f"📌 {pendencia}")
                            with col_b:
                                if st.button("🗑️", key=f"del_{grupo}_{i}"):
                                    pendencias_manuais[grupo].pop(i)
                                    st.rerun()
                if st.button(f"+ Adicionar", key=f"add_{grupo}"):
                    pendencias_manuais[grupo].append("")
                    st.rerun()
                if pendencias_manuais[grupo] and not pendencias_manuais[grupo][-1]:
                    nova = st.text_area("Nova inconformidade", key=f"new_{grupo}", height=68)
                    if nova:
                        pendencias_manuais[grupo][-1] = nova
                        st.rerun()

    if not modo_leitura:
        st.session_state["respostas_analise"] = respostas
        st.session_state["observacoes_analise"] = observacoes
        st.session_state["pendencias_analise"] = pendencias_manuais

    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)
    if preenchidas < total:
        st.warning(f"⚠️ Atenção: {total - preenchidas} perguntas ainda não foram respondidas!")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "2. Analista"
            st.rerun()
    with col2:
        if not modo_leitura:
            if st.button("Prosseguir →", use_container_width=True, type="primary"):
                respostas_atual = st.session_state["respostas_analise"]
                preenchidas_agora = sum(1 for v in respostas_atual.values() if resposta_preenchida(v))
                if preenchidas_agora == total:
                    st.session_state["etapa"] = "4. Revisão"
                    st.rerun()
                else:
                    st.error(f"⚠️ Responda todas as perguntas antes de prosseguir ({total - preenchidas_agora} pendentes)")
        else:
            if st.button("Voltar à página inicial", use_container_width=True):
                st.session_state["etapa"] = "1. Protocolo"
                st.rerun()

    if not modo_leitura and st.session_state.get("botao_flutuante", False):
        st.markdown(f"""
        <div style="position: fixed; bottom: 20px; right: 20px; z-index: 999;">
            <button onclick="document.querySelector('input[type=\"text\"]')?.focus(); window.scrollTo(0, 0);" 
                    style="background-color: {tema['botoes']['primario_fundo']}; color: white; border: none; border-radius: 50px; padding: 12px 24px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3);">
                🎯 Ir para primeira pergunta pendente
            </button>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# ETAPA 4 - REVISÃO
# ============================================
elif st.session_state["etapa"] == "4. Revisão":
    estado = st.session_state.get("analise_estado", "em_andamento")
    usuario_atual = st.session_state.get("usuario", "")
    responsavel = st.session_state.get("analista_responsavel", "")
    eh_revisor = (
        estado == "aguardando_revisao"
        and st.session_state.get("modo_revisao", False)
        and bool(usuario_atual)
        and usuario_atual != responsavel
    )

    # O revisor precisa poder trabalhar na etapa 4 mesmo que a análise
    # tenha sido concluída pelo Analista 1 e analise_ativa=False.
    if not st.session_state.get("analise_ativa", False) and not st.session_state.get("analise_concluida", False):
        if st.session_state.get("respostas_analise"):
            st.session_state["analise_ativa"] = True
        else:
            st.error("❌ Nenhuma análise carregada para revisão.")
            if st.button("← Voltar à Etapa 1"):
                st.session_state["etapa"] = "1. Protocolo"
                st.rerun()
            st.stop()

    st.header("📋 Revisão da análise")

    if eh_revisor:
        st.info(
            "👀 Você é o revisor desta análise. Nesta etapa você pode "
            "editar respostas, observações e inconformidades antes de aprovar."
        )
    elif estado == "aguardando_revisao":
        st.warning("⏳ Esta análise aguarda revisão por outro analista.")
    elif estado == "revisado":
        st.success("✅ Esta análise já foi revisada e aprovada.")

    # Ativação explícita do modo de revisão. Isso evita que a Etapa 4 fique
    # apenas em modo de visualização quando o protocolo foi aberto por outro
    # caminho (por exemplo, pelo menu lateral).
    if estado == "aguardando_revisao" and not eh_revisor:
        if usuario_atual == responsavel:
            st.error(
                "🚫 Este é o Analista 1 responsável pela análise. "
                "Ele não pode revisar a própria análise."
            )
        elif usuario_atual:
            st.warning(
                f"⏳ Revisão pendente. Responsável: **{responsavel or 'não informado'}**"
            )
            if st.button("📝 Assumir revisão como Analista 2", type="primary", use_container_width=True):
                st.session_state["modo_revisao"] = True
                st.session_state["analista_revisor"] = usuario_atual
                st.rerun()

    respostas = dict(st.session_state.get("respostas_analise", {}))
    observacoes = dict(st.session_state.get("observacoes_analise", {}))
    pendencias_manuais = dict(st.session_state.get("pendencias_analise", {}))

    # ========================================================
    # PAINEL DE EDIÇÃO DO REVISOR
    # ========================================================
    if eh_revisor:
        st.success(
            f"🔓 **Modo de revisão ativo — Analista 2: {usuario_atual}** | "
            f"Analista 1: {responsavel or 'não informado'}"
        )
        st.markdown("### ✏️ Ferramentas de revisão")
        st.caption(
            "Altere qualquer resposta que precise de correção. As alterações "
            "são salvas no registro do protocolo e serão utilizadas na geração do parecer."
        )

        # Garantir que os widgets da revisão recebam os valores atuais apenas
        # na primeira renderização. As chaves review_resp_* e review_obs_* são
        # independentes dos widgets da Etapa 3.
        for p in perguntas:
            pid = p["id"]
            opcoes = ["Selecione..."] + p.get("opcoes", [])
            valor_atual = respostas.get(pid, "Selecione...")
            if valor_atual not in opcoes:
                valor_atual = "Selecione..."

            key_resp = f"review_resp_{pid}"
            key_obs = f"review_obs_{pid}"
            if key_resp not in st.session_state:
                st.session_state[key_resp] = valor_atual
            if key_obs not in st.session_state:
                st.session_state[key_obs] = observacoes.get(pid, "")

        grupos_ordenados = []
        for p in perguntas:
            if p["grupo"] not in grupos_ordenados:
                grupos_ordenados.append(p["grupo"])

        for grupo in grupos_ordenados:
            perguntas_grupo = [p for p in perguntas if p["grupo"] == grupo]
            with st.expander(f"📁 {grupo}", expanded=True):
                for p in perguntas_grupo:
                    pid = p["id"]
                    opcoes = ["Selecione..."] + p.get("opcoes", [])
                    key_resp = f"review_resp_{pid}"
                    key_obs = f"review_obs_{pid}"

                    st.markdown(f"**{p['pergunta']}**")
                    st.selectbox(
                        "Resposta",
                        opcoes,
                        key=key_resp,
                        help=f"ID: {pid}"
                    )
                    st.text_area(
                        "📝 Observação (opcional)",
                        key=key_obs,
                        height=68
                    )
                    st.markdown("---")

                st.markdown("**📝 Inconformidades diversas**")
                pendencias_grupo = list(pendencias_manuais.get(grupo, []))
                # Remove apenas itens vazios para facilitar a edição.
                pendencias_grupo = [x for x in pendencias_grupo if str(x).strip()]
                pendencias_manuais[grupo] = pendencias_grupo

                if pendencias_grupo:
                    for i, pendencia in enumerate(pendencias_grupo):
                        key_pend = f"review_pend_{grupo}_{i}"
                        if key_pend not in st.session_state:
                            st.session_state[key_pend] = pendencia
                        st.text_area(
                            f"Inconformidade {i + 1}",
                            key=key_pend,
                            height=68
                        )
                else:
                    st.caption("Nenhuma inconformidade diversa registrada neste grupo.")

                key_nova = f"review_nova_pend_{grupo}"
                st.text_area(
                    "➕ Nova inconformidade",
                    key=key_nova,
                    height=68,
                    placeholder="Digite uma nova inconformidade, se necessário..."
                )

        # Sincroniza os widgets de revisão com os dicionários de trabalho.
        for p in perguntas:
            pid = p["id"]
            respostas[pid] = st.session_state.get(f"review_resp_{pid}", respostas.get(pid, "Selecione..."))
            observacoes[pid] = st.session_state.get(f"review_obs_{pid}", observacoes.get(pid, ""))

        for grupo in list(pendencias_manuais.keys()):
            atualizadas = []
            for i in range(len(pendencias_manuais.get(grupo, []))):
                valor = st.session_state.get(f"review_pend_{grupo}_{i}", "").strip()
                if valor:
                    atualizadas.append(valor)
            nova = st.session_state.get(f"review_nova_pend_{grupo}", "").strip()
            if nova:
                atualizadas.append(nova)
            pendencias_manuais[grupo] = atualizadas

        st.session_state["respostas_analise"] = respostas
        st.session_state["observacoes_analise"] = observacoes
        st.session_state["pendencias_analise"] = pendencias_manuais

        col_salvar, col_aprovar = st.columns(2)
        with col_salvar:
            if st.button("💾 Salvar alterações da revisão", use_container_width=True):
                st.session_state["ultima_edicao_revisao_por"] = usuario_atual
                st.session_state["ultima_edicao_revisao_em"] = agora_brasilia().isoformat()
                salvar_estado_analise_persistente()
                st.success("✅ Alterações da revisão salvas no protocolo.")
                st.rerun()
        with col_aprovar:
            if st.button("✅ Aprovar revisão e prosseguir", use_container_width=True, type="primary"):
                preenchidas_revisao, total_revisao, _ = progresso_percentual(respostas)
                if preenchidas_revisao != total_revisao:
                    st.error(
                        f"⚠️ Não é possível aprovar: existem "
                        f"{total_revisao - preenchidas_revisao} pergunta(s) sem resposta."
                    )
                else:
                    # Grava as últimas alterações antes de mudar o estado.
                    st.session_state["respostas_analise"] = respostas
                    st.session_state["observacoes_analise"] = observacoes
                    st.session_state["pendencias_analise"] = pendencias_manuais
                    st.session_state["analise_estado"] = "revisado"
                    st.session_state["analista_revisor"] = usuario_atual
                    st.session_state["modo_revisao"] = False
                    st.session_state["analise_concluida"] = True
                    st.session_state["analise_ativa"] = False
                    st.session_state["tempo_fim"] = agora_brasilia()
                    st.session_state["ultima_edicao_revisao_por"] = usuario_atual
                    st.session_state["ultima_edicao_revisao_em"] = agora_brasilia().isoformat()
                    salvar_estado_analise_persistente()
                    salvar_analise_analista(
                        st.session_state.get("protocolo", ""),
                        usuario_atual,
                        "revisor_2",
                        respostas,
                        observacoes,
                        pendencias_manuais
                    )
                    st.session_state["etapa"] = "5. Gerar parecer"
                    st.success("✅ Revisão aprovada. Prosseguindo para a geração do parecer.")
                    st.rerun()

    # ========================================================
    # VISÃO/RESUMO DA REVISÃO
    # ========================================================
    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)
    conclusao = definir_conclusao(respostas, pendencias_manuais)
    grupos_inconformes = montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📌 Resumo")
        st.write(f"**Protocolo:** {st.session_state['protocolo']}")
        st.write(f"**Requerente:** {st.session_state['interessado']}")
        st.write(f"**Analista:** {st.session_state['analista']}")
        if estado == "em_andamento":
            st.info("📝 Esta análise está em andamento.")
        elif estado == "aguardando_revisao":
            st.warning("⏳ Esta análise aguarda revisão.")
            if eh_revisor:
                st.success("🔑 Você é o revisor desta análise.")
        elif estado == "revisado":
            st.success("✅ Esta análise já foi revisada.")
        if conclusao == "FAVORÁVEL":
            st.success(f"✅ Conclusão: {conclusao}")
        elif conclusao == "PENDENTE":
            st.warning(f"⚠️ Conclusão: {conclusao}")
        else:
            st.error(f"❌ Conclusão: {conclusao}")
    with col2:
        st.subheader("📊 Estatísticas")
        total_inconformes = sum(len(v) for v in grupos_inconformes.values())
        st.metric("Perguntas", total)
        st.metric("Respondidas", preenchidas)
        st.metric("Inconformidades", total_inconformes)

    fig = gerar_grafico_inconformidades(respostas, grupos_inconformes)
    if fig:
        st.plotly_chart(fig, use_container_width=True)

    # Pré-visualização do parecer já usa as respostas editadas pelo revisor.
    with st.expander("👁️ Pré-visualizar parecer", expanded=False):
        dados_preview = {
            "protocolo": st.session_state.get("protocolo", ""),
            "tipo": st.session_state.get("tipo", ""),
            "interessado": st.session_state.get("interessado", ""),
            "n_lotes": st.session_state.get("n_lotes", 1),
            "matriculas": st.session_state.get("matriculas", "")
        }
        html_parecer = visualizar_parecer_html(
            dados_preview,
            respostas,
            observacoes,
            conclusao,
            st.session_state["analista"],
            st.session_state["n_analise"],
            pendencias_manuais
        )
        st.markdown(html_parecer, unsafe_allow_html=True)

    col_botoes1, col_botoes2, col_botoes3 = st.columns(3)
    with col_botoes1:
        if eh_revisor:
            if st.button("🔄 Recarregar análise original", use_container_width=True):
                dados = carregar_estado_analise_persistente(st.session_state.get("protocolo", ""))
                if dados:
                    aplicar_estado_analise(dados, etapa="4. Revisão")
                    st.session_state["modo_revisao"] = True
                    # As chaves dos widgets são removidas para que recebam os
                    # valores recém-carregados no próximo ciclo.
                    for p in perguntas:
                        st.session_state.pop(f"review_resp_{p['id']}", None)
                        st.session_state.pop(f"review_obs_{p['id']}", None)
                    st.rerun()
        else:
            if st.button("← Voltar para análise", use_container_width=True):
                st.session_state["etapa"] = "3. Análise"
                st.rerun()

    # Analista responsável pode concluir ou pular revisão.
    if estado == "em_andamento" and usuario_atual == responsavel:
        with col_botoes2:
            if not st.session_state.get("analise_concluida", False):
                if st.button("✅ Concluir e enviar para revisão", use_container_width=True, type="primary"):
                    if preenchidas == total:
                        st.session_state["analise_estado"] = "aguardando_revisao"
                        st.session_state["analise_concluida"] = True
                        st.session_state["analise_ativa"] = False
                        st.session_state["tempo_fim"] = agora_brasilia()
                        st.session_state["analista_responsavel"] = st.session_state.get("analista_responsavel") or st.session_state.get("usuario", "")
                        salvar_estado_analise_persistente()
                        salvar_analise_analista(
                            st.session_state.get("protocolo", ""),
                            st.session_state.get("analista", "") or st.session_state.get("usuario", ""),
                            "analista_1",
                            st.session_state.get("respostas_analise", {}),
                            st.session_state.get("observacoes_analise", {}),
                            st.session_state.get("pendencias_analise", {})
                        )
                        st.success("✅ Análise concluída e enviada para revisão. Aguarde a aprovação de outro analista.")
                        st.rerun()
                    else:
                        st.error(f"⚠️ Responda todas as perguntas antes de concluir ({total - preenchidas} pendentes)")
            else:
                st.success("✅ Análise já concluída")
        with col_botoes3:
            if not st.session_state.get("analise_concluida", False):
                if st.button("⚡ Gerar parecer diretamente (pular revisão)", use_container_width=True):
                    if preenchidas == total:
                        st.warning("⚠️ Você está prestes a gerar o parecer sem revisão de outro analista.")
                        if st.button("Confirmar geração direta", use_container_width=True):
                            st.session_state["analise_estado"] = "revisado"
                            st.session_state["analista_revisor"] = usuario_atual
                            st.session_state["modo_revisao"] = False
                            st.session_state["analise_concluida"] = True
                            st.session_state["analise_ativa"] = False
                            st.session_state["tempo_fim"] = agora_brasilia()
                            salvar_estado_analise_persistente()
                            st.success("✅ Parecer gerado diretamente! Você pode prosseguir para a geração do arquivo.")
                            st.session_state["etapa"] = "5. Gerar parecer"
                            st.rerun()
                    else:
                        st.error(f"⚠️ Responda todas as perguntas antes de gerar ({total - preenchidas} pendentes)")
            else:
                st.info("Análise já concluída. Se deseja gerar, vá para a etapa 5.")

    # Revisor: o botão principal de aprovação fica no painel superior para
    # que as alterações sejam explicitamente salvas antes da aprovação.
    elif estado == "aguardando_revisao" and usuario_atual == responsavel:
        with col_botoes3:
            st.warning("⚠️ O analista responsável não pode revisar a própria análise.")

    elif estado == "revisado":
        with col_botoes3:
            if st.button("📄 Gerar Parecer", use_container_width=True, type="primary"):
                st.session_state["etapa"] = "5. Gerar parecer"
                st.rerun()

# ============================================
# ETAPA 5 - GERAR PARECER
# ============================================
elif st.session_state["etapa"] == "5. Gerar parecer":
    if st.session_state.get("analise_estado") != "revisado":
        st.error("❌ O parecer só pode ser gerado após a revisão e aprovação por um segundo analista, ou após a geração direta.")
        if st.button("← Voltar à Revisão"):
            st.session_state["etapa"] = "4. Revisão"
            st.rerun()
        st.stop()

    st.header("📄 Geração do parecer")
    respostas = st.session_state.get("respostas_analise", {})
    observacoes = st.session_state.get("observacoes_analise", {})
    pendencias_manuais = st.session_state.get("pendencias_analise", {})
    dados = {
        "protocolo": st.session_state.get("protocolo", ""),
        "tipo": st.session_state.get("tipo", ""),
        "interessado": st.session_state.get("interessado", ""),
        "n_lotes": st.session_state.get("n_lotes", 1),
        "matriculas": st.session_state.get("matriculas", "")
    }
    conclusao = definir_conclusao(respostas, pendencias_manuais)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📋 Dados do Parecer")
        st.write(f"**Protocolo:** {dados['protocolo']}")
        st.write(f"**Requerente:** {dados['interessado']}")
        st.write(f"**Analista:** {st.session_state['analista']}")
        st.write(f"**Revisor:** {st.session_state.get('analista_revisor', 'N/A')}")
    with col2:
        st.write(f"**Tipo:** {dados['tipo']}")
        st.write(f"**Tipo de Análise:** {st.session_state.get('tipo_analise', '—')}")
        if conclusao == "FAVORÁVEL":
            st.success(f"✅ Conclusão: {conclusao}")
        elif conclusao == "PENDENTE":
            st.warning(f"⚠️ Conclusão: {conclusao}")
        else:
            st.error(f"❌ Conclusão: {conclusao}")
    st.markdown("---")

    with st.expander("👁️ Visualizar Parecer", expanded=True):
        html_parecer = visualizar_parecer_html(dados, respostas, observacoes, conclusao,
                                                st.session_state["analista"], st.session_state["n_analise"],
                                                pendencias_manuais)
        st.markdown(html_parecer, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "4. Revisão"
            st.rerun()
    with col2:
        if st.button("📄 Gerar Parecer", use_container_width=True, type="primary"):
            if not dados["protocolo"]:
                st.error("Protocolo não informado")
            elif not os.path.exists("modelo_parecer.docx"):
                st.error("modelo_parecer.docx não encontrado")
            else:
                try:
                    with st.spinner("Gerando parecer..."):
                        arquivo = gerar_docx(dados, respostas, observacoes, conclusao,
                                            st.session_state["analista"], st.session_state["matricula_analista"],
                                            st.session_state["setor"], st.session_state["n_analise"],
                                            pendencias_manuais)
                        nome = f"PU_{dados['protocolo'].replace('/', '-')}_AN{st.session_state['n_analise']}.docx"
                        st.success("✅ Parecer gerado!")
                        st.download_button("⬇️ Baixar", data=arquivo, file_name=nome, use_container_width=True)
                except Exception as e:
                    st.error(f"Erro: {e}")

# ============================================
# ETAPA 6 - DASHBOARD
# ============================================
elif st.session_state["etapa"] == "6. Dashboard":
    st.header("📊 Dashboard de Métricas")
    if not tem_permissao(2):
        st.error("❌ Acesso negado! Apenas Estagiários Sênior e Analistas Responsáveis podem acessar o Dashboard.")
    else:
        tempo_medio = calcular_tempo_medio_analise()
        if tempo_medio:
            st.metric("⏱️ Tempo Médio de Análise", f"{tempo_medio:.1f} dias")
        fig_tempo = gerar_grafico_tempo_analises()
        if fig_tempo:
            st.plotly_chart(fig_tempo, use_container_width=True)
        st.subheader("🔍 Busca Avançada")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            termo_busca_dash = st.text_input("Protocolo", placeholder="Digite o protocolo...")
        with col_f2:
            filtro_status = st.selectbox("Status", ["Todos", "FAVORÁVEL", "DESFAVORÁVEL", "PENDENTE", "Em análise"])
        with col_f3:
            filtro_analista = st.text_input("Analista", placeholder="Digite o nome...")
        if st.button("🔍 Buscar", use_container_width=True):
            resultados = buscar_protocolos(termo_busca_dash, filtro_status if filtro_status != "Todos" else None,
                                           filtro_analista if filtro_analista else None)
            if resultados:
                st.subheader(f"📋 Resultados encontrados: {len(resultados)}")
                df_resultados = pd.DataFrame(resultados)
                st.dataframe(df_resultados, use_container_width=True)
            else:
                st.info("Nenhum protocolo encontrado")

# ============================================
# ETAPA 7 - COMPARADOR
# ============================================
elif st.session_state["etapa"] == "7. Comparador":
    if not tem_permissao(3):
        st.error("❌ Acesso negado! Apenas Analistas Responsáveis podem acessar o Comparador.")
        if st.button("← Voltar ao menu principal"):
            st.session_state["etapa"] = "1. Protocolo"
            st.rerun()
    else:
        st.header("🔍 Comparador de Análises")
        protocolo_comp = st.text_input("N° Protocolo para comparar", value=st.session_state.get("protocolo", ""))
        if protocolo_comp:
            render_comparador_analises(protocolo_comp)
