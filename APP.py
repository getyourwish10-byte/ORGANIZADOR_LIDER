#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Organizador Inteligente de Pedidos (uniformes) — Gemini + Tesseract + Coluna de Número de Calção
Desenvolvido por: Douglas Oliveira | getyourwish10@gmail.com
"""

from __future__ import annotations

import base64
import copy
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from collections import Counter
from datetime import datetime
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog

__version__ = "1.8.0"
URL_VERSAO_REMOTE = "https://raw.githubusercontent.com/getyourwish10-byte/ORGANIZADOR_LIDER/refs/heads/main/version.json"

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_DISPONIVEL = True
except ImportError:
    OPENPYXL_DISPONIVEL = False

try:
    from PIL import Image, ImageGrab, ImageOps
    PIL_DISPONIVEL = True
except ImportError:
    PIL_DISPONIVEL = False

try:
    import pytesseract
    PYTESSERACT_DISPONIVEL = True
    if not shutil.which("tesseract"):
        for _caminho in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                         r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                         os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe")):
            if os.path.isfile(_caminho):
                pytesseract.pytesseract.tesseract_cmd = _caminho
                break
except ImportError:
    PYTESSERACT_DISPONIVEL = False

try:
    from google import genai
    from google.genai import types
    GEMINI_DISPONIVEL = True
except ImportError:
    GEMINI_DISPONIVEL = False


CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".organizador_pedidos_config.json")
HISTORICO_MAXIMO = 30
COLUNAS = ("NOME", "TAMANHO DE CAMISA", "NÚMERO", "TAMANHO DE CALÇÃO", "NÚMERO DE CALÇÃO")
COLUNAS_TAMANHO = ("TAMANHO DE CAMISA", "TAMANHO DE CALÇÃO")
PREENCHIMENTO = {
    "TAMANHO DE CAMISA": "SEM TAMANHO", 
    "NÚMERO": "SEM NÚMERO", 
    "TAMANHO DE CALÇÃO": "SEM CALÇÃO",
    "NÚMERO DE CALÇÃO": "SEM NÚMERO"
}
VALORES_VAZIOS = {"", "SEM NOME", "SEM TAMANHO", "SEM NÚMERO", "SEM CALÇÃO"}

ORDEM_PADRAO = [
    "G4", "G3", "G2", "GG", "G", "M", "P", "G1",
    "BLG3", "BLG2", "BLGG", "BLG", "BLM", "BLP", "BLPP",
    "13_14_PP", "11_12_GG_INF", "8_10_G_INF", "6_7_M_INF",
    "4_5_P_INF", "2_3_PP_INF", "1_BB", "0_RN",
]

CONVERSOES_PADRAO = {
    "EXG": "G2", "EXGG": "G3", "G1": "GG",
    "BLG1": "BLGG", "BABYLOOK G1": "BLGG",
    "PP": "13_14_PP", "13 A 14 ANOS": "13_14_PP", "13 A 14": "13_14_PP", "13-14 ANOS": "13_14_PP",
    "GG INFANTIL": "11_12_GG_INF", "INFANTIL GG": "11_12_GG_INF", "11 A 12 ANOS": "11_12_GG_INF", "11 A 12": "11_12_GG_INF", "11-12 ANOS": "11_12_GG_INF",
    "G INFANTIL": "8_10_G_INF", "INFANTIL G": "8_10_G_INF", "8 A 10 ANOS": "8_10_G_INF", "8 A 10": "8_10_G_INF", "8-10 ANOS": "8_10_G_INF",
    "M INFANTIL": "6_7_M_INF", "INFANTIL M": "6_7_M_INF", "6 A 7 ANOS": "6_7_M_INF", "6 A 7": "6_7_M_INF", "6-7 ANOS": "6_7_M_INF",
    "P INFANTIL": "4_5_P_INF", "INFANTIL P": "4_5_P_INF", "4 A 5 ANOS": "4_5_P_INF", "4 A 5": "4_5_P_INF", "4-5 ANOS": "4_5_P_INF",
    "PP INFANTIL": "2_3_PP_INF", "INFANTIL PP": "2_3_PP_INF", "2 A 3 ANOS": "2_3_PP_INF", "2 A 3": "2_3_PP_INF", "2-3 ANOS": "2_3_PP_INF",
}

def versao_tupla(versao: str) -> Tuple[int, ...]:
    return tuple(int(n) for n in re.findall(r"\d+", str(versao)))


def executar_com_retry_gemini(funcao_api, *args, max_tentativas=6, espera_base=3, **kwargs):
    for tentativa in range(1, max_tentativas + 1):
        try:
            return funcao_api(*args, **kwargs)
        except Exception as e:
            erro_str = str(e).lower()
            if any(termo in erro_str for termo in ["exhausted", "503", "overloaded", "resource_exhausted", "quota", "unavailable", "congestion"]):
                if tentativa < max_tentativas:
                    time.sleep(espera_base * tentativa)
                    continue
            raise e


# ==========================================================================
# Parser (independente da interface)
# ==========================================================================
LETRA = "A-Z0-9À-Ÿ"
ANTES = rf"(?<![{LETRA}])"
DEPOIS = rf"(?![{LETRA}])"
SEP = r"[\s\-:/;.,]*"
CALCAO = r"(?:CAL[CÇ][AÃ]O|CALS[AÃ]O|CAUCAO|SHORT[ES]?|BERMUDA)"
CAMISA = r"(?:CAMISETA|CAMISA|CAMIZA|BLUSA)"
VALOR_NUMERO = r"(?:\d+|PI|π)"

RE_IGNORAR = re.compile(r"PATROCÍNIO|NUMERAÇÃO|JOGO|MATERIAL|PENHAROL")
RE_SUBSTITUTO = re.compile(r"^\*?\s*SUB\s*\d+")
RE_ENUMERACAO = re.compile(r"^\s*\d+\s*[-.)]\s*")
RE_SEM_NUMERO = re.compile(r"\bS\s*/\s*N\b")
RE_QUANTIDADE = re.compile(r"^\s*\*?\s*(\d+)\s*(?:X|UN|PCT)\b")
RE_KIT_QTD = re.compile(r"^\s*\*?\s*\d+\s")
RE_BABYLOOK = re.compile(r"\bBABY\s*LOOK\b")
RE_TRADICIONAL = re.compile(r"\bTRADICIONAL\b")
RE_NUM_EXPLICITO = re.compile(rf"{ANTES}(?:N[UÚ]MERO|NUM\.?|N[º°O.]|N)[\s\-:;]*({VALOR_NUMERO}){DEPOIS}|#\s*({VALOR_NUMERO}){DEPOIS}")
RE_NUM_BARRA = re.compile(rf"[/\-]\s*({VALOR_NUMERO}){DEPOIS}")
RE_NUM_SOLTO = re.compile(rf"{ANTES}(\d{{1,3}}|PI|π){DEPOIS}")
RE_PALAVRAS_REMOVER = re.compile(
    r"\b(?:ESCREVER|COM|SEM|NOME|AVULS[AO]S?|CAMISETAS?|CAMISAS?|CAMIZAS?|BLUSAS?|BABY\s*LOOK|TRADICIONAL|ADULTO|INFANTIL|"
    r"KITS?|GOLEIRO|DE|E|CONJUNTOS?|CAL[CÇ][AÃ]O|CALS[AÃ]O|CAUCAO|SHORT[ES]?|BERMUDA|TAMANHO|TAM|T|"
    r"N[UÚ]MERO|NUM|N[º°O]|N)\b\.?|#"
)
RE_PONTAS = re.compile(r"^[\-/:,._–—*\s]+|[\-/:,._–—*\s]+$")
RE_CARACTERES_INVALIDOS = re.compile(r"[^A-Z0-9À-Ÿ'\-\s]")
RE_HIFEN_SOLTO = re.compile(r"(^|\s)[\-']+(\s|$)")
RE_FICHA = re.compile(r"(?:tamanho|nome|n[úu]mero)[/\s]*:", re.IGNORECASE)
RE_SEPARADOR_BLOCOS = re.compile(r"_{3,}|-{3,}")
RE_FICHA_INICIO_TAM = re.compile(r"^\s*(?:TAMANHO|TAM)\b", re.IGNORECASE)
RE_FICHA_TEM_DADO = re.compile(r"NOME|N[ÚU]MERO|\d+", re.IGNORECASE)
RE_FICHA_TAM = re.compile(r"\b(?:TAMANHO|TAM)\b\.?")
RE_FICHA_NUM = re.compile(rf"N[ÚU]MERO[^\d:\n]*:?\s*({VALOR_NUMERO}){DEPOIS}")
RE_FICHA_PREFIXO_NOME = re.compile(r"^(nome(\s+na\s+camisa|\s+da\s+camisa)?)[/\s]*:?\s*", re.IGNORECASE)

def _remover(texto: str, m: re.Match) -> str:
    return f"{texto[:m.start()]} {texto[m.end():]}"


class ParserPedidos:
    def __init__(self, ordem: List[str], conversoes: Dict[str, str]):
        self.ordem = list(ordem)
        self.conversoes = dict(conversoes)
        self.posicao = {t: i for i, t in enumerate(self.ordem)}
        termos = sorted(set(self.ordem) | set(self.conversoes), key=lambda t: (-len(t), t))
        alternativas = "|".join(re.escape(t) for t in termos) or r"(?!)"
        termo = rf"{ANTES}(?P<tam>{alternativas}){DEPOIS}"
        self.re_termo = re.compile(termo)
        self.re_calcao = re.compile(rf"{ANTES}(?:TAMANHO\s+(?:D[AEO]\s+)?{CALCAO}|{CALCAO}\s+TAMANHO|{CALCAO}){SEP}{termo}")
        self.re_camisa = re.compile(
            rf"{ANTES}(?:TAMANHO\s+(?:D[AEO]\s+)?{CAMISA}|{CAMISA}\s+TAMANHO|TAMANHO|TAM\.|TAM|T|{CAMISA}|BABY\s*LOOK|TRADICIONAL){SEP}{termo}"
        )

    def converter(self, tamanho: str) -> str:
        if not tamanho or "SEM" in tamanho: return ""
        return self.conversoes.get(tamanho, tamanho)

    def _babylook(self, tamanho: str, eh_tradicional: bool) -> str:
        if eh_tradicional or not tamanho or tamanho.startswith("BL"):
            return self.converter(tamanho)
        for candidato in ("BL" + tamanho, "BL" + self.converter(tamanho)):
            convertido = self.conversoes.get(candidato, candidato)
            if convertido in self.posicao: return convertido
        return self.converter(tamanho)

    def peso(self, pedido: dict) -> Tuple[int, int]:
        camisa, calcao = pedido["TAMANHO DE CAMISA"], pedido["TAMANHO DE CALÇÃO"]
        if camisa in self.posicao: return (self.posicao[camisa], 0)
        if calcao in self.posicao: return (self.posicao[calcao], 1)
        return (999, 99)

    @staticmethod
    def novo_pedido(nome="", camisa="", numero="", calcao="", num_calcao="") -> dict:
        return {
            "NOME": nome.upper(), 
            "TAMANHO DE CAMISA": camisa, 
            "NÚMERO": numero, 
            "TAMANHO DE CALÇÃO": calcao, 
            "NÚMERO DE CALÇÃO": num_calcao
        }

    def processar(self, texto: str, preencher: bool = True) -> Tuple[List[dict], List[str]]:
        texto = re.sub(r"Número Tabela camisa da equipe:?", "", texto, flags=re.IGNORECASE)
        if RE_FICHA.search(texto) or "___" in texto:
            pedidos, nao_reconhecidas = self._processar_fichas(texto)
        else:
            pedidos, nao_reconhecidas = self._processar_linhas(texto)
        if preencher:
            for p in pedidos:
                for coluna, padrao in PREENCHIMENTO.items():
                    if not p[coluna]: p[coluna] = padrao
        pedidos.sort(key=self.peso)
        return pedidos, nao_reconhecidas

    def _processar_linhas(self, texto: str) -> Tuple[List[dict], List[str]]:
        pedidos: List[dict] = []
        nao_reconhecidas: List[str] = []
        contexto_camisa = contexto_calcao = ""
        for original in texto.splitlines():
            original = original.strip()
            if not original: continue
            linha = RE_ENUMERACAO.sub("", original.upper())
            linha = RE_SEM_NUMERO.sub(" ", linha)
            if RE_IGNORAR.search(linha) or RE_SUBSTITUTO.match(linha): continue

            quantidade = 1
            eh_kit, eh_conjunto = "KIT" in linha, "CONJUNTO" in linha
            eh_babylook = bool(RE_BABYLOOK.search(linha))
            eh_tradicional = bool(RE_TRADICIONAL.search(linha))
            
            if eh_tradicional:
                eh_babylook = False

            if eh_kit: linha = RE_KIT_QTD.sub(" ", linha, count=1)
            else:
                m = RE_QUANTIDADE.search(linha)
                if m: quantidade, linha = int(m.group(1)), _remover(linha, m)
            if "AVULSA" in linha or "AVULSO" in linha: contexto_calcao = ""

            tam_calcao = tam_camisa = ""
            m = self.re_calcao.search(linha)
            if m: tam_calcao, linha = m.group("tam"), _remover(linha, m)
            m = self.re_camisa.search(linha)
            if m: tam_camisa, linha = m.group("tam"), _remover(linha, m)

            soltos = []
            while True:
                m = self.re_termo.search(linha)
                if not m: break
                soltos.append(m.group("tam"))
                linha = _remover(linha, m)
            if not tam_camisa and soltos: tam_camisa = soltos.pop(0)
            if not tam_calcao and soltos: tam_calcao = soltos.pop(0)

            contexto_camisa = tam_camisa or contexto_camisa
            contexto_calcao = tam_calcao or contexto_calcao
            camisa_final = tam_camisa or contexto_camisa
            calcao_final = tam_calcao or contexto_calcao
            if eh_conjunto and not calcao_final and camisa_final: calcao_final = camisa_final

            numero = ""
            for regex in (RE_NUM_EXPLICITO, RE_NUM_BARRA, RE_NUM_SOLTO):
                m = regex.search(linha)
                if m:
                    numero, linha = next(g for g in m.groups() if g).upper(), _remover(linha, m)
                    break

            # Verificar se há número específico para o calção
            num_calcao = ""
            re_num_calc = re.compile(rf"{CALCAO}[^\d\n]*({VALOR_NUMERO})")
            m_nc = re_num_calc.search(linha)
            if m_nc:
                num_calcao, linha = m_nc.group(1).upper(), _remover(linha, m_nc)
            elif calcao_final and numero:
                num_calcao = numero # Herança automática para conjuntos

            linha = RE_PALAVRAS_REMOVER.sub(" ", linha)
            linha = RE_PONTAS.sub("", linha)
            linha = RE_CARACTERES_INVALIDOS.sub(" ", linha)
            nome = re.sub(r"\s+", " ", linha).strip()
            nome = RE_HIFEN_SOLTO.sub(r"\1\2", nome).strip()
            if nome.isdigit(): nome = ""

            original_up = original.upper()
            eh_pedido = bool(numero or nome) if eh_kit else bool(numero or nome or "CAMISA" in original_up or "BABY" in original_up or "TRADICIONAL" in original_up)
            if eh_pedido and "AVULSAS" in original_up and not tam_camisa: eh_pedido = False

            if not eh_pedido:
                nao_reconhecidas.append(original)
                continue

            camisa = self._babylook(camisa_final, eh_tradicional) if eh_babylook else self.converter(camisa_final)
            calcao = self.converter(calcao_final)
            num_camisa_final = ("" if "SEM" in numero else numero)
            num_calcao_final = ("" if "SEM" in num_calcao else num_calcao) if calcao else ""
            
            pedidos.extend(self.novo_pedido(nome, camisa, num_camisa_final, calcao, num_calcao_final) for _ in range(quantidade))
        return pedidos, nao_reconhecidas

    def _dividir_blocos(self, texto: str) -> List[str]:
        if RE_SEPARADOR_BLOCOS.search(texto): return [b.strip() for b in RE_SEPARADOR_BLOCOS.split(texto) if b.strip()]
        blocos, atual = [], []
        for linha in (l.strip() for l in texto.splitlines()):
            if not linha: continue
            if RE_FICHA_INICIO_TAM.match(linha) and any(RE_FICHA_TEM_DADO.search(l) for l in atual):
                blocos.append("\n".join(atual))
                atual = [linha]
            else: atual.append(linha)
        if atual: blocos.append("\n".join(atual))
        return blocos

    def _processar_fichas(self, texto: str) -> Tuple[List[dict], List[str]]:
        pedidos: List[dict] = []
        nao_reconhecidas: List[str] = []
        for bloco in self._dividir_blocos(texto):
            tam_camisa = tam_calcao = nome = numero = num_calcao = ""
            eh_conjunto = "CONJUNTO" in bloco.upper()
            for linha in bloco.splitlines():
                original = linha.strip()
                if not original: continue
                up = original.upper()

                m = RE_FICHA_TAM.search(up)
                if m:
                    mt = self.re_termo.search(up[m.end():]) or re.search(r"(?P<tam>[A-Z0-9_]+)", up[m.end():])
                    if mt:
                        if re.search(CALCAO, up): tam_calcao = mt.group("tam")
                        else: tam_camisa = mt.group("tam")
                    continue
                m = RE_FICHA_NUM.search(up)
                if m:
                    if re.search(CALCAO, up): num_calcao = m.group(1)
                    else: numero = m.group(1)
                    continue
                if "NOME" in up:
                    partes = original.split(":", 1)
                    if len(partes) > 1 and partes[1].strip(): nome = partes[1].strip()
                    continue
                if re.fullmatch(r"\d+|PI|π", up):
                    numero = numero or up
                    continue
                m = self.re_termo.match(up)
                if m and not tam_camisa:
                    tam_camisa, original = m.group("tam"), original[m.end():].strip()
                    up = original.upper()
                    if not original: continue
                if not nome and not any(k in up for k in ("TAMANHO", "NUMERO", "NÚMERO", "CONJUNTO")):
                    nome = original

            nome = re.sub(r"[\xa0\t]+", " ", RE_FICHA_PREFIXO_NOME.sub("", nome)).strip()
            camisa = self.converter(tam_camisa)
            calcao = self.converter(tam_calcao) or (camisa if eh_conjunto else "")
            if calcao and not num_calcao:
                num_calcao = numero
            if nome or numero or camisa: 
                pedidos.append(self.novo_pedido(nome, camisa, numero, calcao, num_calcao))
            else: 
                nao_reconhecidas.append(bloco)
        return pedidos, nao_reconhecidas


# ==========================================================================
# Interface (Padrão Tailwind CSS)
# ==========================================================================
PALETAS = {
    "claro": {
        "bg": "#F1F5F9",
        "painel": "#FFFFFF",
        "borda": "#E2E8F0",
        "texto": "#0F172A",
        "texto2": "#64748B",
        "cabecalho": "#1E293B",
        "aviso_bg": "#FEF2F2",
        "aviso_fg": "#DC2626",
        "aviso_txt": "#991B1B",
        "dup_nome": "#FEE2E2",
        "dup_num": "#FEF3C7",
    },
    "escuro": {
        "bg": "#0F172A",
        "painel": "#1E293B",
        "borda": "#334155",
        "texto": "#F8FAFC",
        "texto2": "#94A3B8",
        "cabecalho": "#020617",
        "aviso_bg": "#450A0A",
        "aviso_fg": "#FCA5A5",
        "aviso_txt": "#FECACA",
        "dup_nome": "#7F1D1D",
        "dup_num": "#78350F",
    },
}

class AplicativoPedidosMagico:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.geometry("1280x920")
        self.root.minsize(1050, 720)

        self.ordem_tamanhos = list(ORDEM_PADRAO)
        self.conversoes = dict(CONVERSOES_PADRAO)
        self.preencher_padrao = tk.BooleanVar(value=True)
        self.tema_escuro = tk.BooleanVar(value=False)
        self.gemini_api_key = "AQ.Ab8RN6LoKvoFgnIxTdiKmA08tU93Ldtmv3IYkw0SvngS59H7tA"
        self.cliente_gemini = None
        
        self.carregar_config()
        self.parser = ParserPedidos(self.ordem_tamanhos, self.conversoes)

        self.pedidos_atuais: List[dict] = []
        self.linhas_nao_reconhecidas: List[str] = []
        self.historico: List[List[dict]] = []
        self.futuro: List[List[dict]] = []
        self.arquivo_atual: Optional[str] = None
        self._agrupar = True
        self._ordenacao_reversa: Dict[str, bool] = {}
        self._linhas_visiveis: List[Optional[dict]] = []
        self._tarefa_ocupada = False
        self._cores_atuais = PALETAS["claro"]

        self.criar_interface()
        self._configurar_atalhos()
        self._atualizar_titulo()
        if self.tema_escuro.get(): self.alternar_tema()
        self.root.protocol("WM_DELETE_WINDOW", self._fechar)

        self.root.after(2000, lambda: self.verificar_atualizacoes(silencioso=True))

    def carregar_config(self) -> None:
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                dados = json.load(f)
        except (OSError, ValueError): return
        if not isinstance(dados, dict): return
        if isinstance(dados.get("ordem_tamanhos"), list) and dados["ordem_tamanhos"]:
            self.ordem_tamanhos = [str(t) for t in dados["ordem_tamanhos"]]
        if isinstance(dados.get("conversoes"), dict):
            self.conversoes = {str(k): str(v) for k, v in dados["conversoes"].items()}
        self.preencher_padrao.set(bool(dados.get("preencher_padrao", True)))
        self.tema_escuro.set(bool(dados.get("tema_escuro", False)))
        if dados.get("gemini_api_key"):
            self.gemini_api_key = str(dados.get("gemini_api_key"))

    def salvar_config(self) -> None:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({
                    "ordem_tamanhos": self.ordem_tamanhos, 
                    "conversoes": self.conversoes,
                    "preencher_padrao": self.preencher_padrao.get(), 
                    "tema_escuro": self.tema_escuro.get(),
                    "gemini_api_key": self.gemini_api_key
                }, f, ensure_ascii=False, indent=2)
        except OSError: pass

    def _tratar_erro_gemini(self, e: Exception) -> str:
        erro_str = str(e)
        if "429" in erro_str or "RESOURCE_EXHAUSTED" in erro_str or "quota" in erro_str.lower():
            m = re.search(r"retry in\s*([\d\.]+)\s*s", erro_str, re.IGNORECASE)
            prazo_txt = "em alguns instantes"
            if m:
                segundos = float(m.group(1))
                if segundos < 60:
                    prazo_txt = f"em aproximadamente {int(segundos) + 1} segundos"
                else:
                    minutos = int(segundos // 60)
                    prazo_txt = f"em aproximadamente {minutos} minuto(s)"
            return f"⚠️ Limite de requisições da IA atingido!\n\nVocê excedeu a cota gratuita do plano atual.\nO serviço deverá voltar a funcionar {prazo_txt}.\n\nDica: Aguarde o prazo acima, utilize a leitura offline do Tesseract ou configure outra chave de API."
        return f"Erro de comunicação com a IA:\n\n{e}"

    def _obter_cliente_gemini(self):
        if not GEMINI_DISPONIVEL:
            return None
        if self.cliente_gemini:
            return self.cliente_gemini
        
        if not self.gemini_api_key or self.gemini_api_key == "COLE_SUA_CHAVE_GEMINI_AQUI":
            self.pedir_api_key()
            
        if not self.gemini_api_key or self.gemini_api_key == "COLE_SUA_CHAVE_GEMINI_AQUI":
            return None

        try:
            os.environ["GEMINI_API_KEY"] = self.gemini_api_key
            self.cliente_gemini = genai.Client(api_key=self.gemini_api_key)
            return self.cliente_gemini
        except Exception as e:
            messagebox.showerror("Erro de Autenticação", f"A chave de API fornecida é inválida:\n{e}", parent=self.root)
            self.cliente_gemini = None
            return None

    def pedir_api_key(self) -> None:
        atual = self.gemini_api_key if self.gemini_api_key != "COLE_SUA_CHAVE_GEMINI_AQUI" else ""
        nova = simpledialog.askstring(
            "Chave de API do Gemini", 
            "Insira a sua chave de API do Gemini (começada por AIzaSy...):", 
            initialvalue=atual, 
            show='*', 
            parent=self.root
        )
        if nova is not None:
            self.gemini_api_key = nova.strip()
            self.cliente_gemini = None
            self.salvar_config()
            if self.gemini_api_key:
                messagebox.showinfo("Sucesso", "Chave de API guardada com sucesso!", parent=self.root)

    def _fechar(self) -> None:
        self.salvar_config()
        self.root.destroy()

    def _na_interface(self, funcao: Callable, *args) -> None:
        try: self.root.after(0, funcao, *args)
        except (RuntimeError, tk.TclError): pass

    def set_status(self, msg: str, progresso: Optional[int] = None) -> None:
        if threading.current_thread() is not threading.main_thread():
            self._na_interface(self.set_status, msg, progresso)
            return
        self.status_var.set(f"{datetime.now():%H:%M:%S}  |  {msg}")
        if progresso is not None: self.progress_bar["value"] = progresso

    def _em_segundo_plano(self, alvo: Callable, *args) -> bool:
        if self._tarefa_ocupada:
            messagebox.showinfo("Aguarde", "Já existe uma tarefa em andamento.", parent=self.root)
            return False
        self._tarefa_ocupada = True
        for b in self.botoes_ia: b.config(state="disabled")
        
        self._na_interface(lambda: [
            self.progress_bar.config(mode="indeterminate"),
            self.progress_bar.start(10)
        ])
        
        def executar():
            try: 
                alvo(*args)
            finally:
                self._na_interface(lambda: [
                    self.progress_bar.stop(),
                    self.progress_bar.config(mode="determinate"),
                    self.progress_bar.configure(value=100),
                    setattr(self, '_tarefa_ocupada', False), 
                    [b.config(state="normal") for b in self.botoes_ia]
                ])
        threading.Thread(target=executar, daemon=True).start()
        return True

    def verificar_atualizacoes(self, silencioso: bool = True) -> None:
        self.set_status("🔍 A verificar atualizações...", 15)
        def tarefa():
            try:
                req = urllib.request.Request(URL_VERSAO_REMOTE, headers={"User-Agent": "Organizador"})
                with urllib.request.urlopen(req, timeout=8) as resp:
                    self._na_interface(self._resultado_atualizacao, json.loads(resp.read().decode("utf-8")), silencioso)
            except Exception:
                self.set_status("⚠️ Sem ligação para verificar atualizações.", 0)
        threading.Thread(target=tarefa, daemon=True).start()

    def _resultado_atualizacao(self, dados: dict, silencioso: bool) -> None:
        remota = str(dados.get("version", ""))
        if not remota or versao_tupla(remota) <= versao_tupla(__version__):
            self.set_status(f"✅ Versão {__version__} atualizada.", 100)
            if not silencioso: messagebox.showinfo("Atualizações", "Já tem a versão mais recente!", parent=self.root)
            return
        self.set_status(f"Nova versão {remota} disponível!", 100)
        if messagebox.askyesno("Atualização", f"Nova versão {remota} disponível!\nNovidades:\n{dados.get('changelog', '')}\n\nAtualizar agora?", parent=self.root):
            threading.Thread(target=self._baixar_atualizacao, args=(dados.get("url", ""), dados.get("sha256")), daemon=True).start()

    def _baixar_atualizacao(self, url: str, sha256_esperado: Optional[str]) -> None:
        caminho_atual = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
        caminho_temp = caminho_atual + ".novo"
        try:
            def progresso(bloco, tamanho_bloco, total):
                if total > 0: self.set_status(f"A descarregar... {min(90, 10 + bloco * tamanho_bloco * 80 // total)}%", 50)
            urllib.request.urlretrieve(url, caminho_temp, reporthook=progresso)
            with open(caminho_temp, "rb") as f:
                if sha256_esperado and hashlib.sha256(f.read()).hexdigest().lower() != str(sha256_esperado).lower():
                    raise ValueError("SHA256 não confere")
        except Exception as e:
            self.set_status("Falha na atualização.", 0)
            return
        self._na_interface(self._aplicar_atualizacao, caminho_atual, caminho_temp)

    def _aplicar_atualizacao(self, caminho_atual: str, caminho_temp: str) -> None:
        try:
            if caminho_atual.lower().endswith(".py"):
                shutil.copy2(caminho_atual, caminho_atual + ".bak")
                os.replace(caminho_temp, caminho_atual)
                subprocess.Popen([sys.executable, caminho_atual, *sys.argv[1:]])
                self._fechar()
                return
            bat = os.path.join(tempfile.gettempdir(), "atualizar.bat")
            with open(bat, "w", encoding="utf-8") as f:
                f.write(f'@echo off\r\nping -n 2 127.0.0.1 > nul\r\nmove /y "{caminho_temp}" "{caminho_atual}"\r\nstart "" "{caminho_atual}"\r\ndel "%~f0"\r\n')
            subprocess.Popen(["cmd", "/c", bat], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            self._fechar()
        except Exception: pass

    def _checar_gemini(self) -> bool:
        if not GEMINI_DISPONIVEL:
            messagebox.showwarning("Biblioteca em Falta", "A biblioteca do Gemini não está instalada.", parent=self.root)
            return False
        cli = self._obter_cliente_gemini()
        if cli: 
            return True
        return False

    # ---------------------------------------------------------------- interface
    def criar_interface(self) -> None:
        estilo = ttk.Style()
        estilo.theme_use("clam")
        self.estilo = estilo
        self.botoes_ia = []

        menubar = tk.Menu(self.root)
        m_arq = tk.Menu(menubar, tearoff=0)
        m_arq.add_command(label="Abrir lista (.txt)...", command=self.abrir_txt, accelerator="Ctrl+O")
        m_arq.add_command(label="Colar da área de transferência", command=self.colar_da_area_transferencia)
        m_arq.add_separator()
        m_arq.add_command(label="Exportar Excel (.xlsx)", command=self.salvar_em_xlsx, accelerator="Ctrl+S")
        m_arq.add_separator()
        m_arq.add_command(label="Sair", command=self._fechar)
        menubar.add_cascade(label="Ficheiro", menu=m_arq)

        m_edit = tk.Menu(menubar, tearoff=0)
        m_edit.add_command(label="Desfazer (tabela)", command=self.desfazer, accelerator="Ctrl+Z")
        m_edit.add_command(label="Refazer (tabela)", command=self.refazer, accelerator="Ctrl+Y")
        m_edit.add_separator()
        m_edit.add_command(label="Adicionar linha manual", command=self.adicionar_linha_manual)
        m_edit.add_command(label="Excluir linha(s)", command=self.excluir_linhas_selecionadas, accelerator="Delete")
        menubar.add_cascade(label="Editar", menu=m_edit)

        m_conf = tk.Menu(menubar, tearoff=0)
        m_conf.add_command(label="Configurar Chave API Gemini...", command=self.pedir_api_key)
        m_conf.add_command(label="Tamanhos e conversões...", command=self.abrir_editor_regras)
        m_conf.add_checkbutton(label="Preencher campos vazios", variable=self.preencher_padrao, command=self.salvar_config)
        m_conf.add_checkbutton(label="Tema escuro", variable=self.tema_escuro, command=self.alternar_tema)
        menubar.add_cascade(label="Configurações", menu=m_conf)
        self.root.config(menu=menubar)

        c = PALETAS["claro"]
        self.root.configure(bg=c["bg"])
        
        # CABEÇALHO
        self.header_frame = tk.Frame(self.root, bg=c["cabecalho"], pady=20)
        self.header_frame.pack(fill="x")
        tk.Label(self.header_frame, text="ORGANIZADOR INTELIGENTE DE UNIFORMES — TURBO", font=("Segoe UI", 18, "bold"), bg=c["cabecalho"], fg="#FFFFFF").pack()
        tk.Label(self.header_frame, text=f"Versão {__version__} | Integrado com Gemini e Tesseract Offline", font=("Segoe UI", 10), bg=c["cabecalho"], fg="#94A3B8").pack()

        # RODAPÉ
        rodape = tk.Frame(self.root, bg="#0F172A", pady=10, padx=25)
        rodape.pack(fill="x", side="bottom")
        barra = tk.Frame(rodape, bg="#0F172A")
        barra.pack(fill="x", pady=(0, 5))
        self.status_var = tk.StringVar(value="🟢 A iniciar...")
        tk.Label(barra, textvariable=self.status_var, font=("Segoe UI", 10, "bold"), bg="#0F172A", fg="#F8FAFC", anchor="w").pack(side="left", fill="x", expand=True)
        self.progress_bar = ttk.Progressbar(barra, orient="horizontal", length=220, mode="determinate", maximum=100)
        self.progress_bar.pack(side="right", padx=(10, 0))

        # CONTAINER PRINCIPAL
        self.main_wrap = tk.Frame(self.root, bg=c["bg"])
        self.main_wrap.pack(fill="both", expand=True, padx=25, pady=20)

        # CARD 1: ENTRADA
        self.card_in = tk.Frame(self.main_wrap, bg=c["painel"], highlightthickness=1, highlightbackground=c["borda"], bd=0)
        self.card_in.pack(fill="x", pady=(0, 20))

        box_in_header = tk.Frame(self.card_in, bg=c["painel"])
        box_in_header.pack(fill="x", padx=20, pady=(15, 10))
        tk.Label(box_in_header, text="📝 Colar Lista (Texto ou Imagem)", font=("Segoe UI", 12, "bold"), bg=c["painel"], fg=c["texto"]).pack(side="left")

        box_img_tools = tk.Frame(box_in_header, bg=c["painel"])
        box_img_tools.pack(side="right")
        estilo_btn_pq = dict(fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=8, pady=5)
        
        # Botões Gemini
        btn_ocr1 = tk.Button(box_img_tools, text="🖼️ Img (Gemini)", command=self.ler_imagem_com_gemini, bg="#7C3AED", **estilo_btn_pq)
        btn_ocr1.pack(side="left", padx=(0, 3))
        btn_ocr2 = tk.Button(box_img_tools, text="📋 Clip (Gemini)", command=self.colar_imagem_com_gemini, bg="#7C3AED", **estilo_btn_pq)
        btn_ocr2.pack(side="left", padx=(0, 8))
        self.botoes_ia.extend([btn_ocr1, btn_ocr2])

        # Botões Tesseract Offline
        btn_tess1 = tk.Button(box_img_tools, text="🖥️ Img (Offline)", command=self.ler_imagem_com_tesseract, bg="#0D9488", **estilo_btn_pq)
        btn_tess1.pack(side="left", padx=(0, 3))
        btn_tess2 = tk.Button(box_img_tools, text="📋 Clip (Offline)", command=self.colar_imagem_com_tesseract, bg="#0D9488", **estilo_btn_pq)
        btn_tess2.pack(side="left", padx=(0, 0))
        self.botoes_ia.extend([btn_tess1, btn_tess2])

        self.text_area = scrolledtext.ScrolledText(self.card_in, height=6, font=("Segoe UI", 10), bd=1, relief="solid", bg=c["bg"], fg=c["texto"])
        self.text_area.pack(fill="x", padx=20, pady=(0, 15))

        box_actions = tk.Frame(self.card_in, bg=c["painel"])
        box_actions.pack(fill="x", padx=20, pady=(0, 15))
        box_actions.columnconfigure((0, 1, 2), weight=1, uniform="actions")
        
        estilo_btn_gr = dict(fg="white", font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2", pady=8)
        tk.Button(box_actions, text="✨ SEPARAR E ORGANIZAR (Ctrl+Enter)", command=self.processar_texto, bg="#2563EB", **estilo_btn_gr).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        btn_ia_principal = tk.Button(box_actions, text="🧠 ORGANIZAR COM GEMINI", command=self.processar_texto_com_gemini_thread, bg="#059669", **estilo_btn_gr)
        btn_ia_principal.grid(row=0, column=1, sticky="ew", padx=5)
        self.botoes_ia.append(btn_ia_principal)
        tk.Button(box_actions, text="🗑️ Limpar tudo", command=self.limpar_tudo, bg="#64748B", **estilo_btn_gr).grid(row=0, column=2, sticky="ew", padx=(5, 0))

        # CARD 2: RESULTADOS
        self.card_out = tk.Frame(self.main_wrap, bg=c["painel"], highlightthickness=1, highlightbackground=c["borda"], bd=0)
        self.card_out.pack(fill="both", expand=True)

        box_out_header = tk.Frame(self.card_out, bg=c["painel"])
        box_out_header.pack(fill="x", padx=20, pady=(15, 5))
        tk.Label(box_out_header, text="🔍 Tabela de Pedidos", font=("Segoe UI", 12, "bold"), bg=c["painel"], fg=c["texto"]).pack(side="left")

        box_filter = tk.Frame(box_out_header, bg=c["painel"])
        box_filter.pack(side="right")
        tk.Label(box_filter, text="Filtrar:", bg=c["painel"], fg=c["texto2"], font=("Segoe UI", 10)).pack(side="left", padx=5)
        self.var_filtro = tk.StringVar()
        self.var_filtro.trace_add("write", lambda *_: self._redesenhar_tabela())
        self.entrada_filtro = tk.Entry(box_filter, textvariable=self.var_filtro, font=("Segoe UI", 10), width=30, bg=c["bg"], fg=c["texto"], relief="solid", bd=1)
        self.entrada_filtro.pack(side="left")

        self.lbl_resumo = tk.Label(self.card_out, text="Nenhum pedido gerado ainda.", bg=c["painel"], fg=c["texto2"], font=("Segoe UI", 9, "bold"))
        self.lbl_resumo.pack(fill="x", padx=20, anchor="w", pady=(0, 10))

        self.box_export = tk.Frame(self.card_out, bg=c["painel"])
        self.box_export.pack(fill="x", padx=20, pady=(0, 20), side="bottom")
        self.box_export.columnconfigure((0, 1, 2), weight=1, uniform="exports")
        tk.Button(self.box_export, text="📋 COPIAR PARA EXCEL", command=self.copiar_para_excel, bg="#16A34A", **estilo_btn_gr).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        tk.Button(self.box_export, text="📊 SALVAR COMO EXCEL (.XLSX)", command=self.salvar_em_xlsx, bg="#0EA5E9", **estilo_btn_gr).grid(row=0, column=1, sticky="ew", padx=5)
        tk.Button(self.box_export, text="💾 SALVAR FICHEIRO (.CSV)", command=self.salvar_em_csv, bg="#D97706", **estilo_btn_gr).grid(row=0, column=2, sticky="ew", padx=(5, 0))

        frame_tabela = tk.Frame(self.card_out, bg=c["borda"], bd=0)
        frame_tabela.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        
        scroll_y = ttk.Scrollbar(frame_tabela)
        scroll_y.pack(side="right", fill="y")
        scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")
        self.tree = ttk.Treeview(frame_tabela, columns=COLUNAS, show="headings", selectmode="extended", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        for col in COLUNAS:
            self.tree.heading(col, text=col, command=lambda c_=col: self.ordenar_por_coluna(c_))
            self.tree.column(col, width=220, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=1, pady=1)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree.bind("<Double-1>", self.editar_celula)
        self.tree.bind("<Delete>", lambda _e: self.excluir_linhas_selecionadas())

        self.frame_avisos = tk.Frame(self.card_out, bg=c["aviso_bg"], highlightthickness=1, highlightbackground=c["aviso_fg"], bd=0)
        tk.Label(self.frame_avisos, text="⚠️ Linhas não reconhecidas (revise manualmente):", font=("Segoe UI", 9, "bold"), bg=c["aviso_bg"], fg=c["aviso_fg"]).pack(anchor="w", padx=10, pady=(10, 0))
        self.txt_avisos = tk.Text(self.frame_avisos, height=3, font=("Segoe UI", 9), bd=0, bg=c["aviso_bg"], fg=c["aviso_txt"], wrap="word", state="disabled")
        self.txt_avisos.pack(fill="both", expand=True, padx=10, pady=10)

        self._aplicar_estilo_tabela(c)

    def _aplicar_estilo_tabela(self, p: dict) -> None:
        self.estilo.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=p["borda"], foreground=p["texto"], borderwidth=0, relief="flat")
        self.estilo.map("Treeview.Heading", background=[("active", p["borda"])])
        self.estilo.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background=p["bg"], fieldbackground=p["bg"], foreground=p["texto"], borderwidth=0)
        self.tree.tag_configure("dup_nome", background=p["dup_nome"])
        self.tree.tag_configure("dup_num", background=p["dup_num"])

    def _configurar_atalhos(self) -> None:
        def so_fora_de_campos(acao):
            def handler(_e):
                if isinstance(self.root.focus_get(), (tk.Text, tk.Entry)): return None
                acao()
                return "break"
            return handler
        def ctrl_enter(_e):
            self.processar_texto()
            return "break"
        self.text_area.bind("<Control-Return>", ctrl_enter)
        self.root.bind("<Control-Return>", ctrl_enter)
        self.root.bind("<Control-z>", so_fora_de_campos(self.desfazer))
        self.root.bind("<Control-y>", so_fora_de_campos(self.refazer))
        self.root.bind("<Control-f>", lambda _e: self.entrada_filtro.focus_set())
        self.root.bind("<Control-s>", lambda _e: self.salvar_em_xlsx())
        self.root.bind("<Control-o>", lambda _e: self.abrir_txt())
        self.root.bind("<Control-C>", lambda _e: self.copiar_para_excel())
        self.entrada_filtro.bind("<Escape>", lambda _e: self.var_filtro.set(""))

    def _atualizar_titulo(self) -> None:
        extra = f" — {os.path.basename(self.arquivo_atual)}" if self.arquivo_atual else ""
        self.root.title(f"Organizador Inteligente de Pedidos{extra}")

    def alternar_tema(self) -> None:
        novo = PALETAS["escuro" if self.tema_escuro.get() else "claro"]
        antigo = self._cores_atuais
        mapa = {antigo[k].upper(): novo[k] for k in ("bg", "painel", "borda", "texto", "texto2", "cabecalho", "aviso_bg", "aviso_fg", "aviso_txt", "dup_nome", "dup_num")}
        mapa[antigo["cabecalho"].upper()] = novo["cabecalho"]

        def trocar(widget):
            for opcao in ("bg", "fg", "insertbackground", "highlightbackground"):
                try:
                    atual = str(widget.cget(opcao)).upper()
                    if atual in mapa: widget.configure(**{opcao: mapa[atual]})
                except tk.TclError: pass
            for filho in widget.winfo_children(): trocar(filho)

        trocar(self.root)
        self._aplicar_estilo_tabela(novo)
        self._cores_atuais = novo
        self.salvar_config()
        self.set_status("Tema atualizado.")

    def _salvar_estado_para_undo(self) -> None:
        self.historico.append(copy.deepcopy(self.pedidos_atuais))
        del self.historico[:-HISTORICO_MAXIMO]
        self.futuro.clear()

    def desfazer(self) -> None:
        if not self.historico: return self.set_status("Nada para desfazer.")
        self.futuro.append(copy.deepcopy(self.pedidos_atuais))
        self.pedidos_atuais = self.historico.pop()
        self._redesenhar_tabela()
        self.set_status("Ação desfeita.")

    def refazer(self) -> None:
        if not self.futuro: return self.set_status("Nada para refazer.")
        self.historico.append(copy.deepcopy(self.pedidos_atuais))
        self.pedidos_atuais = self.futuro.pop()
        self._redesenhar_tabela()
        self.set_status("Ação refeita.")

    def _texto_entrada(self) -> Optional[str]:
        texto = self.text_area.get("1.0", tk.END).strip()
        if not texto: messagebox.showwarning("Aviso", "A caixa de texto está vazia!", parent=self.root)
        return texto or None

    def _aplicar_resultado(self, pedidos: List[dict], nao_reconhecidas: List[str], origem: str) -> None:
        self._salvar_estado_para_undo()
        self.pedidos_atuais = pedidos
        self.linhas_nao_reconhecidas = nao_reconhecidas
        self._agrupar = True
        self._redesenhar_tabela()
        self._atualizar_avisos()
        extra = f" {len(nao_reconhecidas)} linha(s) não reconhecida(s)." if nao_reconhecidas else ""
        self.set_status(f"Sucesso! {len(pedidos)} pedido(s) organizado(s) ({origem}).{extra}", 100)

    def processar_texto(self) -> None:
        texto = self._texto_entrada()
        if texto is None: return
        self.set_status("A processar (modo local)...", 30)
        pedidos, nao_reconhecidas = self.parser.processar(texto, self.preencher_padrao.get())
        self._aplicar_resultado(pedidos, nao_reconhecidas, "modo local")

    # ==========================================================================
    # Organização de Texto com Gemini + Sistema Anticongestionamento (Retry)
    # ==========================================================================
    def processar_texto_com_gemini_thread(self) -> None:
        texto = self._texto_entrada()
        if texto is None or not self._checar_gemini(): return
        if self._em_segundo_plano(self._processar_texto_com_gemini_exec, texto):
            self.set_status("A IA (Gemini) está a organizar os dados...", 40)

    def _prompt_ia(self, texto: str) -> str:
        conversoes = ", ".join(f"{k} → {v}" for k, v in self.conversoes.items())
        return (f"Você organiza listas de uniformes. Retorne um JSON estrito no formato: "
                f'{{"pedidos": [{{"nome": "", "camisa": "", "numero": "", "calcao": "", "num_calcao": ""}}]}}\n'
                f"- Nomes em MAIÚSCULAS.\n"
                f"- A palavra 'tradicional' significa camiseta normal (NÃO é baby look, não use o prefixo BL).\n"
                f"- A expressão 'baby look' significa blusa baby look (use prefixo BL se necessário conforme conversões).\n"
                f"- Tamanhos válidos: {', '.join(self.ordem_tamanhos)}.\n"
                f"- Converta: {conversoes}.\n"
                f"- Se tiver CONJUNTO/CALÇÃO, preencha o tamanho do calção e o número do calção (se não especificado separadamente, use o mesmo número da camisa).\n"
                f"Texto:\n{texto}")

    def _normalizar_pedido_ia(self, nome, camisa, numero, calcao, num_calcao) -> dict:
        limpo = lambda v: str(v or "").strip().upper()
        p_num = limpo(numero)
        p_calc = self.parser.converter(limpo(calcao))
        p_num_c = limpo(num_calcao)
        if p_calc and not p_num_c:
            p_num_c = p_num

        pedido = ParserPedidos.novo_pedido(
            limpo(nome), 
            self.parser.converter(limpo(camisa)), 
            p_num, 
            p_calc, 
            p_num_c
        )
        if pedido["NOME"] == "SEM NOME": pedido["NOME"] = ""
        if "SEM" in pedido["NÚMERO"]: pedido["NÚMERO"] = ""
        if "SEM" in pedido["NÚMERO DE CALÇÃO"]: pedido["NÚMERO DE CALÇÃO"] = ""
        if self.preencher_padrao.get():
            for coluna, padrao in PREENCHIMENTO.items(): pedido[coluna] = pedido[coluna] or padrao
        return pedido

    def _interpretar_resposta_ia(self, resposta: str) -> List[dict]:
        resposta = resposta.replace("```json", "").replace("```csv", "").replace("```", "").strip()
        pedidos: List[dict] = []
        try:
            dados = json.loads(resposta)
            itens = dados.get("pedidos", []) if isinstance(dados, dict) else dados
            for it in itens if isinstance(itens, list) else []:
                if isinstance(it, dict): 
                    pedidos.append(self._normalizar_pedido_ia(
                        it.get("nome"), 
                        it.get("camisa"), 
                        it.get("numero"), 
                        it.get("calcao"), 
                        it.get("num_calcao") or it.get("numero_calcao")
                    ))
        except ValueError:
            for linha in resposta.splitlines():
                partes = [p.strip() for p in linha.split(";")]
                if len(partes) >= 4 and partes[0].upper() != "NOME": 
                    nc = partes[4] if len(partes) > 4 else partes[2]
                    pedidos.append(self._normalizar_pedido_ia(partes[0], partes[1], partes[2], partes[3], nc))
        return [p for p in pedidos if any(p[c] not in VALORES_VAZIOS for c in COLUNAS)]

    def _processar_texto_com_gemini_exec(self, texto: str) -> None:
        try:
            cli = self._obter_cliente_gemini()
            if not cli: return
            def chamada():
                return cli.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=self._prompt_ia(texto),
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.0
                    )
                )
            resposta = executar_com_retry_gemini(chamada)
            conteudo = resposta.text.strip()
            pedidos = self._interpretar_resposta_ia(conteudo)
        except Exception as e:
            msg_amigavel = self._tratar_erro_gemini(e)
            self._na_interface(lambda: messagebox.showwarning("Limite Atingido / Erro IA", msg_amigavel, parent=self.root))
            self.set_status("Limite de uso da IA atingido.", 0)
            return
            
        if not pedidos:
            self.set_status("A IA não retornou dados válidos.", 0)
            return
            
        pedidos.sort(key=self.parser.peso)
        self._na_interface(self._aplicar_resultado, pedidos, [], "IA Gemini")

    # ==========================================================================
    # Leitura de Imagens com Gemini (Visão)
    # ==========================================================================
    def ler_imagem_com_gemini(self) -> None:
        if not self._checar_gemini(): return
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if caminho: self._em_segundo_plano(self._visao_gemini_exec, caminho)

    def colar_imagem_com_gemini(self) -> None:
        if not self._checar_gemini(): return
        imagem = self._imagem_da_area_transferencia()
        if imagem is None: return
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            imagem.convert("RGB").save(tmp, "PNG")
        self._em_segundo_plano(self._visao_gemini_exec, tmp.name, True)

    def _imagem_da_area_transferencia(self):
        if not PIL_DISPONIVEL:
            messagebox.showerror("Erro", "Biblioteca Pillow em falta: pip install pillow", parent=self.root)
            return None
        try: conteudo = ImageGrab.grabclipboard()
        except Exception: conteudo = None
        if isinstance(conteudo, list):
            caminhos = [c for c in conteudo if c.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".webp"))]
            conteudo = Image.open(caminhos[0]) if caminhos else None
        if conteudo is None: messagebox.showinfo("Aviso", "Não há nenhuma imagem copiada no clipboard.", parent=self.root)
        return conteudo

    def _visao_gemini_exec(self, caminho: str, apagar_depois: bool = False) -> None:
        prompt = (
            "Esta imagem contém uma lista manuscrita de pedidos de uniformes desportivos. "
            "Transcreva EXATAMENTE linha por linha todo o texto visível e manuscrito com máxima precisão. "
            "Não adicione introduções, comentários nem saudações, apenas a transcrição limpa."
        )
        try:
            self.set_status("A ler a imagem com o Gemini...", 50)
            img = Image.open(caminho)
            cli = self._obter_cliente_gemini()
            if not cli: return
            
            def chamada():
                return cli.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[prompt, img]
                )

            resposta = executar_com_retry_gemini(chamada)
            texto = resposta.text.strip()
        except Exception as e:
            msg_amigavel = self._tratar_erro_gemini(e)
            self._na_interface(lambda: messagebox.showwarning("Limite Atingido / Erro IA", msg_amigavel, parent=self.root))
            self.set_status("Limite de uso da IA atingido.", 0)
            return
        finally:
            if apagar_depois:
                try: os.remove(caminho)
                except OSError: pass
                
        if not texto: return
        self.set_status("Leitura da imagem concluída.", 100)
        self._na_interface(self._mostrar_janela_revisao_ocr, texto)

    # ==========================================================================
    # Leitura de Imagens Offline com Tesseract
    # ==========================================================================
    def ler_imagem_com_tesseract(self) -> None:
        if not PYTESSERACT_DISPONIVEL:
            messagebox.showwarning("Tesseract em Falta", "O Tesseract OCR não está instalado ou disponível no sistema.", parent=self.root)
            return
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if caminho: self._em_segundo_plano(self._tesseract_exec, caminho)

    def colar_imagem_com_tesseract(self) -> None:
        if not PYTESSERACT_DISPONIVEL:
            messagebox.showwarning("Tesseract em Falta", "O Tesseract OCR não está instalado ou disponível no sistema.", parent=self.root)
            return
        imagem = self._imagem_da_area_transferencia()
        if imagem is None: return
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            imagem.convert("RGB").save(tmp, "PNG")
        self._em_segundo_plano(self._tesseract_exec, tmp.name, True)

    def _tesseract_exec(self, caminho: str, apagar_depois: bool = False) -> None:
        try:
            self.set_status("A ler a imagem offline (Tesseract)...", 50)
            img = Image.open(caminho)
            img = ImageOps.grayscale(img)
            texto = pytesseract.image_to_string(img, lang="por")
        except Exception as e:
            self._na_interface(lambda erro=e: messagebox.showerror("Erro Tesseract", f"Erro:\n\n{erro}", parent=self.root))
            self.set_status("Erro ao ler imagem com Tesseract.", 0)
            return
        finally:
            if apagar_depois:
                try: os.remove(caminho)
                except OSError: pass
                
        if not texto.strip():
            self.set_status("Nenhum texto encontrado.", 0)
            self._na_interface(lambda: messagebox.showinfo("Aviso", "O Tesseract não encontrou nenhum texto legível na imagem.", parent=self.root))
            return
            
        self.set_status("Leitura offline concluída.", 100)
        self._na_interface(self._mostrar_janela_revisao_ocr, texto.strip())

    # ---------------------------------------------------------------- tabela e ficheiros
    def ordenar_por_coluna(self, coluna: str) -> None:
        pos = self.parser.posicao
        chaves = {
            "NOME": lambda p: p["NOME"],
            "TAMANHO DE CAMISA": lambda p: pos.get(p["TAMANHO DE CAMISA"], 999),
            "NÚMERO": lambda p: int(p["NÚMERO"]) if str(p["NÚMERO"]).isdigit() else (0 if p["NÚMERO"] in ("PI", "π") else 9999),
            "TAMANHO DE CALÇÃO": lambda p: pos.get(p["TAMANHO DE CALÇÃO"], 999),
            "NÚMERO DE CALÇÃO": lambda p: int(p["NÚMERO DE CALÇÃO"]) if str(p["NÚMERO DE CALÇÃO"]).isdigit() else (0 if p["NÚMERO DE CALÇÃO"] in ("PI", "π") else 9999),
        }
        rev = self._ordenacao_reversa.get(coluna, False)
        self._salvar_estado_para_undo()
        self.pedidos_atuais.sort(key=chaves[coluna], reverse=rev)
        self._ordenacao_reversa[coluna] = not rev
        self._agrupar = coluna in COLUNAS_TAMANHO and not rev
        self._redesenhar_tabela()

    def _redesenhar_tabela(self) -> None:
        self.tree.delete(*self.tree.get_children())
        filtro = self.var_filtro.get().strip().upper()
        agrupar = self._agrupar and not filtro
        nomes = Counter(p["NOME"].strip() for p in self.pedidos_atuais if p["NOME"].strip() not in VALORES_VAZIOS)
        numeros = Counter(p["NÚMERO"].strip() for p in self.pedidos_atuais if p["NÚMERO"].strip() not in VALORES_VAZIOS)

        self._linhas_visiveis = []
        grupo_atual = None
        for idx, p in enumerate(self.pedidos_atuais):
            if filtro and filtro not in " ".join(p.values()).upper(): continue
            if agrupar:
                grupo = p["TAMANHO DE CAMISA"] if p["TAMANHO DE CAMISA"] in self.parser.posicao else p["TAMANHO DE CALÇÃO"]
                if grupo_atual is not None and grupo != grupo_atual:
                    self.tree.insert("", tk.END, values=("", "", "", "", ""))
                    self._linhas_visiveis.append(None)
                grupo_atual = grupo
            tag = "dup_nome" if nomes.get(p["NOME"].strip(), 0) > 1 else "dup_num" if numeros.get(p["NÚMERO"].strip(), 0) > 1 else ""
            self.tree.insert("", tk.END, iid=str(idx), values=[p[c] for c in COLUNAS], tags=(tag,))
            self._linhas_visiveis.append(p)
        self._atualizar_resumo(nomes, numeros)

    def _atualizar_resumo(self, nomes: Counter, numeros: Counter) -> None:
        tot = len(self.pedidos_atuais)
        if not tot: return self.lbl_resumo.config(text="Nenhum pedido gerado ainda.")
        pos = self.parser.posicao
        camisas = Counter(p["TAMANHO DE CAMISA"] for p in self.pedidos_atuais if p["TAMANHO DE CAMISA"] not in VALORES_VAZIOS)
        resumo = "  ".join(f"{t}: {q}" for t, q in sorted(camisas.items(), key=lambda kv: pos.get(kv[0], 999)))
        alts = [a for a, c in zip(["🟥 nomes rep.", "🟨 num rep."], [sum(1 for q in nomes.values() if q > 1), sum(1 for q in numeros.values() if q > 1)]) if c]
        self.lbl_resumo.config(text=f"Total: {tot}  |  {resumo}" + ("  |  " + ", ".join(alts) if alts else ""))

    def _atualizar_avisos(self) -> None:
        if self.linhas_nao_reconhecidas:
            self.frame_avisos.pack(fill="x", padx=20, pady=(0, 20), side="bottom", before=self.box_export)
            self.txt_avisos.config(state="normal")
            self.txt_avisos.delete("1.0", tk.END)
            self.txt_avisos.insert(tk.END, "\n".join(f"• {l}" for l in self.linhas_nao_reconhecidas))
            self.txt_avisos.config(state="disabled")
        else:
            self.frame_avisos.pack_forget()

    def editar_celula(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell": return
        item_id = self.tree.identify_row(event.y)
        coluna_id = self.tree.identify_column(event.x)
        if not item_id.isdigit(): return
        bbox = self.tree.bbox(item_id, coluna_id)
        if not bbox: return
        coluna = COLUNAS[int(coluna_id[1:]) - 1]
        idx = int(item_id)

        entry = tk.Entry(self.tree, font=("Segoe UI", 10), justify="center")
        entry.insert(0, self.pedidos_atuais[idx][coluna])
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        concluido = False

        def salvar(_e=None):
            nonlocal concluido
            if concluido: return
            concluido, valor = True, entry.get().strip().upper()
            entry.destroy()
            if coluna in COLUNAS_TAMANHO: valor = self.parser.conversoes.get(valor, valor)
            if valor == self.pedidos_atuais[idx][coluna]: return
            self._salvar_estado_para_undo()
            self.pedidos_atuais[idx][coluna] = valor
            self._redesenhar_tabela()

        def cancelar(_e=None):
            nonlocal concluido
            concluido = True
            entry.destroy()

        entry.bind("<Return>", salvar)
        entry.bind("<FocusOut>", salvar)
        entry.bind("<Escape>", cancelar)

    def adicionar_linha_manual(self) -> None:
        self._salvar_estado_para_undo()
        self.pedidos_atuais.append(ParserPedidos.novo_pedido())
        self._agrupar = False
        self.var_filtro.set("")
        self._redesenhar_tabela()
        self.tree.see(str(len(self.pedidos_atuais) - 1))

    def excluir_linhas_selecionadas(self) -> None:
        indices = sorted((int(i) for i in self.tree.selection() if i.isdigit()), reverse=True)
        if not indices: return
        self._salvar_estado_para_undo()
        for idx in indices: del self.pedidos_atuais[idx]
        self._redesenhar_tabela()

    def limpar_tudo(self) -> None:
        if not messagebox.askyesno("Limpar", "Pretende limpar o texto e a tabela?", parent=self.root): return
        self._salvar_estado_para_undo()
        self.text_area.delete("1.0", tk.END)
        self.pedidos_atuais, self.linhas_nao_reconhecidas = [], []
        self.var_filtro.set("")
        self._redesenhar_tabela()
        self._atualizar_avisos()

    def abrir_txt(self) -> None:
        caminho = filedialog.askopenfilename(filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if not caminho: return
        for codificacao in ("utf-8-sig", "cp1252"):
            try:
                with open(caminho, encoding=codificacao) as f: conteudo = f.read()
                break
            except UnicodeDecodeError: continue
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, conteudo)

    def colar_da_area_transferencia(self) -> None:
        try: self.text_area.insert(tk.INSERT, self.root.clipboard_get())
        except tk.TclError: pass

    def _linhas_para_exportar(self) -> List[List[str]]:
        return [[p[c] for c in COLUNAS] if p else ["", "", "", "", ""] for p in self._linhas_visiveis]

    def _nome_sugerido(self, extensao: str) -> str:
        base = os.path.splitext(os.path.basename(self.arquivo_atual))[0] if self.arquivo_atual else "pedidos"
        return f"{base}_{datetime.now():%Y-%m-%d}{extensao}"

    def copiar_para_excel(self) -> None:
        if not self._linhas_visiveis: return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join("\t".join(l) for l in self._linhas_para_exportar()))
        self.set_status("✅ Copiado!", 100)

    def salvar_em_csv(self) -> None:
        if not self._linhas_visiveis: return
        caminho = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=self._nome_sugerido(".csv"), filetypes=[("CSV", "*.csv")])
        if not caminho: return
        try:
            with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(COLUNAS)
                w.writerows(self._linhas_para_exportar())
        except OSError as e: messagebox.showerror("Erro", str(e), parent=self.root)

    def salvar_em_xlsx(self) -> None:
        if not OPENPYXL_DISPONIVEL: return messagebox.showerror("Erro", "pip install openpyxl", parent=self.root)
        if not self._linhas_visiveis: return
        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=self._nome_sugerido(".xlsx"), filetypes=[("Excel", "*.xlsx")])
        if not caminho: return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Pedidos"
            ws.append(list(COLUNAS))
            for l in self._linhas_para_exportar(): ws.append(l)
            for c in ws[1]: c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1E293B")
            for r in ws.iter_rows():
                for c in r: c.alignment = Alignment(horizontal="center", vertical="center")
            for i, w in enumerate([32, 20, 12, 20, 18]): ws.column_dimensions[chr(ord("A") + i)].width = w
            for c in ws["C"][1:]: c.number_format = "@"
            for c in ws["E"][1:]: c.number_format = "@"
            ws.freeze_panes = "A2"
            wb.save(caminho)
            if messagebox.askyesno("Salvo", "Pretende abrir agora?", parent=self.root): os.startfile(caminho)
        except OSError as e: messagebox.showerror("Erro", str(e), parent=self.root)

    def _mostrar_janela_revisao_ocr(self, texto: str) -> None:
        janela = tk.Toplevel(self.root)
        janela.title("Revisar Texto Lido")
        janela.geometry("640x500")
        janela.grab_set()
        tk.Label(janela, text="Verifique o texto lido antes de inserir:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        botoes = tk.Frame(janela)
        botoes.pack(side="bottom", fill="x", padx=15, pady=15)
        txt = scrolledtext.ScrolledText(janela, font=("Segoe UI", 10), undo=True)
        txt.pack(fill="both", expand=True, padx=15, pady=(0, 5))
        txt.insert(tk.END, texto)
        def usar(substituir: bool, organizar: bool = False):
            v = txt.get("1.0", tk.END).strip()
            if substituir: self.text_area.delete("1.0", tk.END)
            self.text_area.insert(tk.END, ("\n" if not substituir else "") + v)
            janela.destroy()
            if organizar: self.processar_texto()
        tk.Button(botoes, text="✨ Usar e Organizar", command=lambda: usar(True, True), bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), padx=10, pady=6).pack(side="right", padx=5)
        tk.Button(botoes, text="Substituir", command=lambda: usar(True), bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), padx=10, pady=6).pack(side="right", padx=5)
        tk.Button(botoes, text="Somar", command=lambda: usar(False), bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), padx=10, pady=6).pack(side="right", padx=5)

    def abrir_editor_regras(self) -> None:
        janela = tk.Toplevel(self.root)
        janela.title("Tamanhos e Conversões")
        janela.geometry("760x560")
        janela.grab_set()

        corpo = tk.Frame(janela, padx=20, pady=20)
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=1)
        corpo.columnconfigure(1, weight=2)
        corpo.rowconfigure(1, weight=1)

        tk.Label(corpo, text="Ordem Produção (Um por linha)", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(corpo, text="Regras: ESCRITO = TAMANHO", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=(10, 0))
        txt_ordem = scrolledtext.ScrolledText(corpo, width=18, font=("Consolas", 10))
        txt_ordem.grid(row=1, column=0, sticky="nsew", pady=5)
        txt_ordem.insert("1.0", "\n".join(self.ordem_tamanhos))
        txt_conv = scrolledtext.ScrolledText(corpo, font=("Consolas", 10))
        txt_conv.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=5)
        txt_conv.insert("1.0", "\n".join(f"{k} = {v}" for k, v in self.conversoes.items()))

        def salvar():
            ordem = list(dict.fromkeys(l.strip().upper() for l in txt_ordem.get("1.0", tk.END).splitlines() if l.strip()))
            conversoes = {}
            for linha in txt_conv.get("1.0", tk.END).splitlines():
                if not linha.strip(): continue
                de, sep, para = linha.partition("=")
                if sep: conversoes[de.strip().upper()] = para.strip().upper()
            self.ordem_tamanhos, self.conversoes = ordem, conversoes
            self.parser = ParserPedidos(self.ordem_tamanhos, self.conversoes)
            self.salvar_config()
            self._redesenhar_tabela()
            janela.destroy()

        tk.Button(corpo, text="💾 Salvar Modificações", command=salvar, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), pady=8).grid(row=2, column=1, sticky="e", pady=(10, 0))


if __name__ == "__main__":
    root = tk.Tk()
    app = AplicativoPedidosMagico(root)
    root.mainloop()
