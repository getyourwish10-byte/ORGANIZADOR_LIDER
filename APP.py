#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador Profissional de Orçamentos
Desenvolvido por: Douglas Oliveira | getyourwish10@gmail.com

Otimizado para:
- Carregar logotipos e carimbos dinamicamente via interface (Tkinter)
- Salvar automaticamente o último diretório das logos e carimbos
- Centralizar o bloco de Valores Totais
- GERAÇÃO MULTIPÁGINA: Suporte a listas longas sem cortar o orçamento
- SALVAR/CARREGAR PROJETO: Permite salvar orçamentos em andamento (.json) para edição futura
- TEMPLATES EXCLUSIVOS: Cores exclusivas para cada empresa (Marinho, Laranja, P&B, Azul Royal)
- ORDENAÇÃO ALEATÓRIA INTELIGENTE: Itens embaralhados nos concorrentes com numeração sequencial crescente
- ABERTURA AUTOMÁTICA OPCIONAL E PADRÃO JPG DE ALTA QUALIDADE
"""

from __future__ import annotations

import json
import os
import platform
import re
import subprocess
import textwrap
import random
import copy
from dataclasses import dataclass
from functools import lru_cache
from typing import List, Optional, Tuple, Callable

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from PIL import Image, ImageDraw, ImageFont, ImageEnhance

# Compatibilidade de versão do Pillow para Resampling
try:
    RESAMPLE = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE = Image.LANCZOS


# ==========================================================================
# Utilitários gerais e Configurações
# ==========================================================================

PASTA_PADRAO_ORCAMENTOS = os.path.join(os.path.expanduser("~"), "Desktop", "ORÇAMENTOS")
ARQUIVO_CONFIG = os.path.join(PASTA_PADRAO_ORCAMENTOS, "config_logos.json")

def formatar_valor(valor: float) -> str:
    """Formata um float para o padrão monetário brasileiro: 1.500,00."""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_numero(texto: str) -> float:
    """Converte texto com vírgula/ponto em float, com segurança."""
    texto = str(texto).replace("R$", "").strip()
    if not texto:
        return 0.0
    if "." in texto and "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def limpar_nome_arquivo(texto: str) -> str:
    """Remove caracteres inválidos e troca espaços por '_' para arquivos."""
    texto = str(texto).upper()
    texto = re.sub(r"[^A-Z0-9ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜ _-]", "", texto)
    return re.sub(r"\s+", "_", texto.strip()) or "SEM_NOME"


def abrir_arquivo(caminho: str) -> None:
    """Abre o arquivo gerado com o programa padrão do sistema operacional."""
    if not os.path.exists(caminho):
        return
    try:
        sistema = platform.system()
        if sistema == "Windows":
            os.startfile(caminho)
        elif sistema == "Darwin":
            subprocess.call(["open", caminho])
        else:
            subprocess.call(["xdg-open", caminho])
    except Exception as exc:
        print(f"Não foi possível abrir '{caminho}' automaticamente: {exc}")


def carregar_config_logos() -> dict:
    """Carrega os caminhos das logos e carimbos salvos na última sessão."""
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def salvar_config_logos(logos: dict) -> None:
    """Salva os caminhos das logos e carimbos para não precisar selecionar novamente."""
    os.makedirs(PASTA_PADRAO_ORCAMENTOS, exist_ok=True)
    try:
        with open(ARQUIVO_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(logos, f, ensure_ascii=False, indent=4)
    except Exception as exc:
        print(f"Erro ao salvar config: {exc}")


# ==========================================================================
# Cache de Fontes e Imagens
# ==========================================================================

class Fontes:
    """Cache de fontes TrueType com fallback multiplataforma."""
    _BOLD = ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"]
    _REGULAR = ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf"]
    _cache: dict = {}

    @classmethod
    def _carregar(cls, candidatos: List[str], tamanho: int) -> ImageFont.FreeTypeFont:
        chave = (tuple(candidatos), tamanho)
        if chave in cls._cache:
            return cls._cache[chave]
        for nome in candidatos:
            try:
                fonte = ImageFont.truetype(nome, tamanho)
                cls._cache[chave] = fonte
                return fonte
            except (OSError, IOError):
                continue
        fonte = ImageFont.load_default()
        cls._cache[chave] = fonte
        return fonte

    @classmethod
    def bold(cls, tamanho: int) -> ImageFont.FreeTypeFont:
        return cls._carregar(cls._BOLD, tamanho)

    @classmethod
    def regular(cls, tamanho: int) -> ImageFont.FreeTypeFont:
        return cls._carregar(cls._REGULAR, tamanho)


@lru_cache(maxsize=16)
def carregar_logo(caminho: str) -> Optional[Image.Image]:
    """Carrega uma logo ou carimbo do disco e mantém em cache."""
    if not caminho or not os.path.isfile(caminho):
        return None
    try:
        img = Image.open(caminho)
        return img.convert("RGBA") if img.mode != "RGBA" else img
    except Exception as exc:
        print(f"Erro ao carregar imagem ({caminho}): {exc}")
        return None


# ==========================================================================
# Modelos de Dados
# ==========================================================================

@dataclass
class ItemOrcamento:
    numero: str
    descricao: str
    unidade: str
    quantidade: float
    valor_unitario: float


@dataclass
class DadosCabecalho:
    cidade: str
    data: str
    ac: str
    condicao_pagamento: str
    prazo_entrega: str
    validade: str
    pasta_destino: str


@dataclass
class LogosEmpresas:
    nahora: str
    dakar: str
    lider_comercio: str
    lider_sport: str
    carimbo_nahora: str = ""
    carimbo_dakar: str = ""
    carimbo_lider_comercio: str = ""
    carimbo_lider_sport: str = ""


# ==========================================================================
# Geração das Imagens/PDF
# ==========================================================================

class GeradorOrcamento:
    LARGURA, ALTURA = 2371, 3410

    def __init__(self, cabecalho: DadosCabecalho, itens: List[ItemOrcamento], formato: str, logos: LogosEmpresas):
        self.cab = cabecalho
        self.itens = itens
        self.formato = formato.upper()
        self.logos = logos

    def _nova_pagina(self) -> Image.Image:
        return Image.new("RGB", (self.LARGURA, self.ALTURA), "white")

    def _colar_logo(self, img: Image.Image, caminho_logo: str, x: int, y: int, max_w: int, max_h: int) -> None:
        logo = carregar_logo(caminho_logo)
        if not logo:
            return
        logo = logo.copy()
        logo.thumbnail((max_w, max_h), RESAMPLE)
        w, h = logo.size
        img.paste(logo, (x + (max_w - w) // 2, y + (max_h - h) // 2), logo)

    def _colar_carimbo(self, img: Image.Image, caminho_carimbo: str, x: int, y: int, max_w: int, max_h: int) -> None:
        carimbo = carregar_logo(caminho_carimbo)
        if not carimbo:
            return
        carimbo = carimbo.copy()
        carimbo.thumbnail((max_w, max_h), RESAMPLE)
        w, h = carimbo.size
        img.paste(carimbo, (x + (max_w - w) // 2, y + (max_h - h) // 2), carimbo)

    def _marca_dagua(self, img: Image.Image, caminho_logo: str) -> None:
        logo = carregar_logo(caminho_logo)
        if not logo:
            return
        w0, h0 = logo.size
        largura_alvo = 1300
        altura_alvo = int(h0 * (largura_alvo / w0))
        logo = logo.resize((largura_alvo, altura_alvo), RESAMPLE)
        alpha = ImageEnhance.Brightness(logo.split()[3]).enhance(0.06)
        logo.putalpha(alpha)
        x = (self.LARGURA - largura_alvo) // 2
        y = (self.ALTURA - altura_alvo) // 2
        img.paste(logo, (x, y), logo)

    @staticmethod
    def _quebrar_texto(draw: ImageDraw.ImageDraw, texto: str, fonte, largura_max: int) -> List[str]:
        linhas: List[str] = []
        atual = ""
        for palavra in texto.split():
            while draw.textlength(palavra, font=fonte) > largura_max:
                corte = len(palavra)
                while corte > 1 and draw.textlength(palavra[:corte], font=fonte) > largura_max:
                    corte -= 1
                if atual:
                    linhas.append(atual)
                    atual = ""
                linhas.append(palavra[:corte])
                palavra = palavra[corte:]
            teste = f"{atual} {palavra}".strip()
            if draw.textlength(teste, font=fonte) <= largura_max:
                atual = teste
            else:
                if atual:
                    linhas.append(atual)
                atual = palavra
        if atual:
            linhas.append(atual)
        return linhas or [""]

    @staticmethod
    def _texto_centralizado(draw, texto, fonte, x, largura, y, cor="white"):
        w = draw.textlength(texto, font=fonte)
        draw.text((x + (largura - w) / 2, y), texto, fill=cor, font=fonte)

    @staticmethod
    def _texto_direita(draw, texto, fonte, x_direita, y, cor="white"):
        w = draw.textlength(texto, font=fonte)
        draw.text((x_direita - w, y), texto, fill=cor, font=fonte)

    def _desenhar_tabela(
        self, draw: ImageDraw.ImageDraw, acrescimo: float, y_topo: int, 
        col_x: List[int], col_w: List[int], cor_cabecalho: str, 
        nova_pagina_callback: Callable[[], Tuple[ImageDraw.ImageDraw, int]], 
        modo_compacto: bool = False,
        cor_borda: str = "#94a3b8",
        cor_texto_cab: str = "white",
        itens: Optional[List[ItemOrcamento]] = None
    ) -> Tuple[ImageDraw.ImageDraw, int, float]:
        
        lista_itens = itens if itens is not None else self.itens
        f_head = Fontes.bold(45)
        f_head_pequena = Fontes.bold(34)
        f_norm = Fontes.regular(40)

        headers = (
            ["Nº", "DESCRIÇÃO", "PREÇO", "QTD", "TOTAL"]
            if modo_compacto else ["Nº", "DESCRIÇÃO", "UND", "QTD", "VALOR UNIT.", "VALOR TOTAL"]
        )

        def desenhar_cabecalho(d: ImageDraw.ImageDraw, y_t: int) -> int:
            d.rectangle([150, y_t, 2221, y_t + 100], fill=cor_cabecalho, outline=cor_borda, width=2)
            for i, h in enumerate(headers):
                fonte_col = f_head_pequena if (not modo_compacto and i in (4, 5)) else f_head
                self._texto_centralizado(d, h, fonte_col, col_x[i], col_w[i], y_t + 28, cor=cor_texto_cab)
            for x in col_x[1:]:
                d.line([(x, y_t), (x, y_t + 100)], fill=cor_texto_cab, width=2)
            return y_t + 100

        y = desenhar_cabecalho(draw, y_topo)
        total_geral = 0.0
        largura_desc = col_w[1] - 60

        for item in lista_itens:
            valor_unit = item.valor_unitario + acrescimo
            valor_total = item.quantidade * valor_unit
            qtd_str = str(int(item.quantidade)) if float(item.quantidade).is_integer() else str(item.quantidade)

            linhas_desc = self._quebrar_texto(draw, item.descricao, f_norm, largura_desc)
            altura = max(120, len(linhas_desc) * 50 + 40)

            if y + altura > 2450:
                draw, y_novo = nova_pagina_callback()
                y = desenhar_cabecalho(draw, y_novo)

            total_geral += valor_total
            draw.rectangle([150, y, 2221, y + altura], fill="white", outline=cor_borda, width=2)

            valores = [item.numero, f"R$ {formatar_valor(valor_unit)}", qtd_str, f"R$ {formatar_valor(valor_total)}"] if modo_compacto \
                else [item.numero, item.unidade, qtd_str, f"R$ {formatar_valor(valor_unit)}", f"R$ {formatar_valor(valor_total)}"]

            self._texto_centralizado(draw, valores[0], f_norm, col_x[0], col_w[0], y + (altura - 40) / 2, cor="#111827")

            y_txt = y + (altura - len(linhas_desc) * 45) / 2
            for linha in linhas_desc:
                draw.text((col_x[1] + 30, y_txt), linha, fill="#111827", font=f_norm)
                y_txt += 45

            for i in range(2, len(col_x)):
                self._texto_centralizado(draw, valores[i - 1], f_norm, col_x[i], col_w[i], y + (altura - 40) / 2, cor="#111827")

            for x in col_x[1:]:
                draw.line([(x, y), (x, y + altura)], fill=cor_borda, width=2)
            y += altura

        return draw, y, total_geral

    def _rodape_condicoes(self, draw, y: int, cor_texto: str = "#b91c1c") -> int:
        f_pag = Fontes.bold(34)
        texto = f"CONDIÇÃO DE PAGAMENTO: {self.cab.condicao_pagamento}"
        for linha in textwrap.wrap(texto, width=65):
            self._texto_centralizado(draw, linha, f_pag, 0, self.LARGURA, y, cor=cor_texto)
            y += 45
        return y + 30

    def _bloco_total_prazo(self, draw, y: int, total: float, fundo: str = "#1e293b", texto_cor: str = "white", cor_borda: str = "#94a3b8", cor_fundo_total: str = "#e2e8f0", cor_texto_total: str = "#0f172a") -> int:
        f_bold = Fontes.bold(52)
        f_title = Fontes.bold(85)
        
        draw.rectangle([150, y, 1400, y + 180], fill=fundo, outline=cor_borda, width=2)
        draw.text((180, y + 30), f"PRAZO DE ENTREGA: {self.cab.prazo_entrega}", fill=texto_cor, font=f_bold)
        draw.text((180, y + 100), f"VALIDADE DA PROPOSTA: {self.cab.validade}", fill=texto_cor, font=f_bold)

        tot_str = f"R$ {formatar_valor(total)}"
        x_box = 1550         
        w_box = 2221 - 1550  
        
        draw.rectangle([x_box, y, 2221, y + 180], fill=cor_fundo_total, outline=cor_borda, width=2)
        self._texto_centralizado(draw, tot_str, f_title, x_box, w_box, y + 45, cor=cor_texto_total)
        
        return y + 240

    def _cabecalho_titulo_logo(self, img, draw, caminho_logo: str, caminho_carimbo: str, cor_fundo: str, cor_titulo: str, cor_sub: str, cor_borda_ac: str = "#cbd5e1", cor_fundo_ac: str = "#f3f4f6", cor_texto_ac: str = "#1f2937") -> None:
        f_title = Fontes.bold(85)
        f_bold = Fontes.bold(52)

        if cor_fundo and cor_fundo.lower() not in ["white", "#ffffff"]:
            draw.rectangle([0, 0, self.LARGURA, 260], fill=cor_fundo)
            y_logo, y_titulo, y_sub, max_h_logo = 35, 45, 145, 190
        else:
            y_logo, y_titulo, y_sub, max_h_logo = 50, 50, 150, 230

        self._colar_logo(img, caminho_logo, 150, y_logo, 950, max_h_logo)
        
        if caminho_carimbo:
            self._colar_carimbo(img, caminho_carimbo, 1130, y_logo, 450, max_h_logo)

        self._texto_direita(draw, "ORÇAMENTO", f_title, 2221, y_titulo, cor=cor_titulo)
        self._texto_direita(draw, f"{self.cab.cidade} - {self.cab.data}", f_bold, 2221, y_sub, cor=cor_sub)

        draw.rectangle([150, 320, 2221, 410], fill=cor_fundo_ac, outline=cor_borda_ac, width=2)
        draw.text((180, 345), f"A/C: {self.cab.ac}", fill=cor_texto_ac, font=f_bold)

    def _bloco_bancario_duas_colunas(self, draw, y: int, titulo_esq: str, linhas_esq, titulo_dir: str, linhas_dir, cor_borda: str = "#cbd5e1", cor_titulo: str = "#1e293b", cor_fundo: str = "#f8fafc", cor_texto: str = "#111827") -> int:
        f_banco = Fontes.bold(40)
        draw.rectangle([385, y, 1985, y + 200], fill=cor_fundo, outline=cor_borda, width=2)
        draw.line([(1185, y), (1185, y + 200)], fill=cor_borda, width=2)

        self._texto_centralizado(draw, titulo_esq, f_banco, 385, 800, y + 40, cor=cor_titulo)
        y_atual = y + 110
        for texto, fonte in linhas_esq:
            self._texto_centralizado(draw, texto, fonte, 385, 800, y_atual, cor=cor_texto)
            y_atual += 60

        self._texto_centralizado(draw, titulo_dir, f_banco, 1185, 800, y + 25, cor=cor_titulo)
        y_atual = y + 80
        for texto, fonte in linhas_dir:
            self._texto_centralizado(draw, texto, fonte, 1185, 800, y_atual, cor=cor_texto)
            y_atual += 55
        return y + 240

    def _salvar(self, paginas: List[Image.Image], prefixo_empresa: str) -> str:
        pasta = self.cab.pasta_destino.strip() or PASTA_PADRAO_ORCAMENTOS
        os.makedirs(pasta, exist_ok=True)

        base = f"{limpar_nome_arquivo(self.cab.ac)}_{limpar_nome_arquivo(prefixo_empresa)}_{self.cab.data.replace('/', '-')}"
        ext = f".{self.formato.lower()}"
        
        caminho_base = os.path.join(pasta, base)
        caminho = caminho_base + ext

        contador = 1
        while os.path.exists(caminho) or any(os.path.exists(f"{caminho_base}_{contador}_Pagina_1{ext}") for _ in range(1)):
            caminho_base = os.path.join(pasta, f"{base}_{contador}")
            caminho = caminho_base + ext
            contador += 1

        if self.formato == "PDF":
            paginas[0].save(caminho, "PDF", resolution=300, save_all=True, append_images=paginas[1:])
        else:
            if len(paginas) == 1:
                if self.formato == "PNG":
                    paginas[0].save(caminho, "PNG")
                else:
                    paginas[0].save(caminho, "JPEG", quality=95)
            else:
                for i, pag in enumerate(paginas):
                    cam_pag = f"{caminho_base}_Pagina_{i+1}{ext}"
                    if self.formato == "PNG":
                        pag.save(cam_pag, "PNG")
                    else:
                        pag.save(cam_pag, "JPEG", quality=95)
                caminho = f"{caminho_base}_Pagina_1{ext}"
        
        return caminho

    # -------------------------------------------------------------
    # TEMPLATE 1: NA HORA (TEMA LARANJA)
    # -------------------------------------------------------------
    def gerar_nahora(self, acrescimo: float, itens: Optional[List[ItemOrcamento]] = None) -> str:
        paginas = []
        f_bold = Fontes.bold(52)
        f_banco = Fontes.bold(40)
        
        cor_primaria = "#ea580c"
        cor_borda = "#f97316"
        cor_fundo_destaque = "#fff7ed"

        def criar_pagina(num_pag: int) -> Tuple[Image.Image, ImageDraw.ImageDraw, int]:
            img = self._nova_pagina()
            self._marca_dagua(img, self.logos.nahora)
            draw = ImageDraw.Draw(img)
            
            if num_pag == 1:
                self._cabecalho_titulo_logo(
                    img, draw, self.logos.nahora, self.logos.carimbo_nahora, 
                    cor_fundo="white", cor_titulo="#111827", cor_sub=cor_primaria, 
                    cor_borda_ac=cor_primaria, cor_fundo_ac=cor_fundo_destaque, cor_texto_ac=cor_primaria
                )
                y_topo = 470
            else:
                self._colar_logo(img, self.logos.nahora, 150, 50, 400, 150)
                if self.logos.carimbo_nahora:
                    self._colar_carimbo(img, self.logos.carimbo_nahora, 580, 50, 300, 150)
                f_title_pag = Fontes.bold(60)
                self._texto_direita(draw, f"ORÇAMENTO - PÁGINA {num_pag}", f_title_pag, 2221, 80, cor=cor_primaria)
                draw.line([(150, 220), (2221, 220)], fill=cor_primaria, width=2)
                y_topo = 260

            draw.rectangle([0, 3150, 2371, 3410], fill=cor_primaria)
            draw.text((250, 3230), "Rua P-16 Nº 55 Setor dos Funcionários Goiânia - Goiás  |  Tel: (62) 3087-2424", fill="white", font=f_bold)
            
            return img, draw, y_topo

        img, draw, y_topo = criar_pagina(1)
        paginas.append(img)

        def callback_nova_pagina():
            img_nova, draw_novo, y_t = criar_pagina(len(paginas) + 1)
            paginas.append(img_nova)
            return draw_novo, y_t

        col_x = [150, 300, 1200, 1550, 1850]
        col_w = [150, 900, 350, 300, 371]
        
        draw, y, total = self._desenhar_tabela(
            draw, acrescimo, y_topo, col_x, col_w, 
            cor_cabecalho=cor_primaria, nova_pagina_callback=callback_nova_pagina, 
            modo_compacto=True, cor_borda=cor_borda, itens=itens
        )

        if y > 2450: draw, y = callback_nova_pagina(); y += 50

        y = self._rodape_condicoes(draw, y + 50, cor_texto=cor_primaria)
        y = self._bloco_total_prazo(
            draw, y, total, fundo=cor_primaria, texto_cor="white", 
            cor_borda=cor_borda, cor_fundo_total=cor_fundo_destaque, cor_texto_total=cor_primaria
        )

        draw.rectangle([585, y, 1785, y + 200], fill=cor_fundo_destaque, outline=cor_borda, width=2)
        for i, linha in enumerate(["BANCO DO BRASIL", "AG.: 3483-5  -  C/C.: 55005-1", "CNPJ: 30.339.532/0001-19"]):
            self._texto_centralizado(draw, linha, f_banco, 585, 1200, y + 25 + i * 55, cor=cor_primaria)

        return self._salvar(paginas, "NaHora")

    # -------------------------------------------------------------
    # TEMPLATE 2: DAKAR SPORT (TEMA MARINHO)
    # -------------------------------------------------------------
    def gerar_dakar(self, acrescimo: float, itens: Optional[List[ItemOrcamento]] = None) -> str:
        paginas = []
        f_bold = Fontes.bold(52)

        cor_fundo_header = "#0f172a"
        cor_primaria = "#1e3a8a"
        cor_borda = "#94a3b8"

        def criar_pagina(num_pag: int) -> Tuple[Image.Image, ImageDraw.ImageDraw, int]:
            img = self._nova_pagina()
            self._marca_dagua(img, self.logos.dakar)
            draw = ImageDraw.Draw(img)

            if num_pag == 1:
                self._cabecalho_titulo_logo(
                    img, draw, self.logos.dakar, self.logos.carimbo_dakar, 
                    cor_fundo=cor_fundo_header, cor_titulo="white", cor_sub="#cbd5e1"
                )
                y_topo = 470
            else:
                draw.rectangle([0, 0, self.LARGURA, 220], fill=cor_fundo_header)
                self._colar_logo(img, self.logos.dakar, 150, 40, 400, 140)
                if self.logos.carimbo_dakar:
                    self._colar_carimbo(img, self.logos.carimbo_dakar, 580, 40, 300, 140)
                f_title_pag = Fontes.bold(60)
                self._texto_direita(draw, f"ORÇAMENTO - PÁGINA {num_pag}", f_title_pag, 2221, 80, cor="white")
                y_topo = 280

            draw.rectangle([0, 3150, 2371, 3410], fill=cor_fundo_header)
            draw.text((250, 3230), "DAKAR SPORT  -  CNPJ: 29.332.450/0001-63  -  INSC. EST.: 10.713.610-4", fill="white", font=f_bold)
            
            return img, draw, y_topo

        img, draw, y_topo = criar_pagina(1)
        paginas.append(img)

        def callback_nova_pagina():
            img_nova, draw_novo, y_t = criar_pagina(len(paginas) + 1)
            paginas.append(img_nova)
            return draw_novo, y_t

        col_x = [150, 350, 1150, 1380, 1630, 1950]
        col_w = [200, 800, 230, 250, 320, 271]
        
        draw, y, total = self._desenhar_tabela(
            draw, acrescimo, y_topo, col_x, col_w, 
            cor_cabecalho=cor_primaria, nova_pagina_callback=callback_nova_pagina, cor_borda=cor_borda, itens=itens
        )

        if y > 2450: draw, y = callback_nova_pagina(); y += 50

        y = self._rodape_condicoes(draw, y + 50)
        y = self._bloco_total_prazo(
            draw, y, total, fundo=cor_primaria, texto_cor="white", cor_borda=cor_borda
        )

        y = self._bloco_bancario_duas_colunas(
            draw, y, "PIX (CNPJ)", [("29.332.450/0001-63", f_bold)],
            "AGÊNCIA E CONTA", [("AG: 3483-5", f_bold), ("C/C: 55104-X", f_bold)],
            cor_borda=cor_borda, cor_titulo=cor_primaria
        )

        return self._salvar(paginas, "DakarSport")

    # -------------------------------------------------------------
    # TEMPLATE 3: LÍDER COMÉRCIO (TEMA PRETO E BRANCO)
    # -------------------------------------------------------------
    def gerar_lider_comercio(self, acrescimo: float, itens: Optional[List[ItemOrcamento]] = None) -> str:
        paginas = []
        f_bold = Fontes.bold(52)
        f_sub = Fontes.regular(36)

        cor_primaria = "#000000"
        cor_borda = "#000000"
        
        def criar_pagina(num_pag: int) -> Tuple[Image.Image, ImageDraw.ImageDraw, int]:
            img = self._nova_pagina()
            self._marca_dagua(img, self.logos.lider_comercio)
            draw = ImageDraw.Draw(img)

            if num_pag == 1:
                self._cabecalho_titulo_logo(
                    img, draw, self.logos.lider_comercio, self.logos.carimbo_lider_comercio, 
                    cor_fundo="white", cor_titulo="#000000", cor_sub="#4b5563",
                    cor_borda_ac=cor_primaria, cor_fundo_ac="white", cor_texto_ac="#000000"
                )
                y_topo = 470
            else:
                self._colar_logo(img, self.logos.lider_comercio, 150, 40, 400, 140)
                if self.logos.carimbo_lider_comercio:
                    self._colar_carimbo(img, self.logos.carimbo_lider_comercio, 580, 40, 300, 140)
                f_title_pag = Fontes.bold(60)
                self._texto_direita(draw, f"ORÇAMENTO - PÁGINA {num_pag}", f_title_pag, 2221, 80, cor="#000000")
                draw.line([(150, 220), (2221, 220)], fill="#000000", width=3)
                y_topo = 280

            draw.rectangle([0, 3150, 2371, 3410], fill="white")
            draw.line([(0, 3150), (2371, 3150)], fill="#000000", width=4)
            draw.text((150, 3195), "R JOSE SINIMBU FILHO - Nº 67 QUADRA 140A LOTE 50E", fill="#000000", font=f_sub)
            draw.text((150, 3255), "SETOR NORTE FERROVIÁRIO - 74.063-330", fill="#000000", font=f_sub)
            self._texto_direita(draw, "(62) 99175-7971", f_bold, 2221, 3185, cor="#000000")
            self._texto_direita(draw, "LIDERSPORT44@GMAIL.COM", f_sub, 2221, 3260, cor="#000000")
            
            return img, draw, y_topo

        img, draw, y_topo = criar_pagina(1)
        paginas.append(img)

        def callback_nova_pagina():
            img_nova, draw_novo, y_t = criar_pagina(len(paginas) + 1)
            paginas.append(img_nova)
            return draw_novo, y_t

        col_x = [150, 350, 1150, 1380, 1630, 1950]
        col_w = [200, 800, 230, 250, 320, 271]
        
        draw, y, total = self._desenhar_tabela(
            draw, acrescimo, y_topo, col_x, col_w, 
            cor_cabecalho=cor_primaria, nova_pagina_callback=callback_nova_pagina, cor_borda=cor_borda, itens=itens
        )

        if y > 2450: draw, y = callback_nova_pagina(); y += 50

        y = self._rodape_condicoes(draw, y + 50, cor_texto="#000000")
        y = self._bloco_total_prazo(
            draw, y, total, fundo=cor_primaria, texto_cor="white", 
            cor_borda=cor_borda, cor_fundo_total="white", cor_texto_total="#000000"
        )

        y = self._bloco_bancario_duas_colunas(
            draw, y, "PIX (CNPJ)", [("50.519.580/0001-04", f_bold)],
            "AGÊNCIA E CONTA", [("3483-5 · 55510-X", f_bold), ("LÍDER COMÉRCIO", f_sub)],
            cor_borda=cor_borda, cor_titulo="#000000", cor_fundo="white", cor_texto="#000000"
        )

        return self._salvar(paginas, "LiderComercio")

    # -------------------------------------------------------------
    # TEMPLATE 4: LIDER SPORT (TEMA AZUL ROYAL)
    # -------------------------------------------------------------
    def gerar_lider_sport(self, acrescimo: float, itens: Optional[List[ItemOrcamento]] = None) -> str:
        paginas = []
        f_bold = Fontes.bold(52)
        f_sub = Fontes.regular(36)

        cor_fundo_header = "#1d4ed8"
        cor_primaria = "#2563eb"
        cor_borda = "#93c5fd"

        def criar_pagina(num_pag: int) -> Tuple[Image.Image, ImageDraw.ImageDraw, int]:
            img = self._nova_pagina()
            self._marca_dagua(img, self.logos.lider_sport)
            draw = ImageDraw.Draw(img)

            if num_pag == 1:
                self._cabecalho_titulo_logo(
                    img, draw, self.logos.lider_sport, self.logos.carimbo_lider_sport, 
                    cor_fundo=cor_fundo_header, cor_titulo="white", cor_sub="#bfdbfe",
                    cor_borda_ac=cor_primaria, cor_fundo_ac="#eff6ff", cor_texto_ac=cor_fundo_header
                )
                y_topo = 470
            else:
                draw.rectangle([0, 0, self.LARGURA, 220], fill=cor_fundo_header)
                self._colar_logo(img, self.logos.lider_sport, 150, 40, 400, 140)
                if self.logos.carimbo_lider_sport:
                    self._colar_carimbo(img, self.logos.carimbo_lider_sport, 580, 40, 300, 140)
                f_title_pag = Fontes.bold(60)
                self._texto_direita(draw, f"ORÇAMENTO - PÁGINA {num_pag}", f_title_pag, 2221, 80, cor="white")
                y_topo = 280

            draw.rectangle([0, 3080, 2371, 3410], fill=cor_fundo_header)
            
            f_rodape_bold = Fontes.bold(45)
            f_rodape_sub = Fontes.regular(38)
            
            self._texto_centralizado(draw, "LÍDER SPORT LTDA. ME - CNPJ: 02.667.069/0001-07 - INSC. EST.: 10.327.650-5", f_rodape_bold, 0, self.LARGURA, 3120, cor="white")
            self._texto_centralizado(draw, "GOIÂNIA: (62) 3291-4804 / 3293-1257  |  Rua P-16 Nº 82 - Setor dos Funcionários", f_rodape_sub, 0, self.LARGURA, 3190, cor="#bfdbfe")
            self._texto_centralizado(draw, "BRASÍLIA: (61) 3046-2583  |  CSC 11 - Lote 02 - Lojas 03 e 04 (Sandu Sul)", f_rodape_sub, 0, self.LARGURA, 3250, cor="#bfdbfe")
            self._texto_centralizado(draw, "www.lidersport.com.br", f_rodape_bold, 0, self.LARGURA, 3320, cor="white")
            
            return img, draw, y_topo

        img, draw, y_topo = criar_pagina(1)
        paginas.append(img)

        def callback_nova_pagina():
            img_nova, draw_novo, y_t = criar_pagina(len(paginas) + 1)
            paginas.append(img_nova)
            return draw_novo, y_t

        col_x = [150, 350, 1150, 1380, 1630, 1950]
        col_w = [200, 800, 230, 250, 320, 271]
        
        draw, y, total = self._desenhar_tabela(
            draw, acrescimo, y_topo, col_x, col_w, 
            cor_cabecalho=cor_primaria, nova_pagina_callback=callback_nova_pagina, cor_borda=cor_borda, itens=itens
        )

        if y > 2450: draw, y = callback_nova_pagina(); y += 50

        y = self._rodape_condicoes(draw, y + 50, cor_texto=cor_fundo_header)
        y = self._bloco_total_prazo(
            draw, y, total, fundo=cor_primaria, texto_cor="white", 
            cor_borda=cor_borda, cor_fundo_total="#eff6ff", cor_texto_total=cor_fundo_header
        )

        y = self._bloco_bancario_duas_colunas(
            draw, y, "PIX (CNPJ)", [("02.667.069/0001-07", f_bold)],
            "BANCO DO BRASIL", [("AG.: 3485-1", f_bold), ("C/C.: 33562-2", f_bold)],
            cor_borda=cor_borda, cor_titulo=cor_fundo_header, cor_fundo="#eff6ff", cor_texto=cor_fundo_header
        )

        return self._salvar(paginas, "LiderSport")

    def gerar(self, empresa_menor_preco: str, apenas_menor_preco: bool = False) -> List[str]:
        geradores = {
            "Na Hora": lambda acr, its: self.gerar_nahora(acr, its),
            "Dakar Sport": lambda acr, its: self.gerar_dakar(acr, its),
            "Líder Comércio": lambda acr, its: self.gerar_lider_comercio(acr, its),
            "Lider Sport": lambda acr, its: self.gerar_lider_sport(acr, its)
        }

        caminhos = []

        def criar_itens_embaralhados(itens_originais: List[ItemOrcamento], ordens_excluidas: List[tuple]) -> List[ItemOrcamento]:
            novos_itens = copy.deepcopy(itens_originais)
            if len(novos_itens) <= 1:
                for i, it in enumerate(novos_itens):
                    it.numero = str(i + 1)
                return novos_itens

            while True:
                random.shuffle(novos_itens)
                ordem_atual = tuple(id(x) for x in novos_itens)
                if ordem_atual not in ordens_excluidas:
                    break

            for i, it in enumerate(novos_itens):
                it.numero = str(i + 1)
            return novos_itens

        if apenas_menor_preco:
            caminho = geradores[empresa_menor_preco](0.0, self.itens)
            caminhos.append(caminho)
        else:
            concorrentes = [emp for emp in geradores.keys() if emp != empresa_menor_preco]
            concorrentes_escolhidos = random.sample(concorrentes, 2)

            caminhos.append(geradores[empresa_menor_preco](0.0, self.itens))

            acrescimos_perdedoras = [5.0, 10.0]
            random.shuffle(acrescimos_perdedoras)

            ordens_ja_usadas = [tuple(id(x) for x in self.itens)]

            for i, empresa in enumerate(concorrentes_escolhidos):
                itens_embaralhados = criar_itens_embaralhados(self.itens, ordens_ja_usadas)
                ordens_ja_usadas.append(tuple(id(x) for x in itens_embaralhados))
                caminhos.append(geradores[empresa](acrescimos_perdedoras[i], itens_embaralhados))
            
        return caminhos


# ==========================================================================
# Interface Gráfica (Tkinter)
# ==========================================================================

class OrcamentoApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Gerador Profissional de Orçamentos - Douglas Oliveira")
        self.root.geometry("1000x980")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        bg_color = "#e6edf5"
        self.root.configure(bg=bg_color)

        self.style.configure(".", background=bg_color, font=("Segoe UI", 10))
        self.style.configure("TLabel", background=bg_color, font=("Segoe UI", 10), foreground="#1a252f")
        self.style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"), foreground="#1b4f72")
        self.style.configure("TLabelframe", background=bg_color, bordercolor="#a9cce3", lightcolor="#a9cce3", darkcolor="#a9cce3")
        self.style.configure("TLabelframe.Label", background=bg_color, font=("Segoe UI", 11, "bold"), foreground="#154360")

        self.linhas: List[dict] = []

        # SISTEMA DE ROLAGEM
        container_principal = ttk.Frame(root)
        container_principal.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(container_principal, bg=bg_color, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container_principal, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding=15)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self._bind_mousewheel()

        self._montar_frame_logos(self.scrollable_frame)
        self._montar_frame_meta(self.scrollable_frame)
        self._montar_frame_tabela(self.scrollable_frame)
        self._montar_botoes(self.scrollable_frame)
        self._montar_rodape(self.scrollable_frame)

        self.adicionar_linha()

    def _bind_mousewheel(self) -> None:
        def _on_mousewheel(event):
            if platform.system() == "Windows":
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif platform.system() == "Darwin":
                self.canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                if event.num == 4:
                    self.canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    self.canvas.yview_scroll(1, "units")

        self.root.bind_all("<MouseWheel>", _on_mousewheel)
        self.root.bind_all("<Button-4>", _on_mousewheel)
        self.root.bind_all("<Button-5>", _on_mousewheel)

    def _montar_frame_logos(self, main_frame: ttk.Frame) -> None:
        self.frame_logos = ttk.LabelFrame(main_frame, text=" Logotipos e Carimbos das Empresas ", padding=12)
        self.frame_logos.pack(fill="x", pady=5)

        self.path_nahora = tk.StringVar()
        self.path_carimbo_nahora = tk.StringVar()
        self.path_dakar = tk.StringVar()
        self.path_carimbo_dakar = tk.StringVar()
        self.path_lider_comercio = tk.StringVar()
        self.path_carimbo_lider_comercio = tk.StringVar()
        self.path_lider_sport = tk.StringVar()
        self.path_carimbo_lider_sport = tk.StringVar()
        
        config_salva = carregar_config_logos()
        
        if config_salva.get("nahora") and os.path.exists(config_salva["nahora"]):
            self.path_nahora.set(config_salva["nahora"])
        if config_salva.get("carimbo_nahora") and os.path.exists(config_salva["carimbo_nahora"]):
            self.path_carimbo_nahora.set(config_salva["carimbo_nahora"])
            
        if config_salva.get("dakar") and os.path.exists(config_salva["dakar"]):
            self.path_dakar.set(config_salva["dakar"])
        if config_salva.get("carimbo_dakar") and os.path.exists(config_salva["carimbo_dakar"]):
            self.path_carimbo_dakar.set(config_salva["carimbo_dakar"])
            
        lider_com_path = config_salva.get("lider_comercio") or config_salva.get("lider")
        if lider_com_path and os.path.exists(lider_com_path):
            self.path_lider_comercio.set(lider_com_path)
        if config_salva.get("carimbo_lider_comercio") and os.path.exists(config_salva["carimbo_lider_comercio"]):
            self.path_carimbo_lider_comercio.set(config_salva["carimbo_lider_comercio"])
            
        if config_salva.get("lider_sport") and os.path.exists(config_salva["lider_sport"]):
            self.path_lider_sport.set(config_salva["lider_sport"])
        if config_salva.get("carimbo_lider_sport") and os.path.exists(config_salva["carimbo_lider_sport"]):
            self.path_carimbo_lider_sport.set(config_salva["carimbo_lider_sport"])

        itens_config = [
            ("Logo Na Hora:", self.path_nahora),
            ("Carimbo Na Hora:", self.path_carimbo_nahora),
            ("Logo Dakar Sport:", self.path_dakar),
            ("Carimbo Dakar Sport:", self.path_carimbo_dakar),
            ("Logo Líder Comércio:", self.path_lider_comercio),
            ("Carimbo Líder Comércio:", self.path_carimbo_lider_comercio),
            ("Logo Lider Sport:", self.path_lider_sport),
            ("Carimbo Lider Sport:", self.path_carimbo_lider_sport)
        ]

        for i, (label_text, string_var) in enumerate(itens_config):
            ttk.Label(self.frame_logos, text=label_text, style="Header.TLabel").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ttk.Entry(self.frame_logos, textvariable=string_var, width=65).grid(row=i, column=1, padx=5, pady=2)
            tk.Button(
                self.frame_logos, text="Procurar...", command=lambda v=string_var: self._escolher_imagem(v),
                bg="#eaeded", font=("Segoe UI", 9), relief="raised", cursor="hand2"
            ).grid(row=i, column=2, padx=5, pady=2)

    def _salvar_estado_logos(self) -> None:
        estado = {
            "nahora": self.path_nahora.get().strip(),
            "carimbo_nahora": self.path_carimbo_nahora.get().strip(),
            "dakar": self.path_dakar.get().strip(),
            "carimbo_dakar": self.path_carimbo_dakar.get().strip(),
            "lider_comercio": self.path_lider_comercio.get().strip(),
            "carimbo_lider_comercio": self.path_carimbo_lider_comercio.get().strip(),
            "lider_sport": self.path_lider_sport.get().strip(),
            "carimbo_lider_sport": self.path_carimbo_lider_sport.get().strip()
        }
        salvar_config_logos(estado)

    def _escolher_imagem(self, string_var: tk.StringVar) -> None:
        arquivo = filedialog.askopenfilename(
            title="Selecione a Imagem",
            filetypes=[("Arquivos de Imagem", "*.png *.jpg *.jpeg"), ("Todos os Arquivos", "*.*")]
        )
        if arquivo:
            string_var.set(arquivo)
            self._salvar_estado_logos()

    def _montar_frame_meta(self, main_frame: ttk.Frame) -> None:
        self.frame_meta = ttk.LabelFrame(main_frame, text=" Informações do Cabeçalho ", padding=12)
        self.frame_meta.pack(fill="x", pady=5)

        ttk.Label(self.frame_meta, text="Cidade:", style="Header.TLabel").grid(row=0, column=0, sticky="w", padx=5, pady=6)
        self.ent_cidade = ttk.Entry(self.frame_meta, width=22)
        self.ent_cidade.grid(row=0, column=1, sticky="w", padx=5, pady=6)
        self.ent_cidade.insert(0, "GOIÂNIA")

        ttk.Label(self.frame_meta, text="Data (DD/MM/AAAA):", style="Header.TLabel").grid(row=0, column=2, sticky="w", padx=5, pady=6)
        self.ent_data = ttk.Entry(self.frame_meta, width=22)
        self.ent_data.grid(row=0, column=3, sticky="w", padx=5, pady=6)
        self.ent_data.insert(0, "10/06/2026")

        ttk.Label(self.frame_meta, text="Aos Cuidados de (A/C):", style="Header.TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=6)
        self.ent_ac = ttk.Entry(self.frame_meta, width=50)
        self.ent_ac.grid(row=1, column=1, columnspan=3, sticky="w", padx=5, pady=6)
        self.ent_ac.insert(0, "PREFEITURA MUNICIPAL DE ALVORADA")

        ttk.Label(self.frame_meta, text="Empresa Menor Preço:", style="Header.TLabel", foreground="#0b5345").grid(row=2, column=0, sticky="w", padx=5, pady=6)
        self.var_menor_preco = tk.StringVar(value="Na Hora")
        empresas = ["Na Hora", "Dakar Sport", "Líder Comércio", "Lider Sport"]
        ttk.OptionMenu(self.frame_meta, self.var_menor_preco, empresas[0], *empresas).grid(row=2, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(self.frame_meta, text="Formato de Saída:", style="Header.TLabel", foreground="#7d6608").grid(row=2, column=2, sticky="w", padx=5, pady=6)
        self.var_formato = tk.StringVar(value="JPG")
        ttk.OptionMenu(self.frame_meta, self.var_formato, "JPG", "JPG", "PDF", "PNG").grid(row=2, column=3, sticky="w", padx=5, pady=6)

        ttk.Label(self.frame_meta, text="Condição Pagamento:", style="Header.TLabel").grid(row=3, column=0, sticky="w", padx=5, pady=6)
        self.var_pagamento = tk.StringVar(value="50% do valor de entrada e 50% na retirada")
        ttk.Combobox(
            self.frame_meta, textvariable=self.var_pagamento, width=48,
            values=["50% do valor de entrada e 50% na retirada", "À Vista", "Boleto Bancário"],
        ).grid(row=3, column=1, columnspan=3, sticky="w", padx=5, pady=6)

        ttk.Label(self.frame_meta, text="Prazo de Entrega:", style="Header.TLabel").grid(row=4, column=0, sticky="w", padx=5, pady=6)
        self.ent_prazo = ttk.Entry(self.frame_meta, width=22)
        self.ent_prazo.grid(row=4, column=1, sticky="w", padx=5, pady=6)
        self.ent_prazo.insert(0, "15 - 20 DIAS")

        ttk.Label(self.frame_meta, text="Validade Orçamento:", style="Header.TLabel").grid(row=4, column=2, sticky="w", padx=5, pady=6)
        self.ent_validade = ttk.Entry(self.frame_meta, width=22)
        self.ent_validade.grid(row=4, column=3, sticky="w", padx=5, pady=6)
        self.ent_validade.insert(0, "30 DIAS")

        ttk.Label(self.frame_meta, text="Salvar na Pasta:", style="Header.TLabel", foreground="#117a65").grid(row=5, column=0, sticky="w", padx=5, pady=6)
        self.ent_pasta = ttk.Entry(self.frame_meta, width=40)
        self.ent_pasta.grid(row=5, column=1, columnspan=2, sticky="w", padx=5, pady=6)
        self.ent_pasta.insert(0, PASTA_PADRAO_ORCAMENTOS)

        tk.Button(
            self.frame_meta, text="Procurar...", command=self.escolher_pasta, bg="#eaeded",
            font=("Segoe UI", 9), relief="raised", cursor="hand2",
        ).grid(row=5, column=3, sticky="w", padx=5, pady=6)

        ttk.Label(self.frame_meta, text="Modo de Geração:", style="Header.TLabel", foreground="#8e44ad").grid(row=6, column=0, sticky="w", padx=5, pady=6)
        self.var_modo_geracao = tk.StringVar(value="Gerar 3 (1 Vencedor + 2 Aleatórios)")
        ttk.OptionMenu(
            self.frame_meta, self.var_modo_geracao, "Gerar 3 (1 Vencedor + 2 Aleatórios)", 
            "Gerar 3 (1 Vencedor + 2 Aleatórios)", "Apenas Menor Preço"
        ).grid(row=6, column=1, sticky="w", padx=5, pady=6)

        self.var_abrir_auto = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.frame_meta, text="Abrir arquivos automaticamente após gerar", variable=self.var_abrir_auto
        ).grid(row=6, column=2, columnspan=2, sticky="w", padx=5, pady=6)

    def _montar_frame_tabela(self, main_frame: ttk.Frame) -> None:
        self.frame_tabela = ttk.LabelFrame(main_frame, text=" Itens do Orçamento ", padding=12)
        self.frame_tabela.pack(fill="x", pady=5)

        for i, h in enumerate(["Nº", "Descrição", "UND", "QTD", "Valor Unit.", "Valor Total"]):
            ttk.Label(self.frame_tabela, text=h, font=("Segoe UI", 10, "bold")).grid(row=0, column=i, padx=5, pady=5)

        self.lbl_texto_total = ttk.Label(self.frame_tabela, text="TOTAL GERAL BASE:", font=("Segoe UI", 10, "bold"))
        self.lbl_total_geral = ttk.Label(self.frame_tabela, text="R$ 0,00", font=("Segoe UI", 10, "bold"), foreground="#1b4f72")

    def _montar_botoes(self, main_frame: ttk.Frame) -> None:
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(pady=15)

        tk.Button(
            frame_botoes, text="+ Adicionar Nova Linha", command=self.adicionar_linha,
            bg="#d4edda", fg="#155724", font=("Segoe UI", 10, "bold"), relief="raised", bd=2, padx=12, pady=6, cursor="hand2"
        ).grid(row=0, column=0, padx=10, pady=5)

        tk.Button(
            frame_botoes, text="- Remover Última Linha", command=self.remover_ultima_linha,
            bg="#f8d7da", fg="#721c24", font=("Segoe UI", 10, "bold"), relief="raised", bd=2, padx=12, pady=6, cursor="hand2"
        ).grid(row=0, column=1, padx=10, pady=5)

        tk.Button(
            frame_botoes, text="📂 Carregar Projeto", command=self.carregar_projeto,
            bg="#cce5ff", fg="#004085", font=("Segoe UI", 10, "bold"), relief="raised", bd=2, padx=12, pady=6, cursor="hand2"
        ).grid(row=1, column=0, padx=10, pady=5)

        tk.Button(
            frame_botoes, text="💾 Salvar Projeto", command=self.salvar_projeto,
            bg="#fff3cd", fg="#856404", font=("Segoe UI", 10, "bold"), relief="raised", bd=2, padx=12, pady=6, cursor="hand2"
        ).grid(row=1, column=1, padx=10, pady=5)

        tk.Button(
            frame_botoes, text="🚀 Gerar Orçamentos", command=self.gerar_orcamentos,
            bg="#d1ecf1", fg="#0c5460", font=("Segoe UI", 10, "bold"), relief="raised", bd=2, padx=18, pady=6, cursor="hand2"
        ).grid(row=0, column=2, rowspan=2, padx=15, pady=5, sticky="ns")

    def _montar_rodape(self, main_frame: ttk.Frame) -> None:
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(side="bottom", fill="x", pady=5)
        ttk.Label(
            footer_frame, text="Programa Desenvolvido por: Douglas Oliveira  |  email: getyourwish10@gmail.com",
            font=("Segoe UI", 9, "italic"), foreground="#555555"
        ).pack(anchor="center")

    def salvar_projeto(self):
        dados = {
            "cabecalho": {
                "cidade": self.ent_cidade.get(),
                "data": self.ent_data.get(),
                "ac": self.ent_ac.get(),
                "menor_preco": self.var_menor_preco.get(),
                "formato": self.var_formato.get(),
                "condicao": self.var_pagamento.get(),
                "prazo": self.ent_prazo.get(),
                "validade": self.ent_validade.get(),
                "pasta": self.ent_pasta.get(),
                "modo_geracao": self.var_modo_geracao.get(),
                "abrir_auto": self.var_abrir_auto.get()
            },
            "itens": []
        }
        
        for linha in self.linhas:
            dados["itens"].append({
                "item": linha["item"].get(),
                "desc": linha["desc"].get(),
                "unid": linha["unid"].get(),
                "qtd": linha["qtd"].get(),
                "vlr": linha["vlr"].get()
            })

        pasta_padrao = self.ent_pasta.get().strip() or PASTA_PADRAO_ORCAMENTOS
        os.makedirs(pasta_padrao, exist_ok=True)
        nome_sugerido = limpar_nome_arquivo(self.ent_ac.get()) or "Projeto_Orcamento"
        
        caminho = filedialog.asksaveasfilename(
            initialdir=pasta_padrao,
            initialfile=nome_sugerido,
            defaultextension=".json",
            filetypes=[("Arquivo de Projeto JSON", "*.json"), ("Todos Arquivos", "*.*")],
            title="Salvar Projeto"
        )
        
        if caminho:
            try:
                with open(caminho, 'w', encoding='utf-8') as f:
                    json.dump(dados, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("Sucesso", "Projeto salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar:\n{e}")

    def carregar_projeto(self):
        pasta_padrao = self.ent_pasta.get().strip() or PASTA_PADRAO_ORCAMENTOS
        caminho = filedialog.askopenfilename(
            initialdir=pasta_padrao,
            filetypes=[("Arquivo de Projeto JSON", "*.json"), ("Todos Arquivos", "*.*")],
            title="Carregar Projeto"
        )
        
        if not caminho:
            return

        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                dados = json.load(f)

            cab = dados.get("cabecalho", {})
            self.ent_cidade.delete(0, tk.END); self.ent_cidade.insert(0, cab.get("cidade", ""))
            self.ent_data.delete(0, tk.END); self.ent_data.insert(0, cab.get("data", ""))
            self.ent_ac.delete(0, tk.END); self.ent_ac.insert(0, cab.get("ac", ""))
            
            menor_pr_salvo = cab.get("menor_preco", "Na Hora")
            if menor_pr_salvo == "Líder Comércio": self.var_menor_preco.set("Líder Comércio")
            else: self.var_menor_preco.set(menor_pr_salvo)
            
            self.var_formato.set(cab.get("formato", "JPG"))
            self.var_pagamento.set(cab.get("condicao", ""))
            self.ent_prazo.delete(0, tk.END); self.ent_prazo.insert(0, cab.get("prazo", ""))
            self.ent_validade.delete(0, tk.END); self.ent_validade.insert(0, cab.get("validade", ""))
            self.ent_pasta.delete(0, tk.END); self.ent_pasta.insert(0, cab.get("pasta", ""))
            self.var_modo_geracao.set(cab.get("modo_geracao", "Gerar 3 (1 Vencedor + 2 Aleatórios)"))
            self.var_abrir_auto.set(cab.get("abrir_auto", True))

            for linha in self.linhas:
                for chave in ("item", "desc", "unid", "ent_qtd", "ent_vlr", "lbl_total"):
                    linha[chave].destroy()
            self.linhas.clear()

            itens = dados.get("itens", [])
            if not itens:
                self.adicionar_linha()
            else:
                for d_item in itens:
                    self.adicionar_linha()
                    linha_atual = self.linhas[-1]
                    linha_atual["item"].delete(0, tk.END); linha_atual["item"].insert(0, d_item.get("item", ""))
                    linha_atual["desc"].delete(0, tk.END); linha_atual["desc"].insert(0, d_item.get("desc", ""))
                    linha_atual["unid"].delete(0, tk.END); linha_atual["unid"].insert(0, d_item.get("unid", ""))
                    linha_atual["qtd"].set(d_item.get("qtd", "1"))
                    linha_atual["vlr"].set(d_item.get("vlr", "0,00"))

            self.atualizar_totais()
            messagebox.showinfo("Projeto Carregado", "Orçamento recuperado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Carregar", f"Falha no arquivo:\n{e}")

    def escolher_pasta(self) -> None:
        pasta_atual = self.ent_pasta.get()
        initialdir = pasta_atual if os.path.isdir(pasta_atual) else os.path.expanduser("~")
        pasta_selecionada = filedialog.askdirectory(initialdir=initialdir)
        if pasta_selecionada:
            self.ent_pasta.delete(0, tk.END)
            self.ent_pasta.insert(0, pasta_selecionada)

    def atualizar_totais(self, *_args) -> None:
        total_geral = 0.0
        for linha in self.linhas:
            qtd = limpar_numero(linha["qtd"].get())
            vlr = limpar_numero(linha["vlr"].get())
            total_linha = qtd * vlr
            linha["lbl_total"].config(text=f"R$ {formatar_valor(total_linha)}")
            total_geral += total_linha
        self.lbl_total_geral.config(text=f"R$ {formatar_valor(total_geral)}")

    def _reposicionar_rodape_tabela(self) -> None:
        linha_rodape = len(self.linhas) + 1
        self.lbl_texto_total.grid(row=linha_rodape, column=4, sticky="e", padx=5, pady=15)
        self.lbl_total_geral.grid(row=linha_rodape, column=5, sticky="e", padx=5, pady=15)

    def adicionar_linha(self) -> None:
        row_idx = len(self.linhas) + 1

        ent_item = ttk.Entry(self.frame_tabela, width=5, justify="center")
        ent_item.grid(row=row_idx, column=0, padx=5, pady=5)
        ent_item.insert(0, str(len(self.linhas) + 1))

        ent_desc = ttk.Entry(self.frame_tabela, width=35)
        ent_desc.grid(row=row_idx, column=1, padx=5, pady=5)

        ent_unid = ttk.Entry(self.frame_tabela, width=8, justify="center")
        ent_unid.grid(row=row_idx, column=2, padx=5, pady=5)
        ent_unid.insert(0, "UND")

        var_qtd = tk.StringVar(value="1")
        ent_qtd = ttk.Entry(self.frame_tabela, textvariable=var_qtd, width=10, justify="center")
        ent_qtd.grid(row=row_idx, column=3, padx=5, pady=5)

        var_vlr = tk.StringVar(value="0,00")
        ent_vlr = ttk.Entry(self.frame_tabela, textvariable=var_vlr, width=12, justify="right")
        ent_vlr.grid(row=row_idx, column=4, padx=5, pady=5)

        lbl_total = ttk.Label(self.frame_tabela, text="R$ 0,00", width=15, anchor="e")
        lbl_total.grid(row=row_idx, column=5, padx=5, pady=5)

        var_qtd.trace_add("write", self.atualizar_totais)
        var_vlr.trace_add("write", self.atualizar_totais)

        self.linhas.append({
            "item": ent_item, "desc": ent_desc, "unid": ent_unid,
            "ent_qtd": ent_qtd, "ent_vlr": ent_vlr,
            "qtd": var_qtd, "vlr": var_vlr, "lbl_total": lbl_total,
        })

        self._reposicionar_rodape_tabela()
        self.atualizar_totais()
        self.root.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def remover_ultima_linha(self) -> None:
        if len(self.linhas) <= 1:
            messagebox.showinfo("Aviso", "É necessário manter ao menos um item na tabela.")
            return
        linha = self.linhas.pop()
        for chave in ("item", "desc", "unid", "ent_qtd", "ent_vlr", "lbl_total"):
            linha[chave].destroy()
        self._reposicionar_rodape_tabela()
        self.atualizar_totais()

    def _coletar_dados(self) -> Tuple[DadosCabecalho, List[ItemOrcamento], LogosEmpresas]:
        cidade = self.ent_cidade.get().strip().upper()
        data = self.ent_data.get().strip()
        ac = self.ent_ac.get().strip().upper()
        condicao = self.var_pagamento.get().strip().upper()
        prazo = self.ent_prazo.get().strip().upper()
        validade = self.ent_validade.get().strip().upper()
        pasta = self.ent_pasta.get().strip() or PASTA_PADRAO_ORCAMENTOS

        if not cidade:
            raise ValueError("Informe a cidade.")
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", data):
            raise ValueError("Data inválida. Use o formato DD/MM/AAAA.")
        if not ac:
            raise ValueError("Informe o destinatário (A/C).")

        self._salvar_estado_logos()

        cabecalho = DadosCabecalho(cidade, data, ac, condicao, prazo, validade, pasta)

        logos = LogosEmpresas(
            nahora=self.path_nahora.get().strip(),
            dakar=self.path_dakar.get().strip(),
            lider_comercio=self.path_lider_comercio.get().strip(),
            lider_sport=self.path_lider_sport.get().strip(),
            carimbo_nahora=self.path_carimbo_nahora.get().strip(),
            carimbo_dakar=self.path_carimbo_dakar.get().strip(),
            carimbo_lider_comercio=self.path_carimbo_lider_comercio.get().strip(),
            carimbo_lider_sport=self.path_carimbo_lider_sport.get().strip()
        )

        itens: List[ItemOrcamento] = []
        for linha in self.linhas:
            descricao = linha["desc"].get().strip()
            if not descricao:
                continue
            itens.append(ItemOrcamento(
                numero=linha["item"].get().strip() or str(len(itens) + 1),
                descricao=descricao,
                unidade=linha["unid"].get().strip() or "UND",
                quantidade=limpar_numero(linha["qtd"].get()),
                valor_unitario=limpar_numero(linha["vlr"].get()),
            ))

        if not itens:
            raise ValueError("Adicione ao menos um item com descrição válida.")

        return cabecalho, itens, logos

    def gerar_orcamentos(self) -> None:
        try:
            cabecalho, itens, logos = self._coletar_dados()
        except ValueError as erro:
            messagebox.showwarning("Dados inválidos", str(erro))
            return

        formato = self.var_formato.get().upper()
        empresa_menor_preco = self.var_menor_preco.get()
        apenas_menor_preco = (self.var_modo_geracao.get() == "Apenas Menor Preço")

        try:
            gerador = GeradorOrcamento(cabecalho, itens, formato, logos)
            caminhos = gerador.gerar(empresa_menor_preco, apenas_menor_preco)
        except Exception as erro:
            messagebox.showerror("Erro ao gerar", f"Ocorreu um erro inesperado:\n{erro}")
            return

        txt_gerados = "Orçamento gerado" if apenas_menor_preco else "Os 3 orçamentos (Menor Preço + 2 sorteados) foram gerados"

        mensagem = (
            f"{txt_gerados} com sucesso em formato {formato}!\n\n"
            f"Salvo(s) na pasta:\n{cabecalho.pasta_destino}\n\n"
            f"Empresa de menor preço (0%): {empresa_menor_preco}"
        )
        messagebox.showinfo("Sucesso!", mensagem)

        if self.var_abrir_auto.get():
            for caminho in caminhos:
                abrir_arquivo(caminho)


if __name__ == "__main__":
    root = tk.Tk()
    app = OrcamentoApp(root)
    root.mainloop()
