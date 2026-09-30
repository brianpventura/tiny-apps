"""Regras de venda: MECANISMO generico + PERFIL configuravel por empresa.

O que decide se um pedido e "venda valida", em que canal ele cai e quem e
cliente do proprio grupo muda de uma empresa para outra. Por isso este modulo
separa duas coisas:

  * mecanismo (aqui, no codigo): normalizar texto, ler a data do pedido,
    casar nome de natureza, aplicar as regras de canal em ordem, reconhecer
    transferencia interna. Serve para qualquer conta do Tiny.
  * politica (no perfil): QUAIS situacoes contam, QUAIS naturezas sao de venda,
    QUAIS termos levam a cada canal. Vem de config/regras.padrao.json (padrao
    generico) sobreposto por regras.json da empresa (pasta de
    core.config.config_dir(): TINY_APPS_CONFIG ou data/).

Uso:
    from core import regras_vendas as rv
    p = rv.perfil("empresa_a")                  # regras de uma empresa
    p = rv.perfil("empresa_a", "recorrencia")   # recorte de um relatorio
    p.eh_situacao_valida(5); p.eh_natureza_venda("Venda de Mercadoria")
    p.classificar_canal("Shopify", "Ecommerce")

Ordem de sobreposicao (a de baixo vence; listas substituem, nao somam):
    config/regras.padrao.json
    regras.json da empresa (raiz)
    regras.json -> tenants.<tenant_id>
    regras.json -> relatorios.<nome>
    regras.json -> relatorios.<nome>.tenants.<tenant_id>

Nome de empresa, razao social e empresas do grupo NAO ficam no perfil: vem de
data/tenants.json (clientes_excluidos + razao_social).

Este modulo nao faz I/O no import e nao mexe em sys.stdout.
"""

import argparse
import copy
import hashlib
import json
import unicodedata
from datetime import date, datetime
from functools import lru_cache

from . import config
from .auth_tiny import TENANT_NAMES, TENANTS_FILE

# Vocabulario do Tiny (Swagger /pedidos) -- nao e politica, e o mapa dos codigos.
SITUACOES = {
    0: "Aberta", 1: "Faturada", 2: "Cancelada", 3: "Aprovada",
    4: "Preparando Envio", 5: "Enviada", 6: "Entregue", 7: "Pronto Envio",
    8: "Dados Incompletos", 9: "Nao Entregue",
}

# Nome de exibicao por tenant (campo `nome` de data/tenants.json).
NOMES = TENANT_NAMES

ARQUIVO_PADRAO = config.PADROES_DIR / "regras.padrao.json"
ARQUIVO_EMPRESA = "regras.json"


# ------------------------------------------------------------------ #
#  Helpers genericos                                                  #
# ------------------------------------------------------------------ #

def norm(texto: str) -> str:
    """Minusculas, sem acento, espacos colapsados."""
    s = (texto or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.split())


def para_float(v, default=0.0) -> float:
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return default


def data_de(p: dict) -> date | None:
    """Data do pedido (dataCriacao > data > dataPedido)."""
    for campo in ("dataCriacao", "data", "dataPedido"):
        v = p.get(campo)
        if v:
            try:
                return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
            except ValueError:
                continue
    return None


def extrair_produto(item: dict) -> tuple[str, str, str, str]:
    """(produto_id, sku, nome, tipo_do_item) de um item de pedido.

    O 'tipo' aqui usa vocabulario diferente do cadastro ('P' no item, contra
    S/V/K/M/F em /produtos) e NAO serve para detectar kit.
    """
    prod = item.get("produto") or {}
    if not isinstance(prod, dict):
        prod = {}
    pid = str(prod.get("id") or "").strip()
    sku = str(prod.get("sku") or prod.get("codigo") or prod.get("id") or "").strip()
    nome = str(prod.get("descricao") or prod.get("nome") or "").strip()
    tipo = str(prod.get("tipo") or "").strip()
    return pid, sku, nome, tipo


def parse_data(txt: str) -> date:
    """Tipo de argparse: aceita YYYY-MM-DD ou DD/MM/YYYY."""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(txt.strip(), fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(
        f"data invalida: {txt!r} (use YYYY-MM-DD ou DD/MM/YYYY)")


# ------------------------------------------------------------------ #
#  Clientes do proprio grupo (de data/tenants.json)                   #
# ------------------------------------------------------------------ #

def clientes_excluidos(tenant: str) -> set[str]:
    """Razoes sociais do grupo que nunca contam como cliente.

    Inclui `clientes_excluidos` (as OUTRAS empresas do grupo) e TAMBEM a
    `razao_social` do proprio tenant: sem ela, a empresa vendendo para ela
    mesma passa como faturamento.
    """
    try:
        cfg = json.loads(TENANTS_FILE.read_text(encoding="utf-8"))
        conf = cfg.get(tenant, {})
        nomes = list(conf.get("clientes_excluidos", []))
        propria = conf.get("razao_social")
        if propria:
            nomes.append(propria)
        return {norm(c) for c in nomes if c}
    except Exception:
        return set()


def eh_excluido(nome_cliente: str, excluidos: set[str]) -> bool:
    """Transferencia interna entre empresas do grupo.

    Match por prefixo/substring nos dois sentidos porque o Tiny trunca a razao
    social (ex. 'EMPRESA A COSMETICOS PROFISSI'). Igualdade exata nunca casa.
    """
    n = norm(nome_cliente)
    if not n:
        return False
    return any(n == e or n.startswith(e) or e.startswith(n) for e in excluidos)


# ------------------------------------------------------------------ #
#  Perfil                                                             #
# ------------------------------------------------------------------ #

def _sobrepor(base: dict, extra: dict) -> dict:
    """Merge recursivo: dict funde, qualquer outro valor (inclusive lista) substitui."""
    saida = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if k.startswith("_"):
            continue
        if isinstance(v, dict) and isinstance(saida.get(k), dict):
            saida[k] = _sobrepor(saida[k], v)
        else:
            saida[k] = copy.deepcopy(v)
    return saida


def _sem_comentarios(d):
    if isinstance(d, dict):
        return {k: _sem_comentarios(v) for k, v in d.items() if not k.startswith("_")}
    if isinstance(d, list):
        return [_sem_comentarios(v) for v in d]
    return d


@lru_cache(maxsize=None)
def _arquivos() -> tuple[str, str | None]:
    padrao = ARQUIVO_PADRAO.read_text(encoding="utf-8")
    emp = config.caminho(ARQUIVO_EMPRESA)
    return padrao, (emp.read_text(encoding="utf-8") if emp.exists() else None)


def origem() -> str:
    """De onde vem o perfil da empresa (para banner e diagnostico)."""
    emp = config.caminho(ARQUIVO_EMPRESA)
    return str(emp) if emp.exists() else "(so o padrao generico: config/regras.padrao.json)"


class Regras:
    """Perfil efetivo de regras para um tenant (e, opcionalmente, um relatorio)."""

    def __init__(self, dados: dict, tenant: str | None, relatorio: str | None):
        self.tenant, self.relatorio = tenant, relatorio
        self.dados = dados
        v = dados.get("vendas", {})
        self.situacoes_validas = frozenset(int(s) for s in v.get("situacoes_validas", []))
        nat = v.get("naturezas", {})
        self._nat_modo = nat.get("modo", "prefixo")
        if self._nat_modo not in ("prefixo", "lista", "todas"):
            raise SystemExit(f"  [ERRO] regras: naturezas.modo invalido: {self._nat_modo!r}")
        self._nat_valores = tuple(norm(x) for x in nat.get("valores", []))
        self._nat_incluir = frozenset(norm(x) for x in nat.get("incluir", []))
        self._nat_excluir = frozenset(norm(x) for x in nat.get("excluir", []))
        c = dados.get("canais", {})
        vp = c.get("vendedor_pessoa", {})
        self._grupo_pessoa = vp.get("grupo", "Vendas internas")
        self._pessoa_prevalece = bool(vp.get("prevalece_sobre_integracao", True))
        self._nao_pessoa = frozenset(norm(x) for x in c.get("vendedores_nao_pessoa", []))
        self.vendedores_nao_pessoa = self._nao_pessoa
        self._regras_integracao = [(r["grupo"], tuple(norm(t) for t in r.get("contem", [])))
                                   for r in c.get("integracao", [])]
        self._integ_nao_reconhecida = c.get("integracao_nao_reconhecida", "Outros")
        self._sem_integ_nao_pessoa = c.get("sem_integracao_vendedor_nao_pessoa", "Outros")
        self._sem_integ = c.get("sem_integracao", "Outros")
        grupos = list(c.get("grupos", []))
        for g in ([self._grupo_pessoa, self._integ_nao_reconhecida,
                   self._sem_integ_nao_pessoa, self._sem_integ]
                  + [g for g, _ in self._regras_integracao]):
            if g not in grupos:
                grupos.append(g)
        self.grupos = grupos
        self.extras = {k: v for k, v in dados.items() if k not in ("vendas", "canais")}

    # -- identidade ------------------------------------------------- #
    @property
    def hash(self) -> str:
        """Impressao digital do perfil (sem comentarios). Muda => regra mudou."""
        base = json.dumps({"vendas": self.dados.get("vendas"), "canais": self.dados.get("canais")},
                          sort_keys=True, ensure_ascii=False)
        return hashlib.sha256(base.encode("utf-8")).hexdigest()

    # -- venda valida ----------------------------------------------- #
    def eh_situacao_valida(self, codigo) -> bool:
        try:
            return int(codigo) in self.situacoes_validas
        except (TypeError, ValueError):
            return False

    def eh_natureza_venda(self, nome: str) -> bool:
        n = norm(nome)
        if n in self._nat_excluir:
            return False
        if n in self._nat_incluir:
            return True
        if self._nat_modo == "todas":
            return True
        if self._nat_modo == "lista":
            return n in self._nat_valores
        return bool(n) and any(n.startswith(p) for p in self._nat_valores)

    def descrever_naturezas(self) -> str:
        if self._nat_modo == "todas":
            base = "todas as naturezas"
        elif self._nat_modo == "lista":
            base = f"lista de {len(self._nat_valores)} nomes"
        else:
            base = "nome comecando por " + " / ".join(repr(p) for p in self._nat_valores)
        if self._nat_incluir:
            base += f", + {len(self._nat_incluir)} incluida(s)"
        if self._nat_excluir:
            base += f", - {len(self._nat_excluir)} excluida(s)"
        return base

    # -- canal ------------------------------------------------------ #
    def eh_vendedor_pessoa(self, vend_nome: str, vend_id: str = "") -> bool:
        """True quando o vendedor do pedido e uma pessoa (= venda interna)."""
        n = norm(vend_nome)
        if n:
            return n not in self._nao_pessoa
        return bool(str(vend_id).strip() and str(vend_id).strip() != "0")

    def _grupo_da_integracao(self, ecom_nome: str) -> str | None:
        n = norm(ecom_nome)
        if not n:
            return None
        for grupo, termos in self._regras_integracao:
            if any(t in n for t in termos):
                return grupo
        return self._integ_nao_reconhecida

    def classificar_canal(self, ecom_nome: str, vend_nome: str, vend_id: str = "") -> str:
        """Grupo do canal de venda.

        Com `prevalece_sobre_integracao` (padrao), vendedor pessoa vence a
        integracao: pedido de marketplace com vendedora atribuida cai no grupo
        do vendedor. O nome cru da integracao continua sendo a fonte de verdade
        para auditoria.
        """
        pessoa = self.eh_vendedor_pessoa(vend_nome, vend_id)
        if pessoa and self._pessoa_prevalece:
            return self._grupo_pessoa
        g = self._grupo_da_integracao(ecom_nome)
        if g is not None:
            return g
        if pessoa:
            return self._grupo_pessoa
        if norm(vend_nome) in self._nao_pessoa:
            return self._sem_integ_nao_pessoa
        return self._sem_integ

    def conflito_canal(self, ecom_nome: str, vend_nome: str, vend_id: str = "") -> bool:
        """Pedido com integracao E vendedor pessoa (a regra teve de escolher)."""
        return bool(norm(ecom_nome)) and self.eh_vendedor_pessoa(vend_nome, vend_id)


@lru_cache(maxsize=None)
def perfil(tenant: str | None = None, relatorio: str | None = None) -> Regras:
    """Perfil efetivo. Sem tenant: o nivel da empresa (sem override por tenant)."""
    padrao_txt, emp_txt = _arquivos()
    dados = _sem_comentarios(json.loads(padrao_txt))
    emp = json.loads(emp_txt) if emp_txt else {}
    dados = _sobrepor(dados, {k: v for k, v in emp.items() if k not in ("tenants", "relatorios")})
    if tenant:
        dados = _sobrepor(dados, (emp.get("tenants") or {}).get(tenant, {}))
    if relatorio:
        rel = (emp.get("relatorios") or {}).get(relatorio, {})
        dados = _sobrepor(dados, {k: v for k, v in rel.items() if k != "tenants"})
        if tenant:
            dados = _sobrepor(dados, (rel.get("tenants") or {}).get(tenant, {}))
    return Regras(dados, tenant, relatorio)
