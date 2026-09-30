"""Transform: JSON cru -> linhas normalizadas. FUNCAO PURA.

Este modulo nao abre conexao, nao faz HTTP e nao le arquivo. Recebe os dois
payloads (o item da listagem e o detalhe) e devolve dicionarios prontos para o
etl_load. E isso que torna a fase 4 offline e reprocessavel: mudou uma regra,
roda `--fases 4 --refazer-transform` e o banco inteiro se corrige sem gastar
um unico request.

O que este modulo NAO faz, de proposito:

  * nao descarta pedido por situacao, natureza ou cliente do grupo -- isso vira
    coluna (eh_situacao_valida, eh_natureza_alvo, cliente_excluido). A regra da
    regra de canal ja mudou uma vez no historico; filtro na ingestao custaria
    horas de recoleta;
  * nao descarta item de quantidade <= 0 (como os relatorios antigos faziam) --
    so marca qtd_zero. Sem isso a checagem COUNT(pedido_item) contra
    JSON_LENGTH(payload,'$.itens') nunca fecharia ao inteiro;
  * nao explode kit: o Tiny ja entrega o kit destrinchado em componentes.
"""

import hashlib
import json
from datetime import datetime, date
from decimal import Decimal, InvalidOperation

import _bootstrap  # noqa: F401
import etl_regras as regras

_ZERO = Decimal("0")


# ------------------------------------------------------------------ #
#  Hash canonico do payload                                           #
# ------------------------------------------------------------------ #

def hash_payload(payload) -> str:
    """SHA-256 sobre o JSON canonico (chaves ordenadas, sem espaco).

    Calculado em PYTHON de proposito: o MySQL nao preserva a ordem das chaves ao
    guardar JSON, entao um SHA2() sobre a coluna lida nao seria estavel entre
    gravacao e releitura.
    """
    texto = json.dumps(payload, sort_keys=True, separators=(",", ":"),
                       ensure_ascii=False, default=str)
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def serializar(payload) -> str:
    """Forma canonica gravada na coluna JSON -- casa com hash_payload."""
    return json.dumps(payload, sort_keys=True, separators=(",", ":"),
                      ensure_ascii=False, default=str)


# ------------------------------------------------------------------ #
#  Coercao de tipos                                                   #
# ------------------------------------------------------------------ #

def _vazio(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip().lower() in ("", "null"))


def txt(v, n: int = 255):
    """String truncada, ou None. O truncamento e explicito porque o sql_mode
    estrito (que fica LIGADO) recusa a linha inteira em vez de cortar."""
    if _vazio(v):
        return None
    s = str(v).strip()
    return s[:n] if s else None


def inteiro(v):
    if _vazio(v):
        return None
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def dec(v, default=_ZERO):
    """Dinheiro e quantidade em Decimal -- nunca float, senao o teste de
    regressao nao fecha ao centavo."""
    if _vazio(v):
        return default
    try:
        return Decimal(str(v).replace(",", "."))
    except (InvalidOperation, TypeError, ValueError):
        return default


def data(v):
    """Aceita YYYY-MM-DD e DD/MM/YYYY. O Tiny manda '' para data ausente, e
    INSERT '' INTO DATE erra em sql_mode estrito -- daqui sai None."""
    if _vazio(v):
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    s = str(v).strip()
    for fmt, corte in (("%Y-%m-%d", 10), ("%d/%m/%Y", 10)):
        try:
            return datetime.strptime(s[:corte], fmt).date()
        except ValueError:
            continue
    return None


def data_hora(v):
    if _vazio(v):
        return None
    if isinstance(v, datetime):
        return v
    s = str(v).strip().replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            continue
    d = data(s)
    return datetime(d.year, d.month, d.day) if d else None


def obj(d, chave) -> dict:
    v = (d or {}).get(chave)
    return v if isinstance(v, dict) else {}


def lista(d, chave) -> list:
    v = (d or {}).get(chave)
    return v if isinstance(v, list) else []


# ------------------------------------------------------------------ #
#  Contexto de regras (montado uma vez por tenant)                    #
# ------------------------------------------------------------------ #

def contexto(tenant: str, regras_versao: int = 0) -> dict:
    """Regras de um tenant, montadas uma vez: perfil (situacoes, naturezas,
    canais), empresas do grupo e a versao das regras a gravar no pedido."""
    return {
        "tenant": tenant,
        "excluidos": regras.clientes_excluidos(tenant),
        "situacoes": regras.SITUACOES,
        "perfil": regras.perfil(tenant),
        "regras_versao": regras_versao,
    }


def _merge(lista_json: dict, detalhe_json: dict) -> dict:
    """Detalhe manda; a listagem preenche o que o detalhe nao trouxer.

    A listagem tem o cliente completo e o ecommerce; o detalhe tem os itens e a
    naturezaOperacao. Nenhum dos dois sozinho descreve o pedido.
    """
    base = dict(lista_json or {})
    for k, v in (detalhe_json or {}).items():
        if v in (None, "", [], {}):
            continue
        base[k] = v
    return base


# ------------------------------------------------------------------ #
#  Transform                                                          #
# ------------------------------------------------------------------ #

def transformar(ctx: dict, id_pedido: int, lista_json: dict,
                detalhe_json: dict | None, agora: datetime | None = None) -> dict:
    """Devolve {'pedido': {...}, 'itens': [...], 'enderecos': [...],
    'parcelas': [...], 'pagamentos': [...], 'marcadores': [...], 'dims': {...}}.

    Todas as chaves dos dicts sao nomes de coluna -- o etl_load monta as tuplas
    na ordem certa, para que a ordem viva num lugar so.
    """
    agora = agora or datetime.now()
    tenant = ctx["tenant"]
    tem_detalhe = bool(detalhe_json)
    p = _merge(lista_json, detalhe_json)

    cliente = obj(p, "cliente")
    end_cli = obj(cliente, "endereco")
    end_ent = obj(p, "enderecoEntrega")
    ecom = obj(p, "ecommerce")
    vend = obj(p, "vendedor")
    nat = obj(p, "naturezaOperacao")
    transp = obj(p, "transportador")
    forma_envio = obj(transp, "formaEnvio")
    forma_frete = obj(transp, "formaFrete")
    dep = obj(p, "deposito")
    inter = obj(p, "intermediador")
    lp = obj(p, "listaPreco")
    pag = obj(p, "pagamento")
    forma_rec = obj(pag, "formaRecebimento")
    meio_pag = obj(pag, "meioPagamento")
    cat_fin = obj(pag, "categoria")

    situacao = inteiro(p.get("situacao"))
    nat_nome = txt(nat.get("nome"), 160) or ""
    nat_norm = regras.norm(nat_nome)
    cliente_nome = txt(cliente.get("nome") or cliente.get("fantasia"), 255) or ""
    ecom_nome = txt(ecom.get("nome"), 160) or ""
    vend_nome = txt(vend.get("nome"), 255) or ""
    vend_id = str(vend.get("id") or "").strip()

    data_pedido = regras.data_de(p)
    perfil = ctx["perfil"]
    canal = perfil.classificar_canal(ecom_nome, vend_nome, vend_id)
    # Conflito: pedido com integracao E vendedor pessoa -- a regra de canal
    # teve de escolher um dos dois (por padrao o vendedor vence). A coluna
    # existe para que isso nao passe despercebido.
    conflito = perfil.conflito_canal(ecom_nome, vend_nome, vend_id)

    itens = []
    total_un = _ZERO
    for i, it in enumerate(lista(p, "itens"), start=1):
        prod = obj(it, "produto")
        itens.append({
            "tenant": tenant, "id_pedido": id_pedido, "seq": i,
            "id_produto": inteiro(prod.get("id")),
            "sku": txt(prod.get("sku") or prod.get("codigo"), 80) or "",
            "descricao_item": txt(prod.get("descricao") or prod.get("nome"), 255) or "",
            "tipo_item": txt(prod.get("tipo"), 8),
            "quantidade": dec(it.get("quantidade")),
            "valor_unitario": dec(it.get("valorUnitario")),
            "info_adicional": txt(it.get("infoAdicional"), 4000),
        })
        total_un += dec(it.get("quantidade"))

    pedido = {
        "tenant": tenant,
        "id_pedido": id_pedido,
        "numero_pedido": inteiro(p.get("numeroPedido") or p.get("numero")),
        "id_nota_fiscal": inteiro(p.get("idNotaFiscal")),
        "origem_pedido": inteiro(p.get("origemPedido")),
        "data_pedido": data_pedido,
        "data_prevista": data(p.get("dataPrevista")),
        "data_entrega": data(p.get("dataEntrega")),
        "data_faturamento": data(p.get("dataFaturamento")),
        "data_envio": data_hora(p.get("dataEnvio")),
        "situacao": situacao,
        "situacao_nome": txt(ctx["situacoes"].get(situacao), 32),
        "valor_total_produtos": dec(p.get("valorTotalProdutos")),
        "valor_total_pedido": dec(p.get("valorTotalPedido") or p.get("valor")),
        "valor_desconto": dec(p.get("valorDesconto")),
        "valor_frete": dec(p.get("valorFrete")),
        "valor_outras_despesas": dec(p.get("valorOutrasDespesas")),
        "numero_ordem_compra": txt(p.get("numeroOrdemCompra"), 80),
        "observacoes": txt(p.get("observacoes"), 8000),
        "observacoes_internas": txt(p.get("observacoesInternas"), 8000),
        "id_cliente": inteiro(cliente.get("id")) or 0,
        "cliente_nome": cliente_nome,
        "cliente_cpf_cnpj": txt(cliente.get("cpfCnpj"), 24),
        "cliente_tipo_pessoa": txt(cliente.get("tipoPessoa"), 1),
        "cliente_email": txt(cliente.get("email"), 255),
        "cliente_telefone": txt(cliente.get("telefone"), 40),
        "cliente_celular": txt(cliente.get("celular"), 40),
        "cliente_uf": txt(end_cli.get("uf"), 2),
        "cliente_municipio": txt(end_cli.get("municipio"), 160),
        "id_natureza": inteiro(nat.get("id")) or 0,
        "natureza_nome": nat_nome or None,
        "natureza_norm": nat_norm or None,
        "id_vendedor": inteiro(vend.get("id")) or 0,
        "vendedor_nome": vend_nome or None,
        "id_ecommerce": inteiro(ecom.get("id")) or 0,
        "ecommerce_nome": ecom_nome or None,
        "ecommerce_canal": txt(ecom.get("canalVenda"), 160),
        "numero_pedido_ecommerce": txt(ecom.get("numeroPedidoEcommerce"), 80),
        "numero_pedido_canal": txt(ecom.get("numeroPedidoCanalVenda"), 80),
        "id_transportador": inteiro(transp.get("id")) or 0,
        "transportador_nome": txt(transp.get("nome"), 255),
        "frete_por_conta": txt(transp.get("fretePorConta"), 1),
        "codigo_rastreamento": txt(transp.get("codigoRastreamento"), 160),
        "url_rastreamento": txt(transp.get("urlRastreamento"), 500),
        "id_forma_envio": inteiro(forma_envio.get("id")),
        "forma_envio_nome": txt(forma_envio.get("nome"), 160),
        "id_forma_frete": inteiro(forma_frete.get("id")),
        "forma_frete_nome": txt(forma_frete.get("nome"), 160),
        "id_deposito": inteiro(dep.get("id")) or 0,
        "deposito_nome": txt(dep.get("nome"), 160),
        "id_intermediador": inteiro(inter.get("id")) or 0,
        "intermediador_nome": txt(inter.get("nome"), 255),
        "id_lista_preco": inteiro(lp.get("id")),
        "lista_preco_nome": txt(lp.get("nome"), 160),
        "lista_preco_acrescimo": dec(lp.get("acrescimoDesconto"), None),
        "id_forma_recebimento": inteiro(forma_rec.get("id")),
        "forma_recebimento_nome": txt(forma_rec.get("nome"), 160),
        "id_meio_pagamento": inteiro(meio_pag.get("id")),
        "meio_pagamento_nome": txt(meio_pag.get("nome"), 160),
        "condicao_pagamento": txt(pag.get("condicaoPagamento"), 255),
        "id_categoria_financeira": inteiro(cat_fin.get("id")),
        "categoria_financeira": txt(cat_fin.get("descricao"), 255),
        "eh_situacao_valida": 1 if perfil.eh_situacao_valida(situacao) else 0,
        "eh_natureza_alvo": 1 if perfil.eh_natureza_venda(nat_nome) else 0,
        "cliente_excluido": 1 if regras.eh_excluido(cliente_nome,
                                                    ctx["excluidos"]) else 0,
        "canal": canal,
        "conflito_canal": 1 if conflito else 0,
        "tem_detalhe": 1 if tem_detalhe else 0,
        "regras_versao": ctx["regras_versao"],
        "qtd_itens": len(itens),
        "total_unidades": total_un,
        "criado_em": agora,
        "atualizado_em": agora,
    }

    enderecos = []
    if any(end_cli.get(k) for k in ("endereco", "municipio", "cep", "uf")):
        enderecos.append(_endereco(tenant, id_pedido, "cobranca", end_cli,
                                   cliente_nome, cliente.get("cpfCnpj"),
                                   cliente.get("tipoPessoa"),
                                   cliente.get("telefone"),
                                   cliente.get("inscricaoEstadual")))
    if any(end_ent.get(k) for k in ("endereco", "municipio", "cep", "uf",
                                    "nomeDestinatario")):
        enderecos.append(_endereco(tenant, id_pedido, "entrega", end_ent,
                                   end_ent.get("nomeDestinatario"),
                                   end_ent.get("cpfCnpj"),
                                   end_ent.get("tipoPessoa"),
                                   end_ent.get("telefone"),
                                   end_ent.get("inscricaoEstadual")))

    parcelas = []
    for i, par in enumerate(lista(pag, "parcelas") or lista(p, "parcelas"),
                            start=1):
        fr = obj(par, "formaRecebimento")
        mp = obj(par, "meioPagamento")
        parcelas.append({
            "tenant": tenant, "id_pedido": id_pedido, "seq": i,
            "dias": inteiro(par.get("dias")),
            "data_vencimento": data(par.get("data")),
            "valor": dec(par.get("valor")),
            "observacoes": txt(par.get("observacoes"), 500),
            "id_forma_recebimento": inteiro(fr.get("id")),
            "forma_recebimento_nome": txt(fr.get("nome"), 160),
            "id_meio_pagamento": inteiro(mp.get("id")),
            "meio_pagamento_nome": txt(mp.get("nome"), 160),
        })

    pagamentos = []
    for i, pgi in enumerate(lista(p, "pagamentosIntegrados"), start=1):
        pagamentos.append({
            "tenant": tenant, "id_pedido": id_pedido, "seq": i,
            "valor": dec(pgi.get("valor")),
            "tipo_pagamento": inteiro(pgi.get("tipoPagamento")),
            "cnpj_intermediador": txt(pgi.get("cnpjIntermediador"), 24),
            "codigo_autorizacao": txt(pgi.get("codigoAutorizacao"), 160),
            "codigo_bandeira": inteiro(pgi.get("codigoBandeira")),
        })

    marcadores = []
    for m in lista(p, "marcadores"):
        nome = m.get("descricao") or m.get("nome") if isinstance(m, dict) else m
        nome = txt(nome, 120)
        if nome:
            marcadores.append({"tenant": tenant, "id_pedido": id_pedido,
                               "marcador": nome})

    return {
        "pedido": pedido,
        "itens": itens,
        "enderecos": enderecos,
        "parcelas": parcelas,
        "pagamentos": pagamentos,
        "marcadores": marcadores,
        "dims": _dims(ctx, agora, cliente, end_cli, cliente_nome, vend, vend_id,
                      vend_nome, ecom, ecom_nome, nat, nat_nome, nat_norm,
                      transp, dep, inter, lp, forma_rec, meio_pag, cat_fin,
                      forma_envio, forma_frete, itens),
    }


def _endereco(tenant, id_pedido, tipo, e, nome, cpf, tipo_pessoa, telefone, ie):
    return {
        "tenant": tenant, "id_pedido": id_pedido, "tipo": tipo,
        "nome_destinatario": txt(nome, 255),
        "cpf_cnpj": txt(cpf, 24),
        "tipo_pessoa": txt(tipo_pessoa, 1),
        "telefone": txt(telefone, 40),
        "inscricao_estadual": txt(ie, 40),
        "endereco": txt(e.get("endereco"), 255),
        "numero": txt(e.get("numero"), 30),
        "complemento": txt(e.get("complemento"), 160),
        "bairro": txt(e.get("bairro"), 160),
        "municipio": txt(e.get("municipio"), 160),
        "cep": txt(e.get("cep"), 12),
        "uf": txt(e.get("uf"), 2),
        "pais": txt(e.get("pais") or e.get("país"), 60),
    }


def _dims(ctx, agora, cliente, end_cli, cliente_nome, vend, vend_id, vend_nome,
          ecom, ecom_nome, nat, nat_nome, nat_norm, transp, dep, inter, lp,
          forma_rec, meio_pag, cat_fin, forma_envio, forma_frete, itens) -> dict:
    tenant = ctx["tenant"]
    d = {"cliente": [], "vendedor": [], "ecommerce": [], "natureza": [],
         "transportador": [], "deposito": [], "intermediador": [],
         "referencia": [], "produto": []}

    id_cli = inteiro(cliente.get("id"))
    if id_cli:
        d["cliente"].append({
            "tenant": tenant, "id_cliente": id_cli,
            "nome": cliente_nome, "nome_norm": regras.norm(cliente_nome)[:255],
            "codigo": txt(cliente.get("codigo"), 60),
            "fantasia": txt(cliente.get("fantasia"), 255),
            "tipo_pessoa": txt(cliente.get("tipoPessoa"), 1),
            "cpf_cnpj": txt(cliente.get("cpfCnpj"), 24),
            "inscricao_estadual": txt(cliente.get("inscricaoEstadual"), 40),
            "rg": txt(cliente.get("rg"), 40),
            "telefone": txt(cliente.get("telefone"), 40),
            "celular": txt(cliente.get("celular"), 40),
            "email": txt(cliente.get("email"), 255),
            "endereco": txt(end_cli.get("endereco"), 255),
            "numero": txt(end_cli.get("numero"), 30),
            "complemento": txt(end_cli.get("complemento"), 160),
            "bairro": txt(end_cli.get("bairro"), 160),
            "municipio": txt(end_cli.get("municipio"), 160),
            "cep": txt(end_cli.get("cep"), 12),
            "uf": txt(end_cli.get("uf"), 2),
            "pais": txt(end_cli.get("pais") or end_cli.get("país"), 60),
            "eh_cliente_grupo": 1 if regras.eh_excluido(
                cliente_nome, ctx["excluidos"]) else 0,
            "primeiro_visto": agora, "ultimo_visto": agora,
        })

    id_v = inteiro(vend.get("id"))
    if id_v:
        d["vendedor"].append({
            "tenant": tenant, "id_vendedor": id_v, "nome": vend_nome,
            "eh_pessoa": 1 if ctx["perfil"].eh_vendedor_pessoa(vend_nome, vend_id) else 0,
            "ultimo_visto": agora})

    id_e = inteiro(ecom.get("id"))
    if id_e:
        d["ecommerce"].append({
            "tenant": tenant, "id_ecommerce": id_e, "nome": ecom_nome,
            "canal_venda": txt(ecom.get("canalVenda"), 160),
            "ultimo_visto": agora})

    id_n = inteiro(nat.get("id"))
    if id_n or nat_nome:
        d["natureza"].append({
            "tenant": tenant, "id_natureza": id_n or 0, "nome": nat_nome,
            "nome_norm": nat_norm[:160],
            "eh_natureza_alvo": 1 if ctx["perfil"].eh_natureza_venda(nat_nome) else 0,
            "ultimo_visto": agora})

    id_t = inteiro(transp.get("id"))
    if id_t:
        d["transportador"].append({
            "tenant": tenant, "id_transportador": id_t,
            "nome": txt(transp.get("nome"), 255) or "", "ultimo_visto": agora})

    id_d = inteiro(dep.get("id"))
    if id_d:
        d["deposito"].append({
            "tenant": tenant, "id_deposito": id_d,
            "nome": txt(dep.get("nome"), 160) or "", "ultimo_visto": agora})

    id_i = inteiro(inter.get("id"))
    if id_i:
        d["intermediador"].append({
            "tenant": tenant, "id_intermediador": id_i,
            "nome": txt(inter.get("nome"), 255) or "",
            "cnpj": txt(inter.get("cnpj"), 24),
            "cnpj_instituicao": txt(inter.get("cnpjPagamentoInstituicao"), 24),
            "ultimo_visto": agora})

    for dominio, o in (("forma_recebimento", forma_rec),
                       ("meio_pagamento", meio_pag),
                       ("forma_envio", forma_envio),
                       ("forma_frete", forma_frete),
                       ("lista_preco", lp)):
        rid = inteiro(o.get("id"))
        if rid:
            d["referencia"].append({
                "tenant": tenant, "dominio": dominio, "id_ref": rid,
                "nome": txt(o.get("nome"), 255) or "", "ultimo_visto": agora})
    rid = inteiro(cat_fin.get("id"))
    if rid:
        d["referencia"].append({
            "tenant": tenant, "dominio": "categoria_financeira", "id_ref": rid,
            "nome": txt(cat_fin.get("descricao"), 255) or "",
            "ultimo_visto": agora})

    # Esqueleto do produto vindo do item. A fase 3 completa tipo/categoria --
    # que e a UNICA forma de saber se e kit.
    vistos = set()
    for it in itens:
        pid = it["id_produto"]
        if pid and pid not in vistos:
            vistos.add(pid)
            d["produto"].append({
                "tenant": tenant, "id_produto": pid, "sku": it["sku"],
                "descricao": it["descricao_item"], "atualizado_em": agora})

    return d
