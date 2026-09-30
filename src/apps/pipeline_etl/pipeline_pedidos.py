"""
Pipeline ETL de pedidos -- API v3 Tiny -> MySQL local (todas as empresas de data/tenants.json).

Extrai GET /pedidos/{idPedido}, trata (padroniza e limpa) e carrega num
warehouse local. Substitui o padrao de cache JSON dos relatorios: coleta uma
vez, responde para sempre.

DECISOES DE NEGOCIO (acordadas com o usuario):
  - Staging RAW + tabelas normalizadas. O JSON integro de cada pedido fica em
    pedido_raw, e o transform e reprocessavel 100% offline a partir dele: mudar
    uma regra custa ZERO request, so `--fases 4 --refazer-transform`.
  - Todas as empresas no mesmo banco, com `tenant` compondo a chave.
  - TUDO CRU. Pedido cancelado, transferencia interna, natureza fora de venda e
    brinde a R$ 0,00 sao GRAVADOS; as regras viram colunas (eh_situacao_valida,
    eh_natureza_alvo, cliente_excluido, canal) e views (v_venda_item). Filtrar
    na ingestao custaria horas de recoleta quando a regra mudar -- e regra de
    canal muda (um marketplace migra de integrador e volta, por exemplo).
  - Item de quantidade <= 0 NAO e descartado (so marcado qtd_zero), senao a
    checagem contra JSON_LENGTH(payload,'$.itens') nao fecharia ao inteiro.
  - Kits NAO sao explodidos: o Tiny ja entrega o kit destrinchado.
  - Valor por produto e sempre BRUTO -- nao existe desconto em nivel de item.
  - As regras de negocio sao IMPORTADAS de core/regras_vendas.py, nunca recopiadas.

RESILIENCIA
  - 5xx/timeout/conexao caida: ja resolvido em core/tiny_client.py:_get.
  - 429: o _get respeita Retry-After; aqui a defesa e nao criar pressao nova
    (4 workers fixos, zero sleep no loop) e o GET_LOCK do MySQL, que impede dois
    processos do mesmo tenant (= 4 req/s na mesma conta = 429 em cadeia).
  - TokenExpiredError: parada LIMPA com tudo commitado e saida com codigo 2.
    Retomar e `--fases 2,3,4,5`, sem repetir um unico request.
  - Ctrl+C / queda do MySQL / maquina desligada: perde-se no maximo `--lote`
    pedidos em voo, porque o dado e o estado commitam na mesma transacao.

FASES
  0 esquema + seed (offline)        3 catalogo de produtos (1 req/produto novo)
  1 listagem da janela (~1 req/100) 4 transform + load (OFFLINE, nao valida token)
  2 detalhe (1 req/pedido)          5 verificacao + regressao (OFFLINE)

Uso:
    python pipeline_pedidos.py --sondar-schema <tenant>  # 4 req: JSON real -> data/
    python pipeline_pedidos.py --fases 0                 # so cria/atualiza o esquema
    python pipeline_pedidos.py --inicio 2026-08-21 --fim 2026-09-20 --fases 0,1
    python pipeline_pedidos.py <tenant> --limite 20      # smoke test (~1 min)
    python pipeline_pedidos.py --inicio 2026-08-21 --fim 2026-09-20
    python pipeline_pedidos.py --fases 4 --refazer-transform   # offline
    python pipeline_pedidos.py --so-verificar
    python pipeline_pedidos.py --regressao
    python pipeline_pedidos.py --incremental             # janela deslizante 45d
    python pipeline_pedidos.py <tenant> --sondar-filtro  # 4 req: dataAtualizacao

    Conexao em data/db.json (host/porta/usuario/senha/banco) ou nas variaveis
    TINY_DB_*. A senha nunca aparece no banner nem no log.

Codigos de saida:
    0 ok | 1 token invalido na largada | 2 token expirou no meio (retomavel)
    3 outro processo do mesmo tenant rodando | 4 verificacao reprovou
    5 o perfil de regras mudou e faltou --aplicar-regras (nada foi reprocessado)

Regras de negocio: perfil em config/regras.padrao.json + regras.json da empresa
(TINY_APPS_CONFIG ou data/). Mudou o perfil? A fase 4 para e mostra o que mudou;
rode com --aplicar-regras para reprocessar o warehouse offline.
"""

import argparse
import io
import sys
from datetime import date, datetime, timedelta

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                              line_buffering=True)

import _bootstrap  # noqa: E402,F401
import etl_catalogo  # noqa: E402
import etl_db  # noqa: E402
import etl_esquema  # noqa: E402
import etl_extract  # noqa: E402
import etl_load  # noqa: E402
import etl_regras as regras  # noqa: E402
import etl_sonda  # noqa: E402
import etl_transform  # noqa: E402
import etl_verificar  # noqa: E402
from core.auth_tiny import validate_tenants  # noqa: E402
from core.tiny_client import TinyClient  # noqa: E402

TENANTS_PADRAO = regras.TENANTS_PADRAO
NOMES = regras.NOMES

SAIDA_OK = 0
SAIDA_TOKEN_LARGADA = 1
SAIDA_TOKEN_MEIO = 2
SAIDA_OCUPADO = 3
SAIDA_VERIFICACAO = 4
SAIDA_REGRAS = 5


def _banner(titulo: str) -> None:
    print("\n{}\n#  {}\n{}".format("#" * 78, titulo, "#" * 78))


def _abrir_run(conn, tenant, args, fases, inicio, fim) -> int:
    with etl_db.transacao(conn):
        etl_db.executar(
            conn,
            "INSERT INTO etl_run (tenant, comando, fases, janela_ini, "
            "janela_fim, inicio) VALUES (%s,%s,%s,%s,%s,%s)",
            (tenant, " ".join(sys.argv)[:4000],
             ",".join(str(f) for f in sorted(fases)), inicio, fim,
             datetime.now()))
    return etl_db.um_valor(conn, "SELECT LAST_INSERT_ID()")


def _fechar_run(conn, run_id, status, **contas) -> None:
    if not run_id:
        return
    try:
        with etl_db.transacao(conn):
            etl_db.executar(
                conn,
                "UPDATE etl_run SET fim=%s, status=%s, pedidos_listados=%s, "
                "pedidos_detalhados=%s, pedidos_carregados=%s, erros=%s, "
                "mensagem=%s WHERE id=%s",
                (datetime.now(), status, contas.get("listados", 0),
                 contas.get("detalhados", 0), contas.get("carregados", 0),
                 contas.get("erros", 0), (contas.get("mensagem") or None)[:500]
                 if contas.get("mensagem") else None, run_id))
    except Exception:
        pass


def _processar_tenant(args, tenant, fases, inicio, fim, cfg) -> dict:
    """Roda as fases pedidas para um tenant. Devolve as contagens."""
    conn = etl_db.conectar(cfg)
    ref = {"conn": conn, "cfg": cfg}
    contas = {"listados": 0, "detalhados": 0, "carregados": 0, "erros": 0}
    run_id = None

    if not etl_db.travar_tenant(conn, tenant):
        print("  [{}] [ERRO] outro processo ja esta rodando este tenant.".format(
            tenant))
        print("         Dois processos = dois rate limiters = 4 req/s na mesma "
              "conta = 429 em cadeia.")
        conn.close()
        raise SystemExit(SAIDA_OCUPADO)

    online = bool(fases & {1, 2, 3})
    client = TinyClient(tenant) if (online or args.sondar_filtro
                                    or args.sondar_schema) else None
    try:
        run_id = _abrir_run(conn, tenant, args, fases, inicio, fim)

        if args.sondar_schema:
            etl_sonda.sondar_schema(client, tenant, inicio, fim,
                                    quantos=args.sondar_quantos)
            _fechar_run(conn, run_id, "ok", **contas)
            return contas
        if args.sondar_filtro:
            etl_sonda.sondar_filtro(client, conn, tenant)
            _fechar_run(conn, run_id, "ok", **contas)
            return contas

        if 1 in fases:
            print("\n  -- fase 1: listagem da janela (sem filtro nenhum)")
            contas["listados"] = etl_extract.fase1_listar(
                client, ref, tenant, inicio, fim, max_janela=args.max_janela,
                refazer=args.refazer_lista or args.incremental)
        if 2 in fases:
            print("\n  -- fase 2: detalhe (1 request por pedido)")
            ok, err = etl_extract.fase2_detalhar(
                client, ref, tenant, lote=args.lote, workers=args.workers,
                limite=args.limite, max_tentativas=args.max_tentativas)
            contas["detalhados"], contas["erros"] = ok, err
        if 3 in fases and not args.sem_catalogo:
            print("\n  -- fase 3: catalogo dos produtos vendidos")
            etl_catalogo.fase3_catalogo(client, ref, tenant,
                                        workers=args.workers, lote=args.lote)
        if 4 in fases:
            print("\n  -- fase 4: transform + load (offline)")
            versao = regras.versao_regras(ref["conn"], tenant,
                                          aplicar=args.aplicar_regras)
            print("  [{}] regras: versao {} ({})".format(
                tenant, versao, regras.origem_perfil()))
            carregados, ruins = etl_load.fase4_carregar(
                ref, tenant, versao, lote=args.lote_load,
                refazer=args.refazer_transform, limite=args.limite)
            contas["carregados"] = carregados
            contas["erros"] += ruins

        _fechar_run(ref["conn"], run_id, "ok", **contas)
        return contas

    except regras.RegrasMudaram as e:
        _fechar_run(ref["conn"], run_id, "erro",
                    mensagem=str(e)[:500], **contas)
        print("\n  [{}] [PARADO] {}".format(tenant, e))
        raise SystemExit(SAIDA_REGRAS)
    except etl_extract.ParadaToken:
        _fechar_run(ref["conn"], run_id, "parado_token", mensagem=
                    "refresh token expirou no meio da coleta", **contas)
        print("\n  [{}] [ERRO] o refresh token expirou no meio da coleta.".format(
            tenant))
        print("         Tudo o que foi baixado ja esta COMMITADO no banco.")
        print("         Para retomar, sem repetir um unico request:")
        print("           .venv\\Scripts\\python.exe src\\apps\\auth\\_reauth.py {}"
              .format(tenant))
        print("           .venv\\Scripts\\python.exe src\\apps\\pipeline_etl\\"
              "pipeline_pedidos.py {} --inicio {} --fim {} --fases 2,3,4,5"
              .format(tenant, inicio, fim))
        raise SystemExit(SAIDA_TOKEN_MEIO)
    except KeyboardInterrupt:
        _fechar_run(ref["conn"], run_id, "interrompido",
                    mensagem="Ctrl+C", **contas)
        print("\n  [{}] interrompido. O que ja foi baixado esta no banco; "
              "rode de novo para continuar.".format(tenant))
        raise SystemExit(SAIDA_TOKEN_MEIO)
    except SystemExit:
        raise
    except Exception as e:
        _fechar_run(ref["conn"], run_id, "erro", mensagem=str(e), **contas)
        raise
    finally:
        try:
            ref["conn"].close()   # solta o GET_LOCK
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Pipeline ETL de pedidos do Tiny para MySQL local.")
    ap.add_argument("tenants", nargs="*", default=None,
                    help="tenant_ids de data/tenants.json (padrao: todos)")
    ap.add_argument("--inicio", type=regras.parse_data, default=None,
                    help="YYYY-MM-DD ou DD/MM/YYYY (padrao: hoje - 30 dias)")
    ap.add_argument("--fim", type=regras.parse_data, default=None,
                    help="YYYY-MM-DD ou DD/MM/YYYY (padrao: hoje)")
    ap.add_argument("--fases", default="0,1,2,3,4,5",
                    help="0=esquema 1=lista 2=detalhe 3=catalogo 4=load 5=verifica")
    ap.add_argument("--limite", type=int, default=None,
                    help="limita as fases 2 e 4 a N pedidos (smoke test)")
    ap.add_argument("--lote", type=int, default=200,
                    help="pedidos por commit na fase 2 (default 200)")
    ap.add_argument("--lote-load", type=int, default=200,
                    help="pedidos por transacao na fase 4 (default 200)")
    ap.add_argument("--workers", type=int, default=4,
                    help="threads; subir NAO acelera (o limiter e o gargalo)")
    ap.add_argument("--max-tentativas", type=int, default=5,
                    help="tentativas por pedido antes de marcar estado='erro'")
    ap.add_argument("--max-janela", type=int, default=5000,
                    help="acima disso a janela mensal e subdividida")
    ap.add_argument("--sem-catalogo", action="store_true",
                    help="pula a fase 3")
    ap.add_argument("--incremental", action="store_true",
                    help="ignora --inicio/--fim e usa a janela de --dias-retro")
    ap.add_argument("--dias-retro", type=int, default=45,
                    help="tamanho da janela deslizante do incremental")
    ap.add_argument("--refazer-lista", action="store_true",
                    help="ignora etl_janela.status='ok' e re-pagina")
    ap.add_argument("--refazer-transform", action="store_true",
                    help="fase 4 tambem sobre o que ja esta 'transformado'")
    ap.add_argument("--aplicar-regras", action="store_true",
                    help="aceita a mudanca do perfil de regras e reprocessa "
                         "o warehouse offline (sem ela, a fase 4 para)")
    ap.add_argument("--sondar-schema", action="store_true",
                    help="grava um pedido cru em data/ e relata (4 requests)")
    ap.add_argument("--sondar-quantos", type=int, default=3,
                    help="pedidos na amostra da sonda de schema")
    ap.add_argument("--sondar-filtro", action="store_true",
                    help="testa o filtro dataAtualizacao (4 requests)")
    ap.add_argument("--so-verificar", action="store_true",
                    help="atalho para --fases 5")
    ap.add_argument("--regressao", action="store_true",
                    help="so o teste de regressao de data/verificacao.json e sai")
    ap.add_argument("--forcar-esquema", action="store_true",
                    help="reaplica os sql/ mesmo com o sha256 igual")
    ap.add_argument("--host", default=None)
    ap.add_argument("--porta", type=int, default=None)
    ap.add_argument("--usuario", default=None)
    ap.add_argument("--banco", default=None)
    args = ap.parse_args()

    tenants = args.tenants or TENANTS_PADRAO
    desconhecidos = [t for t in tenants if t not in NOMES]
    if desconhecidos:
        print("  [ERRO] tenant desconhecido: {}".format(", ".join(desconhecidos)))
        print("         validos: {}".format(", ".join(NOMES)))
        return 1

    hoje = date.today()
    if args.incremental:
        inicio, fim = hoje - timedelta(days=args.dias_retro), hoje
    else:
        inicio = args.inicio or (hoje - timedelta(days=30))
        fim = args.fim or hoje
    if inicio > fim:
        print("  [ERRO] --inicio depois de --fim.")
        return 1

    if args.so_verificar:
        fases = {5}
    elif args.regressao:
        fases = set()
    else:
        fases = {int(x) for x in args.fases.replace(" ", "").split(",") if x}
    if args.sondar_schema or args.sondar_filtro:
        fases = fases & {0}

    cfg = etl_db.carregar_config({"host": args.host, "porta": args.porta,
                                  "usuario": args.usuario, "banco": args.banco})

    print("\n  Pipeline ETL de pedidos (Tiny API v3 -> MySQL)")
    print("  Banco    : {}".format(etl_db.descrever(cfg)))
    print("  Empresas : {}".format(", ".join(NOMES.get(t, t) for t in tenants)))
    print("  Periodo  : {} a {}{}".format(
        inicio, fim, "  (incremental)" if args.incremental else ""))
    print("  Fases    : {}".format(
        ",".join(str(f) for f in sorted(fases)) if fases else "(nenhuma)"))
    if args.limite:
        print("  Limite   : {} pedidos (smoke test)".format(args.limite))

    # Fase 0 e sempre a primeira: sem esquema nao ha onde gravar.
    conn0 = None
    if 0 in fases or args.sondar_filtro or fases & {1, 2, 3, 4, 5} or args.regressao:
        if 0 in fases:
            print("\n  -- fase 0: esquema + seed")
            conn0 = etl_esquema.fase0(cfg, tenants, forcar=args.forcar_esquema)
        else:
            conn0 = etl_db.conectar(cfg)

    if args.regressao:
        linhas = etl_verificar.regressao(
            conn0, tenants, etl_verificar.carregar_verificacao().get("regressao"))
        etl_verificar.imprimir(linhas)
        caminho = etl_verificar.salvar_csv(linhas, tenants)
        print("\n  Relatorio: {}".format(caminho))
        reprovou = any(l["status"] == "FALHOU" for l in linhas)
        conn0.close()
        return SAIDA_VERIFICACAO if reprovou else SAIDA_OK

    online = bool(fases & {1, 2, 3}) or args.sondar_schema or args.sondar_filtro
    if online:
        if not validate_tenants(tenants, interactive=False):
            print("\n  Autenticacao invalida. Rode:")
            for t in tenants:
                print("    .venv\\Scripts\\python.exe src\\apps\\auth\\_reauth.py {}"
                      .format(t))
            if conn0:
                conn0.close()
            return SAIDA_TOKEN_LARGADA
    elif fases:
        print("  (fases offline -- nao valida token)")

    if conn0:
        conn0.close()

    total = {"listados": 0, "detalhados": 0, "carregados": 0, "erros": 0}
    if fases - {0, 5} or args.sondar_schema or args.sondar_filtro:
        for tenant in tenants:
            _banner("{} ({})".format(NOMES.get(tenant, tenant).upper(), tenant))
            contas = _processar_tenant(args, tenant, fases, inicio, fim, cfg)
            for k in total:
                total[k] += contas.get(k, 0)

    saida = SAIDA_OK
    if 5 in fases:
        _banner("VERIFICACAO")
        conn = etl_db.conectar(cfg)
        try:
            if not etl_verificar.fase5_verificar(conn, tenants):
                saida = SAIDA_VERIFICACAO
        finally:
            conn.close()

    if fases - {0}:
        print("\n  Total: {} listados, {} detalhados, {} carregados, "
              "{} erros.".format(total["listados"], total["detalhados"],
                                 total["carregados"], total["erros"]))
    print("\n=== Concluido ===")
    return saida


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except KeyboardInterrupt:
        print("\n  Interrompido.")
        sys.exit(SAIDA_TOKEN_MEIO)
