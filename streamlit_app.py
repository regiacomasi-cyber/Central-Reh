# -*- coding: utf-8 -*-
"""Analisador REH — versão web (Streamlit). Mesmas contas do programa desktop.
Rodar local: streamlit run streamlit_app.py
"""
import io
import os
import tempfile

import pandas as pd
import streamlit as st

import app as motor

st.set_page_config(page_title="Analisador REH", layout="wide")
st.title("Analisador REH")

if "dados" not in st.session_state:
    st.session_state.dados = []
    st.session_state.resultado = []
    st.session_state.total = 0
    st.session_state.conc = {}
    st.session_state.saude = {}

abas = st.tabs(["1 - Vendas", "2 - Curva ABC", "3 - Preços", "4 - Ads",
                "5 - Oportunidades", "6 - Concorrência", "7 - Saúde"])

with abas[0]:
    st.subheader("Mandar planilha de vendas (Excel ou CSV)")
    arq = st.file_uploader("Relatório da Shopee (Product_Sales ou Order)", type=["xlsx", "csv"])
    c1, c2, c3 = st.columns(3)
    de = c1.text_input("De (DD/MM/AAAA)", "")
    ate = c2.text_input("Até (DD/MM/AAAA)", "")
    if c3.button("Carregar exemplo"):
        st.session_state.dados = motor.exemplo()
        st.session_state.resultado = []
        st.success(f"{len(st.session_state.dados)} produtos de exemplo.")
    if arq and st.button("Importar arquivo"):
        suf = ".xlsx" if arq.name.lower().endswith("xlsx") else ".csv"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suf) as tmp:
            tmp.write(arq.read())
            caminho = tmp.name
        try:
            if suf == ".xlsx":
                from datetime import datetime as _dt

                def _db(s):
                    try:
                        return _dt.strptime(s.strip(), "%d/%m/%Y").date()
                    except (ValueError, AttributeError):
                        return None

                dados = motor.ler_excel(caminho, _db(de), _db(ate))
            else:
                dados = motor.ler_csv(caminho)
            st.session_state.dados = dados
            st.session_state.resultado = []
            st.session_state.saude = dict(motor.ULTIMOS_PEDIDOS)
            st.success(f"{len(dados)} produtos importados!")
        except Exception as e:
            st.error(str(e))
        finally:
            os.unlink(caminho)
    if st.session_state.dados:
        st.info(f"{len(st.session_state.dados)} produtos carregados.")

with abas[1]:
    st.subheader("Curva ABC")
    if st.session_state.dados and st.button("Gerar análise"):
        res, tot = motor.classificar(st.session_state.dados)
        st.session_state.resultado = res
        st.session_state.total = tot
    res = st.session_state.resultado
    if res:
        df = pd.DataFrame([{"Produto": x["produto"], "Qtd": x["qtd"],
                            "Faturamento": x["fat"], "Classe": x["classe"],
                            "%": x["part"]} for x in res])
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Faturamento", f"R$ {st.session_state.total:,.2f}".replace(",", "."))
        for cl, col in (("A", c2), ("B", c3), ("C", c4)):
            g = [x for x in res if x["classe"] == cl]
            col.metric(f"Classe {cl}", f"{len(g)} itens",
                       f"R$ {sum(x['fat'] for x in g):,.0f}".replace(",", "."))
        st.bar_chart(df.set_index("Produto")["Faturamento"].head(15))
        filtro = st.selectbox("Ver curva", ["Todas", "A", "B", "C"])
        if filtro != "Todas":
            df = df[df["Classe"] == filtro]
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button("Baixar CSV", csv, "curva_abc.csv", "text/csv")
        buf = io.BytesIO()
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resumo"
        ws.append(["Curva", "Qtd", "Faturamento"])
        for cl in "ABC":
            g = [x for x in res if x["classe"] == cl]
            ws.append([cl, len(g), round(sum(x["fat"] for x in g), 2)])
        for cl in "ABC":
            wa = wb.create_sheet(f"Curva {cl}")
            wa.append(["Produto", "Qtd", "Faturamento", "%", "% acum"])
            for x in [r for r in res if r["classe"] == cl]:
                wa.append([x["produto"], x["qtd"], x["fat"], x["part"], x["acum"]])
        wb.save(buf)
        st.download_button("Baixar Excel A/B/C", buf.getvalue(),
                           "curva_abc.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with abas[2]:
    st.subheader("Precificação (taxas out/2026)")
    res = st.session_state.resultado or st.session_state.dados
    if not res:
        st.warning("Importe primeiro.")
    else:
        plat = st.selectbox("Plataforma", ["Shopee", "ML Clássico", "ML Premium"])
        c1, c2, c3, c4 = st.columns(4)
        ads = c1.number_input("Reserva ads %", value=10.0)
        luc = c2.number_input("Lucro %", value=30.0)
        imp = c3.number_input("Imposto %", value=4.0)
        dev = c4.number_input("Perdas %", value=3.0)
        df = pd.DataFrame([{"Produto": x["produto"],
                            "Preço médio": round(x["fat"]/x["qtd"], 2) if x["qtd"] else 0,
                            "Custo": x.get("custo", 0) or 0.0} for x in res])
        edit = st.data_editor(df, use_container_width=True, num_rows="fixed",
                              column_config={"Custo": st.column_config.NumberColumn("Custo (edite)")})
        if st.button("Calcular"):
            for i, x in enumerate(res):
                x["custo"] = float(edit.iloc[i]["Custo"])
            linhas = []
            for x in res:
                pm = x["fat"]/x["qtd"] if x["qtd"] else 0
                cu = x.get("custo", 0) or 0
                if plat == "Shopee":
                    tx = motor.taxa_shopee(pm)
                    ideal = motor.preco_ideal_shopee(cu, ads/100, luc/100, imp=imp/100, dev=dev/100)
                else:
                    com = 0.17 if "Premium" in plat else 0.12
                    tx = motor.taxa_ml(pm, "Premium" in plat)
                    ideal = motor.preco_ideal_ml(cu, ads/100, luc/100, com, imp=imp/100, dev=dev/100)
                margem = ((pm-cu-tx)/pm - imp/100 - dev/100)*100 if pm else 0
                linhas.append({"Produto": x["produto"][:50], "Preço": pm, "Custo": cu,
                              "Taxas": round(tx, 2), "Lucro %": round(margem, 1),
                              "Preço ideal": ideal})
            st.dataframe(pd.DataFrame(linhas), use_container_width=True)

with abas[3]:
    st.subheader("Estratégia Ads")
    res = st.session_state.resultado
    if not res:
        st.warning("Gere a Curva ABC primeiro.")
    else:
        plat = st.selectbox("Onde", ["Shopee", "Mercado Livre", "Nos dois"], key="ads_plat")
        vt = st.number_input("Verba/dia R$", value=200.0, min_value=0.0)
        lg = st.number_input("Lucro líquido %", value=22.0, min_value=1.0)
        if st.button("Montar estratégia"):
            plats = ["Shopee", "Mercado Livre"] if plat == "Nos dois" else [plat]
            for p in plats:
                st.markdown(f"### {p} — verba R$ {vt/len(plats):.2f}/dia — meta {1/(lg/100):.1f}x")
                for x in [r for r in res if r["classe"] == "A"][:8]:
                    st.write(f"• {x['produto'][:60]} — unitário")
                bs = [r for r in res if r["classe"] == "B"]
                if bs:
                    st.write(f"Testes B ({len(bs)}): " + ", ".join(x["produto"][:30] for x in bs[:5]))

with abas[4]:
    st.subheader("Oportunidades (prioridades)")
    res = st.session_state.resultado
    if not res:
        st.warning("Gere a Curva ABC primeiro.")
    else:
        rows = []
        for x in res:
            if not x.get("custo"):
                rows.append((x["produto"], x["classe"], "Lucro desconhecido", "Informar custo", x["part"]))
            if len(x["produto"]) < 40:
                rows.append((x["produto"], x["classe"], "Título curto", "Estender p/ 80+ letras", x["part"]))
            if x["classe"] == "A" and 0 < x.get("est", 0) < x["qtd"]*0.3:
                rows.append((x["produto"], x["classe"], "Estoque baixo (campeão)", "Repor urgente", x["part"]*2))
            if x["classe"] == "C" and x.get("est", 0) > max(x["qtd"]*2, 1):
                rows.append((x["produto"], x["classe"], "Encalhado", "Kit com campeão", x["part"]))
        rows.sort(key=lambda r: -r[4])
        st.dataframe(pd.DataFrame(rows[:25], columns=["Produto", "CL", "Problema", "Ação", "Impacto"]),
                     use_container_width=True)

with abas[5]:
    st.subheader("Concorrência")
    res = st.session_state.resultado or st.session_state.dados
    if not res:
        st.warning("Importe primeiro.")
    else:
        nomes = [x["produto"] for x in res]
        prod = st.selectbox("Seu produto", nomes)
        key = st.text_input("Chave SerpApi (grátis em serpapi.com)", type="password")
        if st.button("Buscar sozinho"):
            if not key:
                st.warning("Cole a chave SerpApi primeiro.")
            else:
                from urllib.parse import quote_plus
                import json as _js
                termo = " ".join(prod.split()[:8])
                url = (f"https://serpapi.com/search.json?engine=google_shopping&q={quote_plus(termo)}"
                       f"&gl=br&hl=pt-BR&num=10&api_key={quote_plus(key)}")
                try:
                    data = _js.loads(motor.http_get(url, timeout=20))
                    ach = motor.App._parse_shopping(data)
                    st.session_state.conc.setdefault(prod, []).extend(ach)
                    st.success(f"{len(ach)} concorrentes adicionados!")
                except Exception as e:
                    st.error(f"Falhou: {e}")
        conc = st.session_state.conc.get(prod, [])
        if conc:
            dfc = pd.DataFrame([{"Nome": c["nome"][:60], "Preço": c["preco"],
                                "Vend/mês": c["vend"], "Nota": c["nota"]} for c in conc])
            st.dataframe(dfc, use_container_width=True)
            top = sorted(conc, key=lambda c: c["vend"], reverse=True)[:4]
            precos = sorted(c["preco"] for c in top)
            meu = next(x for x in res if x["produto"] == prod)
            pm = meu["fat"]/meu["qtd"] if meu["qtd"] else 0
            med = precos[len(precos)//2]
            st.info(f"Seu preço {pm:.2f} x média top {med:.2f}. "
                    + ("Acima: considere igualar ou kit." if pm > med else "Na faixa ou abaixo: confira a margem."))

with abas[6]:
    st.subheader("Saúde das vendas")
    s = st.session_state.saude
    if not s:
        st.warning("Importe o relatório de PEDIDOS (Order...xlsx).")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Pedidos", s["total"])
        c2.metric("Cancelados", s["cancelados"], f"R$ {s['v_cancelado']:,.0f}".replace(",", "."))
        c3.metric("Em risco", f"R$ {s['v_risco']:,.0f}".replace(",", "."))
        mots = sorted(s["motivos"].items(), key=lambda x: -x[1])[:5]
        st.write("**Top motivos:**")
        for mot, q in mots:
            st.write(f"• {q}x — {mot[:100]}")
