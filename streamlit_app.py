# -*- coding: utf-8 -*-
"""Central REH — versão web. Mesma inteligência do programa desktop."""
import io
import os
import re
import tempfile
from datetime import datetime as _dt

import pandas as pd
import streamlit as st

import app as motor

st.set_page_config(page_title="Central REH", layout="wide")
st.markdown(
    "<div style='background-color:#3D0A1A;padding:18px 22px;border-radius:12px;margin-bottom:6px'>"
    "<div style='font-family:Georgia,serif;font-size:30px;font-weight:bold;color:#FAF3E6'>Central REH</div>"
    "<div style='font-size:14px;color:#E3B341'><i>Análise completa para o seu negócio</i></div>"
    "</div>", unsafe_allow_html=True)

for k, v in {"dados": [], "resultado": [], "total": 0, "conc": {},
             "saude": {}, "hist": [], "p_de": "", "p_ate": "", "ajustes": {}}.items():
    if k not in st.session_state:
        st.session_state[k] = v

PASSOS = ["1 - Vendas", "2 - Curva ABC", "3 - Preços", "4 - Ads",
          "5 - Oportunidades", "6 - Concorrência", "7 - Cancelados", "8 - Mês x Mês"]
DICAS = {
    "1 - Vendas": "Mande 1 ou várias planilhas (segure Ctrl). Cada uma vira um mês no histórico.",
    "2 - Curva ABC": "Verde = campeões (A). Amarelo = podem crescer (B). Cinza = parados (C).",
    "3 - Preços": "Informe o custo de cada produto para lucro real. Taxas oficiais de out/2026.",
    "4 - Ads": "Diga verba e lucro líquido: sai o plano por campanha e por produto.",
    "5 - Oportunidades": "Sua fila de trabalho: prioridades, kits, ruptura, Full e campanhas.",
    "6 - Concorrência": "Busque os 4 mais vendidos e compare preço, vendas e oferta.",
    "7 - Cancelados": "Quem cancela, quanto se perde e o que corrigir (precisa do Order).",
    "8 - Mês x Mês": "Comparativo entre 2 meses + onde trabalhar primeiro.",
}
pag = st.sidebar.radio("Passo a passo", PASSOS)
st.sidebar.caption("Central REH · dados ficam só na sua sessão")
if len(st.session_state.hist) > 1:
    st.sidebar.markdown("**Planilha em análise**")
    pers = [h["periodo"] for h in st.session_state.hist]
    atual = st.session_state.get("planilha_ativa", pers[-1])
    if atual not in pers:
        atual = pers[-1]
    nova = st.sidebar.selectbox("Trocar de mês (vale p/ todas as abas)", pers,
                                index=pers.index(atual), label_visibility="collapsed")
    if nova != st.session_state.get("planilha_ativa"):
        st.session_state.planilha_ativa = nova
        st.rerun()
    st.session_state.planilha_ativa = nova
    h_ativa = next(h for h in st.session_state.hist if h["periodo"] == nova)
    if h_ativa.get("dados"):
        import copy as _copy2
        dados = _copy2.deepcopy(h_ativa["dados"])
        for x in dados:
            aj = st.session_state.ajustes.get(x["produto"])
            if aj:
                x["custo"] = aj.get("custo", x.get("custo", 0) or 0)
                x["est"] = aj.get("est", x.get("est", 0) or 0)
        st.session_state.dados = dados
        st.session_state.saude = _copy2.deepcopy(h_ativa.get("saude", {}))
        res2, tot2 = motor.classificar(dados)
        st.session_state.resultado = res2
        st.session_state.total = tot2
st.info("💡 " + DICAS[pag])


def _db(s):
    try:
        return _dt.strptime(s.strip(), "%d/%m/%Y").date()
    except (ValueError, AttributeError):
        return None


def _periodo_arq(nome):
    m = re.search(r"(\d{8})_(\d{8})", nome)
    if m:
        try:
            return (_dt.strptime(m.group(1), "%Y%m%d").strftime("%d/%m/%Y") + " a " +
                    _dt.strptime(m.group(2), "%Y%m%d").strftime("%d/%m/%Y"))
        except ValueError:
            pass
    return ""


@st.cache_data(show_spinner="Lendo planilha...")
def _parse_file(file_bytes, filename, de, ate):
    suf = ".xlsx" if filename.lower().endswith("xlsx") else ".csv"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suf) as tmp:
        tmp.write(file_bytes)
        caminho = tmp.name
    try:
        if suf == ".xlsx":
            dados = motor.ler_excel(caminho, _db(de), _db(ate))
        else:
            dados = motor.ler_csv(caminho)
        return dados, dict(motor.ULTIMOS_PEDIDOS)
    finally:
        os.unlink(caminho)


def _din(v):
    return f"R$ {v:,.2f}".replace(",", ".")


# ==================== 1 VENDAS ====================
if pag == "1 - Vendas":
    st.header("Mandar planilha")
    arqs = st.file_uploader("Relatório Shopee (Product_Sales ou Order)", type=["xlsx", "csv"],
                            accept_multiple_files=True)
    c1, c2 = st.columns(2)
    st.session_state.p_de = c1.text_input("De (DD/MM/AAAA)", st.session_state.p_de)
    st.session_state.p_ate = c2.text_input("Até (DD/MM/AAAA)", st.session_state.p_ate)
    if st.button("Carregar exemplo", use_container_width=False):
        st.session_state.dados = motor.exemplo()
        st.session_state.resultado = []
        st.success("Exemplo carregado.")
    if arqs and st.button(f"Importar {len(arqs)} arquivo(s)", type="primary"):
        ok = 0
        for up in arqs:
            try:
                dados, saude = _parse_file(up.getvalue(), up.name,
                                           st.session_state.p_de, st.session_state.p_ate)
                per = _periodo_arq(up.name) or \
                    f"{st.session_state.p_de} a {st.session_state.p_ate}".strip(" a") or "sem período"
                st.session_state.dados = dados
                st.session_state.saude = saude
                res, tot = motor.classificar(dados)
                st.session_state.resultado = res
                st.session_state.total = tot
                import copy as _copy
                st.session_state.hist = [h for h in st.session_state.hist if h["periodo"] != per]
                st.session_state.hist.append({
                    "periodo": per, "total": tot,
                    "dados": _copy.deepcopy(dados),
                    "produtos": {x["produto"]: {"qtd": x["qtd"], "fat": x["fat"],
                                               "classe": x["classe"]} for x in res},
                    "saude": _copy.deepcopy(saude)})
                ok += 1
            except Exception as e:
                st.error(f"{up.name}: {e}")
        if ok:
            st.success(f"{ok} importado(s). Análise pronta + histórico com {len(st.session_state.hist)} período(s).")
    if st.session_state.dados:
        st.info(f"{len(st.session_state.dados)} produtos na análise atual.")

# ==================== 2 CURVA ====================
elif pag == "2 - Curva ABC":
    st.header("Curva ABC")
    pers = [h["periodo"] for h in st.session_state.hist]
    if len(pers) > 1:
        ativa = st.selectbox("Planilha em análise", pers, index=len(pers)-1,
                             help="Troque para analisar outra planilha em todas as abas.")
        h_ativa = next(h for h in st.session_state.hist if h["periodo"] == ativa)
        if h_ativa.get("dados"):
            import copy as _copy2
            dados = _copy2.deepcopy(h_ativa["dados"])
            for x in dados:
                aj = st.session_state.ajustes.get(x["produto"])
                if aj:
                    x["custo"] = aj.get("custo", x.get("custo", 0) or 0)
                    x["est"] = aj.get("est", x.get("est", 0) or 0)
            st.session_state.dados = dados
            st.session_state.saude = _copy2.deepcopy(h_ativa.get("saude", {}))
            res2, tot2 = motor.classificar(st.session_state.dados)
            st.session_state.resultado = res2
            st.session_state.total = tot2
    res = st.session_state.resultado
    if not res:
        st.warning("Importe na etapa 1 - Vendas.")
    else:
        tot = st.session_state.total
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Faturamento", _din(tot))
        for cl, col, cor in (("A", c2, "🟢"), ("B", c3, "🟡"), ("C", c4, "⚪")):
            g = [x for x in res if x["classe"] == cl]
            col.metric(f"{cor} Classe {cl}", f"{len(g)} itens",
                       f"R$ {sum(x['fat'] for x in g):,.0f}".replace(",", "."))
        df = pd.DataFrame([{"Produto": x["produto"], "Qtd": x["qtd"],
                            "Faturamento": x["fat"], "Classe": x["classe"],
                            "%": x["part"]} for x in res])
        st.subheader("Top 10 por faturamento")
        st.bar_chart(df.set_index("Produto")["Faturamento"].head(10))
        f = st.selectbox("Ver curva", ["Todas", "A", "B", "C"])
        if f != "Todas":
            df = df[df["Classe"] == f]
        st.dataframe(df, use_container_width=True, height=320)
        d1, d2 = st.columns(2)
        d1.download_button("Baixar CSV", df.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           "curva_abc.csv", "text/csv", use_container_width=True)
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
        d2.download_button("Baixar Excel A/B/C", buf.getvalue(), "curva_abc.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        with st.expander("DRE real (taxas pagas)"):
            linhas = []
            for x in sorted(res, key=lambda d: d["fat"], reverse=True)[:20]:
                cu = x.get("custo", 0) or 0
                tx = x["taxas"] if x.get("taxas", 0) else (
                    motor.taxa_shopee(x["fat"]/x["qtd"])*x["qtd"] if x["qtd"] else 0)
                lucro = x["fat"] - tx - cu*x["qtd"]
                linhas.append({"Produto": x["produto"][:50], "Fatura": round(x["fat"], 2),
                              "Taxas": round(tx, 2), "Lucro": round(lucro, 2),
                              "Margem %": round(lucro/x["fat"]*100, 1) if x["fat"] else 0})
            st.dataframe(pd.DataFrame(linhas), use_container_width=True)
        st.subheader("Comparar meses e evolução")
        hist = st.session_state.hist
        if len(hist) < 2:
            st.info("Importe outro mês para comparar (o histórico guarda cada análise).")
        else:
            pers = [h["periodo"] for h in hist]
            p1 = st.selectbox("De", pers, index=len(pers)-2)
            p2 = st.selectbox("Para", pers, index=len(pers)-1)
            ha = next(h for h in hist if h["periodo"] == p1)
            hb = next(h for h in hist if h["periodo"] == p2)
            pa, pb = ha["produtos"], hb["produtos"]
            var = (hb["total"]-ha["total"])/ha["total"]*100 if ha["total"] else 0
            st.metric("Faturamento", f"R$ {hb['total']:,.0f}".replace(",", "."),
                      f"{var:+.1f}% vs {p1[:10]}")
            topa = max(pa.items(), key=lambda kv: kv[1]["fat"])[0]
            topb = max(pb.items(), key=lambda kv: kv[1]["fat"])[0]
            st.write("**Campeão:** " + ("continua " + topa[:60] if topa == topb
                                        else f"antes {topa[:45]} → agora {topb[:45]}"))
            sub = sorted([p for p in pb if p in pa and pa[p]["classe"] == "C" and pb[p]["classe"] in "AB"])
            cai = sorted([p for p in pb if p in pa and pa[p]["classe"] == "A" and pb[p]["classe"] in "BC"])
            e1, e2 = st.columns(2)
            with e1:
                st.write(f"**Subiram ({len(sub)})**")
                st.dataframe(pd.DataFrame([{"Produto": p[:50], "Foi para": pb[p]["classe"]}
                                           for p in sub[:20]]), use_container_width=True)
            with e2:
                st.write(f"**Caíram ({len(cai)})**")
                st.dataframe(pd.DataFrame([{"Produto": p[:50], "Caiu para": pb[p]["classe"]}
                                           for p in cai[:20]]), use_container_width=True)
            t5 = {}
            for p, v in sorted(pa.items(), key=lambda kv: -kv[1]["fat"])[:5] + \
                    sorted(pb.items(), key=lambda kv: -kv[1]["fat"])[:5]:
                t5.setdefault(p, ["—", "—"])
            for p, v in sorted(pa.items(), key=lambda kv: -kv[1]["fat"])[:5]:
                t5[p][0] = round(v["fat"])
            for p, v in sorted(pb.items(), key=lambda kv: -kv[1]["fat"])[:5]:
                t5[p][1] = round(v["fat"])
            st.write("**Top 5 de cada mês**")
            st.dataframe(pd.DataFrame([{"Produto": k[:55], p1[:16]: v[0], p2[:16]: v[1]}
                                       for k, v in t5.items()]), use_container_width=True)
            st.write("**Evolução do faturamento**")
            st.line_chart(pd.DataFrame({"Faturamento": [h["total"] for h in hist]},
                                       index=[h["periodo"][:10] for h in hist]))

# ==================== 3 PREÇOS ====================
elif pag == "3 - Preços":
    st.header("Precificação (taxas out/2026)")
    res = st.session_state.resultado or st.session_state.dados
    if not res:
        st.warning("Importe primeiro.")
    else:
        plat = st.selectbox("Precificar para", ["Shopee", "ML Clássico", "ML Premium"])
        c1, c2, c3, c4 = st.columns(4)
        ads = c1.number_input("Reserva ads %", value=10.0)
        luc = c2.number_input("Lucro %", value=30.0)
        imp = c3.number_input("Imposto %", value=4.0)
        dev = c4.number_input("Perdas %", value=3.0)
        cpf = st.checkbox("CPF +450 pedidos (Shopee +R$3)")
        peso = st.selectbox("Peso (ML)", ["leve (até 0,5 kg)", "medio (0,5 a 2 kg)"])
        pw = "leve" if peso.startswith("leve") else "medio"
        if plat != "Shopee":
            st.caption(f"Comissão ML automática: {'17%' if 'Premium' in plat else '12%'}")
        df = pd.DataFrame([{"Produto": x["produto"],
                            "Preço médio": round(x["fat"]/x["qtd"], 2) if x["qtd"] else 0,
                            "Custo": x.get("custo", 0) or 0.0,
                            "Estoque": x.get("est", 0) or 0} for x in res])
        edit = st.data_editor(df, use_container_width=True, num_rows="fixed", height=300)
        if st.button("Calcular tudo", type="primary"):
            for i, x in enumerate(res):
                x["custo"] = float(edit.iloc[i]["Custo"])
                x["est"] = int(edit.iloc[i]["Estoque"])
                st.session_state.ajustes[x["produto"]] = {"custo": x["custo"], "est": x["est"]}
            linhas = []
            for x in res:
                pm = x["fat"]/x["qtd"] if x["qtd"] else 0
                cu = x.get("custo", 0) or 0
                if plat == "Shopee":
                    tx, fr = motor.taxa_shopee(pm, cpf, 0), 0.0
                    ideal = motor.preco_ideal_shopee(cu, ads/100, luc/100, cpf, 0, imp/100, dev/100)
                else:
                    com = 0.17 if "Premium" in plat else 0.12
                    tx = motor.taxa_ml(pm, "Premium" in plat, com, pw)
                    fr = tx - pm*com
                    ideal = motor.preco_ideal_ml(cu, ads/100, luc/100, com, pw, imp=imp/100, dev=dev/100)
                margem = ((pm-cu-tx)/pm - imp/100 - dev/100)*100 if pm else 0
                linhas.append({"Produto": x["produto"][:50], "Preço": pm, "Custo": cu,
                              "Taxas": round(tx, 2), "Frete": round(fr, 2),
                              "Lucro %": round(margem, 1), "Preço ideal": ideal})
            st.dataframe(pd.DataFrame(linhas), use_container_width=True, height=350)
            ver = [l for l in linhas if l["Lucro %"] < 8 and l["Custo"] > 0]
            if ver:
                st.error(f"Atenção: {len(ver)} com lucro abaixo de 8% — " +
                         ", ".join(l["Produto"][:30] for l in ver[:3]))
        with st.expander("Shopee x ML (top 15)"):
            rows = []
            for x in sorted(res, key=lambda d: d["fat"], reverse=True)[:15]:
                pm = x["fat"]/x["qtd"] if x["qtd"] else 0
                cu = x.get("custo", 0) or 0
                tsh, tml = motor.taxa_shopee(pm), motor.taxa_ml(pm, False)
                msh = ((pm-cu-tsh)/pm - imp/100 - dev/100)*100 if pm else 0
                mml = ((pm-cu-tml)/pm - imp/100 - dev/100)*100 if pm else 0
                mel = "—" if not cu else ("Empate" if abs(msh-mml) < 2 else ("Shopee" if msh > mml else "ML"))
                rows.append({"Produto": x["produto"][:45], "Preço": pm,
                             "SH líquido": round(pm-tsh, 2), "SH margem %": round(msh, 1),
                             "ML líquido": round(pm-tml, 2), "ML margem %": round(mml, 1),
                             "Melhor": mel})
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

# ==================== 4 ADS ====================
elif pag == "4 - Ads":
    st.header("Estratégia Ads")
    res = st.session_state.resultado
    if not res:
        st.warning("Gere a Curva ABC primeiro.")
    else:
        c1, c2, c3 = st.columns(3)
        plat = c1.selectbox("Onde", ["Shopee", "Mercado Livre", "Nos dois"])
        mlt = c2.selectbox("Anúncio ML", ["Clássico (taxa menor)", "Premium (12x sem juros)"])
        vt = c3.number_input("Verba/dia R$", value=200.0, min_value=0.0)
        lg = st.number_input("Lucro líquido %", value=22.0, min_value=1.0)
        pa = st.number_input("% A", value=60.0)
        if st.button("Montar estratégia", type="primary"):
            plats = ["Shopee", "Mercado Livre"] if plat == "Nos dois" else [plat]
            for p in plats:
                vb = vt/len(plats)
                st.markdown(f"### {p} — {motor.dinheiro(vb)}/dia — meta {1/(lg/100):.1f}x")
                st.write("**Campanha 1 — Campeões (unitário)**")
                for x in [r for r in res if r["classe"] == "A"][:8]:
                    pm = x["fat"]/x["qtd"] if x["qtd"] else 0
                    cu = x.get("custo", 0) or 0
                    fee = motor.taxa_shopee(pm) if p == "Shopee" else motor.taxa_ml(pm, "Premium" in mlt)
                    if cu and pm:
                        M0 = (pm-cu-fee)/pm - 0.04 - 0.03
                        if M0-lg/100 > 0.02:
                            st.write(f"• {x['produto'][:55]} — meta {1/(M0-lg/100):.1f}x (eq {1/M0:.1f}x)")
                        else:
                            st.write(f"• {x['produto'][:55]} — **inviável p/ {lg:.0f}%**")
                    else:
                        st.write(f"• {x['produto'][:55]} — meta {1/(lg/100):.1f}x (sem custo)")
                bs = [r for r in res if r["classe"] == "B"]
                if bs:
                    st.write(f"**Campanha 2 — Testes B ({len(bs)} produtos)**")
                st.write("**Campanha 3 — Kits** (C só em kit com campeão)")
            st.info("Rotina: 2ª pausar sem venda (30+ cliques) · 4ª dobrar verba de quem vende · 6ª trocar criativo.")

# ==================== 5 OPORTUNIDADES ====================
elif pag == "5 - Oportunidades":
    st.header("Oportunidades")
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
        st.subheader("Plano do especialista: onde investir e onde não investir")
        st.caption("Com base nas vendas dos últimos meses + margem de cada produto.")
        hist = st.session_state.hist
        plano = []
        movs = {}
        if len(hist) >= 2:
            pa = hist[-2]["produtos"]
            pb = hist[-1]["produtos"]
            for p in pb:
                if p in pa:
                    movs[p] = pb[p]["fat"]-pa[p]["fat"]
        for x in res:
            pm = x["fat"]/x["qtd"] if x["qtd"] else 0
            cu = x.get("custo", 0) or 0
            mg = ((pm-cu)/pm if pm and cu else 0)*100
            d = movs.get(x["produto"], 0)
            if x["classe"] == "A" and (not cu or mg >= 15) and d >= 0:
                plano.append((x["produto"][:50], "INVESTIR MAIS", "Campeão com margem e demanda. Dobrar verba aqui.",
                              f"{mg:.0f}%" if cu else "s/ custo"))
            elif x["classe"] == "A" and d < 0:
                plano.append((x["produto"][:50], "INVESTIR (com ajuste)", "Campeão caindo: revisar preço/foto e manter.",
                              f"{mg:.0f}%" if cu else "s/ custo"))
            elif x["classe"] == "B" and d > 0:
                plano.append((x["produto"][:50], "TESTAR", "Subindo: verba de teste para virar campeão.",
                              f"{mg:.0f}%" if cu else "s/ custo"))
            elif x["classe"] == "C" or (cu and mg < 8):
                plano.append((x["produto"][:50], "NÃO INVESTIR", "Só em kit com campeão. Verba solo é prejuízo.",
                              f"{mg:.0f}%" if cu else "s/ custo"))
            else:
                plano.append((x["produto"][:50], "MANTER", "Segue o plano atual, sem mexer.",
                              f"{mg:.0f}%" if cu else "s/ custo"))
        dfp = pd.DataFrame(plano, columns=["Produto", "Veredito", "Por quê", "Margem"])
        cores = {"INVESTIR MAIS": "background-color:#DFF5DF;color:#1E7E34;font-weight:bold",
                 "INVESTIR (com ajuste)": "background-color:#DFF5DF",
                 "TESTAR": "background-color:#FFF3D6",
                 "NÃO INVESTIR": "background-color:#FFCDD2;color:#B71C1C;font-weight:bold",
                 "MANTER": ""}
        st.dataframe(dfp.style.map(
            lambda v: cores.get(v, ""), subset=["Veredito"]), use_container_width=True)
        e1, e2 = st.columns(2)
        with e1:
            st.subheader("Montar kit")
            la = [x["produto"] for x in res if x["classe"] == "A"][:15]
            lc = [x["produto"] for x in res if x["classe"] == "C"]
            if la and lc:
                a = st.selectbox("Campeão (A)", la)
                c = st.selectbox("Parado (C)", lc)
                desc = st.number_input("Desconto %", value=12.0)
                xa = next(x for x in res if x["produto"] == a)
                xc = next(x for x in res if x["produto"] == c)
                pa = xa["fat"]/xa["qtd"] if xa["qtd"] else 0
                pc = xc["fat"]/xc["qtd"] if xc["qtd"] else 0
                kit = round((pa+pc)*(1-desc/100), 2)
                fee = motor.taxa_shopee(kit)
                mg = ((kit-(xa.get("custo", 0) or 0)-(xc.get("custo", 0) or 0)-fee)/kit-0.04-0.03)*100 if kit else 0
                st.success(f"**Kit {a[:40]} + {c[:35]}**\n\nDe {motor.dinheiro(pa+pc)} por **{motor.dinheiro(kit)}** — lucro {mg:.1f}%")
                st.write("ÓTIMO." if mg >= 15 else ("OK, teste com pouca verba." if mg >= 8 else "RUIM: diminua o desconto."))
        with e2:
            st.subheader("Ruptura de estoque")
            dias = 30
            rr = []
            for x in res:
                if not x.get("est"):
                    continue
                daily = x["qtd"]/dias if dias else 0
                if daily <= 0:
                    continue
                cob = x["est"]/daily
                rr.append({"Produto": x["produto"][:45], "Venda/dia": round(daily, 1),
                           "Estoque": x["est"],
                           "Cobertura": f"{cob:.0f} dias ({'CRÍTICO' if cob < 7 else ('ATENÇÃO' if cob < 14 else 'ok')})",
                           "Repor p/ 30d": max(0, int(30*daily-x["est"]+0.5))})
            if rr:
                st.dataframe(pd.DataFrame(sorted(rr, key=lambda r: -r["Repor p/ 30d"])),
                             use_container_width=True)
            else:
                st.info("Informe estoques na aba Preços para ver a cobertura.")
        st.subheader("Enviar para o Full (curva A)")
        from app import proxima_campanha as _pc
        nome_camp, dias_camp, extra = _pc()
        st.write(f"Campanha chegando: **{nome_camp}** em {dias_camp} dias → reserva extra de {extra*100:.0f}%.")
        st.subheader("Enviar para o Full (curva A)")
        from app import proxima_campanha as _pc
        nome_camp, dias_camp, extra = _pc()
        st.write(f"Campanha chegando: **{nome_camp}** em {dias_camp} dias → reserva extra de {extra*100:.0f}%.")
        fl = {}
        for x in [r for r in res if r.get("classe") == "A"]:
            daily = x["qtd"]/30
            if daily <= 0:
                continue
            est = x.get("est", 0) or 0
            row = fl.setdefault(x["produto"][:50], {"Produto": x["produto"][:50],
                                                    "Venda/dia": round(daily, 1), "Estoque": est})
            for d in (15, 30):
                q = max(0, int(d*daily*(1+extra)-est+0.5))
                row[f"Enviar {d}d"] = q
                row[f"Atende {d}d"] = f"{(est+q)/daily:.0f} dias"
        if fl:
            dff = pd.DataFrame(sorted(fl.values(), key=lambda r: -r["Enviar 30d"]))
            t15, t30 = int(dff["Enviar 15d"].sum()), int(dff["Enviar 30d"].sum())
            k1, k2, k3 = st.columns(3)
            k1.metric("Enviar (15 dias)", f"{t15} un")
            k2.metric("Enviar (30 dias)", f"{t30} un")
            k3.metric("Campeões", len(dff))
            def _cor_full(v):
                if isinstance(v, str) and "dias" in v:
                    try:
                        n = int(v.split()[0])
                    except ValueError:
                        return ""
                    if n < 14:
                        return "background-color:#FFCDD2"
                    if n < 30:
                        return "background-color:#FFF3D6"
                    return "background-color:#DFF5DF"
                return ""

            st.dataframe(dff.style.map(_cor_full, subset=["Atende 15d", "Atende 30d"]),
                         use_container_width=True)
            st.caption("Vermelho = acaba em <14 dias (urgente) · Amarelo = <30 dias · Verde = coberto.")
        st.subheader("Campanhas")
        st.write(f"Próxima: **{nome_camp}** em {dias_camp} dias. "
                 "30 dias antes: verba e campeões · 15 dias: estoque e fotos · "
                 "7 dias: cupons, kits e preços · No dia: monitore de hora em hora.")

# ==================== 6 CONCORRÊNCIA ====================
elif pag == "6 - Concorrência":
    st.header("Concorrência")
    res = st.session_state.resultado or st.session_state.dados
    if not res:
        st.warning("Importe primeiro.")
    else:
        nomes = [x["produto"] for x in res]
        prod = st.selectbox("Seu produto", nomes)
        try:
            qk = st.query_params.get("key", "")
            if isinstance(qk, list):
                qk = qk[0] if qk else ""
        except Exception:
            qk = ""
        key = st.text_input("Chave SerpApi (grátis em serpapi.com)", value=qk, type="password")
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
                    st.success(f"{len(ach)} concorrentes! Confira as vendas e ajuste.")
                except Exception as e:
                    st.error(f"Falhou: {e}")
        conc = st.session_state.conc.get(prod, [])
        if conc:
            dfc = pd.DataFrame([{"Nome": c["nome"][:60], "Preço": c["preco"],
                                "Vend/mês": c["vend"], "Nota": c["nota"],
                                "Link": c.get("link", "")} for c in conc])
            st.dataframe(dfc, use_container_width=True)
            for c in sorted(conc, key=lambda x: -x["vend"])[:4]:
                if c.get("link"):
                    st.link_button(f"Abrir: {c['nome'][:45]}", c["link"])
            top = sorted(conc, key=lambda x: -x["vend"])[:4]
            precos = sorted(c["preco"] for c in top)
            meu = next(x for x in res if x["produto"] == prod)
            pm = meu["fat"]/meu["qtd"] if meu["qtd"] else 0
            med = precos[len(precos)//2]
            st.info(f"Seu preço {pm:.2f} x média top {med:.2f}. " +
                    ("Acima: iguale ou monte kit." if pm > med else "Na faixa ou abaixo: confira a margem."))

# ==================== 7 CANCELADOS ====================
elif pag == "7 - Cancelados":
    st.header("Cancelados")
    s = st.session_state.saude
    if not s or not s.get("por_produto"):
        st.warning("Importe o relatório de PEDIDOS (Order...xlsx).")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Cancelados", s["cancelados"])
        c2.metric("R$ perdido", f"R$ {s['v_cancelado']:,.0f}".replace(",", "."))
        mots = sorted(s.get("motivos", {}).items(), key=lambda kv: -kv[1])[:3]
        c3.metric("Top motivo", f"{mots[0][1]}x" if mots else "—", (mots[0][0][:25] if mots else ""))
        pp = s["por_produto"]
        res = st.session_state.resultado or []
        classes = {x["produto"]: x.get("classe", "") for x in res}
        rows = []
        for prod, v in pp.items():
            mot = max(v["motivos"].items(), key=lambda kv: kv[1])[0] if v["motivos"] else "—"
            rows.append({"Produto": prod[:55], "CL": classes.get(prod, "—"), "Qtd": v["qtd"],
                         "R$ perdido": round(v["valor"], 2), "Motivo": mot[:60],
                         "Solução": motor.solucao_motivo(mot)[:70]})
        rows.sort(key=lambda r: -r["Qtd"])
        st.dataframe(pd.DataFrame(rows[:30]), use_container_width=True)

elif pag == "8 - Mês x Mês":
    st.header("Comparativo entre 2 meses + onde trabalhar primeiro")
    hist = st.session_state.hist
    if len(hist) < 2:
        st.warning("Importe 2 meses para comparar (o histórico guarda cada análise).")
    else:
        pers = [h["periodo"] for h in hist]
        c1, c2 = st.columns(2)
        p1 = c1.selectbox("De", pers, index=len(pers)-2)
        p2 = c2.selectbox("Para", pers, index=len(pers)-1)
        ha = next(h for h in hist if h["periodo"] == p1)
        hb = next(h for h in hist if h["periodo"] == p2)
        pa, pb = ha["produtos"], hb["produtos"]
        ta, tb = ha["total"], hb["total"]

        def _luc(hh):
            tot = 0.0
            for x in hh["produtos"].values():
                pm = x["fat"]/x["qtd"] if x["qtd"] else 0
                if pm:
                    tot += x["fat"] - motor.taxa_shopee(pm)*x["qtd"]
            return round(tot, 2)

        la, lb = _luc(ha), _luc(hb)
        qa = {c: sum(1 for x in pa.values() if x["classe"] == c) for c in "ABC"}
        qb = {c: sum(1 for x in pb.values() if x["classe"] == c) for c in "ABC"}
        sa, sb = ha.get("saude", {}) or {}, hb.get("saude", {}) or {}
        pts = {"a": 0, "b": 0}
        rows = []
        for met, va, vb, ganha_a in [
            ("Faturamento", f"R$ {ta:,.0f}".replace(",", "."), f"R$ {tb:,.0f}".replace(",", "."), ta >= tb),
            ("Lucro estimado", f"R$ {la:,.0f}".replace(",", "."), f"R$ {lb:,.0f}".replace(",", "."), la >= lb),
            ("Qtd curva A", qa["A"], qb["A"], qa["A"] >= qb["A"]),
            ("Qtd curva C (menos é melhor)", qa["C"], qb["C"], qa["C"] <= qb["C"]),
            ("Cancelados", sa.get("cancelados", "—"), sb.get("cancelados", "—"),
             (sa.get("cancelados", 0) <= sb.get("cancelados", 0)) if (sa.get("total") and sb.get("total")) else None),
        ]:
            if ganha_a is None:
                rows.append({"MEDIDA": met, p1[:18]: va, p2[:18]: vb})
            elif ganha_a:
                pts["a"] += 1
                rows.append({"MEDIDA": met, p1[:18]: f"{va} ★", p2[:18]: vb})
            else:
                pts["b"] += 1
                rows.append({"MEDIDA": met, p1[:18]: va, p2[:18]: f"{vb} ★"})
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
        if pts["a"] == pts["b"]:
            st.success("EMPATE TÉCNICO nos critérios.")
        else:
            w = p1 if pts["a"] > pts["b"] else p2
            st.success(f"MELHOR MÊS: {w} ({max(pts.values())} x {min(pts.values())} nos critérios)")
        topa = max(pa.items(), key=lambda kv: kv[1]["fat"])[0] if pa else "—"
        topb = max(pb.items(), key=lambda kv: kv[1]["fat"])[0] if pb else "—"
        st.subheader("Campeão de vendas")
        if topa == topb:
            st.write(f"**{topa[:60]}** continua campeão. Proteja com estoque e ads.")
        else:
            st.write(f"Antes: {topa[:55]} → **Agora: {topb[:55]}** (virou o jogo, priorize ele).")
        subiram = sorted([p for p in pb if p in pa and pa[p]["classe"] == "C" and pb[p]["classe"] in "AB"])
        caíram = sorted([p for p in pb if p in pa and pa[p]["classe"] == "A" and pb[p]["classe"] in "BC"])
        e1, e2 = st.columns(2)
        with e1:
            st.write(f"**Subiram ({len(subiram)})**")
            st.dataframe(pd.DataFrame([{"Produto": p[:50], "Foi para": pb[p]["classe"]}
                                       for p in subiram[:20]]), use_container_width=True)
        with e2:
            st.write(f"**Caíram ({len(caíram)})**")
            st.dataframe(pd.DataFrame([{"Produto": p[:50], "Caiu para": pb[p]["classe"]}
                                       for p in caíram[:20]]), use_container_width=True)
        st.subheader("ONDE TRABALHAR PRIMEIRO")
        plano = []
        for i, p in enumerate(caíram[:5], 1):
            plano.append(f"**{i}. Recuperar {p[:50]}** (caiu {pa[p]['classe']} → {pb[p]['classe']}): conferir preço, foto e estoque.")
        j = len(plano) + 1
        db = {x["produto"]: x for x in hb.get("dados", [])}
        sem_est = [p for p in pb if pb[p]["classe"] == "A" and not (db.get(p, {}).get("est", 0) or 0)]
        if sem_est:
            plano.append(f"**{j}. Proteger o campeão {topb[:45]}**: informar estoque e manter ads. "
                         "(Isto vale para os sem estoque: " + ", ".join(p[:25] for p in sem_est[:3]) + ").")
            j += 1
        novos_c = [p for p in pb if (p not in pa or pa[p]["classe"] in "AB") and pb[p]["classe"] == "C"][:3]
        for p in novos_c:
            plano.append(f"**{j}. Girar {p[:45]}** (novo na C): montar kit com campeão.")
            j += 1
        mov = sorted(((pb[p]["fat"]-pa[p]["fat"], p) for p in pb if p in pa), reverse=True)
        if mov and mov[0][0] > 0:
            v, p = mov[0]
            plano.append(f"**{j}. Acelerar {p[:45]}** (maior alta: +R$ {v:,.0f}".replace(",", ".") + ").")
        for item in plano:
            st.write(f"• {item}")
        if not plano:
            st.write("Nada crítico: mantenha o plano atual.")
