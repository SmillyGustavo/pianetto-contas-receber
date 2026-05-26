"""Pianetto Transportes - Dashboard Contas a Receber (Streamlit Cloud).

Deploy: share.streamlit.io
Fonte: PostgreSQL (titulo id_tipo=1)
Acesso publico - sem senha.
"""
from __future__ import annotations

import io
from datetime import date, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import psycopg2
from psycopg2.extras import RealDictCursor
import streamlit as st

# =================== CONFIG ===================
st.set_page_config(
    page_title="Pianetto - Contas a Receber",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

PIANETTO_CDS = [38754, 38753, 2, 1715, 6691, 4353, 8800, 1538, 14376, 1792, 14171]
MESES_PT = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
            7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
DIAS_SEM = ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"]


# =================== DB ===================
@st.cache_data(ttl=900, show_spinner="Consultando banco...")
def carregar_dados(dt_ini: date, dt_fim: date) -> pd.DataFrame:
    cfg = st.secrets["postgres"]
    conn_kwargs = dict(
        host=cfg["host"], port=int(cfg["port"]), dbname=cfg["dbname"],
        user=cfg["user"], password=cfg["password"],
        sslmode=cfg.get("sslmode", "disable"), connect_timeout=30,
    )
    query = """
    SELECT
        t.cd_titulo,
        t.nr_titulo,
        t.nr_parcela,
        t.dt_vencimento,
        m.dt_movimento                          AS dt_emissao,
        t.dt_quitacao,
        t.vl_titulo,
        COALESCE(t.vl_pago, 0)                  AS vl_pago,
        (t.vl_titulo - COALESCE(t.vl_pago, 0))  AS vl_saldo,
        t.ds_observacao,
        p.nm_pessoa                             AS cliente,
        p.nr_cnpj_cpf                           AS cnpj_cliente,
        fil.nm_pessoa                           AS filial
    FROM titulo t
    JOIN movimento m         ON m.cd_movimento = t.cd_movimento
    LEFT JOIN pessoa p       ON p.cd_pessoa = t.cd_pessoa
    LEFT JOIN pessoa fil     ON fil.cd_pessoa = m.cd_pessoa_filial
    WHERE m.cd_pessoa_filial = ANY(%s)
      AND t.id_tipo = 1
      AND t.dt_vencimento BETWEEN %s AND %s
    ORDER BY t.dt_vencimento, p.nm_pessoa;
    """
    with psycopg2.connect(**conn_kwargs) as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SET statement_timeout = '180s';")
            cur.execute(query, (PIANETTO_CDS, dt_ini, dt_fim))
            rows = cur.fetchall()

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["vl_titulo"] = df["vl_titulo"].astype(float)
    df["vl_pago"]   = df["vl_pago"].astype(float)
    df["vl_saldo"]  = df["vl_saldo"].astype(float)
    df["dt_vencimento"] = pd.to_datetime(df["dt_vencimento"]).dt.date
    df["dt_emissao"]    = pd.to_datetime(df["dt_emissao"]).dt.date
    df["dt_quitacao"]   = pd.to_datetime(df["dt_quitacao"]).dt.date
    df["status"] = df["dt_quitacao"].map(lambda v: "Quitado" if pd.notna(v) else "Em aberto")
    return df


@st.cache_data(ttl=900, show_spinner="Carregando contas a pagar...")
def carregar_pagar(dt_ini: date, dt_fim: date) -> pd.DataFrame:
    """Contas a pagar EM ABERTO (id_tipo=2, dt_quitacao IS NULL, saldo > 0)."""
    cfg = st.secrets["postgres"]
    conn_kwargs = dict(
        host=cfg["host"], port=int(cfg["port"]), dbname=cfg["dbname"],
        user=cfg["user"], password=cfg["password"],
        sslmode=cfg.get("sslmode", "disable"), connect_timeout=30,
    )
    query = """
    SELECT
        t.dt_vencimento,
        t.vl_titulo,
        COALESCE(t.vl_pago, 0)                  AS vl_pago,
        (t.vl_titulo - COALESCE(t.vl_pago, 0))  AS vl_saldo,
        p.nm_pessoa                             AS favorecido,
        fil.nm_pessoa                           AS filial
    FROM titulo t
    JOIN movimento m         ON m.cd_movimento = t.cd_movimento
    LEFT JOIN pessoa p       ON p.cd_pessoa = t.cd_pessoa
    LEFT JOIN pessoa fil     ON fil.cd_pessoa = m.cd_pessoa_filial
    WHERE m.cd_pessoa_filial = ANY(%s)
      AND t.id_tipo = 2
      AND t.dt_quitacao IS NULL
      AND (t.vl_titulo - COALESCE(t.vl_pago, 0)) > 0
      AND t.dt_vencimento BETWEEN %s AND %s
    ORDER BY t.dt_vencimento;
    """
    with psycopg2.connect(**conn_kwargs) as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SET statement_timeout = '180s';")
            cur.execute(query, (PIANETTO_CDS, dt_ini, dt_fim))
            rows = cur.fetchall()
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["vl_titulo"] = df["vl_titulo"].astype(float)
    df["vl_pago"]   = df["vl_pago"].astype(float)
    df["vl_saldo"]  = df["vl_saldo"].astype(float)
    df["dt_vencimento"] = pd.to_datetime(df["dt_vencimento"]).dt.date
    return df


# =================== HELPERS ===================
def fmt_brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def mask_cnpj(v: str) -> str:
    if not isinstance(v, str) or len(v) < 4:
        return "***"
    return "***" + v[-4:]


def to_excel(df: pd.DataFrame, sheet: str = "Dados") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet, index=False)
    return buf.getvalue()


# =================== APP ===================
def main():
    # ===== Sidebar =====
    st.sidebar.title("⚙️ Filtros")
    hoje = date.today()
    dias = st.sidebar.slider("Janela (dias a frente)", 7, 90, 30, step=1)
    dt_ini = st.sidebar.date_input("Inicio", hoje)
    dt_fim = dt_ini + timedelta(days=dias)
    st.sidebar.caption(f"Periodo: {dt_ini.strftime('%d/%m/%Y')} → {dt_fim.strftime('%d/%m/%Y')}")

    st.sidebar.divider()
    st.sidebar.subheader("🔒 Privacidade")
    mascarar_cnpj = st.sidebar.toggle(
        "Mascarar CPF/CNPJ", value=True,
        help="Exibe apenas os 4 ultimos digitos. Recomendado em dashboard publico."
    )
    mascarar_cli = st.sidebar.toggle(
        "Anonimizar clientes", value=False,
        help="Substitui nome por 'Cliente A', 'Cliente B'... preservando ranking."
    )

    st.sidebar.divider()
    if st.sidebar.button("🔄 Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.markdown(
        "**Fonte:** `titulo` (id_tipo=1)  \n"
        "**Filiais:** 11 Pianetto (CNPJ raiz 43.976.512)  \n"
        f"**Cache TTL:** 15 min  \n"
        f"**Hoje:** {hoje.strftime('%d/%m/%Y')}"
    )

    # ===== Header =====
    st.title("💰 Pianetto Transportes")
    st.subheader(f"Contas a Receber — {dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}")

    df = carregar_dados(dt_ini, dt_fim)

    if df.empty:
        st.warning("Nenhum titulo no periodo selecionado.")
        return

    # Anonimizacao (apos cache - operacao barata)
    df = df.copy()
    if mascarar_cnpj:
        df["cnpj_cliente"] = df["cnpj_cliente"].map(mask_cnpj)
    if mascarar_cli:
        # mapeamento estavel por valor total (maior = Cliente A)
        ordem = (df.groupby("cliente")["vl_saldo"].sum()
                   .sort_values(ascending=False).index.tolist())
        mapa = {c: f"Cliente {chr(65 + (i // 26))}{i % 26 + 1:02d}" for i, c in enumerate(ordem)}
        df["cliente"] = df["cliente"].map(mapa).fillna("(sem cliente)")

    # Filtros adicionais
    filiais = sorted(df["filial"].dropna().unique().tolist())
    clientes = sorted(df["cliente"].dropna().unique().tolist())

    sel_filiais = st.sidebar.multiselect("Filiais", filiais, default=filiais, key="filiais")
    sel_clientes = st.sidebar.multiselect("Clientes", clientes, default=clientes, key="clientes")
    sel_status = st.sidebar.multiselect(
        "Status", ["Em aberto", "Quitado"], default=["Em aberto", "Quitado"], key="status"
    )

    mask = (df["filial"].isin(sel_filiais)
            & df["cliente"].isin(sel_clientes)
            & df["status"].isin(sel_status))
    df = df[mask].copy()
    if df.empty:
        st.warning("Filtros nao retornaram resultados.")
        return

    # ===== Carrega Contas a Pagar (mesma janela, em aberto) =====
    df_pag = carregar_pagar(dt_ini, dt_fim)
    if not df_pag.empty:
        df_pag = df_pag[df_pag["filial"].isin(sel_filiais)].copy()

    # ===== KPIs =====
    total_titulo = df["vl_titulo"].sum()
    total_pago   = df["vl_pago"].sum()
    total_saldo  = df["vl_saldo"].sum()
    qtd_tit      = len(df)
    qtd_quit     = (df["status"] == "Quitado").sum()
    qtd_aberto   = (df["status"] == "Em aberto").sum()

    df_aberto = df[df["status"] == "Em aberto"]
    venc_hoje = df_aberto[df_aberto["dt_vencimento"] == hoje]["vl_saldo"].sum()
    s7  = df_aberto[df_aberto["dt_vencimento"] <= hoje + timedelta(days=7)]["vl_saldo"].sum()
    s14 = df_aberto[df_aberto["dt_vencimento"] <= hoje + timedelta(days=14)]["vl_saldo"].sum()
    s21 = df_aberto[df_aberto["dt_vencimento"] <= hoje + timedelta(days=21)]["vl_saldo"].sum()

    # Totais AP
    total_pagar       = df_pag["vl_saldo"].sum() if not df_pag.empty else 0.0
    qtd_pagar         = len(df_pag) if not df_pag.empty else 0
    saldo_liquido     = total_saldo - total_pagar
    pagar_hoje = pagar_7 = pagar_14 = pagar_21 = 0.0
    if not df_pag.empty:
        pagar_hoje = df_pag[df_pag["dt_vencimento"] == hoje]["vl_saldo"].sum()
        pagar_7  = df_pag[df_pag["dt_vencimento"] <= hoje + timedelta(days=7)]["vl_saldo"].sum()
        pagar_14 = df_pag[df_pag["dt_vencimento"] <= hoje + timedelta(days=14)]["vl_saldo"].sum()
        pagar_21 = df_pag[df_pag["dt_vencimento"] <= hoje + timedelta(days=21)]["vl_saldo"].sum()

    # ===== KPIs Linha 1 - Fluxo de Caixa =====
    st.markdown("##### 💰 Fluxo de Caixa do periodo (em aberto)")
    cf1, cf2, cf3 = st.columns(3)
    cf1.metric("Saldo a Receber", fmt_brl(total_saldo), f"{qtd_aberto} titulos",
               delta_color="normal")
    cf2.metric("Saldo a Pagar", fmt_brl(total_pagar), f"{qtd_pagar} titulos",
               delta_color="inverse")
    delta_label = "Positivo" if saldo_liquido >= 0 else "Negativo"
    cf3.metric("Saldo Liquido (R − P)", fmt_brl(saldo_liquido), delta_label,
               delta_color="normal" if saldo_liquido >= 0 else "inverse")

    st.markdown("##### 📊 Detalhe Contas a Receber")
    # Linha 2 - macro AR
    c1, c2, c3 = st.columns(3)
    c1.metric("Valor total dos titulos", fmt_brl(total_titulo), f"{qtd_tit} titulos")
    c2.metric("Ja recebido", fmt_brl(total_pago), f"{qtd_quit} quitados")
    c3.metric("Saldo a receber", fmt_brl(total_saldo), f"{qtd_aberto} em aberto")

    # Linha 3 - prazo AR
    c4, c5, c6, c7 = st.columns(4)
    c4.metric("Vence hoje (aberto)", fmt_brl(venc_hoje))
    c5.metric("Proximos 7 dias", fmt_brl(s7))
    c6.metric("Proximos 14 dias", fmt_brl(s14))
    c7.metric("Proximos 21 dias", fmt_brl(s21))

    st.divider()

    # ===== Tabs =====
    tab_cf, tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["💰 Fluxo de Caixa", "📅 Por Dia", "👥 Por Cliente",
         "🔥 Dia × Cliente", "🏬 Por Filial", "📋 Detalhe"]
    )

    # --- TAB CF: FLUXO DE CAIXA (Receber vs Pagar) ---
    with tab_cf:
        st.subheader("Receber × Pagar por dia (saldos em aberto)")

        # AR aberto por dia
        ar_dia = (df_aberto.groupby("dt_vencimento")["vl_saldo"].sum()
                  .reset_index()
                  .rename(columns={"vl_saldo": "receber"}))
        # AP aberto por dia
        if df_pag.empty:
            ap_dia = pd.DataFrame(columns=["dt_vencimento", "pagar"])
        else:
            ap_dia = (df_pag.groupby("dt_vencimento")["vl_saldo"].sum()
                      .reset_index()
                      .rename(columns={"vl_saldo": "pagar"}))

        full = pd.DataFrame({"dt_vencimento": pd.date_range(dt_ini, dt_fim, freq="D").date})
        cf = (full.merge(ar_dia, on="dt_vencimento", how="left")
                   .merge(ap_dia, on="dt_vencimento", how="left")
                   .fillna(0))
        cf["liquido"] = cf["receber"] - cf["pagar"]
        cf["acumulado"] = cf["liquido"].cumsum()
        cf["dia"] = cf["dt_vencimento"].map(lambda d: DIAS_SEM[d.weekday()])

        # KPIs do periodo
        tot_r = float(cf["receber"].sum())
        tot_p = float(cf["pagar"].sum())
        tot_l = tot_r - tot_p
        k1, k2, k3 = st.columns(3)
        k1.metric("Total a Receber (periodo)", fmt_brl(tot_r))
        k2.metric("Total a Pagar (periodo)", fmt_brl(tot_p))
        k3.metric("Saldo Liquido", fmt_brl(tot_l),
                  "Positivo" if tot_l >= 0 else "Negativo",
                  delta_color="normal" if tot_l >= 0 else "inverse")

        # Grafico: Receber (positivo, verde), Pagar (negativo, vermelho), linha acumulado
        fig = go.Figure()
        fig.add_bar(
            x=cf["dt_vencimento"], y=cf["receber"],
            name="A Receber", marker_color="#2E8B57",
            text=cf["receber"].map(lambda v: fmt_brl(v) if v else ""),
            textposition="outside",
            hovertemplate="Receber: R$ %{y:,.2f}<extra></extra>",
        )
        fig.add_bar(
            x=cf["dt_vencimento"], y=-cf["pagar"],
            name="A Pagar", marker_color="#C0392B",
            text=cf["pagar"].map(lambda v: f"-{fmt_brl(v)}" if v else ""),
            textposition="outside",
            hovertemplate="Pagar: R$ %{customdata:,.2f}<extra></extra>",
            customdata=cf["pagar"],
        )
        fig.add_scatter(
            x=cf["dt_vencimento"], y=cf["acumulado"],
            name="Liquido acumulado", mode="lines+markers",
            line=dict(color="#1F4E78", width=2.5),
            yaxis="y2",
            hovertemplate="Acumulado: R$ %{y:,.2f}<extra></extra>",
        )
        fig.add_hline(y=0, line_width=1, line_color="#666")
        fig.add_vline(x=hoje.isoformat(), line_dash="dash", line_color="#FFC000")
        fig.add_annotation(
            x=hoje.isoformat(), y=1.02, yref="paper",
            text="Hoje", showarrow=False, xanchor="left",
            font=dict(color="#FFC000", size=12),
            bgcolor="rgba(255,255,255,0.7)",
        )
        fig.update_layout(
            barmode="relative", height=520,
            title="Fluxo de caixa diario (Receber positivo · Pagar negativo)",
            xaxis=dict(tickformat="%d/%m"),
            yaxis=dict(title="Valor diario (R$)", zeroline=True),
            yaxis2=dict(title="Saldo acumulado (R$)", overlaying="y",
                        side="right", showgrid=False),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right"),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Tabela diaria")
        tbl = cf.rename(columns={
            "dt_vencimento": "Data", "dia": "Dia",
            "receber": "A Receber", "pagar": "A Pagar",
            "liquido": "Liquido", "acumulado": "Acumulado",
        })[["Data", "Dia", "A Receber", "A Pagar", "Liquido", "Acumulado"]].copy()

        # Estilizando a tabela para destacar liquido negativo
        st.dataframe(
            tbl,
            use_container_width=True, hide_index=True,
            column_config={
                "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "A Receber": st.column_config.NumberColumn(format="R$ %.2f"),
                "A Pagar": st.column_config.NumberColumn(format="R$ %.2f"),
                "Liquido": st.column_config.NumberColumn(format="R$ %.2f"),
                "Acumulado": st.column_config.NumberColumn(format="R$ %.2f"),
            },
        )

        st.download_button(
            "⬇️ Baixar Excel - Fluxo de Caixa",
            data=to_excel(tbl, "FluxoCaixa"),
            file_name=f"pianetto_fluxo_caixa_{hoje.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Dias criticos (liquido negativo)
        criticos = cf[cf["liquido"] < 0]
        if not criticos.empty:
            st.warning(
                f"⚠️ **{len(criticos)} dia(s)** com saldo liquido negativo no periodo. "
                f"Maior deficit: **{fmt_brl(criticos['liquido'].min())}** em "
                f"{criticos.loc[criticos['liquido'].idxmin(), 'dt_vencimento'].strftime('%d/%m/%Y')}."
            )

    # --- TAB 1: POR DIA ---
    with tab1:
        st.subheader("Valores por dia de vencimento")
        por_dia = (df.groupby("dt_vencimento")
                     .agg(qtd=("vl_saldo", "size"),
                          vl_titulo=("vl_titulo", "sum"),
                          vl_pago=("vl_pago", "sum"),
                          vl_saldo=("vl_saldo", "sum"))
                     .reset_index())
        full = pd.DataFrame({"dt_vencimento": pd.date_range(dt_ini, dt_fim, freq="D").date})
        por_dia = full.merge(por_dia, on="dt_vencimento", how="left").fillna(0)
        por_dia["qtd"] = por_dia["qtd"].astype(int)
        por_dia["dia"] = por_dia["dt_vencimento"].map(lambda d: DIAS_SEM[d.weekday()])

        # grafico empilhado: pago + saldo aberto
        fig = go.Figure()
        fig.add_bar(
            x=por_dia["dt_vencimento"], y=por_dia["vl_pago"],
            name="Recebido", marker_color="#2E8B57",
            text=por_dia["vl_pago"].map(lambda v: fmt_brl(v) if v else ""),
            textposition="inside",
        )
        fig.add_bar(
            x=por_dia["dt_vencimento"], y=por_dia["vl_saldo"],
            name="Saldo a receber", marker_color="#1F4E78",
            text=por_dia["vl_saldo"].map(lambda v: fmt_brl(v) if v else ""),
            textposition="outside",
        )
        fig.update_layout(
            barmode="stack", height=480,
            title="Valor por dia (recebido + saldo)",
            xaxis=dict(tickformat="%d/%m"),
            yaxis_title="Valor (R$)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right"),
        )
        # destaque hoje
        fig.add_vline(x=hoje.isoformat(), line_dash="dash", line_color="#FFC000")
        fig.add_annotation(
            x=hoje.isoformat(), y=1.02, yref="paper",
            text="Hoje", showarrow=False, xanchor="left",
            font=dict(color="#FFC000", size=12),
            bgcolor="rgba(255,255,255,0.7)",
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Tabela por dia")
        tbl = por_dia.rename(columns={
            "dt_vencimento": "Data", "dia": "Dia", "qtd": "Qtd",
            "vl_titulo": "Valor Total", "vl_pago": "Recebido", "vl_saldo": "Saldo Aberto",
        })[["Data", "Dia", "Qtd", "Valor Total", "Recebido", "Saldo Aberto"]].copy()
        for col in ["Valor Total", "Recebido", "Saldo Aberto"]:
            tbl[col] = tbl[col].map(fmt_brl)
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # --- TAB 2: POR CLIENTE ---
    with tab2:
        st.subheader("Ranking de clientes")
        por_c = (df.groupby(["cliente", "cnpj_cliente"], dropna=False)
                   .agg(qtd=("vl_saldo", "size"),
                        vl_titulo=("vl_titulo", "sum"),
                        vl_pago=("vl_pago", "sum"),
                        vl_saldo=("vl_saldo", "sum"),
                        primeiro_venc=("dt_vencimento", "min"),
                        ultimo_venc=("dt_vencimento", "max"))
                   .reset_index()
                   .sort_values("vl_saldo", ascending=False))
        denom = total_saldo if total_saldo > 0 else 1.0
        por_c["%"] = (por_c["vl_saldo"] / denom * 100).round(2)
        por_c["acum %"] = por_c["%"].cumsum().round(2)

        col_a, col_b = st.columns([3, 2])
        with col_a:
            topn = st.slider("Top N", 5, 50, 15, key="topn_cli")
            top_df = por_c.head(topn).iloc[::-1]
            fig = px.bar(
                top_df, x="vl_saldo", y="cliente", orientation="h",
                labels={"vl_saldo": "Saldo a receber (R$)", "cliente": ""},
                title=f"Top {topn} clientes - saldo a receber",
                text=top_df["vl_saldo"].map(fmt_brl),
                color="vl_saldo", color_continuous_scale="Blues",
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=max(420, 24 * topn), coloraxis_showscale=False,
                              margin=dict(l=10, r=140, t=60, b=20))
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            fig2 = px.pie(
                por_c.head(10), values="vl_saldo", names="cliente",
                title="Top 10 - share do saldo",
                hole=0.45,
            )
            fig2.update_traces(textposition="inside", textinfo="percent+label")
            fig2.update_layout(height=520, showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("##### Tabela completa")
        tbl = por_c.rename(columns={
            "cliente": "Cliente", "cnpj_cliente": "CNPJ/CPF",
            "qtd": "Qtd",
            "vl_titulo": "Valor Total", "vl_pago": "Recebido", "vl_saldo": "Saldo",
            "primeiro_venc": "1º Venc", "ultimo_venc": "Último Venc",
        })
        tbl_disp = tbl.copy()
        for col in ["Valor Total", "Recebido", "Saldo"]:
            tbl_disp[col] = tbl_disp[col].map(fmt_brl)
        tbl_disp["%"] = tbl_disp["%"].map(lambda v: f"{v:.2f}%")
        tbl_disp["acum %"] = tbl_disp["acum %"].map(lambda v: f"{v:.2f}%")
        st.dataframe(tbl_disp, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Baixar Excel - Por Cliente",
            data=to_excel(tbl, "PorCliente"),
            file_name=f"pianetto_receber_por_cliente_{hoje.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # --- TAB 3: DIA x CLIENTE (heatmap) ---
    with tab3:
        st.subheader("Matriz Dia × Cliente")
        topn_h = st.slider("Top N clientes", 5, 40, 20, key="topn_heat")

        top_c = (df.groupby("cliente")["vl_saldo"].sum()
                   .sort_values(ascending=False).head(topn_h).index.tolist())
        df_h = df[df["cliente"].isin(top_c)].copy()
        pivot = df_h.pivot_table(
            index="cliente", columns="dt_vencimento",
            values="vl_saldo", aggfunc="sum", fill_value=0,
        )
        pivot = pivot.loc[top_c]

        fig = go.Figure(data=go.Heatmap(
            z=pivot.values,
            x=[d.strftime("%d/%m") for d in pivot.columns],
            y=pivot.index,
            colorscale="Blues",
            hovertemplate="Cliente: %{y}<br>Dia: %{x}<br>Saldo: R$ %{z:,.2f}<extra></extra>",
            colorbar=dict(title="R$"),
        ))
        fig.update_layout(
            title=f"Heatmap - Top {topn_h} clientes × dias",
            height=max(480, 26 * topn_h),
            xaxis=dict(tickangle=-45),
            yaxis=dict(autorange="reversed"),
            margin=dict(l=10, r=10, t=60, b=80),
        )
        st.plotly_chart(fig, use_container_width=True)

    # --- TAB 4: FILIAL ---
    with tab4:
        st.subheader("Distribuicao por filial")
        por_fil = (df.groupby("filial", dropna=False)
                     .agg(qtd=("vl_saldo", "size"),
                          vl_titulo=("vl_titulo", "sum"),
                          vl_pago=("vl_pago", "sum"),
                          vl_saldo=("vl_saldo", "sum"))
                     .reset_index()
                     .sort_values("vl_saldo", ascending=False))
        denom = total_saldo if total_saldo > 0 else 1.0
        por_fil["%"] = (por_fil["vl_saldo"] / denom * 100).round(2)

        col_a, col_b = st.columns(2)
        with col_a:
            fig = px.bar(
                por_fil.iloc[::-1], x="vl_saldo", y="filial", orientation="h",
                text=por_fil.iloc[::-1]["vl_saldo"].map(fmt_brl),
                title="Saldo a receber por filial",
                color="vl_saldo", color_continuous_scale="Teal",
                labels={"vl_saldo": "Saldo (R$)", "filial": ""},
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=440, coloraxis_showscale=False, margin=dict(r=140))
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            fig = px.pie(por_fil, values="vl_saldo", names="filial", hole=0.4,
                         title="Share por filial")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(height=440, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        tbl = por_fil.rename(columns={
            "filial": "Filial", "qtd": "Qtd",
            "vl_titulo": "Valor Total", "vl_pago": "Recebido", "vl_saldo": "Saldo",
        }).copy()
        for col in ["Valor Total", "Recebido", "Saldo"]:
            tbl[col] = tbl[col].map(fmt_brl)
        tbl["%"] = tbl["%"].map(lambda v: f"{v:.2f}%")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # --- TAB 5: DETALHE ---
    with tab5:
        st.subheader(f"Detalhe — {len(df)} titulos")
        det = df[[
            "dt_vencimento", "status", "filial", "cliente", "cnpj_cliente",
            "nr_titulo", "nr_parcela", "dt_emissao", "dt_quitacao",
            "vl_titulo", "vl_pago", "vl_saldo", "ds_observacao",
        ]].copy()
        det["dias_p_venc"] = det["dt_vencimento"].map(lambda d: (d - hoje).days)
        det = det.rename(columns={
            "dt_vencimento": "Vencimento", "status": "Status",
            "filial": "Filial", "cliente": "Cliente",
            "cnpj_cliente": "CNPJ/CPF",
            "nr_titulo": "Nº Titulo", "nr_parcela": "Parc",
            "dt_emissao": "Emissão", "dt_quitacao": "Quitacao",
            "vl_titulo": "Vl Titulo", "vl_pago": "Vl Recebido",
            "vl_saldo": "Vl Saldo", "ds_observacao": "Observacao",
            "dias_p_venc": "Dias p/ Venc",
        })
        det = det.sort_values(["Vencimento", "Cliente"])

        st.dataframe(
            det,
            use_container_width=True, hide_index=True,
            column_config={
                "Vl Titulo": st.column_config.NumberColumn(format="R$ %.2f"),
                "Vl Recebido": st.column_config.NumberColumn(format="R$ %.2f"),
                "Vl Saldo": st.column_config.NumberColumn(format="R$ %.2f"),
                "Vencimento": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Emissão": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Quitacao": st.column_config.DateColumn(format="DD/MM/YYYY"),
            },
        )
        st.download_button(
            "⬇️ Baixar Excel - Detalhe",
            data=to_excel(det, "Detalhe"),
            file_name=f"pianetto_receber_detalhe_{hoje.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ===== Rodape =====
    st.divider()
    st.caption(
        "📊 Pianetto Transportes - Contas a Receber + Fluxo de Caixa  |  "
        "Dados consolidados das 11 filiais (CNPJ raiz 43.976.512)  |  "
        "Saldo a Pagar: titulos em aberto (id_tipo=2, dt_quitacao IS NULL)  |  "
        f"Atualizado em {hoje.strftime('%d/%m/%Y')}"
    )


if __name__ == "__main__":
    main()

fix: separa add_vline e add_annotation
