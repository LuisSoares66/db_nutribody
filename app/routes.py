import io
import csv
import os
import json
from datetime import datetime
from app.excel_repo import list_visitas, save_visita, delete_visita

from flask import Blueprint, render_template, request, redirect, url_for, flash, Response, abort

from app.excel_repo import (
    list_hospitais, get_hospital, save_hospital, delete_hospital,
    list_contatos, save_contato, delete_contato,
    get_dados, save_dados,
    list_produtos, save_produto, delete_produto,
    DATA_FILE
)

from .product_catalog import load_produtos_catalog

bp = Blueprint("main", __name__)


def _ensure_excel_exists():
    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(
            f"Arquivo Excel não encontrado em: {DATA_FILE}\n"
            f"Coloque o backup_nutri_hospital.xlsx dentro da pasta /data."
        )

def _parse_date(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except Exception:
        return None


@bp.route("/")
def index():
    return redirect(url_for("main.hospitais"))

@bp.route("/ping")
def ping():
    return "OK"


# =========================
# HOSPITAIS
# =========================
@bp.route("/hospitais")
def hospitais():
    _ensure_excel_exists()
    ordem = (request.args.get("ordem") or "nome").strip().lower()
    if ordem not in ("nome", "cidade"):
        ordem = "nome"
    hospitais_excel = list_hospitais(order_by=ordem)
    return render_template("hospitais.html", hospitais=hospitais_excel, ordem_atual=ordem)

# =========================
# PRODUTOS DO HOSPITAL
# =========================
# =========================
# PRODUTOS DO HOSPITAL
# =========================
@bp.route("/hospitais/<int:hospital_id>/produtos", methods=["GET", "POST"])
def hospital_produtos(hospital_id):
    _ensure_excel_exists()

    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    # -------- CATÁLOGO (data/produtos.xlsx) --------
    catalog_rows = []
    fabricantes = []

    try:
        catalog = load_produtos_catalog()
        catalog_rows = catalog.to_dict("records")

        fabricantes = sorted({
            (r.get("fabricante") or "").strip()
            for r in catalog_rows
            if (r.get("fabricante") or "").strip()
        })

    except Exception as e:
        flash(f"Erro ao carregar produtos.xlsx: {e}", "error")
        catalog_rows = []
        fabricantes = []

    # -------- SALVAR PRODUTO --------
    if request.method == "POST":

        fabricante = (request.form.get("fabricante") or "").strip()
        produto = (request.form.get("produto") or "").strip()

        try:
            quantidade = int(request.form.get("quantidade") or 0)
        except:
            quantidade = 0

        if not produto:
            flash("Selecione um produto.", "error")
            return redirect(url_for("main.hospital_produtos", hospital_id=hospital_id))

        payload = {
            "id": None,
            "hospital_id": hospital_id,
            "nome_hospital": hospital.get("nome_hospital") or "",
            "marca_planilha": fabricante,
            "produto": produto,
            "quantidade": quantidade,

            # campos ocultos preenchidos via JS
            "embalagem": (request.form.get("embalagem") or "").strip(),
            "referencia": (request.form.get("referencia") or "").strip(),
            "kcal": (request.form.get("kcal") or "").strip(),
            "ptn": (request.form.get("ptn") or "").strip(),
            "lip": (request.form.get("lip") or "").strip(),
            "fibras": (request.form.get("fibras") or "").strip(),
            "sodio": (request.form.get("sodio") or "").strip(),
            "ferro": (request.form.get("ferro") or "").strip(),
            "potassio": (request.form.get("potassio") or "").strip(),
            "vit_b12": (request.form.get("vit_b12") or "").strip(),
            "gordura_saturada": (request.form.get("gordura_saturada") or "").strip(),
        }

        save_produto(payload)
        flash("Produto salvo com sucesso ✅", "success")

        return redirect(url_for("main.hospital_produtos", hospital_id=hospital_id))

    # -------- LISTAGEM --------
    itens = list_produtos(hospital_id)
    itens.sort(key=lambda x: int(x.get("id") or 0), reverse=True)

    return render_template(
        "hospital_produtos.html",
        hospital=hospital,
        itens=itens,
        fabricantes=fabricantes,
        catalog_json=json.dumps(catalog_rows, ensure_ascii=False),
    )


# -------- EXCLUIR PRODUTO (AJAX) --------
@bp.route("/hospitais/<int:hospital_id>/produtos/<int:produto_id>/delete", methods=["POST"])
def produto_delete(hospital_id, produto_id):
    _ensure_excel_exists()

    hospital = get_hospital(hospital_id)
    if not hospital:
        return "Hospital não encontrado", 404

    delete_produto(produto_id)

    return "", 204


@bp.route("/hospitais/novo", methods=["GET", "POST"])
def novo_hospital():
    _ensure_excel_exists()
    if request.method == "POST":
        nome = (request.form.get("nome_hospital") or "").strip()
        if not nome:
            flash("Informe o nome do hospital.", "error")
            return redirect(url_for("main.novo_hospital"))

        payload = {
            "id": None,
            "nome_hospital": nome,
            "endereco": (request.form.get("endereco") or "").strip(),
            "numero": (request.form.get("numero") or "").strip(),
            "complemento": (request.form.get("complemento") or "").strip(),
            "cep": (request.form.get("cep") or "").strip(),
            "cidade": (request.form.get("cidade") or "").strip(),
            "estado": (request.form.get("estado") or "").strip(),
        }
        new_id = save_hospital(payload)
        flash("Hospital cadastrado no Excel ✅", "success")
        return redirect(url_for("main.hospital_info", hospital_id=new_id))

    return render_template("hospital_form.html")

@bp.route("/hospitais/<int:hospital_id>/info", methods=["GET", "POST"])
def hospital_info(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado no Excel.", "error")
        return redirect(url_for("main.hospitais"))

    if request.method == "POST":
        nome = (request.form.get("nome_hospital") or "").strip()
        if not nome:
            flash("Nome do hospital é obrigatório.", "error")
            return redirect(url_for("main.hospital_info", hospital_id=hospital_id))

        hospital["nome_hospital"] = nome
        hospital["endereco"] = (request.form.get("endereco") or "").strip()
        hospital["numero"] = (request.form.get("numero") or "").strip()
        hospital["complemento"] = (request.form.get("complemento") or "").strip()
        hospital["cep"] = (request.form.get("cep") or "").strip()
        hospital["cidade"] = (request.form.get("cidade") or "").strip()
        hospital["estado"] = (request.form.get("estado") or "").strip()

        dv = _parse_date(request.form.get("data_visita"))
        dr = _parse_date(request.form.get("data_retorno"))
        hospital["data_visita"] = dv.isoformat() if dv else ""
        hospital["data_retorno"] = dr.isoformat() if dr else ""

        save_hospital(hospital)
        flash("Informações atualizadas no Excel ✅", "success")
        return redirect(url_for("main.hospital_info", hospital_id=hospital_id))

    visitas = list_visitas(hospital_id)
    return render_template("hospital_info.html", hospital=hospital, visitas=visitas)

@bp.route("/hospitais/<int:hospital_id>/excluir", methods=["POST"])
def excluir_hospital(hospital_id):
    _ensure_excel_exists()
    if not get_hospital(hospital_id):
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    delete_hospital(hospital_id)
    flash("Hospital e vínculos removidos do Excel ✅", "success")
    return redirect(url_for("main.hospitais"))

@bp.route("/hospitais/<int:hospital_id>/contatos", methods=["GET", "POST"])
def contatos(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    if request.method == "POST":
        contato_id = (request.form.get("contato_id") or "").strip()
        payload = {
            "id": int(contato_id) if contato_id else None,
            "hospital_id": hospital_id,
            "hospital_nome": hospital.get("nome_hospital") or "",
            "nome_contato": (request.form.get("nome_contato") or "").strip(),
            "cargo": (request.form.get("cargo") or "").strip(),
            "telefone": (request.form.get("telefone") or "").strip(),
        }
        if not payload["nome_contato"]:
            flash("Informe o nome do contato.", "error")
            return redirect(url_for("main.contatos", hospital_id=hospital_id))

        save_contato(payload)
        flash("Contato salvo no Excel ✅", "success")
        return redirect(url_for("main.contatos", hospital_id=hospital_id))

    contatos_excel = list_contatos(hospital_id)
    contatos_excel.sort(key=lambda x: int(x.get("id") or 0), reverse=True)
    return render_template("contatos.html", hospital=hospital, contatos=contatos_excel)

@bp.route("/hospitais/<int:hospital_id>/contatos/<int:contato_id>/excluir", methods=["POST"])
def excluir_contato(hospital_id, contato_id):
    _ensure_excel_exists()
    delete_contato(contato_id)
    flash("Contato removido ✅", "success")
    return redirect(url_for("main.contatos", hospital_id=hospital_id))

@bp.route("/hospitais/<int:hospital_id>/dados", methods=["GET", "POST"])
def dados_hospital(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    dados = get_dados(hospital_id) or {"id": None, "hospital_id": hospital_id}

    if request.method == "POST":
        payload = dict(dados)
        payload["hospital_id"] = hospital_id

        campos = [
            "especialidade", "leitos", "leitos_uti",
            "fatores_decisorios", "prioridades_atendimento", "certificacao",
            "emtn", "emtn_membros",
            "comissao_feridas", "comissao_feridas_membros",
            "nutricao_enteral_dia", "pacientes_tno_dia",
            "altas_orientadas", "quem_orienta_alta",
            "protocolo_evolucao_dieta", "protocolo_evolucao_dieta_qual",
            "protocolo_lesao_pressao", "maior_desafio", "dieta_padrao",
            "bomba_infusao_modelo", "fornecedor",
            "convenio_empresas", "convenio_empresas_modelo_pagamento",
            "reembolso", "modelo_compras", "contrato_tipo", "nova_etapa_negociacao",
        ]
        for c in campos:
            payload[c] = (request.form.get(c) or "").strip()

        save_dados(payload)
        flash("Dados atualizados no Excel ✅", "success")
        return redirect(url_for("main.dados_hospital", hospital_id=hospital_id))

    return render_template("dados_hospitais.html", hospital=hospital, dados=dados)


@bp.route("/hospitais/<int:hospital_id>/relatorios", methods=["GET"])
def relatorios(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    contatos_excel = list_contatos(hospital_id)
    dados = get_dados(hospital_id)
    produtos_excel = list_produtos(hospital_id)

    return render_template(
        "relatorios.html",
        hospital=hospital,
        contatos=contatos_excel,
        dados=dados,
        produtos=produtos_excel
    )

# =========================
# RELATÓRIO GERAL DE VISITAS
# =========================
from datetime import date, datetime
import io
import json
import pandas as pd
from flask import Response, send_file
from app.excel_repo import list_visitas, save_visita, delete_visita
# -------------------------
# helpers
# -------------------------
def _parse_iso_date(s: str):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def _fmt_br(d):
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")

def _build_visitas_rows():
    """
    Retorna lista de dicts:
    {
      nome_hospital, data_visita_iso, data_retorno_iso,
      data_visita_br, data_retorno_br,
      proxima_data (date|None),
      vencida (bool)
    }
    """
    hoje = date.today()
    hospitais = list_hospitais()

    rows = []
    for h in hospitais:
        nome = (h.get("nome_hospital") or "").strip()

        dv = _parse_iso_date(h.get("data_visita"))
        dr = _parse_iso_date(h.get("data_retorno"))

        proxima = dr or dv
        vencida = bool(dr and dr < hoje)

        rows.append({
            "nome_hospital": nome,
            "data_visita_iso": dv.isoformat() if dv else "",
            "data_retorno_iso": dr.isoformat() if dr else "",
            "data_visita_br": _fmt_br(dv),
            "data_retorno_br": _fmt_br(dr),
            "proxima_data": proxima,
            "vencida": vencida,
        })
    return rows


# =========================
# RELATÓRIO GERAL (VISITAS)
# =========================
@bp.route("/relatorios/visitas", methods=["GET"])
def relatorio_visitas():
    _ensure_excel_exists()

    ordem = (request.args.get("ordem") or "nome").strip().lower()
    hoje = date.today()

    rows = _build_visitas_rows()

    # ordenação crescente
    if ordem == "data_visita":
        rows.sort(key=lambda r: (r["data_visita_iso"] or "9999-12-31", r["nome_hospital"]))
    elif ordem == "data_retorno":
        rows.sort(key=lambda r: (r["data_retorno_iso"] or "9999-12-31", r["nome_hospital"]))
    else:
        rows.sort(key=lambda r: (r["nome_hospital"] or "").lower())

    # dashboard: próximas visitas (pega as 10 mais próximas >= hoje)
    proximas = []
    for r in rows:
        d = r["proxima_data"]
        if d and d >= hoje:
            proximas.append((d, r["nome_hospital"]))
    proximas.sort(key=lambda x: x[0])
    proximas = proximas[:10]

    dash_labels = [f"{d.strftime('%d/%m')}" for d, _ in proximas]
    dash_values = []
    # contamos quantas visitas caem na mesma data
    for lab in dash_labels:
        dash_values.append(dash_labels.count(lab))
    # remove duplicados mantendo ordem
    uniq_labels = []
    uniq_values = []
    for i, lab in enumerate(dash_labels):
        if lab not in uniq_labels:
            uniq_labels.append(lab)
            uniq_values.append(dash_values[i])

    return render_template(
        "relatorio_visitas.html",
        hospitais=rows,
        ordem_atual=ordem,
        hoje_br=hoje.strftime("%d/%m/%Y"),
        dash_labels=json.dumps(uniq_labels, ensure_ascii=False),
        dash_values=json.dumps(uniq_values, ensure_ascii=False),
    )


# =========================
# EXPORTAR EXCEL
# =========================
@bp.route("/relatorios/visitas.xlsx", methods=["GET"])
def relatorio_visitas_excel():
    _ensure_excel_exists()

    ordem = (request.args.get("ordem") or "nome").strip().lower()
    rows = _build_visitas_rows()

    if ordem == "data_visita":
        rows.sort(key=lambda r: (r["data_visita_iso"] or "9999-12-31", r["nome_hospital"]))
    elif ordem == "data_retorno":
        rows.sort(key=lambda r: (r["data_retorno_iso"] or "9999-12-31", r["nome_hospital"]))
    else:
        rows.sort(key=lambda r: (r["nome_hospital"] or "").lower())

    df = pd.DataFrame([{
        "Hospital": r["nome_hospital"],
        "Data da Visita": r["data_visita_br"] or "",
        "Data de Retorno": r["data_retorno_br"] or "",
        "Vencida?": "SIM" if r["vencida"] else "NÃO",
    } for r in rows])

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitas")
    bio.seek(0)

    filename = f"relatorio_visitas_{date.today().isoformat()}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================
# EXPORTAR PDF
# =========================
@bp.route("/relatorios/visitas.pdf", methods=["GET"])
def relatorio_visitas_pdf():
    _ensure_excel_exists()

    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    ordem = (request.args.get("ordem") or "nome").strip().lower()
    rows = _build_visitas_rows()

    if ordem == "data_visita":
        rows.sort(key=lambda r: (r["data_visita_iso"] or "9999-12-31", r["nome_hospital"]))
    elif ordem == "data_retorno":
        rows.sort(key=lambda r: (r["data_retorno_iso"] or "9999-12-31", r["nome_hospital"]))
    else:
        rows.sort(key=lambda r: (r["nome_hospital"] or "").lower())

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
    styles = getSampleStyleSheet()

    story = []
    story.append(Paragraph("Relatório Geral de Visitas", styles["Title"]))
    story.append(Paragraph(f"Gerado em: {date.today().strftime('%d/%m/%Y')}", styles["Normal"]))
    story.append(Spacer(1, 12))

    data = [["Hospital", "Data da Visita", "Data de Retorno", "Vencida?"]]
    for r in rows:
        data.append([
            r["nome_hospital"],
            r["data_visita_br"] or "-",
            r["data_retorno_br"] or "-",
            "SIM" if r["vencida"] else "NÃO",
        ])

    table = Table(data, repeatRows=1)
    ts = TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0b7d3e")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.lightgrey]),
    ])

    # pinta linha vencida
    for i, r in enumerate(rows, start=1):
        if r["vencida"]:
            ts.add("TEXTCOLOR", (0,i), (-1,i), colors.red)

    table.setStyle(ts)
    story.append(table)

    doc.build(story)
    buffer.seek(0)

    filename = f"relatorio_visitas_{date.today().isoformat()}.pdf"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/pdf")

@bp.route("/hospitais/<int:hospital_id>/relatorios/csv")
def relatorio_csv(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        return "Hospital não encontrado", 404

    contatos_excel = list_contatos(hospital_id)
    dados = get_dados(hospital_id) or {}
    produtos_excel = list_produtos(hospital_id)

    out = io.StringIO()
    w = csv.writer(out, delimiter=";")

    w.writerow(["HOSPITAL"])
    w.writerow([hospital.get("id"), hospital.get("nome_hospital"), hospital.get("cidade"), hospital.get("estado")])
    w.writerow([])

    w.writerow(["CONTATOS"])
    for c in contatos_excel:
        w.writerow([c.get("nome_contato"), c.get("cargo"), c.get("telefone")])
    w.writerow([])

    w.writerow(["DADOS"])
    w.writerow(["especialidade", dados.get("especialidade", "")])
    w.writerow(["leitos", dados.get("leitos", "")])
    w.writerow(["leitos_uti", dados.get("leitos_uti", "")])
    w.writerow([])

    w.writerow(["PRODUTOS"])
    for p in produtos_excel:
        w.writerow([p.get("marca_planilha", ""), p.get("produto", ""), p.get("quantidade", 0)])

    return Response(
        out.getvalue().encode("utf-8-sig"),
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="hospital_{hospital_id}.csv"'}
    )
    
@bp.route("/hospitais/<int:hospital_id>/visitas/add", methods=["POST"])
def add_visita(hospital_id):
    _ensure_excel_exists()
    hospital = get_hospital(hospital_id)
    if not hospital:
        flash("Hospital não encontrado.", "error")
        return redirect(url_for("main.hospitais"))

    payload = {
        "id": None,
        "hospital_id": hospital_id,
        "nome_hospital": hospital.get("nome_hospital") or "",
        "data_visita": (request.form.get("data_visita") or "").strip(),
        "data_retorno": (request.form.get("data_retorno") or "").strip(),
        "observacao": (request.form.get("observacao") or "").strip(),
        "criado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_visita(payload)
    flash("Visita/retorno adicionado ✅", "success")
    return redirect(url_for("main.hospital_info", hospital_id=hospital_id))

@bp.route("/hospitais/<int:hospital_id>/visitas/<int:visita_id>/delete", methods=["POST"])
def del_visita(hospital_id, visita_id):
    _ensure_excel_exists()
    delete_visita(visita_id)
    flash("Visita removida ✅", "success")
    return redirect(url_for("main.hospital_info", hospital_id=hospital_id))