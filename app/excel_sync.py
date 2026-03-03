import os
import pandas as pd
from datetime import datetime
from sqlalchemy import event
from .extensions import db
from .models import Hospital, Contato, DadosHospital, ProdutoHospital


# ==============================
# Pasta data/
# ==============================
def get_data_path(filename):
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    data_dir = os.path.join(base, "data")
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, filename)


def fmt_date(d):
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")


# ==============================
# EXPORT HOSPITAIS
# ==============================
def export_hospitais():
    path = get_data_path("hospitais.xlsx")

    hospitais = Hospital.query.order_by(Hospital.id).all()

    rows = []
    for h in hospitais:
        rows.append({
            "id": h.id,
            "nome_hospital": h.nome_hospital,
            "endereco": h.endereco,
            "numero": h.numero,
            "complemento": h.complemento,
            "cep": h.cep,
            "cidade": h.cidade,
            "estado": h.estado,
            "data_visita": fmt_date(h.data_visita),
            "data_retorno": fmt_date(h.data_retorno),
        })

    pd.DataFrame(rows).to_excel(path, index=False)


# ==============================
# EXPORT CONTATOS
# ==============================
def export_contatos():
    path = get_data_path("contatos.xlsx")

    contatos = Contato.query.order_by(Contato.id).all()

    rows = []
    for c in contatos:
        rows.append({
            "id": c.id,
            "hospital_id": c.hospital_id,
            "hospital_nome": c.hospital_nome,
            "nome_contato": c.nome_contato,
            "cargo": c.cargo,
            "telefone": c.telefone,
        })

    pd.DataFrame(rows).to_excel(path, index=False)


# ==============================
# EXPORT DADOS HOSPITAIS
# ==============================
def export_dados():
    path = get_data_path("dadoshospitais.xlsx")

    dados = DadosHospital.query.order_by(DadosHospital.id).all()

    rows = []
    for d in dados:
        rows.append({
            "id": d.id,
            "hospital_id": d.hospital_id,
            "especialidade": d.especialidade,
            "leitos": d.leitos,
            "leitos_uti": d.leitos_uti,
            "fatores_decisorios": d.fatores_decisorios,
            "prioridades_atendimento": d.prioridades_atendimento,
            "certificacao": d.certificacao,
            "emtn": d.emtn,
            "emtn_membros": d.emtn_membros,
            "comissao_feridas": d.comissao_feridas,
            "comissao_feridas_membros": d.comissao_feridas_membros,
            "nutricao_enteral_dia": d.nutricao_enteral_dia,
            "pacientes_tno_dia": d.pacientes_tno_dia,
            "altas_orientadas": d.altas_orientadas,
            "quem_orienta_alta": d.quem_orienta_alta,
            "protocolo_evolucao_dieta": d.protocolo_evolucao_dieta,
            "protocolo_evolucao_dieta_qual": d.protocolo_evolucao_dieta_qual,
            "protocolo_lesao_pressao": d.protocolo_lesao_pressao,
            "maior_desafio": d.maior_desafio,
            "dieta_padrao": d.dieta_padrao,
            "bomba_infusao_modelo": d.bomba_infusao_modelo,
            "fornecedor": d.fornecedor,
            "convenio_empresas": d.convenio_empresas,
            "convenio_empresas_modelo_pagamento": d.convenio_empresas_modelo_pagamento,
            "reembolso": d.reembolso,
            "modelo_compras": d.modelo_compras,
            "contrato_tipo": d.contrato_tipo,
            "nova_etapa_negociacao": d.nova_etapa_negociacao,
        })

    pd.DataFrame(rows).to_excel(path, index=False)


# ==============================
# EXPORT PRODUTOS HOSPITAIS
# ==============================
def export_produtos():
    path = get_data_path("produtoshospitais.xlsx")

    produtos = ProdutoHospital.query.order_by(ProdutoHospital.id).all()

    rows = []
    for p in produtos:
        rows.append({
            "id": p.id,
            "hospital_id": p.hospital_id,
            "nome_hospital": p.nome_hospital,
            "marca_planilha": p.marca_planilha,
            "produto": p.produto,
            "quantidade": p.quantidade,
            "embalagem": p.embalagem,
            "referencia": p.referencia,
            "kcal": p.kcal,
            "ptn": p.ptn,
            "lip": p.lip,
            "fibras": p.fibras,
            "sodio": p.sodio,
            "ferro": p.ferro,
            "potassio": p.potassio,
            "vit_b12": p.vit_b12,
            "gordura_saturada": p.gordura_saturada,
        })

    pd.DataFrame(rows).to_excel(path, index=False)


# ==============================
# EXPORT GERAL
# ==============================
def export_all():
    export_hospitais()
    export_contatos()
    export_dados()
    export_produtos()


# ==============================
# AUTO EXECUTAR APÓS COMMIT
# ==============================
# app/excel_sync.py
from sqlalchemy import event
from sqlalchemy.orm import Session as SASession

def register_excel_autosync(app):
    """
    Exporta Excel automaticamente após qualquer commit.
    """
    @event.listens_for(SASession, "after_commit")
    def _after_commit(session):
        # garante que é o session do Flask-SQLAlchemy
        try:
            if session is not db.session:
                return
        except Exception:
            pass

        try:
            with app.app_context():
                export_all()
                app.logger.info("Excel exportado com sucesso em /data")
        except Exception as e:
            # se falhar, não derruba a rota — mas LOGA
            try:
                app.logger.exception(f"Falha ao exportar Excel automaticamente: {e}")
            except Exception:
                print("Falha ao exportar Excel automaticamente:", e)