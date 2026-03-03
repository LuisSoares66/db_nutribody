import os
from dotenv import load_dotenv
from flask import Flask

from .extensions import db  # vem de extensions.py

# Carrega variáveis do .env (na raiz do projeto)
load_dotenv()

def create_app():
    app = Flask(__name__)

    # Ideal: secret em variável de ambiente
    app.secret_key = os.environ.get("SECRET_KEY", "nutrihospital_offline_secret")

    # --- BANCO (Render Postgres) ---
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL não definida no .env")

    # garante sslmode=require
    if "sslmode=" not in db_url:
        db_url += ("&" if "?" in db_url else "?") + "sslmode=require"

    app.config["SQLALCHEMY_DATABASE_URI"] = db_url
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "pool_pre_ping": True,
        "pool_recycle": 120,
        "pool_timeout": 30,
        # importante no Render:
        "connect_args": {"sslmode": "require"},
    }

    # inicializa o SQLAlchemy
    db.init_app(app)
    from . import models
    # rotas
    from app.routes import bp
    app.register_blueprint(bp)

    # cria tabelas no banco (Render) caso não existam
    with app.app_context():
        db.create_all()

    return app

