# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from flask import Flask
from flask_login import LoginManager
from flask_sqlalchemy import SQLAlchemy
from importlib import import_module
from werkzeug.utils import secure_filename

from flask_ngrok import run_with_ngrok


db = SQLAlchemy()
login_manager = LoginManager()


def register_extensions(app):
    db.init_app(app)
    login_manager.init_app(app)


def register_blueprints(app):
    for module_name in ('authentication', 'home'):
        module = import_module('apps.{}.routes'.format(module_name))
        app.register_blueprint(module.blueprint)


def configure_database(app):

    @app.before_first_request
    def initialize_database():
        db.create_all()

    @app.teardown_request
    def shutdown_session(exception=None):
        db.session.remove()



def create_app(config):

    # Creamos un folder a donde van a ir a para los archivos temporales cuando los subamos.
    UPLOAD_FOLDER = 'apps/uploads/tempData'

    # Definimos el tipo de extensiones que vamos a acceptar.
    # That way you can make sure that users are not able to upload HTML files that would cause XSS problems (see Cross-Site Scripting (XSS)).
    # Also make sure to disallow .php files
    ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif','xlsx'}

    app = Flask(__name__)
    run_with_ngrok(app)
    app.config.from_object(config)

    # Creamos el app config que vamos a utilizar en las rutas para poder enviar los archivos temporales.
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

    register_extensions(app)
    register_blueprints(app)
    configure_database(app)
    return app
