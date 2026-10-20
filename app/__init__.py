# Factory de la aplicación Flask.
# Usamos el patrón "application factory" para poder registrar
# rutas (blueprints) de forma organizada.

from flask import Flask


def create_app():
    # Creamos la instancia principal de Flask
    app = Flask(__name__)

    # Importamos el blueprint de rutas y lo registramos en la app.
    # Se importa aquí dentro (y no arriba del todo) para evitar
    # problemas de importación circular.
    from .routes import main
    app.register_blueprint(main)

    return app