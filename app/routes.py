# Rutas (endpoints) de la aplicación.
# "/" muestra el buscador. Si se envía una búsqueda (parámetro "oferta"
# en la URL, vía GET), se muestra también el resultado.

from flask import Blueprint, render_template, request
from .excel_reader import buscar_oferta

main = Blueprint("main", __name__)


@main.route("/")
def index():
    # request.args.get() lee parámetros de la URL, ej: /?oferta=01OFE...
    # Si no hay parámetro "oferta", devuelve cadena vacía.
    oferta_buscada = request.args.get("oferta", "").strip()

    resultado = None
    no_encontrado = False
    error = None

    if oferta_buscada:
        try:
            resultado = buscar_oferta(oferta_buscada)
            if resultado is None:
                no_encontrado = True
        except FileNotFoundError:
            # No hay ningún .xlsx en la carpeta Excels/
            error = "No se ha encontrado el Excel en la carpeta Excels/. Añade el fichero del día e inténtalo de nuevo."

    return render_template(
        "index.html",
        oferta_buscada=oferta_buscada,
        resultado=resultado,
        no_encontrado=no_encontrado,
        error=error,
    )