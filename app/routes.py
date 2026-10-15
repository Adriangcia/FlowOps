# Rutas (endpoints) de la aplicación.
# "/" muestra el buscador único. Lo que se escribe (parámetro "q" por GET)
# se busca por este orden: oferta -> artículo -> pedido de compra.
# Se queda con el primero que dé resultados.

from flask import Blueprint, render_template, request
from .excel_reader import (
    buscar_oferta,
    buscar_articulo,
    buscar_pedido_compra,
    MIN_CARACTERES_PEDIDO,
)

main = Blueprint("main", __name__)


@main.route("/")
def index():
    # request.args.get() lee parámetros de la URL, ej: /?q=01OFE...
    # Si no hay parámetro, devuelve cadena vacía.
    consulta = request.args.get("q", "").strip()

    resultado = None            # resultado si la consulta es una oferta
    resultado_articulo = None   # resultado si la consulta es un artículo
    resultado_pedido = None     # resultado si la consulta es un pedido de compra
    no_encontrado = False
    error = None

    if consulta:
        try:
            # 1º intentamos como oferta (Ref.Documento)
            resultado = buscar_oferta(consulta)

            # 2º si no es una oferta, intentamos como artículo
            if resultado is None:
                resultado_articulo = buscar_articulo(consulta)

            # 3º si tampoco es un artículo, intentamos como pedido de compra
            if resultado is None and resultado_articulo is None:
                resultado_pedido = buscar_pedido_compra(consulta)
                no_encontrado = resultado_pedido is None
        except FileNotFoundError as e:
            # Falta el Excel diario en la carpeta Excels/
            error = f"{e}. Añade el fichero del día e inténtalo de nuevo."

    return render_template(
        "index.html",
        consulta=consulta,
        resultado=resultado,
        resultado_articulo=resultado_articulo,
        resultado_pedido=resultado_pedido,
        no_encontrado=no_encontrado,
        min_caracteres_pedido=MIN_CARACTERES_PEDIDO,
        error=error,
    )