# Rutas (endpoints) de la aplicación.
# "/" muestra dos buscadores:
# - Buscador principal (parámetro "q"): se busca por este orden
#   oferta -> artículo -> pedido de compra, y nos quedamos con el primero
#   que dé resultados.
# - Consulta de stock (parámetro "stock"): stock de un artículo según el
#   Excel de stock.

from flask import Blueprint, render_template, request
from .excel_reader import (
    buscar_oferta,
    buscar_articulo,
    buscar_pedido_compra,
    buscar_stock,
    info_fichero_stock,
    MIN_CARACTERES_PEDIDO,
    MIN_CARACTERES_STOCK,
)

main = Blueprint("main", __name__)


@main.route("/")
def index():
    # request.args.get() lee parámetros de la URL, ej: /?q=01OFE...
    # Si no hay parámetro, devuelve cadena vacía.
    consulta = request.args.get("q", "").strip()
    consulta_stock = request.args.get("stock", "").strip()

    resultado = None            # resultado si la consulta es una oferta
    resultado_articulo = None   # resultado si la consulta es un artículo
    resultado_pedido = None     # resultado si la consulta es un pedido de compra
    resultado_stock = None      # resultado de la consulta de stock
    info_stock = None           # fichero de stock que se ha leído (nombre, fecha, filas)
    no_encontrado = False
    stock_no_encontrado = False
    error = None

    try:
        if consulta:
            # 1º intentamos como oferta (Ref.Documento)
            resultado = buscar_oferta(consulta)

            # 2º si no es una oferta, intentamos como artículo
            if resultado is None:
                resultado_articulo = buscar_articulo(consulta)

            # 3º si tampoco es un artículo, intentamos como pedido de compra
            if resultado is None and resultado_articulo is None:
                resultado_pedido = buscar_pedido_compra(consulta)
                no_encontrado = resultado_pedido is None

        elif consulta_stock:
            # Consulta de stock: el artículo puede no estar en el Excel de
            # stock, y eso significa que no hay stock (no es un error)
            resultado_stock = buscar_stock(consulta_stock)
            stock_no_encontrado = resultado_stock is None
            info_stock = info_fichero_stock()

    except FileNotFoundError as e:
        # Falta algún Excel necesario en la carpeta Excels/
        error = f"{e}. Añade el fichero del día e inténtalo de nuevo."

    return render_template(
        "index.html",
        consulta=consulta,
        consulta_stock=consulta_stock,
        resultado=resultado,
        resultado_articulo=resultado_articulo,
        resultado_pedido=resultado_pedido,
        resultado_stock=resultado_stock,
        info_stock=info_stock,
        no_encontrado=no_encontrado,
        stock_no_encontrado=stock_no_encontrado,
        min_caracteres_pedido=MIN_CARACTERES_PEDIDO,
        min_caracteres_stock=MIN_CARACTERES_STOCK,
        error=error,
    )