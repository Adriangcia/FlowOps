# Módulo encargado de leer y procesar los Excels de FlowOps:
# - El Excel diario (PendienteServir): ofertas, pedidos, artículos.
# - El Excel de stock (stockDisponible.xlsx): stock por artículo y propietario.
# Toda la lógica de lectura, limpieza y cruce de datos vive aquí.

import pandas as pd
from pathlib import Path

# Ruta a la carpeta Excels/, calculada de forma relativa al propio fichero.
EXCELS_DIR = Path(__file__).resolve().parent.parent / "Excels"

# Nombre fijo del fichero de stock (se sobreescribe en cada descarga).
NOMBRE_FICHERO_STOCK = "stockDisponible.xlsx"

# Traducción de los valores de la columna "Propietario" a destinos de envío
# legibles. Si un valor no está en este diccionario, no se muestra destino
# (normalmente es un proveedor de servicios externo, no un destino de envío).
TRADUCCION_DESTINOS = {
    "TCN:TC-1 (CAVERIN)": "TCNET",
    "GIJ:EPS-1 (GIJON)": "Gijón",
    "TC-9 (CAVERIN DIRECTOS)": "Directo Cliente",
}


def obtener_ruta_excel():
    """
    Busca el Excel diario (PendienteServir) dentro de Excels/.
    Es el único .xlsx que NO se llama stockDisponible.xlsx
    (ese nombre está reservado para el Excel de stock).
    """
    ficheros = [
        f for f in EXCELS_DIR.glob("*.xlsx")
        if f.name != NOMBRE_FICHERO_STOCK
    ]

    if not ficheros:
        raise FileNotFoundError(
            "No se ha encontrado el Excel diario (PendienteServir) en Excels/"
        )

    return ficheros[0]


def obtener_ruta_stock():
    """
    Localiza el Excel de stock, que siempre se llama stockDisponible.xlsx
    (se sobreescribe en cada descarga, así que el nombre es fijo).
    """
    ruta = EXCELS_DIR / NOMBRE_FICHERO_STOCK

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se ha encontrado {NOMBRE_FICHERO_STOCK} en Excels/"
        )

    return ruta


def cargar_datos():
    """
    Carga el Excel diario y devuelve un DataFrame con la hoja 'DIA'.
    """
    ruta = obtener_ruta_excel()
    df = pd.read_excel(ruta, sheet_name="DIA")
    return df


def cargar_stock():
    """
    Carga el Excel de stock y devuelve un DataFrame con la hoja
    'GET_PRODUCTOS_4MAIL', quedándonos solo con las columnas relevantes.
    """
    ruta = obtener_ruta_stock()
    df = pd.read_excel(ruta, sheet_name="GET_PRODUCTOS_4MAIL")
    return df[["ARTICULO", "NOMBRE PROPIETARIO", "STOCK"]]


def limpiar_valor(valor):
    """
    Normaliza un valor de celda para mostrarlo en el front:
    - Si es un valor nulo (NaN) -> "-"
    - Si es una cadena vacía o solo espacios -> "-"
    - En cualquier otro caso, se devuelve el valor tal cual.
    """
    if pd.isna(valor):
        return "-"

    if isinstance(valor, str) and valor.strip() == "":
        return "-"

    return valor


def formatear_pdv(numero_pedido):
    """
    Convierte el formato del NºPedido (PDV) de "2026/013340"
    a "2026-013340" para mostrarlo en el front.
    """
    if pd.isna(numero_pedido):
        return "-"

    return str(numero_pedido).replace("/", "-")


def limpiar_articulo(valor):
    """
    Normaliza el código de Artículo forzándolo siempre a texto.
    Algunos partnumbers son puramente numéricos y pandas los interpreta
    como int/float; los convertimos a string para que sean consistentes
    con los códigos alfanuméricos (ej: 'ININ-OS').
    """
    if pd.isna(valor):
        return "-"

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor)


def traducir_destino(propietario):
    """
    Traduce el valor de la columna "Propietario" a un destino de envío
    legible, usando TRADUCCION_DESTINOS. Si el valor no está mapeado
    (normalmente un proveedor de servicios externo), devuelve None
    para que el front no muestre nada.
    """
    if pd.isna(propietario):
        return None

    return TRADUCCION_DESTINOS.get(str(propietario).strip())


def _a_numero(valor):
    """
    Convierte un valor a número (float) para poder sumarlo y compararlo.
    Si no es un número válido o está vacío, devuelve 0.0
    """
    if pd.isna(valor):
        return 0.0
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def _hay_warning_preparacion(unidades, pick_order, albaran):
    """
    Determina si una línea de un pedido de compra tiene una alerta de
    preparación: Uds.Pick Order y/o Uds.Albarán no coinciden con las
    Unidades compradas.

    Si Pick Order o Albarán están vacíos ("-"), se considera que el
    pedido todavía no ha llegado a esa fase del proceso y NO es una
    alerta real.
    """
    unidades_num = _a_numero(unidades)

    if pick_order != "-":
        if _a_numero(pick_order) != unidades_num:
            return True

    if albaran != "-":
        if _a_numero(albaran) != unidades_num:
            return True

    return False


def _construir_mapa_stock():
    """
    Carga el Excel de stock y lo convierte en un diccionario
    {(articulo, propietario): stock} para consultas rápidas O(1)
    al recorrer los artículos de una oferta.
    """
    df_stock = cargar_stock()

    mapa = {}
    for _, fila in df_stock.iterrows():
        clave = (
            limpiar_articulo(fila["ARTICULO"]),
            str(fila["NOMBRE PROPIETARIO"]).strip() if pd.notna(fila["NOMBRE PROPIETARIO"]) else "",
        )
        mapa[clave] = _a_numero(fila["STOCK"])

    return mapa


def _fila_a_articulo(fila, incluir_destino=False, incluir_preparacion=False, mapa_stock=None):
    """
    Construye el diccionario de un artículo a partir de una fila del Excel.
    - incluir_destino: añade el campo "destino" (traducido de Propietario)
    - incluir_preparacion: añade Uds.Pick Order, Uds.Albarán y si hay warning
    - mapa_stock: si se pasa, añade el campo "stock" cruzando por
      Artículo + Propietario (valor ORIGINAL, sin traducir). Si la
      combinación no aparece en el Excel de stock, se asume stock = 0.
    """
    articulo = {
        "articulo": limpiar_articulo(fila["Artículo"]),
        "descripcion": limpiar_valor(fila["Descripción"]),
        "unidades": limpiar_valor(fila["Unidades"]),
    }

    if incluir_destino:
        articulo["destino"] = traducir_destino(fila["Propietario"])

    if incluir_preparacion:
        pick_order = limpiar_valor(fila["Uds.Pick Order"])
        albaran = limpiar_valor(fila["Uds.Albarán"])

        articulo["pick_order"] = pick_order
        articulo["albaran"] = albaran
        articulo["warning_preparacion"] = _hay_warning_preparacion(
            fila["Unidades"], pick_order, albaran
        )

    if mapa_stock is not None:
        propietario_original = (
            str(fila["Propietario"]).strip() if pd.notna(fila["Propietario"]) else ""
        )
        clave = (articulo["articulo"], propietario_original)
        articulo["stock"] = mapa_stock.get(clave, 0.0)

    return articulo


def buscar_oferta(ref_documento):
    """
    Busca todas las líneas del Excel correspondientes a una Oferta
    (campo Ref.Documento) y devuelve los datos estructurados:
    - pdv: Nº de Pedido de venta, formateado
    - pedidos_compra: lista de pedidos de compra únicos, cada uno con sus
      artículos (destino, Pick Order, Albarán, warnings y stock disponible)
    - articulos: lista plana de TODOS los artículos de la oferta
      (incluye los que no tienen pedido de compra, ej: servicios propios)
    - comprobacion_unidades: por cada Artículo de la oferta, unidades
      ofertadas vs compradas, y si hay discrepancia (warning)
    """
    df = cargar_datos()

    resultado = df[df["Ref.Documento"] == ref_documento]

    if resultado.empty:
        return None

    pdv = formatear_pdv(resultado.iloc[0]["NºPedido"])

    # Cargamos el mapa de stock una sola vez para toda la búsqueda.
    # Si no existe el Excel de stock, seguimos sin esa información
    # en vez de romper toda la búsqueda.
    try:
        mapa_stock = _construir_mapa_stock()
    except FileNotFoundError:
        mapa_stock = None

    # --- Tabla general: todos los artículos de la oferta, sin excepción ---
    articulos = [_fila_a_articulo(fila) for _, fila in resultado.iterrows()]

    # --- Pedidos de compra agrupados, cada uno con sus artículos ---
    resultado = resultado.copy()
    resultado["_pedido_compra_limpio"] = resultado["Pedidos de compra"].apply(limpiar_valor)

    con_pedido = resultado[resultado["_pedido_compra_limpio"] != "-"]

    pedidos_compra = []
    for pedido, grupo in sorted(con_pedido.groupby("_pedido_compra_limpio")):
        pedidos_compra.append({
            "nombre": pedido,
            "articulos": [
                _fila_a_articulo(
                    fila,
                    incluir_destino=True,
                    incluir_preparacion=True,
                    mapa_stock=mapa_stock,
                )
                for _, fila in grupo.iterrows()
            ],
        })

    # --- Comprobación de unidades ofertadas vs compradas, por Artículo ---
    resultado["_unidades_num"] = resultado["Unidades"].apply(_a_numero)

    comprobacion_unidades = []
    for articulo_codigo, grupo in sorted(resultado.groupby(resultado["Artículo"].apply(limpiar_articulo))):
        ofertadas = grupo["_unidades_num"].sum()
        grupo_con_pedido = grupo[grupo["_pedido_compra_limpio"] != "-"]
        compradas = grupo_con_pedido["_unidades_num"].sum()

        comprobacion_unidades.append({
            "articulo": articulo_codigo,
            "ofertadas": ofertadas,
            "compradas": compradas,
            "warning": ofertadas != compradas,
        })

    return {
        "oferta": ref_documento,
        "pdv": pdv,
        "pedidos_compra": pedidos_compra,
        "articulos": articulos,
        "comprobacion_unidades": comprobacion_unidades,
    }