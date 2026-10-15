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

# Búsqueda por Pedido de compra: nº mínimo de caracteres que hay que escribir
# (para no devolver medio Excel con textos como "oca" o "20") y nº máximo de
# pedidos de compra que se muestran de una vez.
MIN_CARACTERES_PEDIDO = 4
MAX_PEDIDOS_MOSTRADOS = 20

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


# Caché de Excels ya leídos: {(ruta, hoja): {"firma": (...), "df": DataFrame}}
# Leer un .xlsx grande tarda varios segundos, así que se lee una sola vez y
# se reutiliza hasta que el fichero cambia (nueva fecha de modificación o
# tamaño distinto, por ejemplo al sobreescribirlo con el de cada mañana).
_CACHE_EXCELS = {}


def _leer_excel_con_cache(ruta, hoja):
    """
    Lee una hoja de un Excel, reutilizando la lectura anterior si el
    fichero no ha cambiado desde entonces.
    IMPORTANTE: devuelve el DataFrame compartido de la caché, así que quien
    lo use NO debe modificarlo (si necesita añadir columnas, que haga .copy()).
    """
    info = ruta.stat()
    firma = (info.st_mtime, info.st_size)
    clave = (str(ruta), hoja)

    entrada = _CACHE_EXCELS.get(clave)
    if entrada is not None and entrada["firma"] == firma:
        return entrada["df"]

    df = pd.read_excel(ruta, sheet_name=hoja)
    _CACHE_EXCELS[clave] = {"firma": firma, "df": df}
    return df


def cargar_datos():
    """
    Carga el Excel diario y devuelve un DataFrame con la hoja 'DIA'.
    """
    ruta = obtener_ruta_excel()
    return _leer_excel_con_cache(ruta, "DIA")


def cargar_stock():
    """
    Carga el Excel de stock y devuelve un DataFrame con la hoja
    'GET_PRODUCTOS_4MAIL', quedándonos solo con las columnas relevantes.
    """
    ruta = obtener_ruta_stock()
    df = _leer_excel_con_cache(ruta, "GET_PRODUCTOS_4MAIL")
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


def _formatear_numero(numero):
    """
    Para mostrar en el front: si el número es entero (ej. 3.0) lo devuelve
    como int (3), para que no aparezca como "3.0".
    """
    if isinstance(numero, float) and numero.is_integer():
        return int(numero)

    return numero


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


def _primer_valor_por_oferta(df_filas, columna):
    """
    Devuelve {oferta: valor} a partir de un conjunto de filas, tomando para
    cada oferta el primer valor informado de la columna indicada (y no el de
    la primera fila, porque alguna línea puede tenerlo vacío).
    Las ofertas sin ningún valor informado no aparecen en el diccionario.
    """
    valores = {}
    for oferta, valor in zip(df_filas["Ref.Documento"], df_filas[columna]):
        valor = limpiar_valor(valor)
        if valor != "-" and oferta not in valores:
            valores[oferta] = str(valor).strip()

    return valores


def _cliente_final_por_oferta(df_filas):
    """
    Devuelve {oferta: cliente final} (columna "Cliente final").
    """
    return _primer_valor_por_oferta(df_filas, "Cliente final")


def _construir_mapa_stock():
    """
    Carga el Excel de stock y lo convierte en un diccionario
    {(articulo, propietario): stock} para consultas rápidas O(1)
    al recorrer los artículos.
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


def _cargar_mapa_stock_o_none():
    """
    Intenta construir el mapa de stock. Si no existe el Excel de stock,
    devuelve None para seguir funcionando sin esa información en vez de
    romper toda la búsqueda.
    """
    try:
        return _construir_mapa_stock()
    except FileNotFoundError:
        return None


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
        articulo["stock"] = _formatear_numero(mapa_stock.get(clave, 0.0))

    return articulo


def buscar_oferta(ref_documento):
    """
    Busca todas las líneas del Excel correspondientes a una Oferta
    (campo Ref.Documento) y devuelve los datos estructurados:
    - oferta: referencia de la oferta tal cual aparece en el Excel
    - cliente: "Nombre del cliente" del Excel
    - cliente_final: valor de la columna "Cliente final" ("-" si está vacío)
    - pdv: Nº de Pedido de venta, formateado
    - pedidos_compra: lista de pedidos de compra únicos, cada uno con sus
      artículos (destino, Pick Order, Albarán, warnings y stock disponible)
    - articulos: lista plana de TODOS los artículos de la oferta
      (incluye los que no tienen pedido de compra, ej: servicios propios)
    - comprobacion_unidades: por cada Artículo de la oferta, unidades
      ofertadas vs compradas, y si hay discrepancia (warning)

    La búsqueda es exacta pero sin distinguir mayúsculas/minúsculas ni
    espacios al principio/final. Si no hay coincidencias, devuelve None.
    """
    df = cargar_datos()

    buscado = str(ref_documento).strip().lower()
    if not buscado:
        return None

    refs_norm = df["Ref.Documento"].fillna("").astype(str).str.strip().str.lower()
    resultado = df[refs_norm == buscado]

    if resultado.empty:
        return None

    oferta = limpiar_valor(resultado.iloc[0]["Ref.Documento"])
    cliente = next(iter(_primer_valor_por_oferta(resultado, "Nombre del cliente").values()), "-")
    cliente_final = next(iter(_cliente_final_por_oferta(resultado).values()), "-")
    pdv = formatear_pdv(resultado.iloc[0]["NºPedido"])

    # Cargamos el mapa de stock una sola vez para toda la búsqueda
    mapa_stock = _cargar_mapa_stock_o_none()

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
            "ofertadas": _formatear_numero(ofertadas),
            "compradas": _formatear_numero(compradas),
            # La comparación se hace con los números reales, antes de formatear
            "warning": bool(ofertadas != compradas),
        })

    return {
        "oferta": oferta,
        "cliente": cliente,
        "cliente_final": cliente_final,
        "pdv": pdv,
        "pedidos_compra": pedidos_compra,
        "articulos": articulos,
        "comprobacion_unidades": comprobacion_unidades,
    }


def buscar_articulo(codigo_articulo):
    """
    Busca todas las líneas del Excel diario que contienen un Artículo y
    devuelve los datos estructurados para mostrar "desde el artículo":
    - articulo: código del artículo tal cual aparece en el Excel
    - descripcion: descripción de la primera línea encontrada
    - total_ofertas: nº de ofertas distintas que contienen el artículo
    - lineas: una entrada por línea, con Oferta, Cliente, Cliente final,
      PDV, Pedido de compra, Unidades, Destino y Stock en ese destino
      (cruzado por Artículo + Propietario)

    La búsqueda es exacta pero sin distinguir mayúsculas/minúsculas ni
    espacios al principio/final. Si no hay coincidencias, devuelve None.
    """
    df = cargar_datos()

    buscado = str(codigo_articulo).strip().lower()
    if buscado in ("", "-"):
        return None

    # Normalizamos la columna Artículo a texto (hay partnumbers numéricos)
    # para poder compararla con lo que se escribe en el buscador.
    articulos_norm = df["Artículo"].apply(limpiar_articulo).str.strip().str.lower()
    resultado = df[articulos_norm == buscado]

    if resultado.empty:
        return None

    mapa_stock = _cargar_mapa_stock_o_none()

    # Cliente y Cliente final de cada oferta implicada (mirando TODAS las
    # líneas de esa oferta, no solo las del artículo buscado)
    ofertas_implicadas = df[df["Ref.Documento"].isin(resultado["Ref.Documento"])]
    clientes = _primer_valor_por_oferta(ofertas_implicadas, "Nombre del cliente")
    clientes_finales = _cliente_final_por_oferta(ofertas_implicadas)

    lineas = []
    for _, fila in resultado.iterrows():
        pedido_compra = limpiar_valor(fila["Pedidos de compra"])
        tiene_pedido = pedido_compra != "-"

        linea = _fila_a_articulo(fila, incluir_destino=True, mapa_stock=mapa_stock)
        # Si no hay Excel de stock, el campo no existe: lo dejamos en None
        # para que el front muestre "-" en vez de una celda vacía.
        linea.setdefault("stock", None)
        linea["oferta"] = limpiar_valor(fila["Ref.Documento"])
        linea["cliente"] = clientes.get(fila["Ref.Documento"], "-")
        linea["cliente_final"] = clientes_finales.get(fila["Ref.Documento"], "-")
        linea["pdv"] = formatear_pdv(fila["NºPedido"])
        linea["pedido_compra"] = pedido_compra

        # Las líneas sin pedido de compra son servicios propios: no tienen
        # destino ni stock que mostrar (se verán como "-" en el front).
        if not tiene_pedido:
            linea["destino"] = None
            linea["stock"] = None

        lineas.append(linea)

    # Ordenamos por oferta y, dentro de cada oferta, por pedido de compra
    lineas.sort(key=lambda l: (str(l["oferta"]), str(l["pedido_compra"])))

    return {
        "articulo": lineas[0]["articulo"],
        "descripcion": lineas[0]["descripcion"],
        "total_ofertas": len({l["oferta"] for l in lineas}),
        "lineas": lineas,
    }


def buscar_pedido_compra(texto):
    """
    Busca los Pedidos de compra que contienen el texto escrito y, para cada
    uno, a qué ofertas (y cliente final) pertenece. Pensado para cuando te
    pasan un albarán con el pedido de compra y hay que saber de qué cliente es.

    En el Excel el pedido de compra viene como "OCA:2026/D005290 - Proveedor".
    La búsqueda no distingue mayúsculas/minúsculas y basta con escribir una
    parte (ej: "D005290" o "2026/D005290"). Si se pega el texto completo
    (con " - Proveedor"), solo se usa la parte del código.

    Devuelve None si no hay coincidencias o se escriben menos de
    MIN_CARACTERES_PEDIDO caracteres. Si no, un diccionario con:
    - total_pedidos: nº de pedidos de compra que coinciden
    - truncado: True si hay más pedidos que MAX_PEDIDOS_MOSTRADOS
    - pedidos: lista (como mucho MAX_PEDIDOS_MOSTRADOS) con nombre,
      total_ofertas y lineas (Oferta, Cliente, Cliente final, PDV,
      Artículo, Unidades y Destino de cada línea)

    "Cliente" es el "Nombre del cliente" del Excel (siempre informado);
    "Cliente final" solo lo está en una parte de las ofertas.
    """
    df = cargar_datos()

    buscado = str(texto).strip().lower()
    # Si pegan "OCA:2026/D005290 - PROVEEDOR", nos quedamos con el código
    buscado = buscado.split(" - ", 1)[0].strip()

    if len(buscado) < MIN_CARACTERES_PEDIDO:
        return None

    pedidos_limpios = df["Pedidos de compra"].apply(limpiar_valor)
    pedidos_norm = pedidos_limpios.astype(str).str.strip().str.lower()
    coincide = pedidos_norm.str.contains(buscado, regex=False) & (pedidos_limpios != "-")

    resultado = df[coincide].copy()

    if resultado.empty:
        return None

    # Texto completo del pedido de compra y su código (lo que va antes de " - ").
    # Agrupamos por código porque el nombre del proveedor puede variar un poco
    # entre líneas del mismo pedido.
    resultado["_pedido"] = pedidos_limpios[coincide].astype(str).str.strip()
    resultado["_codigo"] = resultado["_pedido"].str.split(" - ", n=1).str[0].str.strip()

    # Cliente y Cliente final de cada oferta implicada (mirando TODAS las
    # líneas de esa oferta, no solo las de este pedido de compra)
    ofertas_implicadas = df[df["Ref.Documento"].isin(resultado["Ref.Documento"])]
    clientes = _primer_valor_por_oferta(ofertas_implicadas, "Nombre del cliente")
    clientes_finales = _cliente_final_por_oferta(ofertas_implicadas)

    # Pedidos de compra que coinciden, ordenados por código. Solo
    # construimos el detalle de los que se van a mostrar.
    grupos = sorted(resultado.groupby("_codigo"))

    pedidos = []
    for _codigo, grupo in grupos[:MAX_PEDIDOS_MOSTRADOS]:
        lineas = []
        for _, fila in grupo.iterrows():
            lineas.append({
                "oferta": limpiar_valor(fila["Ref.Documento"]),
                "cliente": clientes.get(fila["Ref.Documento"], "-"),
                "cliente_final": clientes_finales.get(fila["Ref.Documento"], "-"),
                "pdv": formatear_pdv(fila["NºPedido"]),
                "articulo": limpiar_articulo(fila["Artículo"]),
                "unidades": limpiar_valor(fila["Unidades"]),
                "destino": traducir_destino(fila["Propietario"]),
            })

        # Ordenamos por oferta y, dentro de cada oferta, por artículo
        lineas.sort(key=lambda l: (str(l["oferta"]), str(l["articulo"])))

        pedidos.append({
            "nombre": grupo["_pedido"].iloc[0],
            "total_ofertas": len({l["oferta"] for l in lineas}),
            "lineas": lineas,
        })

    return {
        "total_pedidos": len(grupos),
        "truncado": len(grupos) > MAX_PEDIDOS_MOSTRADOS,
        "pedidos": pedidos,
    }