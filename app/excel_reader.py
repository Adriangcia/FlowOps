# Módulo encargado de leer y procesar los Excels de FlowOps:
# - El Excel diario (PendienteServir): ofertas, pedidos, artículos.
# - El Excel de stock (stockDisponible.xlsx): stock por artículo y propietario.
# Toda la lógica de lectura, limpieza y cruce de datos vive aquí.

import re
import unicodedata
from datetime import datetime

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

# Consulta de stock: con una sola palabra hay que escribir al menos
# MIN_CARACTERES_STOCK caracteres, y se muestran como mucho
# MAX_ARTICULOS_STOCK artículos de una vez.
MIN_CARACTERES_STOCK = 4
MAX_ARTICULOS_STOCK = 20

# Palabras que se ignoran al buscar por descripción (artículos, preposiciones...)
PALABRAS_VACIAS = {
    "a", "al", "con", "de", "del", "el", "en", "la", "las", "lo", "los",
    "o", "para", "por", "un", "una", "y",
}

# La columna "STOCK PROPIETARIOS" del Excel de stock es un texto del tipo
# "TCN:TC-3 (CAVERIN): 1, GIJ:BUFFER_DAIMLER-GIJ: 3, GIJ:2M.2M GIJ: 3".
# Este patrón saca cada pareja "propietario: unidades". Se busca el
# ": <número>" seguido de ", " o del final para no romper nombres que
# lleven comas o dos puntos dentro.
PATRON_STOCK_PROPIETARIOS = re.compile(r"(.+?): (-?\d+(?:[.,]\d+)?)(?:, |$)")

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


def cargar_stock_completo():
    """
    Carga el Excel de stock con TODAS las columnas (hoja 'GET_PRODUCTOS_4MAIL').
    """
    ruta = obtener_ruta_stock()
    return _leer_excel_con_cache(ruta, "GET_PRODUCTOS_4MAIL")


def cargar_stock():
    """
    Devuelve el Excel de stock solo con las columnas que usa el cruce de
    stock de los pedidos de compra: ARTICULO, NOMBRE PROPIETARIO y STOCK.
    """
    return cargar_stock_completo()[["ARTICULO", "NOMBRE PROPIETARIO", "STOCK"]]


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


def _texto_referencia(valor):
    """
    Pasa una referencia (ej: Ref.Pedido Cliente) a texto para mostrarla:
    - vacío / NaN -> "-"
    - número entero guardado como decimal (4500641581.0) -> "4500641581"
    - cualquier otro valor -> texto sin espacios sobrantes
    """
    valor = limpiar_valor(valor)

    if valor == "-":
        return "-"

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor).strip()


def _referencias_cliente(df_filas):
    """
    Devuelve la "Ref.Pedido Cliente" de un conjunto de filas (las de una
    oferta). Lo normal es un único valor, pero si la oferta tiene varios
    distintos se devuelven todos separados por coma. "-" si no hay ninguno.
    """
    referencias = []
    for valor in df_filas["Ref.Pedido Cliente"]:
        referencia = _texto_referencia(valor)
        if referencia != "-" and referencia not in referencias:
            referencias.append(referencia)

    return ", ".join(referencias) if referencias else "-"


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
    - ref_pedido_cliente: "Ref.Pedido Cliente" del Excel ("-" si está vacía)
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
    ref_pedido_cliente = _referencias_cliente(resultado)
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
        "ref_pedido_cliente": ref_pedido_cliente,
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


def _parsear_stock_propietarios(texto):
    """
    Convierte el texto de la columna "STOCK PROPIETARIOS"
    (ej: "TCN:TC-3 (CAVERIN): 1, GIJ:2M.2M GIJ: 3") en una lista de
    {"propietario": ..., "unidades": ...}. Si el texto no tiene el formato
    esperado, devuelve una lista vacía.
    """
    propietarios = []
    for coincidencia in PATRON_STOCK_PROPIETARIOS.finditer(str(texto)):
        unidades = float(coincidencia.group(2).replace(",", "."))
        propietarios.append({
            "propietario": coincidencia.group(1).strip(),
            "unidades": _formatear_numero(unidades),
        })

    return propietarios


def _stock_de_articulo(filas):
    """
    Construye el stock de UN artículo a partir de sus filas del Excel de
    stock (hay una fila por propietario). Devuelve:
    - articulo: código del artículo tal cual aparece en el Excel de stock
    - descripcion: descripción del artículo
    - propietarios: lista de {"propietario", "unidades"} con el stock que
      tiene cada propietario, ordenada de más a menos unidades
    - total: suma de las unidades de todos los propietarios
    - fuente: "stock_propietarios" si los datos salen de la columna
      "STOCK PROPIETARIOS"; "stock" si esa columna estaba vacía y se ha usado
      como respaldo la columna "STOCK" de cada fila
    """
    articulo = str(filas.iloc[0]["ARTICULO"]).strip()
    descripcion = next(
        (str(d).strip() for d in filas["DESCRIPCION"] if limpiar_valor(d) != "-"),
        "-",
    )

    # 1) Lo que dice la columna STOCK PROPIETARIOS (es el mismo texto en todas
    #    las filas de un mismo artículo, así que usamos el primero informado)
    textos = [str(t) for t in filas["STOCK PROPIETARIOS"] if limpiar_valor(t) != "-"]
    propietarios = _parsear_stock_propietarios(textos[0]) if textos else []
    fuente = "stock_propietarios"

    # 2) Respaldo: si esa columna está vacía (o no se entiende), usamos la
    #    columna STOCK de cada fila, que trae el stock de su propietario
    if not propietarios:
        fuente = "stock"
        acumulado = {}
        for nombre, stock in zip(filas["NOMBRE PROPIETARIO"], filas["STOCK"]):
            unidades = _a_numero(stock)
            if unidades > 0 and limpiar_valor(nombre) != "-":
                nombre = str(nombre).strip()
                acumulado[nombre] = acumulado.get(nombre, 0.0) + unidades

        propietarios = [
            {"propietario": nombre, "unidades": _formatear_numero(unidades)}
            for nombre, unidades in acumulado.items()
        ]

    # Más stock primero y, si empatan, por nombre
    propietarios.sort(key=lambda p: (-_a_numero(p["unidades"]), p["propietario"]))

    total = _formatear_numero(sum(_a_numero(p["unidades"]) for p in propietarios))

    return {
        "articulo": articulo,
        "descripcion": descripcion,
        "propietarios": propietarios,
        "total": total,
        "fuente": fuente,
    }


def _normalizar_busqueda(texto):
    """
    Prepara un texto para compararlo en las búsquedas por palabras:
    minúsculas, sin acentos y sin guiones (así "RJ-45" y "RJ45" coinciden).
    """
    texto = unicodedata.normalize("NFKD", str(texto).lower())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.replace("-", "").replace("_", "")


def _palabras_busqueda(texto):
    """
    Trocea lo escrito en palabras para buscar por descripción. Devuelve una
    lista de (palabra_para_mostrar, palabra_para_buscar), sin repetidas, sin
    palabras vacías ("a", "de", "para"...) y sin letras sueltas. Se mantiene
    el orden escrito: la primera es la palabra principal.
    """
    palabras = []
    vistas = set()
    for trozo in str(texto).split():
        mostrar = trozo.strip(",;:()[]{}\"'").lower()
        buscar = _normalizar_busqueda(mostrar)

        if len(buscar) < 2 or mostrar in PALABRAS_VACIAS or buscar in vistas:
            continue

        vistas.add(buscar)
        palabras.append((mostrar, buscar))

    return palabras


def _niveles_coincidencia(n):
    """
    Niveles de coincidencia, del mejor al peor, para una búsqueda de n palabras
    (la primera es la principal). Cada nivel es (tipo, minimo, cumple), y
    cumple(indices) dice si entra un artículo que contiene esas palabras:
    1) "todas": contiene las n palabras.
    2) "parcial": contiene la principal y al menos "minimo" palabras en total
       (se prueba con n-1, n-2... hasta 2).
    3) "principal": contiene solo la palabra principal.
    """
    niveles = [("todas", n, lambda indices: len(indices) == n)]

    for minimo in range(n - 1, 1, -1):
        niveles.append((
            "parcial",
            minimo,
            lambda indices, m=minimo: 0 in indices and len(indices) >= m,
        ))

    if n >= 2:
        niveles.append(("principal", 1, lambda indices: 0 in indices))

    return niveles


def buscar_stock(texto):
    """
    Consulta SOLO el Excel de stock (no mira el Excel diario) y devuelve el
    stock por propietario de los artículos que coinciden con lo escrito:
    1) Si hay un artículo con ese código exacto (sin distinguir mayúsculas
       ni espacios), se devuelve solo ese.
    2) Si no, se busca por palabras en el código y en la descripción del
       artículo (sin distinguir mayúsculas, acentos ni guiones). La primera
       palabra es la principal. Se prueba por niveles y nos quedamos con el
       primero que dé resultados:
         - todas las palabras;
         - la principal y al menos otras (n-1, n-2... hasta 2 en total);
         - solo la principal.
       Con una sola palabra hay que escribir al menos MIN_CARACTERES_STOCK
       caracteres.
       Dentro del nivel, primero los artículos con más stock.

    Devuelve None si no hay ninguna coincidencia (se interpreta como stock 0).
    Si no, un diccionario con:
    - aproximada: True si no era un código exacto (caso 2)
    - palabras: las palabras buscadas, tal como se escribieron
    - nivel: "todas", "parcial" o "principal" (None si era código exacto)
    - minimas: nº mínimo de palabras del nivel (solo útil en "parcial")
    - total_articulos: nº de artículos que coinciden
    - truncado: True si hay más artículos que MAX_ARTICULOS_STOCK
    - articulos: lista (como mucho MAX_ARTICULOS_STOCK) con el stock de cada
      artículo (ver _stock_de_articulo)

    Si no existe el fichero stockDisponible.xlsx, lanza FileNotFoundError.
    """
    df = cargar_stock_completo()

    escrito = str(texto).strip().lower()
    if escrito in ("", "-"):
        return None

    # Una columna auxiliar con el código normalizado (sin tocar el DataFrame
    # original, que es el de la caché)
    df = df.assign(_art=df["ARTICULO"].astype(str).str.strip().str.lower())

    # 1) Código exacto
    exactas = df[df["_art"] == escrito]
    if not exactas.empty:
        return {
            "aproximada": False,
            "palabras": [],
            "nivel": None,
            "minimas": None,
            "total_articulos": 1,
            "truncado": False,
            "articulos": [_stock_de_articulo(exactas)],
        }

    # 2) Búsqueda por palabras en el código y la descripción
    palabras = _palabras_busqueda(escrito)
    if not palabras:
        return None

    if len(palabras) == 1 and len(palabras[0][1]) < MIN_CARACTERES_STOCK:
        return None

    buscar = [palabra_buscar for _, palabra_buscar in palabras]

    # Para cada artículo, qué palabras de la búsqueda contiene
    candidatos = []
    for _, filas in df.groupby("_art"):
        codigo = str(filas.iloc[0]["ARTICULO"]).strip()
        descripcion = next(
            (str(d) for d in filas["DESCRIPCION"] if limpiar_valor(d) != "-"),
            "",
        )
        pajar = _normalizar_busqueda(f"{codigo} {descripcion}")

        indices = {i for i, palabra in enumerate(buscar) if palabra in pajar}
        if indices:
            candidatos.append((filas, indices))

    # Nos quedamos con el primer nivel que tenga algún artículo
    for tipo, minimo, cumple in _niveles_coincidencia(len(buscar)):
        elegidos = [filas for filas, indices in candidatos if cumple(indices)]
        if elegidos:
            break
    else:
        return None

    # Primero los que tienen más stock y, si empatan, por código
    articulos = [_stock_de_articulo(filas) for filas in elegidos]
    articulos.sort(key=lambda a: (-_a_numero(a["total"]), a["articulo"].lower()))

    return {
        "aproximada": True,
        "palabras": [palabra_mostrar for palabra_mostrar, _ in palabras],
        "nivel": tipo,
        "minimas": minimo,
        "total_articulos": len(articulos),
        "truncado": len(articulos) > MAX_ARTICULOS_STOCK,
        "articulos": articulos[:MAX_ARTICULOS_STOCK],
    }


def info_fichero_stock():
    """
    Datos del Excel de stock que se está leyendo (para mostrarlos en
    pantalla y poder comprobar que es el fichero actualizado):
    nombre, fecha de última modificación y nº de filas.
    Si no existe el fichero, lanza FileNotFoundError.
    """
    ruta = obtener_ruta_stock()
    modificado = datetime.fromtimestamp(ruta.stat().st_mtime)

    return {
        "nombre": ruta.name,
        "modificado": modificado.strftime("%d/%m/%Y %H:%M"),
        "filas": len(cargar_stock_completo()),
    }