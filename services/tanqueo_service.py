# ============================================================
# INTELLIFEET
# SERVICIO - TANQUEOS / COMBUSTIBLE
# ============================================================

from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from zoneinfo import ZoneInfo
import os
import uuid

from werkzeug.utils import secure_filename

from extensions import db
from models import Tanqueo, Vehiculo, Maquinaria

from services.control_diario_service import (
    obtener_usuario,
    validar_rol_admin,
    obtener_activos_operador,
    validar_activo_operador
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

BOGOTA = ZoneInfo("America/Bogota")

UPLOAD_FOLDER = os.path.join(
    "uploads",
    "tanqueos"
)

ALLOWED_EXTENSIONS = {
    "jpg",
    "jpeg",
    "png",
    "webp",
    "pdf"
}

ESTADOS_VALIDOS = {
    "REGISTRADO",
    "VALIDADO",
    "RECHAZADO"
}

TIPOS_COMBUSTIBLE = (
    "DIESEL",
    "GASOLINA",
    "ACPM",
    "OTRO"
)


# ============================================================
# UTILIDADES
# ============================================================

def archivo_permitido(nombre):
    if not nombre or "." not in nombre:
        return False

    extension = nombre.rsplit(".", 1)[1].lower()
    return extension in ALLOWED_EXTENSIONS


def obtener_extension(nombre):
    if not nombre or "." not in nombre:
        return ""

    return nombre.rsplit(".", 1)[1].lower()


def convertir_decimal(valor, campo, permitir_cero=True):
    """
    Convierte un valor recibido por formulario a Decimal.
    """
    if valor is None or str(valor).strip() == "":
        return None

    try:
        numero = Decimal(str(valor).replace(",", "."))
    except (InvalidOperation, ValueError):
        raise ValueError(
            f"El valor de {campo} no es numérico"
        )

    if not permitir_cero and numero <= 0:
        raise ValueError(
            f"{campo} debe ser mayor a cero"
        )

    if permitir_cero and numero < 0:
        raise ValueError(
            f"{campo} no puede ser negativo"
        )

    return numero


def parsear_fecha(valor):
    """
    Acepta YYYY-MM-DD o un objeto date.
    """
    if isinstance(valor, date):
        return valor

    if not valor:
        raise ValueError(
            "Debe indicar la fecha del tanqueo"
        )

    try:
        return datetime.strptime(
            str(valor)[:10],
            "%Y-%m-%d"
        ).date()
    except ValueError:
        raise ValueError(
            "Formato de fecha inválido. Use YYYY-MM-DD"
        )


# ============================================================
# GUARDAR IMAGEN / SOPORTE
# ============================================================

def guardar_archivo(archivo, fecha):
    """
    Guarda el soporte del tanqueo organizado por año/mes.

    Devuelve la ruta relativa que se almacena en la BD.
    """
    if not archivo or not getattr(archivo, "filename", None):
        raise ValueError(
            "Debe adjuntar la imagen o soporte del tanqueo"
        )

    if not archivo_permitido(archivo.filename):
        raise ValueError(
            "Tipo de archivo no permitido. "
            "Use JPG, JPEG, PNG, WEBP o PDF."
        )

    nombre_original = secure_filename(
        archivo.filename
    )

    if not nombre_original:
        raise ValueError(
            "Nombre de archivo inválido"
        )

    extension = obtener_extension(
        nombre_original
    )

    carpeta = os.path.join(
        UPLOAD_FOLDER,
        str(fecha.year),
        f"{fecha.month:02d}"
    )

    os.makedirs(carpeta, exist_ok=True)

    nombre_archivo = (
        f"tanqueo_{fecha.strftime('%Y%m%d')}_"
        f"{uuid.uuid4().hex}.{extension}"
    )

    ruta_fisica = os.path.join(
        carpeta,
        nombre_archivo
    )

    archivo.save(ruta_fisica)

    ruta_relativa = os.path.join(
        "tanqueos",
        str(fecha.year),
        f"{fecha.month:02d}",
        nombre_archivo
    ).replace("\\", "/")

    return {
        "archivo_path": ruta_relativa,
        "archivo_nombre": nombre_original,
        "archivo_tipo": extension,
        "ruta_fisica": ruta_fisica
    }


def eliminar_archivo(ruta_fisica):
    """
    Limpia el archivo si falla la creación del registro.
    """
    try:
        if ruta_fisica and os.path.exists(ruta_fisica):
            os.remove(ruta_fisica)
    except OSError:
        # No queremos que una limpieza fallida rompa la respuesta.
        pass


# ============================================================
# CREAR TANQUEO
# ============================================================

def crear_tanqueo(
    usuario_id,
    fecha,
    galones,
    vehiculo_id=None,
    maquinaria_id=None,
    precio_galon=None,
    tipo_combustible="DIESEL",
    estacion=None,
    numero_factura=None,
    km_actual=None,
    horometro_actual=None,
    archivo=None,
    observaciones=None
):
    """
    Crea un reporte de tanqueo para un activo asignado al operador.
    """
    usuario = obtener_usuario(usuario_id)

    if not usuario:
        raise ValueError(
            "Usuario no encontrado"
        )

    if not usuario.activo:
        raise ValueError(
            "El usuario no está activo"
        )

    # --------------------------------------------------------
    # ACTIVO ASIGNADO
    # --------------------------------------------------------
    if not vehiculo_id and not maquinaria_id:
        raise ValueError(
            "Debe seleccionar el vehículo o maquinaria tanqueada"
        )

    if vehiculo_id and maquinaria_id:
        raise ValueError(
            "Seleccione un solo activo"
        )

    asignacion = validar_activo_operador(
        usuario_id,
        vehiculo_id=vehiculo_id,
        maquinaria_id=maquinaria_id
    )

    if not asignacion:
        raise ValueError(
            "El activo seleccionado no está asignado a este operador"
        )

    # --------------------------------------------------------
    # VALORES
    # --------------------------------------------------------
    fecha = parsear_fecha(fecha)

    galones = convertir_decimal(
        galones,
        "galones",
        permitir_cero=False
    )

    if galones > Decimal("10000"):
        raise ValueError(
            "La cantidad de galones no es válida"
        )

    precio_galon = convertir_decimal(
        precio_galon,
        "precio por galón"
    )

    km_actual = convertir_decimal(
        km_actual,
        "kilometraje"
    )

    horometro_actual = convertir_decimal(
        horometro_actual,
        "horómetro"
    )

    # --------------------------------------------------------
    # LECTURA AUTOMÁTICA SI NO LA ENVÍA EL OPERADOR
    # --------------------------------------------------------
    if vehiculo_id and km_actual is None:
        vehiculo = Vehiculo.query.get(vehiculo_id)
        if vehiculo and vehiculo.km_actual is not None:
            km_actual = Decimal(str(vehiculo.km_actual))

    if maquinaria_id and horometro_actual is None:
        maquinaria = Maquinaria.query.get(maquinaria_id)
        if (
            maquinaria
            and maquinaria.horometro_actual is not None
        ):
            horometro_actual = Decimal(
                str(maquinaria.horometro_actual)
            )

    valor_total = None
    if precio_galon is not None:
        valor_total = (
            galones * precio_galon
        ).quantize(Decimal("0.01"))

    tipo_combustible = (
        str(tipo_combustible or "DIESEL")
        .strip()
        .upper()
    )

    if tipo_combustible not in TIPOS_COMBUSTIBLE:
        tipo_combustible = "OTRO"

    # --------------------------------------------------------
    # ARCHIVO
    # --------------------------------------------------------
    archivo_info = guardar_archivo(
        archivo,
        fecha
    )

    tanqueo = Tanqueo(
        fecha=fecha,
        usuario_id=usuario_id,
        vehiculo_id=vehiculo_id,
        maquinaria_id=maquinaria_id,
        tipo_combustible=tipo_combustible,
        galones=galones,
        precio_galon=precio_galon,
        valor_total=valor_total,
        estacion=(
            str(estacion).strip()[:150]
            if estacion
            else None
        ),
        numero_factura=(
            str(numero_factura).strip()[:50]
            if numero_factura
            else None
        ),
        km_actual=km_actual,
        horometro_actual=horometro_actual,
        archivo_path=archivo_info["archivo_path"],
        archivo_nombre=archivo_info["archivo_nombre"],
        archivo_tipo=archivo_info["archivo_tipo"],
        estado="REGISTRADO",
        observaciones=observaciones
    )

    try:
        db.session.add(tanqueo)
        db.session.commit()
    except Exception:
        db.session.rollback()
        eliminar_archivo(
            archivo_info["ruta_fisica"]
        )
        raise

    return tanqueo


# ============================================================
# CONSULTAS
# ============================================================

def _aplicar_filtros(
    query,
    fecha_desde=None,
    fecha_hasta=None,
    usuario_id=None,
    vehiculo_id=None,
    maquinaria_id=None,
    estado=None
):
    if fecha_desde:
        query = query.filter(
            Tanqueo.fecha >= parsear_fecha(fecha_desde)
        )

    if fecha_hasta:
        query = query.filter(
            Tanqueo.fecha <= parsear_fecha(fecha_hasta)
        )

    if usuario_id:
        query = query.filter(
            Tanqueo.usuario_id == int(usuario_id)
        )

    if vehiculo_id:
        query = query.filter(
            Tanqueo.vehiculo_id == int(vehiculo_id)
        )

    if maquinaria_id:
        query = query.filter(
            Tanqueo.maquinaria_id == int(maquinaria_id)
        )

    if estado:
        estado = str(estado).strip().upper()
        if estado in ESTADOS_VALIDOS:
            query = query.filter(
                Tanqueo.estado == estado
            )

    return query


def obtener_tanqueos_operador(
    usuario_id,
    fecha_desde=None,
    fecha_hasta=None
):
    query = Tanqueo.query.filter(
        Tanqueo.usuario_id == usuario_id
    )

    query = _aplicar_filtros(
        query,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta
    )

    return query.order_by(
        Tanqueo.fecha.desc(),
        Tanqueo.id.desc()
    ).all()


def obtener_tanqueos_admin(
    fecha_desde=None,
    fecha_hasta=None,
    usuario_id=None,
    vehiculo_id=None,
    maquinaria_id=None,
    estado=None
):
    query = _aplicar_filtros(
        Tanqueo.query,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta,
        usuario_id=usuario_id,
        vehiculo_id=vehiculo_id,
        maquinaria_id=maquinaria_id,
        estado=estado
    )

    return query.order_by(
        Tanqueo.fecha.desc(),
        Tanqueo.id.desc()
    ).all()


def obtener_tanqueo(id):
    return Tanqueo.query.get(id)


def cambiar_estado(tanqueo, estado, usuario_id, observacion=None):
    estado = str(estado).strip().upper()

    if estado not in {"VALIDADO", "RECHAZADO"}:
        raise ValueError(
            "Estado inválido"
        )

    tanqueo.estado = estado
    tanqueo.validado_por = usuario_id
    tanqueo.validado_at = datetime.now(BOGOTA)
    tanqueo.observacion_validacion = observacion

    db.session.commit()

    return tanqueo


# ============================================================
# MÉTRICAS DE RENDIMIENTO
# ============================================================

def _numero(valor, decimales=2):
    if valor is None:
        return 0.0

    return round(
        float(valor),
        decimales
    )


def _seguro_dividir(numerador, denominador, decimales=2):
    if not denominador:
        return None

    return round(
        float(numerador) / float(denominador),
        decimales
    )


def calcular_metricas(
    fecha_desde=None,
    fecha_hasta=None,
    usuario_id=None,
    vehiculo_id=None,
    maquinaria_id=None
):
    """
    Construye métricas reales a partir de los tanqueos.

    Para cada activo:
    - total de galones y costo
    - distancia/horas recorridas entre tanqueos
    - rendimiento = recorrido / galones usados
    - costo por KM u hora
    """
    query = Tanqueo.query.filter(
        Tanqueo.estado != "RECHAZADO"
    )

    query = _aplicar_filtros(
        query,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta,
        usuario_id=usuario_id,
        vehiculo_id=vehiculo_id,
        maquinaria_id=maquinaria_id
    )

    tanqueos = query.order_by(
        Tanqueo.vehiculo_id,
        Tanqueo.maquinaria_id,
        Tanqueo.fecha.asc(),
        Tanqueo.id.asc()
    ).all()

    # --------------------------------------------------------
    # AGRUPAR POR ACTIVO
    # --------------------------------------------------------
    grupos = {}

    for tanqueo in tanqueos:
        if tanqueo.vehiculo_id:
            clave = ("VEHICULO", tanqueo.vehiculo_id)
            nombre = (
                tanqueo.vehiculo.placa
                if tanqueo.vehiculo
                else f"Vehículo {tanqueo.vehiculo_id}"
            )
            unidad = "KM"
        else:
            clave = ("MAQUINARIA", tanqueo.maquinaria_id)
            nombre = (
                tanqueo.maquinaria.codigo
                if tanqueo.maquinaria
                else f"Maquinaria {tanqueo.maquinaria_id}"
            )
            unidad = "HORAS"

        grupo = grupos.setdefault(
            clave,
            {
                "tipo_activo": clave[0],
                "activo_id": clave[1],
                "activo_nombre": nombre,
                "unidad": unidad,
                "registros": []
            }
        )

        grupo["registros"].append(tanqueo)

    # --------------------------------------------------------
    # CALCULAR POR ACTIVO
    # --------------------------------------------------------
    por_activo = []

    total_galones = Decimal("0")
    total_costo = Decimal("0")
    total_recorrido = Decimal("0")
    total_galones_usados = Decimal("0")
    total_tanqueos = len(tanqueos)

    for grupo in grupos.values():
        registros = grupo["registros"]

        galones_activo = sum(
            (r.galones or Decimal("0") for r in registros),
            Decimal("0")
        )

        costo_activo = sum(
            (r.valor_total or Decimal("0") for r in registros),
            Decimal("0")
        )

        recorrido = Decimal("0")
        galones_usados = Decimal("0")

        # Solo se puede medir consumo entre dos lecturas distintas.
        for anterior, actual in zip(
            registros,
            registros[1:]
        ):
            lectura_anterior = (
                anterior.km_actual
                if grupo["unidad"] == "KM"
                else anterior.horometro_actual
            )

            lectura_actual = (
                actual.km_actual
                if grupo["unidad"] == "KM"
                else actual.horometro_actual
            )

            if (
                lectura_anterior is None
                or lectura_actual is None
            ):
                continue

            delta = lectura_actual - lectura_anterior

            if delta <= 0:
                continue

            recorrido += delta
            galones_usados += (
                actual.galones or Decimal("0")
            )

        rendimiento = _seguro_dividir(
            recorrido,
            galones_usados
        )

        costo_unidad = _seguro_dividir(
            costo_activo,
            recorrido
        )

        por_activo.append({
            "tipo_activo": grupo["tipo_activo"],
            "activo_id": grupo["activo_id"],
            "activo_nombre": grupo["activo_nombre"],
            "unidad": grupo["unidad"],
            "tanqueos": len(registros),
            "total_galones": _numero(galones_activo),
            "total_costo": _numero(costo_activo),
            "recorrido": _numero(recorrido),
            "galones_usados": _numero(galones_usados),
            "rendimiento": rendimiento,
            "costo_por_unidad": costo_unidad,
            "promedio_galones": _seguro_dividir(
                galones_activo,
                len(registros)
            ),
            "ultima_fecha": (
                registros[-1].fecha.isoformat()
                if registros[-1].fecha
                else None
            )
        })

        total_galones += galones_activo
        total_costo += costo_activo
        total_recorrido += recorrido
        total_galones_usados += galones_usados

    # --------------------------------------------------------
    # SERIE MENSUAL
    # --------------------------------------------------------
    meses = {}

    for tanqueo in tanqueos:
        if not tanqueo.fecha:
            continue

        clave = tanqueo.fecha.strftime("%Y-%m")

        item = meses.setdefault(
            clave,
            {
                "mes": clave,
                "galones": Decimal("0"),
                "costo": Decimal("0"),
                "tanqueos": 0
            }
        )

        item["galones"] += (
            tanqueo.galones or Decimal("0")
        )
        item["costo"] += (
            tanqueo.valor_total or Decimal("0")
        )
        item["tanqueos"] += 1

    por_mes = [
        {
            "mes": datos["mes"],
            "galones": _numero(datos["galones"]),
            "costo": _numero(datos["costo"]),
            "tanqueos": datos["tanqueos"]
        }
        for _, datos in sorted(meses.items())
    ]

    # --------------------------------------------------------
    # POR OPERADOR
    # --------------------------------------------------------
    operadores = {}

    for tanqueo in tanqueos:
        clave = tanqueo.usuario_id
        item = operadores.setdefault(
            clave,
            {
                "usuario_id": clave,
                "usuario": (
                    tanqueo.usuario.nombre
                    if tanqueo.usuario
                    else None
                ),
                "galones": Decimal("0"),
                "costo": Decimal("0"),
                "tanqueos": 0
            }
        )

        item["galones"] += (
            tanqueo.galones or Decimal("0")
        )
        item["costo"] += (
            tanqueo.valor_total or Decimal("0")
        )
        item["tanqueos"] += 1

    por_operador = [
        {
            "usuario_id": datos["usuario_id"],
            "usuario": datos["usuario"],
            "galones": _numero(datos["galones"]),
            "costo": _numero(datos["costo"]),
            "tanqueos": datos["tanqueos"]
        }
        for _, datos in operadores.items()
    ]

    # --------------------------------------------------------
    # RESPUESTA
    # --------------------------------------------------------
    activos_con_rendimiento = [
        a for a in por_activo
        if a["rendimiento"] is not None
    ]

    mejor = None
    peor = None

    if activos_con_rendimiento:
        ordenados = sorted(
            activos_con_rendimiento,
            key=lambda x: x["rendimiento"],
            reverse=True
        )
        mejor = ordenados[0]
        peor = ordenados[-1]

    return {
        "resumen": {
            "tanqueos": total_tanqueos,
            "activos": len(por_activo),
            "total_galones": _numero(total_galones),
            "total_costo": _numero(total_costo),
            "recorrido_total": _numero(total_recorrido),
            "rendimiento_promedio": _seguro_dividir(
                total_recorrido,
                total_galones_usados
            ),
            "costo_por_unidad": _seguro_dividir(
                total_costo,
                total_recorrido
            ),
            "precio_galon_promedio": _seguro_dividir(
                total_costo,
                total_galones
            )
        },
        "por_activo": por_activo,
        "por_mes": por_mes,
        "por_operador": por_operador,
        "mejor_rendimiento": mejor,
        "peor_rendimiento": peor
    }
