# ============================================================
# INTELLIFEET
# RUTAS - TANQUEOS
# ============================================================

from flask import Blueprint, request, jsonify
from flask_jwt_extended import jwt_required, get_jwt_identity

from services.tanqueo_service import (
    obtener_usuario,
    validar_rol_admin,
    obtener_activos_operador,
    crear_tanqueo,
    obtener_tanqueos_operador,
    obtener_tanqueos_admin,
    obtener_tanqueo,
    cambiar_estado,
    calcular_metricas
)


tanqueos_bp = Blueprint("tanqueos", __name__)


def usuario_actual():
    identity = get_jwt_identity()

    if isinstance(identity, dict):
        usuario_id = identity.get("id")
    else:
        usuario_id = identity

    if not usuario_id:
        return None

    try:
        usuario_id = int(usuario_id)
    except (TypeError, ValueError):
        return None

    return obtener_usuario(usuario_id)


def respuesta_error(mensaje, status=400, detail=None):
    respuesta = {"error": mensaje}

    if detail:
        respuesta["detail"] = detail

    return jsonify(respuesta), status


@tanqueos_bp.get("/mis-activos")
@jwt_required()
def mis_activos():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    try:
        return jsonify(
            obtener_activos_operador(usuario.id)
        ), 200

    except Exception as e:
        return respuesta_error(
            "Error al consultar los activos asignados",
            500,
            str(e)
        )


@tanqueos_bp.post("")
@jwt_required()
def crear():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    try:
        archivo = (
            request.files.get("archivo")
            or request.files.get("imagen")
        )

        tanqueo = crear_tanqueo(
            usuario_id=usuario.id,
            fecha=request.form.get("fecha"),
            galones=request.form.get("galones"),
            vehiculo_id=request.form.get("vehiculo_id"),
            maquinaria_id=request.form.get("maquinaria_id"),
            precio_galon=request.form.get("precio_galon"),
            tipo_combustible=request.form.get(
                "tipo_combustible",
                "DIESEL"
            ),
            estacion=request.form.get("estacion"),
            numero_factura=request.form.get(
                "numero_factura"
            ),
            km_actual=request.form.get("km_actual"),
            horometro_actual=request.form.get(
                "horometro_actual"
            ),
            archivo=archivo,
            observaciones=request.form.get(
                "observaciones"
            )
        )

        return jsonify(tanqueo.to_dict()), 201

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al registrar el tanqueo",
            500,
            str(e)
        )


@tanqueos_bp.get("/mis-tanqueos")
@jwt_required()
def mis_tanqueos():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    try:
        tanqueos = obtener_tanqueos_operador(
            usuario.id,
            fecha_desde=request.args.get("fecha_desde"),
            fecha_hasta=request.args.get("fecha_hasta")
        )

        return jsonify([
            t.to_dict() for t in tanqueos
        ]), 200

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al consultar los tanqueos",
            500,
            str(e)
        )


@tanqueos_bp.get("/mis-metricas")
@jwt_required()
def mis_metricas():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    try:
        metricas = calcular_metricas(
            fecha_desde=request.args.get("fecha_desde"),
            fecha_hasta=request.args.get("fecha_hasta"),
            usuario_id=usuario.id
        )

        return jsonify(metricas), 200

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al calcular las métricas",
            500,
            str(e)
        )


@tanqueos_bp.get("")
@jwt_required()
def listar():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    if not validar_rol_admin(usuario):
        return respuesta_error(
            "No tiene permisos para consultar este módulo",
            403
        )

    try:
        tanqueos = obtener_tanqueos_admin(
            fecha_desde=request.args.get("fecha_desde"),
            fecha_hasta=request.args.get("fecha_hasta"),
            usuario_id=request.args.get("usuario_id"),
            vehiculo_id=request.args.get("vehiculo_id"),
            maquinaria_id=request.args.get("maquinaria_id"),
            estado=request.args.get("estado")
        )

        return jsonify([
            t.to_dict() for t in tanqueos
        ]), 200

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al consultar los tanqueos",
            500,
            str(e)
        )


@tanqueos_bp.get("/metricas")
@jwt_required()
def metricas():
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    if not validar_rol_admin(usuario):
        return respuesta_error(
            "No tiene permisos para consultar este módulo",
            403
        )

    try:
        datos = calcular_metricas(
            fecha_desde=request.args.get("fecha_desde"),
            fecha_hasta=request.args.get("fecha_hasta"),
            usuario_id=request.args.get("usuario_id"),
            vehiculo_id=request.args.get("vehiculo_id"),
            maquinaria_id=request.args.get("maquinaria_id")
        )

        return jsonify(datos), 200

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al calcular las métricas",
            500,
            str(e)
        )


@tanqueos_bp.get("/<int:tanqueo_id>")
@jwt_required()
def obtener(tanqueo_id):
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    tanqueo = obtener_tanqueo(tanqueo_id)

    if not tanqueo:
        return respuesta_error("Tanqueo no encontrado", 404)

    if (
        not validar_rol_admin(usuario)
        and tanqueo.usuario_id != usuario.id
    ):
        return respuesta_error(
            "No tiene permiso para ver este tanqueo",
            403
        )

    return jsonify(tanqueo.to_dict()), 200


@tanqueos_bp.put("/<int:tanqueo_id>/validar")
@jwt_required()
def validar(tanqueo_id):
    return _cambiar_estado(tanqueo_id, "VALIDADO")


@tanqueos_bp.put("/<int:tanqueo_id>/rechazar")
@jwt_required()
def rechazar(tanqueo_id):
    return _cambiar_estado(tanqueo_id, "RECHAZADO")


def _cambiar_estado(tanqueo_id, estado):
    usuario = usuario_actual()

    if not usuario:
        return respuesta_error("Usuario no encontrado", 404)

    if not validar_rol_admin(usuario):
        return respuesta_error(
            "No tiene permisos para realizar esta acción",
            403
        )

    tanqueo = obtener_tanqueo(tanqueo_id)

    if not tanqueo:
        return respuesta_error("Tanqueo no encontrado", 404)

    try:
        datos = request.get_json(silent=True) or {}

        actualizado = cambiar_estado(
            tanqueo,
            estado,
            usuario.id,
            datos.get("observacion")
        )

        return jsonify(actualizado.to_dict()), 200

    except ValueError as e:
        return respuesta_error(str(e), 400)

    except Exception as e:
        return respuesta_error(
            "Error al actualizar el estado",
            500,
            str(e)
        )
