from flask import Blueprint, request, jsonify
from flask_jwt_extended import get_jwt_identity, jwt_required

from services.usuario_services import UsuarioService

from models import TipoOperador


usuarios_bp = Blueprint(
    "usuarios",
    __name__
)


# =====================================================
# LISTAR
# =====================================================

@usuarios_bp.route(
    "/",
    methods=["GET"]
)
@jwt_required()
def listar():

    return jsonify(
        UsuarioService.listar()
    )


# =====================================================
# LISTAR TIPOS DE OPERADOR
# =====================================================

@usuarios_bp.route(
    "/tipos-operador",
    methods=["GET"]
)
@jwt_required()
def listar_tipos_operador():

    tipos = (
        TipoOperador.query
        .filter_by(activo=True)
        .order_by(
            TipoOperador.id
        )
        .all()
    )

    return jsonify({
        "success": True,
        "data": [
            tipo.to_dict()
            for tipo in tipos
        ]
    })


# =====================================================
# LISTAR DOCUMENTOS SEGÚN TIPO DE OPERADOR
# =====================================================

@usuarios_bp.route(
    "/tipos-operador/<int:tipo_operador_id>/documentos",
    methods=["GET"]
)
@jwt_required()
def listar_documentos_tipo_operador(
    tipo_operador_id
):

    return UsuarioService.listar_documentos_tipo_operador(
        tipo_operador_id
    )


# =====================================================
# OBTENER POR ID
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>",
    methods=["GET"]
)
@jwt_required()
def obtener(usuario_id):

    return jsonify(
        UsuarioService.obtener(
            usuario_id
        )
    )


# =====================================================
# CREAR
# =====================================================

@usuarios_bp.route(
    "/",
    methods=["POST"]
)
@jwt_required()
def crear():

    data = request.get_json()

    return UsuarioService.crear(
        data
    )


# =====================================================
# ACTUALIZAR
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>",
    methods=["PUT"]
)
@jwt_required()
def actualizar(usuario_id):

    data = request.get_json()

    return UsuarioService.actualizar(
        usuario_id,
        data
    )


# =====================================================
# DESACTIVAR (MANTIENE EL REGISTRO)
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>/desactivar",
    methods=["PUT"]
)
@jwt_required()
def desactivar(usuario_id):

    return UsuarioService.desactivar(
        usuario_id
    )


# =====================================================
# IMPACTO DE ELIMINACIÓN DEFINITIVA
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>/impacto-eliminacion",
    methods=["GET"]
)
@jwt_required()
def impacto_eliminacion(usuario_id):

    return UsuarioService.impacto_eliminacion(
        usuario_id
    )


# =====================================================
# ELIMINAR DEFINITIVAMENTE (SOLO OPERADORES)
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>/eliminar",
    methods=["DELETE"]
)
@jwt_required()
def eliminar_definitivamente(usuario_id):

    identidad = get_jwt_identity()

    if identidad is not None and str(identidad) == str(usuario_id):

        return jsonify({
            "success": False,
            "message": "No puede eliminar su propio usuario."
        }), 400

    return UsuarioService.eliminar_definitivamente(
        usuario_id
    )


# =====================================================
# COMPATIBILIDAD: DELETE ANTES DESACTIVABA
# =====================================================

@usuarios_bp.route(
    "/<int:usuario_id>",
    methods=["DELETE"]
)
@jwt_required()
def eliminar(usuario_id):

    return UsuarioService.desactivar(
        usuario_id
    )