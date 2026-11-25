from datetime import datetime

from sqlalchemy import func

from extensions import db
from models import (
    ActivoOperador,
    Usuario,
    Vehiculo,
    Maquinaria,
    Rol
)

from datetime import datetime
from zoneinfo import ZoneInfo

datetime.now(ZoneInfo("America/Bogota"))

# Roles que pueden usar varios vehículos de la empresa a la vez.
# Los operadores conservan la regla de un único activo asignado.
ROLES_ADMIN = ("admin", "administrador", "supervisor")
ROLES_ASIGNABLES = ("operador",) + ROLES_ADMIN


def _rol_de(usuario):
    """Devuelve el nombre del rol en minúsculas."""

    if not usuario or not usuario.rol:
        return ""

    return (usuario.rol.nombre or "").lower().strip()


def _es_administrador(usuario):
    """Indica si el usuario pertenece a un rol administrativo."""

    return _rol_de(usuario) in ROLES_ADMIN


class ActivoOperadorService:

    # ==================================================
    # LISTAR
    # ==================================================

    @staticmethod
    def listar():

    # ===============================
    # ASIGNACIONES
    # ===============================

        asignaciones = (
            ActivoOperador.query
            .filter(
                ActivoOperador.activo == True
            )
            .order_by(
                ActivoOperador.fecha_asignacion.desc()
            )
            .all()
        )

        lista_asignaciones = []

        for a in asignaciones:

            lista_asignaciones.append({

                "id": a.id,

                "activo": a.activo,

                "fecha_asignacion": (
                    a.fecha_asignacion.isoformat()
                    if a.fecha_asignacion else None
                ),

                "fecha_fin": (
                    a.fecha_fin.isoformat()
                    if a.fecha_fin else None
                ),

                "observaciones": a.observaciones,

                "usuario": {
                    "id": a.usuario.id,
                    "nombre": a.usuario.nombre
                },

                "vehiculo": (
                    {
                        "id": a.vehiculo.id,
                        "placa": a.vehiculo.placa,
                        "marca": a.vehiculo.marca,
                        "modelo": a.vehiculo.modelo
                    }
                    if a.vehiculo else None
                ),

                "maquinaria": (
                    {
                        "id": a.maquinaria.id,
                        "codigo": a.maquinaria.codigo,
                        "marca": a.maquinaria.marca,
                        "modelo": a.maquinaria.modelo
                    }
                    if a.maquinaria else None
                )

            })

    # ===============================
# OPERADORES
# ===============================

        usuarios = (
            Usuario.query
            .join(Rol)
            .filter(
                Usuario.activo == True,
                func.lower(Rol.nombre).in_(ROLES_ASIGNABLES)
            )
            .order_by(Usuario.nombre)
            .all()
        )

        # Los operadores solo pueden tener un activo. Los roles
        # administrativos pueden acumular varios.
        usuarios_con_activo = {
            fila[0]
            for fila in db.session.query(
                ActivoOperador.usuario_id
            ).filter(
                ActivoOperador.activo == True
            ).all()
        }

        lista_usuarios = []

        for u in usuarios:

            rol = _rol_de(u)

            if rol == "operador" and u.id in usuarios_con_activo:
                continue

            lista_usuarios.append({

            "id": u.id,
            "nombre": u.nombre,
            "username": u.username,
            "telefono": u.telefono,
            "rol": rol,
            "multiples_activos": rol in ROLES_ADMIN

        })

    # ===============================
    # VEHÍCULOS
    # ===============================

        # Un activo asignado a un operador no se ofrece para
        # evitar cambiar su flujo. Sí se comparte entre roles
        # administrativos.
        ids_admin = [
            u.id
            for u in Usuario.query.join(Rol).filter(
                func.lower(Rol.nombre).in_(ROLES_ADMIN)
            ).all()
        ]

        vehiculos_en_operador = (
            db.session.query(ActivoOperador.vehiculo_id)
            .filter(
                ActivoOperador.activo == True,
                ActivoOperador.vehiculo_id.isnot(None),
                ~ActivoOperador.usuario_id.in_(ids_admin)
            )
        )

        vehiculos = []

        datos = (
            Vehiculo.query
            .filter(
                Vehiculo.activo == True,
                ~Vehiculo.id.in_(vehiculos_en_operador)
            )
            .order_by(Vehiculo.placa)
            .all()
        )

        for v in datos:

            vehiculos.append({

                "id": v.id,

                "placa": v.placa,

                "marca": v.marca,

                "modelo": v.modelo,

                "tipo_vehiculo_id": v.tipo_vehiculo_id

            })

    # ===============================
    # MAQUINARIA
    # ===============================

        maquinarias_en_operador = (
            db.session.query(ActivoOperador.maquinaria_id)
            .filter(
                ActivoOperador.activo == True,
                ActivoOperador.maquinaria_id.isnot(None),
                ~ActivoOperador.usuario_id.in_(ids_admin)
            )
        )

        maquinarias = []

        datos = (
            Maquinaria.query
            .filter(
                Maquinaria.activo == True,
                ~Maquinaria.id.in_(maquinarias_en_operador)
            )
            .order_by(Maquinaria.codigo)
            .all()
        )

        for m in datos:

            maquinarias.append({

                "id": m.id,

                "codigo": m.codigo,

                "marca": m.marca,

                "modelo": m.modelo,

                "tipo_maquinaria_id": m.tipo_maquinaria_id

            })

        return {

            "asignaciones": lista_asignaciones,

            "usuarios": lista_usuarios,

            "vehiculos": vehiculos,

            "maquinarias": maquinarias

        }
    # ==================================================
    # OBTENER
    # ==================================================

    @staticmethod
    def obtener(id):

        dato = ActivoOperador.query.get(id)

        if not dato:
            raise Exception("Asignación no encontrada.")

        return dato

    # ==================================================
    # CREAR
    # ==================================================

    @staticmethod
    def crear(data):

        usuario_id = data.get("usuario_id")
        vehiculo_id = data.get("vehiculo_id")
        maquinaria_id = data.get("maquinaria_id")
        observaciones = data.get("observaciones")

        if not usuario_id:
            raise Exception("Debe seleccionar un operador.")

        if not vehiculo_id and not maquinaria_id:
            raise Exception("Debe seleccionar un vehículo o maquinaria.")

        if vehiculo_id and maquinaria_id:
            raise Exception("Solo puede asignar un tipo de activo.")

        try:
            usuario_id = int(usuario_id)
        except (TypeError, ValueError):
            raise Exception("Debe seleccionar un operador.")

        usuario = Usuario.query.get(usuario_id)

        if not usuario:
            raise Exception("El usuario seleccionado no existe.")

        es_admin = _es_administrador(usuario)

        # --------------------------------------------
        # No repetir el mismo activo para el mismo usuario
        # --------------------------------------------

        duplicado = ActivoOperador.query.filter_by(
            usuario_id=usuario_id,
            vehiculo_id=vehiculo_id,
            maquinaria_id=maquinaria_id,
            activo=True
        ).first()

        if duplicado:
            raise Exception(
                "Ese activo ya está asignado a este usuario."
            )

        # --------------------------------------------
        # Operadores: un único activo (flujo original)
        # --------------------------------------------

        if not es_admin:

            existe = ActivoOperador.query.filter_by(
                usuario_id=usuario_id,
                activo=True
            ).first()

            if existe:
                raise Exception(
                    "El operador ya tiene un activo asignado."
                )

        # --------------------------------------------
        # Vehículo ocupado por un operador
        # --------------------------------------------

        if vehiculo_id:

            ocupado = ActivoOperador.query.filter_by(
                vehiculo_id=vehiculo_id,
                activo=True
            ).first()

            if ocupado and (
                not es_admin
                or not _es_administrador(ocupado.usuario)
            ):
                raise Exception(
                    "Ese vehículo ya está asignado."
                )

        # --------------------------------------------
        # Maquinaria ocupada por un operador
        # --------------------------------------------

        if maquinaria_id:

            ocupado = ActivoOperador.query.filter_by(
                maquinaria_id=maquinaria_id,
                activo=True
            ).first()

            if ocupado and (
                not es_admin
                or not _es_administrador(ocupado.usuario)
            ):
                raise Exception(
                    "Esa maquinaria ya está asignada."
                )

        nuevo = ActivoOperador(

            usuario_id=usuario_id,

            vehiculo_id=vehiculo_id,

            maquinaria_id=maquinaria_id,

            observaciones=observaciones,

            activo=True

        )

        db.session.add(nuevo)

        db.session.commit()

        return nuevo

    # ==================================================
    # ACTUALIZAR
    # ==================================================

    @staticmethod
    def actualizar(id, data):

        asignacion = ActivoOperador.query.get(id)

        if not asignacion:
            raise Exception("Asignación no encontrada.")

        usuario_id = data.get("usuario_id")
        vehiculo_id = data.get("vehiculo_id")
        maquinaria_id = data.get("maquinaria_id")

        if not vehiculo_id and not maquinaria_id:
            raise Exception("Debe seleccionar un activo.")

        if not usuario_id:
            usuario_id = asignacion.usuario_id

        try:
            usuario_id = int(usuario_id)
        except (TypeError, ValueError):
            raise Exception("Debe seleccionar un operador.")

        usuario = Usuario.query.get(usuario_id)

        if not usuario:
            raise Exception("El usuario seleccionado no existe.")

        es_admin = _es_administrador(usuario)

        # --------------------------------------------
        # No repetir el mismo activo para el mismo usuario
        # --------------------------------------------

        duplicado = (
            ActivoOperador.query
            .filter_by(
                usuario_id=usuario_id,
                vehiculo_id=vehiculo_id,
                maquinaria_id=maquinaria_id,
                activo=True
            )
            .filter(ActivoOperador.id != id)
            .first()
        )

        if duplicado:
            raise Exception(
                "Ese activo ya está asignado a este usuario."
            )

        # --------------------------------------------
        # Operadores: un único activo (flujo original)
        # --------------------------------------------

        if not es_admin:

            existe = (
                ActivoOperador.query
                .filter(
                    ActivoOperador.usuario_id == usuario_id,
                    ActivoOperador.activo == True,
                    ActivoOperador.id != id
                )
                .first()
            )

            if existe:
                raise Exception(
                    "El operador ya tiene otro activo asignado."
                )

        # --------------------------------------------
        # Validar vehículo (no quitar activo a un operador)
        # --------------------------------------------

        if vehiculo_id:

            ocupado = (
                ActivoOperador.query
                .filter(
                    ActivoOperador.vehiculo_id == vehiculo_id,
                    ActivoOperador.activo == True,
                    ActivoOperador.id != id
                )
                .first()
            )

            if ocupado and (
                not es_admin
                or not _es_administrador(ocupado.usuario)
            ):
                raise Exception(
                    "Ese vehículo ya está asignado."
                )

        # --------------------------------------------
        # Validar maquinaria (no quitar activo a un operador)
        # --------------------------------------------

        if maquinaria_id:

            ocupado = (
                ActivoOperador.query
                .filter(
                    ActivoOperador.maquinaria_id == maquinaria_id,
                    ActivoOperador.activo == True,
                    ActivoOperador.id != id
                )
                .first()
            )

            if ocupado and (
                not es_admin
                or not _es_administrador(ocupado.usuario)
            ):
                raise Exception(
                    "Esa maquinaria ya está asignada."
                )

        asignacion.usuario_id = usuario_id
        asignacion.vehiculo_id = vehiculo_id
        asignacion.maquinaria_id = maquinaria_id
        asignacion.observaciones = data.get("observaciones")

        db.session.commit()

        return asignacion

    # ==================================================
    # DESACTIVAR
    # ==================================================

    @staticmethod
    def eliminar(id):

        asignacion = ActivoOperador.query.get(id)

        if not asignacion:
            raise Exception("Asignación no encontrada.")

        asignacion.activo = False
        asignacion.fecha_fin = datetime.now(ZoneInfo("America/Bogota"))

        db.session.commit()

        return asignacion