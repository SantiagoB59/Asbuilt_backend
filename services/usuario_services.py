import os

from flask import jsonify

from sqlalchemy import bindparam, or_, text

from werkzeug.security import generate_password_hash

from extensions import db

from models import (
    ActivoOperador,
    Alerta,
    AnomaliaFoto,
    ControlDiario,
    Inspeccion,
    InspeccionAnomalia,
    Rol,
    Tanqueo,
    TipoOperador,
    TipoOperadorDocumento,
    Usuario,
    UsuarioDocumento,
)


class UsuarioService:

    # =====================================================
    # LISTAR
    # =====================================================

    @staticmethod
    def listar():

        usuarios = Usuario.query.order_by(Usuario.nombre).all()

        roles = Rol.query.order_by(Rol.nombre).all()

        lista_usuarios = []

        for u in usuarios:

            lista_usuarios.append(
                {
                    "id": u.id,
                    "nombre": u.nombre,
                    "username": u.username,
                    "email": u.email,
                    "telefono": u.telefono,
                    "activo": u.activo,
                    "created_at": u.created_at,
                    "role_id": u.role_id,
                    "tipo_operador_id": u.tipo_operador_id,
                    "rol": {"id": u.rol.id, "nombre": u.rol.nombre},
                }
            )

        lista_roles = []

        for r in roles:

            lista_roles.append({"id": r.id, "nombre": r.nombre})

        return {
            "success": True,
            "data": {"usuarios": lista_usuarios, "roles": lista_roles},
        }

    # =====================================================
    # LISTAR DOCUMENTOS SEGÚN TIPO DE OPERADOR
    # =====================================================

    @staticmethod
    def listar_documentos_tipo_operador(tipo_operador_id):

        tipo_operador = TipoOperador.query.get(tipo_operador_id)

        if not tipo_operador:

            return (
                jsonify(
                    {"success": False, "message": "Tipo de operador no encontrado"}
                ),
                404,
            )

        documentos = (
            TipoOperadorDocumento.query.filter_by(tipo_operador_id=tipo_operador_id)
            .order_by(TipoOperadorDocumento.id)
            .all()
        )

        return jsonify(
            {
                "success": True,
                "data": {
                    "tipo_operador": tipo_operador.to_dict(),
                    "documentos": [
                        {
                            "id": documento.documento_tipo_id,
                            "documento_tipo_id": documento.documento_tipo_id,
                            "nombre": (
                                documento.documento_tipo.nombre
                                if documento.documento_tipo
                                else None
                            ),
                            "obligatorio": documento.obligatorio,
                        }
                        for documento in documentos
                    ],
                },
            }
        )

    # =====================================================
    # OBTENER
    # =====================================================

    @staticmethod
    def obtener(usuario_id):

        usuario = Usuario.query.get(usuario_id)

        if not usuario:

            return jsonify({"success": False, "message": "Usuario no encontrado"}), 404

        return {
            "success": True,
            "data": {
                "id": usuario.id,
                "nombre": usuario.nombre,
                "username": usuario.username,
                "email": usuario.email,
                "telefono": usuario.telefono,
                "activo": usuario.activo,
                "role_id": usuario.role_id,
                "rol": usuario.rol.nombre,
            },
        }

    @staticmethod
    def crear(data):
        if Usuario.query.filter_by(username=data["username"]).first():
            return jsonify({
                "success": False,
                "message": "El usuario ya existe"
            }), 400

        if data.get("email"):
            existe = Usuario.query.filter_by(email=data["email"]).first()

            if existe:
                return jsonify({
                    "success": False,
                    "message": "El correo ya existe"
                }), 400

        # =========================================================
        # OBTENER ROL
        # =========================================================

        if "role_id" in data:
            rol = Rol.query.get(data["role_id"])
        else:
            rol = Rol.query.filter_by(
                nombre=data.get("rol")
            ).first()

        if not rol:
            return jsonify({
                "success": False,
                "message": "Rol inválido"
            }), 400

        # =========================================================
        # TIPO DE OPERADOR
        # =========================================================

        tipo_operador_id = data.get("tipo_operador_id")
        tipo_operador = None

        if rol.id == 3:

            if not tipo_operador_id:
                return jsonify({
                    "success": False,
                    "message": "Debe seleccionar el tipo de operador"
                }), 400

            tipo_operador = TipoOperador.query.get(tipo_operador_id)

            if not tipo_operador:
                return jsonify({
                    "success": False,
                    "message": "El tipo de operador no es válido"
                }), 400

        else:
            tipo_operador_id = None

        # =========================================================
        # CREAR USUARIO
        # =========================================================

        usuario = Usuario(
            nombre=data["nombre"],
            username=data["username"],
            email=data.get("email"),
            telefono=data.get("telefono"),
            role_id=rol.id,
            tipo_operador_id=tipo_operador_id,
            activo=data.get("activo", True)
        )

        usuario.password_hash = generate_password_hash(
            data["password"]
        )

        db.session.add(usuario)

        # Obtener ID antes de crear documentos
        db.session.flush()

        # =========================================================
        # CREAR DOCUMENTOS DEL OPERADOR
        # =========================================================

        documentos_creados = []

        if rol.id == 3:

            documentos_requeridos = (
                TipoOperadorDocumento.query
                .filter_by(
                    tipo_operador_id=tipo_operador.id,
                    obligatorio=True
                )
                .all()
            )

            # Documentos enviados desde Angular
            documentos_data = data.get("documentos", [])

            # Crear diccionario para buscar fácilmente
            fechas_documentos = {}

            for documento in documentos_data:

                documento_tipo_id = documento.get(
                    "documento_tipo_id"
                )

                fecha_vencimiento = documento.get(
                    "fecha_vencimiento"
                )

                if documento_tipo_id:
                    fechas_documentos[
                        int(documento_tipo_id)
                    ] = fecha_vencimiento

            # Crear documentos requeridos
            for requerido in documentos_requeridos:

                fecha_vencimiento = fechas_documentos.get(
                    requerido.documento_tipo_id
                )

                documento = UsuarioDocumento(
                    usuario_id=usuario.id,
                    documento_tipo_id=requerido.documento_tipo_id,
                    fecha_vencimiento=fecha_vencimiento,
                    archivo_url=None,
                    activo=True
                )

                db.session.add(documento)

                documentos_creados.append(documento)

        # =========================================================
        # GUARDAR TODO
        # =========================================================

        db.session.commit()

        # =========================================================
        # RESPUESTA
        # =========================================================

        return jsonify({
            "success": True,
            "message": "Usuario creado correctamente",
            "data": {
                "id": usuario.id,
                "nombre": usuario.nombre,
                "username": usuario.username,
                "email": usuario.email,
                "telefono": usuario.telefono,
                "activo": usuario.activo,
                "role_id": usuario.role_id,
                "rol": rol.nombre,
                "tipo_operador_id": usuario.tipo_operador_id,
                "documentos": [
                    documento.to_dict()
                    for documento in documentos_creados
                ]
            }
        }), 201
    # =====================================================
    # ACTUALIZAR
    # =====================================================

    @staticmethod
    def actualizar(usuario_id, data):

        usuario = Usuario.query.get(usuario_id)

        if not usuario:

            return jsonify({"success": False, "message": "Usuario no encontrado"}), 404

        # ==========================================
        # VALIDAR USERNAME
        # ==========================================

        existe = Usuario.query.filter(
            Usuario.username == data["username"], Usuario.id != usuario_id
        ).first()

        if existe:

            return (
                jsonify(
                    {"success": False, "message": "El nombre de usuario ya existe"}
                ),
                400,
            )

        # ==========================================
        # VALIDAR EMAIL
        # ==========================================

        if data.get("email"):

            existe = Usuario.query.filter(
                Usuario.email == data["email"], Usuario.id != usuario_id
            ).first()

            if existe:

                return (
                    jsonify({"success": False, "message": "El correo ya existe"}),
                    400,
                )

        # ==========================================
        # OBTENER ROL
        # ==========================================

        if "role_id" in data:

            rol = Rol.query.get(data["role_id"])

        else:

            rol = Rol.query.filter_by(nombre=data.get("rol")).first()

        if not rol:

            return jsonify({"success": False, "message": "Rol inválido"}), 400

        # ==========================================
        # TIPO DE OPERADOR
        # ==========================================

        tipo_operador_id = data.get("tipo_operador_id")

        tipo_operador = None

        # ------------------------------------------
        # SI ES OPERADOR
        # ------------------------------------------

        if rol.id == 3:

            if not tipo_operador_id:

                return (
                    jsonify(
                        {
                            "success": False,
                            "message": ("Debe seleccionar el tipo " "de operador"),
                        }
                    ),
                    400,
                )

            tipo_operador = TipoOperador.query.get(tipo_operador_id)

            if not tipo_operador:

                return (
                    jsonify(
                        {
                            "success": False,
                            "message": ("El tipo de operador " "no es válido"),
                        }
                    ),
                    400,
                )

        else:

            # Si deja de ser operador,
            # quitamos el tipo de operador.
            tipo_operador_id = None

        # ==========================================
        # ACTUALIZAR DATOS DEL USUARIO
        # ==========================================

        usuario.nombre = data["nombre"]

        usuario.username = data["username"]

        usuario.email = data.get("email")

        usuario.telefono = data.get("telefono")

        usuario.role_id = rol.id

        usuario.tipo_operador_id = tipo_operador_id

        usuario.activo = data.get("activo", True)

        # ==========================================
        # ACTUALIZAR CONTRASEÑA
        # ==========================================

        if data.get("password"):

            usuario.password_hash = generate_password_hash(data["password"])

        # ==========================================
        # DOCUMENTOS DEL OPERADOR
        # ==========================================

        if rol.id == 3:

            documentos_requeridos = TipoOperadorDocumento.query.filter_by(
                tipo_operador_id=tipo_operador.id, obligatorio=True
            ).all()

            # IDs de documentos que ya tiene
            documentos_existentes = {
                documento.documento_tipo_id
                for documento in UsuarioDocumento.query.filter_by(
                    usuario_id=usuario.id
                ).all()
            }

            # ------------------------------------------
            # CREAR SOLO LOS DOCUMENTOS QUE FALTEN
            # ------------------------------------------

            for requerido in documentos_requeridos:

                if requerido.documento_tipo_id not in documentos_existentes:

                    documento = UsuarioDocumento(
                        usuario_id=usuario.id,
                        documento_tipo_id=(requerido.documento_tipo_id),
                        fecha_vencimiento=None,
                        archivo_url=None,
                        activo=True,
                    )

                    db.session.add(documento)

        # ==========================================
        # GUARDAR
        # ==========================================

        db.session.commit()

        return jsonify(
            {"success": True, "message": "Usuario actualizado correctamente"}
        )

    # =====================================================
    # DESACTIVAR (NO BORRA EL REGISTRO)
    # =====================================================

    @staticmethod
    def desactivar(usuario_id):

        usuario = Usuario.query.get(usuario_id)

        if not usuario:

            return jsonify({"success": False, "message": "Usuario no encontrado"}), 404

        usuario.activo = False

        db.session.commit()

        return jsonify(
            {"success": True, "message": "Usuario desactivado correctamente"}
        )

    # =====================================================
    # COMPATIBILIDAD: ELIMINAR ANTES DESACTIVABA
    # =====================================================

    @staticmethod
    def eliminar(usuario_id):

        return UsuarioService.desactivar(usuario_id)

    # =====================================================
    # AYUDAS PARA ELIMINACIÓN DEFINITIVA
    # =====================================================

    @staticmethod
    def _es_operador(usuario):

        nombre_rol = (
            usuario.rol.nombre
            if usuario.rol
            else ""
        )

        return str(nombre_rol).lower() == "operador"

    @staticmethod
    def _eliminar_archivo(ruta_relativa):

        if not ruta_relativa:
            return

        limpia = str(ruta_relativa).replace("\\", "/").lstrip("/")

        if limpia.startswith("uploads/"):
            ruta = limpia
        else:
            ruta = os.path.join(
                "uploads",
                *limpia.split("/")
            )

        try:
            if os.path.exists(ruta):
                os.remove(ruta)
        except OSError:
            # La limpieza de archivos no debe bloquear la eliminación.
            pass

    @staticmethod
    def _ids_anomalias(usuario_id, ids_inspecciones):

        return [
            fila[0]
            for fila in db.session.query(InspeccionAnomalia.id)
            .filter(
                or_(
                    InspeccionAnomalia.usuario_id == usuario_id,
                    InspeccionAnomalia.inspeccion_id.in_(ids_inspecciones),
                )
            )
            .all()
        ]

    # =====================================================
    # IMPACTO DE LA ELIMINACIÓN DEFINITIVA
    # =====================================================

    @staticmethod
    def impacto_eliminacion(usuario_id):

        usuario = Usuario.query.get(usuario_id)

        if not usuario:

            return jsonify({"success": False, "message": "Usuario no encontrado"}), 404

        if not UsuarioService._es_operador(usuario):

            return jsonify({
                "success": False,
                "message": "La eliminación definitiva solo está disponible para operadores."
            }), 400

        ids_inspecciones = [
            fila[0]
            for fila in db.session.query(Inspeccion.id)
            .filter(Inspeccion.usuario_id == usuario_id)
            .all()
        ]

        ids_anomalias = UsuarioService._ids_anomalias(
            usuario_id,
            ids_inspecciones
        )

        data = {
            "usuario": {
                "id": usuario.id,
                "nombre": usuario.nombre,
                "username": usuario.username,
            },
            "eliminar": {
                "documentos": UsuarioDocumento.query
                    .filter_by(usuario_id=usuario_id).count(),
                "asignaciones": ActivoOperador.query
                    .filter_by(usuario_id=usuario_id).count(),
                "inspecciones": len(ids_inspecciones),
                "anomalias": len(ids_anomalias),
                "controles_diarios": ControlDiario.query
                    .filter_by(usuario_id=usuario_id).count(),
                "tanqueos": Tanqueo.query
                    .filter_by(usuario_id=usuario_id).count(),
            },
            "desvincular": {
                "alertas": Alerta.query
                    .filter_by(usuario_id=usuario_id).count(),
                "revisiones_inspeccion": Inspeccion.query
                    .filter_by(responsable_revision=usuario_id).count(),
                "validaciones_control": ControlDiario.query
                    .filter_by(validado_por=usuario_id).count(),
                "validaciones_tanqueo": Tanqueo.query
                    .filter_by(validado_por=usuario_id).count(),
            },
        }

        return jsonify({"success": True, "data": data})

    # =====================================================
    # ELIMINAR DEFINITIVAMENTE
    # =====================================================

    @staticmethod
    def eliminar_definitivamente(usuario_id):

        usuario = Usuario.query.get(usuario_id)

        if not usuario:

            return jsonify({"success": False, "message": "Usuario no encontrado"}), 404

        if not UsuarioService._es_operador(usuario):

            return jsonify({
                "success": False,
                "message": "La eliminación definitiva solo está disponible para operadores."
            }), 400

        try:

            # ---------------------------------------------
            # 1) DESVINCULAR REFERENCIAS OPCIONALES
            # ---------------------------------------------

            Alerta.query.filter_by(usuario_id=usuario_id).update(
                {Alerta.usuario_id: None},
                synchronize_session=False
            )

            Inspeccion.query.filter_by(responsable_revision=usuario_id).update(
                {Inspeccion.responsable_revision: None},
                synchronize_session=False
            )

            ControlDiario.query.filter_by(validado_por=usuario_id).update(
                {ControlDiario.validado_por: None},
                synchronize_session=False
            )

            Tanqueo.query.filter_by(validado_por=usuario_id).update(
                {Tanqueo.validado_por: None},
                synchronize_session=False
            )

            # ---------------------------------------------
            # 2) IDENTIFICAR INSPECCIONES Y ANOMALÍAS
            # ---------------------------------------------

            ids_inspecciones = [
                fila[0]
                for fila in db.session.query(Inspeccion.id)
                .filter(Inspeccion.usuario_id == usuario_id)
                .all()
            ]

            ids_anomalias = UsuarioService._ids_anomalias(
                usuario_id,
                ids_inspecciones
            )

            # ---------------------------------------------
            # 3) ARCHIVOS DE SOPORTE (SE BORRAN TRAS COMMIT)
            # ---------------------------------------------

            archivos_a_eliminar = []

            for tanqueo in Tanqueo.query.filter_by(usuario_id=usuario_id).all():
                if tanqueo.archivo_path:
                    archivos_a_eliminar.append(tanqueo.archivo_path)

            for control in ControlDiario.query.filter_by(usuario_id=usuario_id).all():
                if control.archivo_path:
                    archivos_a_eliminar.append(control.archivo_path)

            # ---------------------------------------------
            # 4) DEPENDENCIAS DE INSPECCIONES
            # ---------------------------------------------

            if ids_anomalias:

                AnomaliaFoto.query.filter(
                    AnomaliaFoto.anomalia_id.in_(ids_anomalias)
                ).delete(synchronize_session=False)

                InspeccionAnomalia.query.filter(
                    InspeccionAnomalia.id.in_(ids_anomalias)
                ).delete(synchronize_session=False)

            if ids_inspecciones:

                db.session.execute(
                    text(
                        "DELETE FROM inspeccion_reportes "
                        "WHERE inspeccion_id IN :ids"
                    ).bindparams(bindparam("ids", expanding=True)),
                    {"ids": ids_inspecciones}
                )

            # ---------------------------------------------
            # 5) REGISTROS PROPIOS DEL OPERADOR
            # ---------------------------------------------

            UsuarioDocumento.query.filter_by(
                usuario_id=usuario_id
            ).delete(synchronize_session=False)

            ActivoOperador.query.filter_by(
                usuario_id=usuario_id
            ).delete(synchronize_session=False)

            ControlDiario.query.filter_by(
                usuario_id=usuario_id
            ).delete(synchronize_session=False)

            Tanqueo.query.filter_by(
                usuario_id=usuario_id
            ).delete(synchronize_session=False)

            for inspeccion in Inspeccion.query.filter_by(
                usuario_id=usuario_id
            ).all():
                db.session.delete(inspeccion)

            # ---------------------------------------------
            # 6) USUARIO
            # ---------------------------------------------

            db.session.delete(usuario)

            db.session.commit()

            for archivo in archivos_a_eliminar:
                UsuarioService._eliminar_archivo(archivo)

        except Exception as exc:

            db.session.rollback()

            return jsonify({
                "success": False,
                "message": (
                    "No fue posible eliminar el operador: "
                    f"{exc}"
                )
            }), 500

        return jsonify({
            "success": True,
            "message": "Operador eliminado definitivamente."
        })
