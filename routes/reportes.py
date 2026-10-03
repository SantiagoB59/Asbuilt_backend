from flask import (
    Blueprint,
    jsonify,
    request,
    send_file
)
from openpyxl.drawing.image import Image
from datetime import (
    datetime,
    timedelta,
    date,
)

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)
from openpyxl.drawing.image import Image
import os

from models import (
    Alerta,
    Vehiculo,
    Mantenimiento
)
from models import (
    Mantenimiento,
    MaquinariaMantenimiento,
    Vehiculo,
    Maquinaria,
    PlanItem,
    VehiculoPlanItem,
    MaquinariaPlanItem
)

reportes_bp = Blueprint(
    'reportes',
    __name__
)

# =====================================================
# HELPERS FECHAS
# =====================================================

def obtener_rango_fechas():

    tipo = request.args.get(
        'tipo',
        'mensual'
    )

    hoy = datetime.now()

    # =========================================
    # DIARIO
    # =========================================

    if tipo == 'diario':

        inicio = hoy.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

        fin = hoy

    # =========================================
    # SEMANAL
    # =========================================

    elif tipo == 'semanal':

        inicio = hoy - timedelta(days=7)

        fin = hoy

    # =========================================
    # MENSUAL
    # =========================================

    elif tipo == 'mensual':

        inicio = hoy - timedelta(days=30)

        fin = hoy

    # =========================================
    # RANGO PERSONALIZADO
    # =========================================

    else:

        fecha_inicio = request.args.get(
            'fecha_inicio'
        )

        fecha_fin = request.args.get(
            'fecha_fin'
        )

        inicio = datetime.strptime(
            fecha_inicio,
            '%Y-%m-%d'
        )

        fin = datetime.strptime(
            fecha_fin,
            '%Y-%m-%d'
        )

    return inicio, fin


# =====================================================
# ESTILOS EXCEL
# =====================================================

header_fill = PatternFill(
    start_color='1E3A8A',
    end_color='1E3A8A',
    fill_type='solid'
)

title_fill = PatternFill(
    start_color='2563EB',
    end_color='2563EB',
    fill_type='solid'
)

white_font = Font(
    color='FFFFFF',
    bold=True
)

title_font = Font(
    color='FFFFFF',
    bold=True,
    size=18
)

subtitle_font = Font(
    color='FFFFFF',
    bold=True,
    size=13
)

center = Alignment(
    horizontal='center',
    vertical='center'
)

thin = Side(
    border_style='thin',
    color='D1D5DB'
)

border = Border(
    left=thin,
    right=thin,
    top=thin,
    bottom=thin
)


# =====================================================
# REPORTE ALERTAS JSON
# =====================================================
from datetime import datetime, time
from flask import request, jsonify
from models import Maquinaria

@reportes_bp.route('/alertas', methods=['GET'])
def reporte_alertas():

    inicio, fin = obtener_rango_fechas()

    # =========================================
    # 🔥 NORMALIZAR RANGO (CLAVE DEL PROBLEMA)
    # =========================================
    if hasattr(inicio, "date"):
        inicio = datetime.combine(inicio.date(), time.min)
    else:
        inicio = datetime.combine(inicio, time.min)

    if hasattr(fin, "date"):
        fin = datetime.combine(fin.date(), time.max)
    else:
        fin = datetime.combine(fin, time.max)

    categoria = request.args.get('categoria')
    vehiculo_id = request.args.get('vehiculo_id')
    maquinaria_id = request.args.get('maquinaria_id')

    query = Alerta.query.filter(
        Alerta.created_at >= inicio,
        Alerta.created_at <= fin
    )

    # =========================================
    # FILTRO CATEGORÍA
    # =========================================
    if categoria:
        query = query.filter(
            Alerta.tipo == categoria
        )

    # =========================================
    # FILTRO VEHÍCULO
    # =========================================
    if vehiculo_id:
        query = query.filter(
            Alerta.vehiculo_id == vehiculo_id
        )
    if maquinaria_id:
        query = query.filter(
            Alerta.maquinaria_id == maquinaria_id
        )

    alertas = query.all()


    data = []

    for alerta in alertas:

        vehiculo = None
        maquinaria = None

        if alerta.vehiculo_id:
            vehiculo = Vehiculo.query.get(alerta.vehiculo_id)

        if alerta.maquinaria_id:
            maquinaria = Maquinaria.query.get(alerta.maquinaria_id)

        data.append({
            'id': alerta.id,
            'vehiculo': (
                vehiculo.placa
                if vehiculo
                else maquinaria.codigo if maquinaria else None
            ),
            'tipo': alerta.tipo,
            'categoria': alerta.categoria,
            'prioridad': alerta.prioridad,
            'estado': alerta.estado,
            'mensaje': alerta.mensaje,
            'fecha': alerta.created_at.isoformat() if alerta.created_at else None
        })

    return jsonify(data)
# =====================================================
# EXPORTAR ALERTAS EXCEL
# =====================================================
@reportes_bp.route(
    '/alertas/excel',
    methods=['GET']
)
def exportar_alertas_excel():

    inicio, fin = obtener_rango_fechas()

    categoria = request.args.get(
        'categoria'
    )

    vehiculo_id = request.args.get(
        'vehiculo_id'
    )

    query = Alerta.query.filter(
        Alerta.created_at >= inicio,
        Alerta.created_at <= fin
    )

    # =========================================
    # FILTRO CATEGORÍA
    # =========================================

    if categoria:

        query = query.filter(
            Alerta.tipo == categoria
        )

    # =========================================
    # FILTRO VEHÍCULO
    # =========================================

    if vehiculo_id:

        query = query.filter(
            Alerta.vehiculo_id == vehiculo_id
        )

    alertas = query.all()

    wb = Workbook()

    ws = wb.active

    ws.title = 'Alertas'

    # =================================================
    # HEADER EMPRESA
    # =================================================

    ws.merge_cells('A1:G1')

    ws['A1'] = 'REPORTE CORPORATIVO DE ALERTAS INTELLIFEET'

    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = center

    ws.merge_cells('A2:G2')

    ws['A2'] = 'REPORTE CORPORATIVO DE ALERTAS'

    ws['A2'].font = subtitle_font
    ws['A2'].fill = header_fill
    ws['A2'].alignment = center

    ws.merge_cells('A3:G3')

    ws['A3'] = (
        f'Generado: '
        f'{datetime.now().strftime("%Y-%m-%d %H:%M")}'
    )

    ws['A3'].alignment = center

    # =================================================
    # ENCABEZADOS
    # =================================================

    headers = [

        'ID',
        'VEHÍCULO',
        'TIPO',
        'CATEGORÍA',
        'PRIORIDAD',
        'ESTADO',
        'FECHA'
    ]

    row_num = 5

    for col_num, header in enumerate(headers, 1):

        cell = ws.cell(
            row=row_num,
            column=col_num
        )

        cell.value = header
        cell.font = white_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # =================================================
    # DATA
    # =================================================

    current_row = 6

    for alerta in alertas:

        vehiculo = None
        maquinaria = None

        if alerta.vehiculo_id:
            vehiculo = Vehiculo.query.get(alerta.vehiculo_id)

        if alerta.maquinaria_id:
            maquinaria = Maquinaria.query.get(alerta.maquinaria_id)

        # Si la alerta está resuelta, en el Excel la prioridad
        # aparecerá como RESUELTA.
        prioridad_excel = (
            'RESUELTA'
            if alerta.estado == 'RESUELTA'
            else alerta.prioridad
        )

        values = [

            alerta.id,

            vehiculo.placa
            if vehiculo
            else maquinaria.codigo if maquinaria else 'N/A',

            alerta.tipo,

            alerta.categoria,

            prioridad_excel,

            alerta.estado,

            alerta.created_at.strftime(
                '%Y-%m-%d %H:%M'
            ) if alerta.created_at else ''
        ]

        for col_num, value in enumerate(values, 1):

            cell = ws.cell(
                row=current_row,
                column=col_num
            )

            cell.value = value
            cell.border = border
            cell.alignment = center

            # =========================================
            # COLOR PRIORIDAD
            # =========================================

            if col_num == 5:

                if value == 'RESUELTA':

                    cell.fill = PatternFill(
                        start_color='86EFAC',
                        end_color='86EFAC',
                        fill_type='solid'
                    )

                elif value == 'CRITICA':

                    cell.fill = PatternFill(
                        start_color='FCA5A5',
                        end_color='FCA5A5',
                        fill_type='solid'
                    )

                elif value == 'ALTA':

                    cell.fill = PatternFill(
                        start_color='FDBA74',
                        end_color='FDBA74',
                        fill_type='solid'
                    )

                elif value == 'MEDIA':

                    cell.fill = PatternFill(
                        start_color='FDE68A',
                        end_color='FDE68A',
                        fill_type='solid'
                    )

                elif value == 'BAJA':

                    cell.fill = PatternFill(
                        start_color='86EFAC',
                        end_color='86EFAC',
                        fill_type='solid'
                    )

            # =========================================
            # COLOR ESTADO
            # =========================================

            if col_num == 6:

                if value == 'RESUELTA':

                    cell.fill = PatternFill(
                        start_color='86EFAC',
                        end_color='86EFAC',
                        fill_type='solid'
                    )

                elif value == 'PENDIENTE':

                    cell.fill = PatternFill(
                        start_color='FCA5A5',
                        end_color='FCA5A5',
                        fill_type='solid'
                    )

        current_row += 1

    # =================================================
    # TAMAÑO COLUMNAS
    # =================================================

    columnas = {

        'A': 10,
        'B': 22,
        'C': 20,
        'D': 20,
        'E': 15,
        'F': 15,
        'G': 25
    }

    for col, width in columnas.items():

        ws.column_dimensions[col].width = width

    # =================================================
    # GENERAR ARCHIVO
    # =================================================

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(

        output,

        download_name=
            f'REPORTE_ALERTAS_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',

        as_attachment=True,

        mimetype=
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
# =====================================================
# SEMÁFORO ALERTAS
# =====================================================

@reportes_bp.route(
    '/semaforo-alertas',
    methods=['GET']
)
def semaforo_alertas():

    criticas = Alerta.query.filter_by(
        prioridad='CRITICA',
        estado='ACTIVA'
    ).count()

    altas = Alerta.query.filter_by(
        prioridad='ALTA',
        estado='ACTIVA'
    ).count()

    medias = Alerta.query.filter_by(
        prioridad='MEDIA',
        estado='ACTIVA'
    ).count()

    bajas = Alerta.query.filter_by(
        prioridad='BAJA',
        estado='ACTIVA'
    ).count()

    total = (
        criticas
        + altas
        + medias
        + bajas
    )

    return jsonify({

        'total': total,

        'criticas': criticas,

        'altas': altas,

        'medias': medias,

        'bajas': bajas
    })

# =====================================================
# REPORTE MANTENIMIENTOS
# =====================================================
@reportes_bp.route(
    '/mantenimientos',
    methods=['GET']
)
def reporte_mantenimientos():

    inicio, fin = obtener_rango_fechas()

    vehiculo_id = request.args.get('vehiculo_id')
    maquinaria_id = request.args.get('maquinaria_id')
    tipo_activo = request.args.get('tipo_activo')

    data = []

    # =========================================
    # MANTENIMIENTOS VEHÍCULOS
    # =========================================

    if tipo_activo != 'MAQUINARIA':

        query = (
            Mantenimiento.query
            .filter(
                Mantenimiento.fecha >= inicio,
                Mantenimiento.fecha <= fin
            )
        )

        if vehiculo_id:
            query = query.filter(
                Mantenimiento.vehiculo_id == vehiculo_id
            )

        mantenimientos = query.all()

        for m in mantenimientos:

            vehiculo = Vehiculo.query.get(m.vehiculo_id)

            data.append({

                'id': m.id,
                'tipo_activo': 'VEHICULO',
                'vehiculo': vehiculo.placa if vehiculo else None,
                'tipo': m.tipo,
                'fecha': str(m.fecha),
                'km': m.km,
                'valor': getattr(m, 'valor', None),
                'observaciones': m.observaciones

            })

    # =========================================
    # MANTENIMIENTOS MAQUINARIA
    # =========================================

    if tipo_activo != 'VEHICULO':

        query = (
            MaquinariaMantenimiento.query
            .filter(
                MaquinariaMantenimiento.fecha >= inicio,
                MaquinariaMantenimiento.fecha <= fin
            )
        )

        if maquinaria_id:
            query = query.filter(
                MaquinariaMantenimiento.maquinaria_id == maquinaria_id
            )

        mantenimientos_maquinaria = query.all()

        for m in mantenimientos_maquinaria:

            maquinaria = Maquinaria.query.get(m.maquinaria_id)

            data.append({

                'id': m.id,
                'tipo_activo': 'MAQUINARIA',
                'vehiculo': maquinaria.codigo if maquinaria else None,
                'tipo': m.tipo,
                'fecha': str(m.fecha),
                'km': m.horas,
                'valor': m.costo,
                'observaciones': m.observaciones

            })

    # =========================================
    # ORDENAR POR FECHA DESCENDENTE
    # =========================================

    data.sort(
        key=lambda x: x['fecha'],
        reverse=True
    )

    return jsonify(data)
# =====================================================
# FORMATO PROFESIONAL MANTENIMIENTO VEHÍCULO
# =====================================================



@reportes_bp.route(
    '/mantenimiento-formato/<int:vehiculo_id>',
    methods=['GET']
)
def exportar_formato_mantenimiento(vehiculo_id):

    from io import BytesIO
    import os

    from flask import send_file

    from openpyxl import Workbook
    from openpyxl.drawing.image import Image
    from openpyxl.styles import (
        Font,
        PatternFill,
        Border,
        Side,
        Alignment
    )
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.utils.cell import range_boundaries

    # ============================================================
    # VEHÍCULO
    # ============================================================

    vehiculo = Vehiculo.query.get_or_404(
        vehiculo_id
    )

    # ============================================================
    # MANTENIMIENTOS
    # ============================================================

    mantenimientos = (
        Mantenimiento.query
        .filter(
            Mantenimiento.vehiculo_id == vehiculo_id
        )
        .order_by(
            Mantenimiento.fecha.desc()
        )
        .all()
    )

    # ============================================================
    # EXCEL
    # ============================================================

    wb = Workbook()

    ws = wb.active

    ws.title = 'R. MANTENIMIENTO'

    # Ocultar cuadrícula
    ws.sheet_view.showGridLines = False

    # ============================================================
    # CONFIGURACIÓN DE PÁGINA
    # ============================================================

    ws.page_setup.orientation = 'landscape'

    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    ws.page_setup.fitToWidth = 1

    ws.page_setup.fitToHeight = 0

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.35,
        bottom=0.35,
        header=0.15,
        footer=0.15
    )

    ws.print_options.horizontalCentered = True

    # ============================================================
    # PALETA CORPORATIVA INTELLIFEET
    # ============================================================

    NAVY = '0B1F33'

    NAVY_2 = '132F4C'

    GOLD = 'B58A32'

    GOLD_LIGHT = 'F3E8C8'

    BLUE_LIGHT = 'EAF2F8'

    BLUE_SOFT = 'DCEAF5'

    WHITE = 'FFFFFF'

    GRAY_LIGHT = 'F5F7FA'

    GRAY = 'E8ECF1'

    GRAY_TEXT = '475569'

    DARK = '17202A'

    GREEN = '15803D'

    GREEN_LIGHT = 'DCFCE7'

    RED = 'B91C1C'

    RED_LIGHT = 'FEE2E2'

    ORANGE = 'B45309'

    ORANGE_LIGHT = 'FEF3C7'

    # ============================================================
    # ANCHOS DE COLUMNAS
    # ============================================================

    columnas = {

        'A': 13,   # Fecha

        'B': 15,   # Kilometraje

        'C': 10,   # Sistema

        'D': 20,   # Descripción

        'E': 20,

        'F': 20,

        'G': 20,

        'H': 18,   # Insumos

        'I': 18,

        'J': 16,   # Responsable

        'K': 14,   # Preventivo

        'L': 14,   # Correctivo

        'M': 28    # Soporte

    }

    for col, width in columnas.items():

        ws.column_dimensions[
            col
        ].width = width

    # ============================================================
    # ESTILOS DE FONDO
    # ============================================================

    fill_navy = PatternFill(
        fill_type='solid',
        fgColor=NAVY
    )

    fill_navy_2 = PatternFill(
        fill_type='solid',
        fgColor=NAVY_2
    )

    fill_gold = PatternFill(
        fill_type='solid',
        fgColor=GOLD
    )

    fill_gold_light = PatternFill(
        fill_type='solid',
        fgColor=GOLD_LIGHT
    )

    fill_blue_light = PatternFill(
        fill_type='solid',
        fgColor=BLUE_LIGHT
    )

    fill_blue_soft = PatternFill(
        fill_type='solid',
        fgColor=BLUE_SOFT
    )

    fill_gray = PatternFill(
        fill_type='solid',
        fgColor=GRAY_LIGHT
    )

    fill_white = PatternFill(
        fill_type='solid',
        fgColor=WHITE
    )

    fill_green = PatternFill(
        fill_type='solid',
        fgColor=GREEN_LIGHT
    )

    fill_red = PatternFill(
        fill_type='solid',
        fgColor=RED_LIGHT
    )

    fill_orange = PatternFill(
        fill_type='solid',
        fgColor=ORANGE_LIGHT
    )

    # ============================================================
    # BORDES
    # ============================================================

    thin_gray = Side(
        border_style='thin',
        color='CBD5E1'
    )

    medium_navy = Side(
        border_style='medium',
        color=NAVY
    )

    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray
    )

    border_navy = Border(
        left=medium_navy,
        right=medium_navy,
        top=medium_navy,
        bottom=medium_navy
    )

    # ============================================================
    # ALINEACIONES
    # ============================================================

    center = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )

    left = Alignment(
        horizontal='left',
        vertical='center',
        wrap_text=True
    )

    # ============================================================
    # FUENTES
    # ============================================================

    font_title = Font(
        name='Aptos Display',
        bold=True,
        size=16,
        color=WHITE
    )

    font_subtitle = Font(
        name='Aptos',
        bold=True,
        size=11,
        color=WHITE
    )

    font_section = Font(
        name='Aptos',
        bold=True,
        size=10,
        color=WHITE
    )

    font_label = Font(
        name='Aptos',
        bold=True,
        size=9,
        color=NAVY
    )

    font_value = Font(
        name='Aptos',
        size=10,
        color=DARK
    )

    font_header = Font(
        name='Aptos',
        bold=True,
        size=9,
        color=WHITE
    )

    font_normal = Font(
        name='Aptos',
        size=9,
        color=DARK
    )

    font_km = Font(
        name='Aptos Display',
        bold=True,
        size=10,
        color=NAVY
    )

    # ============================================================
    # FUNCIÓN PARA BORDES COMPLETOS EN RANGOS COMBINADOS
    # ============================================================

    def aplicar_borde_rango(
        hoja,
        rango,
        border_style
    ):

        min_col, min_row, max_col, max_row = (
            range_boundaries(rango)
        )

        for row in range(
            min_row,
            max_row + 1
        ):

            for col in range(
                min_col,
                max_col + 1
            ):

                cell = hoja.cell(
                    row=row,
                    column=col
                )

                left_border = (
                    border_style.left
                    if col == min_col
                    else Side(style=None)
                )

                right_border = (
                    border_style.right
                    if col == max_col
                    else Side(style=None)
                )

                top_border = (
                    border_style.top
                    if row == min_row
                    else Side(style=None)
                )

                bottom_border = (
                    border_style.bottom
                    if row == max_row
                    else Side(style=None)
                )

                cell.border = Border(
                    left=left_border,
                    right=right_border,
                    top=top_border,
                    bottom=bottom_border
                )

    # ============================================================
    # ALTURAS DE FILAS
    # ============================================================

    for i in range(1, 60):

        ws.row_dimensions[
            i
        ].height = 25

    ws.row_dimensions[1].height = 48

    ws.row_dimensions[2].height = 28

    ws.row_dimensions[3].height = 28

    ws.row_dimensions[4].height = 6

    ws.row_dimensions[5].height = 26

    ws.row_dimensions[14].height = 38

    ws.row_dimensions[15].height = 38

    # ============================================================
    # CABECERA CORPORATIVA
    # ============================================================

    for row in range(1, 4):

        for col in range(1, 14):

            ws.cell(
                row=row,
                column=col
            ).fill = fill_navy

    # ============================================================
    # LOGO
    # ============================================================

    try:

        rutas_logo = [

            os.path.join(
                os.getcwd(),
                'static',
                'intellifeet.png'
            ),

            os.path.join(
                os.path.dirname(__file__),
                '..',
                'static',
                'intellifeet.png'
            ),

            os.path.join(
                os.path.dirname(__file__),
                'static',
                'intellifeet.png'
            )

        ]

        ruta_logo = next(
            (
                os.path.abspath(ruta)
                for ruta in rutas_logo
                if os.path.exists(ruta)
            ),
            None
        )

        if ruta_logo:

            logo = Image(
                ruta_logo
            )

            # ----------------------------------------------------
            # LOGO MÁS GRANDE
            # ----------------------------------------------------

            max_width = 230

            max_height = 90

            if logo.width and logo.height:

                escala = min(
                    max_width / logo.width,
                    max_height / logo.height
                )

                logo.width = int(
                    logo.width * escala
                )

                logo.height = int(
                    logo.height * escala
                )

            ws.add_image(
                logo,
                'A1'
            )

    except Exception as e:

        print(
            'No se pudo cargar logo IntelliFeet:',
            e
        )

    # ============================================================
    # TÍTULO
    # ============================================================

    ws.merge_cells(
        'C1:J1'
    )

    ws['C1'] = (
        'GESTIÓN DIRECCIÓN DE PROYECTOS'
    )

    ws['C1'].font = font_title

    ws['C1'].alignment = center

    ws['C1'].fill = fill_navy

    # ============================================================
    # SUBTÍTULO
    # ============================================================

    ws.merge_cells(
        'C2:J2'
    )

    ws['C2'] = (
        'REPORTE DE MANTENIMIENTO VEHICULAR'
    )

    ws['C2'].font = font_subtitle

    ws['C2'].alignment = center

    ws['C2'].fill = fill_navy

    # ============================================================
    # DESCRIPCIÓN
    # ============================================================

    ws.merge_cells(
        'C3:J3'
    )

    ws['C3'] = (
        'CONTROL DE MANTENIMIENTO • '
        'TRAZABILIDAD OPERACIONAL'
    )

    ws['C3'].font = Font(
        name='Aptos',
        bold=True,
        size=9,
        color=GOLD_LIGHT
    )

    ws['C3'].alignment = center

    ws['C3'].fill = fill_navy

    # ============================================================
    # CONTROL DOCUMENTAL
    # ============================================================

    info_header = [

        (
            'K1:M1',
            'VERSIÓN: 006'
        ),

        (
            'K2:M2',
            'CÓDIGO: PROY-R-022'
        ),

        (
            'K3:M3',
            'PÁGINA: 1 DE 1'
        )

    ]

    for rango, texto in info_header:

        ws.merge_cells(
            rango
        )

        cell = rango.split(':')[0]

        ws[cell] = texto

        ws[cell].fill = fill_navy_2

        ws[cell].font = Font(
            name='Aptos',
            bold=True,
            size=8,
            color=WHITE
        )

        ws[cell].alignment = center

        aplicar_borde_rango(
            ws,
            rango,
            border
        )

    # ============================================================
    # LÍNEA DORADA
    # ============================================================

    for col in range(1, 14):

        ws.cell(
            row=4,
            column=col
        ).fill = fill_gold

    # ============================================================
    # INFORMACIÓN GENERAL
    # ============================================================

    rango_info = 'A5:M5'

    ws.merge_cells(
        rango_info
    )

    ws['A5'] = (
        'INFORMACIÓN GENERAL DEL VEHÍCULO'
    )

    ws['A5'].fill = fill_navy

    ws['A5'].font = font_section

    ws['A5'].alignment = center

    aplicar_borde_rango(
        ws,
        rango_info,
        border_navy
    )

    # ============================================================
    # TIPO VEHÍCULO
    # ============================================================

    tipo_vehiculo = ''

    if vehiculo.tipo_vehiculo:

        tipo_vehiculo = getattr(
            vehiculo.tipo_vehiculo,
            'nombre',
            str(
                vehiculo.tipo_vehiculo
            )
        )

    # ============================================================
    # DATOS VEHÍCULO
    # ============================================================

    datos = [

        (
            'CLASE VEHÍCULO',
            tipo_vehiculo
        ),

        (
            'MODELO',
            str(
                vehiculo.modelo or ''
            )
        ),

        (
            'PLACA',
            str(
                vehiculo.placa or ''
            )
        ),

        (
            'MARCA',
            str(
                vehiculo.marca or ''
            )
        ),

        (
            'KILOMETRAJE ACTUAL',
            (
                f"{float(vehiculo.km_actual):,.0f} km"
                if vehiculo.km_actual is not None
                else ''
            )
        ),

        (
            'ESTADO',
            str(
                getattr(
                    vehiculo,
                    'estado',
                    'ACTIVO'
                )
                or 'ACTIVO'
            )
        )

    ]

    posiciones = [

        (
            'A7:B7',
            'C7:E7'
        ),

        (
            'F7:G7',
            'H7:M7'
        ),

        (
            'A9:B9',
            'C9:E9'
        ),

        (
            'F9:G9',
            'H9:M9'
        ),

        (
            'A11:B11',
            'C11:E11'
        ),

        (
            'F11:G11',
            'H11:M11'
        )

    ]

    for i, (
        label_pos,
        value_pos
    ) in enumerate(
        posiciones
    ):

        label, value = datos[i]

        # --------------------------------------------------------
        # LABEL
        # --------------------------------------------------------

        ws.merge_cells(
            label_pos
        )

        label_cell = (
            label_pos.split(':')[0]
        )

        ws[label_cell] = label

        ws[label_cell].fill = (
            fill_gold_light
        )

        ws[label_cell].font = (
            font_label
        )

        ws[label_cell].alignment = (
            center
        )

        aplicar_borde_rango(
            ws,
            label_pos,
            border
        )

        # --------------------------------------------------------
        # VALOR
        # --------------------------------------------------------

        ws.merge_cells(
            value_pos
        )

        value_cell = (
            value_pos.split(':')[0]
        )

        ws[value_cell] = value

        ws[value_cell].fill = (
            fill_gray
        )

        ws[value_cell].font = (
            font_value
        )

        ws[value_cell].alignment = (
            center
        )

        aplicar_borde_rango(
            ws,
            value_pos,
            border
        )

    # ============================================================
    # ENCABEZADO TABLA
    # ============================================================

    headers = [

        'FECHA',

        'KILOMETRAJE\nDEL MANTENIMIENTO',

        'SISTEMA',

        'DESCRIPCIÓN DETALLADA DEL\nMANTENIMIENTO',

        'INSUMOS Y\nREPUESTOS',

        'ENTIDAD Y/O\nRESPONSABLE',

        'MANTENIMIENTO\nPREVENTIVO',

        'MANTENIMIENTO\nCORRECTIVO',

        'SOPORTE'

    ]

    merges = [

        'A14:A15',

        'B14:B15',

        'C14:C15',

        'D14:G15',

        'H14:I15',

        'J14:J15',

        'K14:K15',

        'L14:L15',

        'M14:M15'

    ]

    for i, merge in enumerate(
        merges
    ):

        ws.merge_cells(
            merge
        )

        cell = merge.split(':')[0]

        ws[cell] = headers[i]

        ws[cell].fill = fill_navy_2

        ws[cell].font = font_header

        ws[cell].alignment = center

        aplicar_borde_rango(
            ws,
            merge,
            border_navy
        )

    # ============================================================
    # CONVENCIONES DE SISTEMAS
    # ============================================================

    mapa_sistemas = {

        'SISTEMA DE LUBRICACIÓN': 'SL',

        'SISTEMA DE COMBUSTIBLE': 'SC',

        'SISTEMA ELÉCTRICO': 'SEL',

        'SISTEMA DE FRENOS': 'SF',

        'SISTEMA DE TRANSMISIÓN': 'ST',

        'SISTEMA DE DIRECCIÓN': 'SD',

        'SISTEMA DE MOTOR': 'SM',

        'SISTEMA DE SUSPENSIÓN': 'SS',

        'SISTEMA DE ESCAPE': 'SES',

        'SISTEMA DE LLANTAS': 'SLL'

    }

    # ============================================================
    # TABLA MANTENIMIENTOS
    # ============================================================

    fila_actual = 16

    for m in mantenimientos[:18]:

        # ========================================================
        # FECHA
        # ========================================================

        if m.fecha:

            if hasattr(
                m.fecha,
                'strftime'
            ):

                ws[
                    f'A{fila_actual}'
                ] = (
                    m.fecha.strftime(
                        '%d/%m/%Y'
                    )
                )

            else:

                ws[
                    f'A{fila_actual}'
                ] = str(
                    m.fecha
                )

        else:

            ws[
                f'A{fila_actual}'
            ] = ''

        # ========================================================
        # KILOMETRAJE DEL MANTENIMIENTO
        # ========================================================

        km_mantenimiento = getattr(
            m,
            'km',
            None
        )

        # --------------------------------------------------------
        # RESPALDO
        # --------------------------------------------------------

        if km_mantenimiento is None:

            km_mantenimiento = getattr(
                m,
                'kilometraje',
                None
            )

        if km_mantenimiento is not None:

            try:

                # Guardamos como número REAL de Excel
                ws[
                    f'B{fila_actual}'
                ] = float(
                    km_mantenimiento
                )

                # Formato visual
                ws[
                    f'B{fila_actual}'
                ].number_format = (
                    '#,##0 "km"'
                )

            except Exception:

                ws[
                    f'B{fila_actual}'
                ] = str(
                    km_mantenimiento
                )

        else:

            ws[
                f'B{fila_actual}'
            ] = ''

        # ========================================================
        # SISTEMA
        # ========================================================

        sistema_nombre = getattr(
            m.plan_item,
            'sistema',
            ''
        ) or ''

        sistema_key = str(
            sistema_nombre
        ).strip().upper()

        ws[
            f'C{fila_actual}'
        ] = (
            mapa_sistemas.get(
                sistema_key,
                sistema_nombre
            )
        )

        # ========================================================
        # DESCRIPCIÓN
        # ========================================================

        rango_descripcion = (
            f'D{fila_actual}:G{fila_actual}'
        )

        ws.merge_cells(
            rango_descripcion
        )

        descripcion = ''

        if getattr(
            m.plan_item,
            'nombre',
            None
        ):

            descripcion += str(
                m.plan_item.nombre
            )

        if getattr(
            m,
            'observaciones',
            None
        ):

            if descripcion:

                descripcion += '\n\n'

            descripcion += (
                'Observaciones: '
                f'{m.observaciones}'
            )

        ws[
            f'D{fila_actual}'
        ] = descripcion

        aplicar_borde_rango(
            ws,
            rango_descripcion,
            border
        )

        # ========================================================
        # INSUMOS / REPUESTOS
        # ========================================================

        rango_insumos = (
            f'H{fila_actual}:I{fila_actual}'
        )

        ws.merge_cells(
            rango_insumos
        )

        ws[
            f'H{fila_actual}'
        ] = (
            str(
                getattr(
                    m,
                    'proveedor',
                    ''
                ) or ''
            )
        )

        aplicar_borde_rango(
            ws,
            rango_insumos,
            border
        )

        # ========================================================
        # RESPONSABLE
        # ========================================================

        ws[
            f'J{fila_actual}'
        ] = (
            str(
                getattr(
                    m,
                    'responsable',
                    ''
                ) or ''
            )
        )

        # ========================================================
        # TIPO DE MANTENIMIENTO
        # ========================================================

        tipo_mantenimiento = str(
            getattr(
                m,
                'type',
                getattr(
                    m,
                    'tipo',
                    ''
                )
            ) or ''
        ).upper()

        if tipo_mantenimiento == 'PREVENTIVO':

            ws[
                f'K{fila_actual}'
            ] = '✓'

            ws[
                f'K{fila_actual}'
            ].fill = fill_green

            ws[
                f'K{fila_actual}'
            ].font = Font(
                name='Aptos',
                bold=True,
                size=14,
                color=GREEN
            )

        elif tipo_mantenimiento == 'CORRECTIVO':

            ws[
                f'L{fila_actual}'
            ] = '✓'

            ws[
                f'L{fila_actual}'
            ].fill = fill_orange

            ws[
                f'L{fila_actual}'
            ].font = Font(
                name='Aptos',
                bold=True,
                size=14,
                color=ORANGE
            )

        # ========================================================
        # SOPORTE
        # ========================================================

        if m.soporte:

            ruta_imagen = os.path.join(
                os.getcwd(),
                m.soporte
            )

            if os.path.exists(
                ruta_imagen
            ):

                try:

                    img = Image(
                        ruta_imagen
                    )

                    # Mantener tamaño profesional
                    img.width = 180
                    img.height = 140

                    ws.add_image(
                        img,
                        f'M{fila_actual}'
                    )

                except Exception as e:

                    print(
                        'Error cargando soporte:',
                        e
                    )

            else:

                ws[
                    f'M{fila_actual}'
                ] = (
                    'Archivo no disponible'
                )

        # ========================================================
        # ESTILOS GENERALES DE FILA
        # ========================================================

        for col in range(
            1,
            14
        ):

            cell = ws.cell(
                row=fila_actual,
                column=col
            )

            # ----------------------------------------------------
            # Borde general
            # ----------------------------------------------------

            if cell.border == Border():

                cell.border = border

            # ----------------------------------------------------
            # Alineación
            # ----------------------------------------------------

            cell.alignment = center

            # ----------------------------------------------------
            # Fuente
            # ----------------------------------------------------

            if not cell.font.bold:

                cell.font = font_normal

            # ----------------------------------------------------
            # Fondo alternado
            # ----------------------------------------------------

            if (
                fila_actual % 2 == 0
                and cell.fill.fill_type is None
            ):

                cell.fill = fill_gray

        # ========================================================
        # DESCRIPCIÓN A LA IZQUIERDA
        # ========================================================

        ws[
            f'D{fila_actual}'
        ].alignment = left

        # ========================================================
        # KILOMETRAJE
        # ========================================================

        ws[
            f'B{fila_actual}'
        ].alignment = center

        ws[
            f'B{fila_actual}'
        ].font = font_km

        # ========================================================
        # FECHA
        # ========================================================

        ws[
            f'A{fila_actual}'
        ].alignment = center

        # ========================================================
        # ALTURA
        # ========================================================

        ws.row_dimensions[
            fila_actual
        ].height = 110

        fila_actual += 1

    # ============================================================
    # FILAS VACÍAS
    # ============================================================

    while fila_actual <= 34:

        for col in range(
            1,
            14
        ):

            cell = ws.cell(
                row=fila_actual,
                column=col
            )

            cell.border = border

            cell.alignment = center

            cell.font = font_normal

            if (
                fila_actual % 2 == 0
            ):

                cell.fill = fill_gray

        ws.row_dimensions[
            fila_actual
        ].height = 30

        fila_actual += 1

    # ============================================================
    # TABLA DE CONVENCIONES
    # ============================================================

    rango_convenciones = 'A36:M36'

    ws.merge_cells(
        rango_convenciones
    )

    ws['A36'] = (
        'TABLA DE CONVENCIONES • '
        'SISTEMAS DE MANTENIMIENTO'
    )

    ws['A36'].fill = fill_navy

    ws['A36'].font = font_section

    ws['A36'].alignment = center

    aplicar_borde_rango(
        ws,
        rango_convenciones,
        border_navy
    )

    # ============================================================
    # CONVENCIONES IZQUIERDA
    # ============================================================

    convenciones = [

        (
            'SISTEMA DE LUBRICACIÓN',
            'SL'
        ),

        (
            'SISTEMA DE COMBUSTIBLE',
            'SC'
        ),

        (
            'SISTEMA ELÉCTRICO',
            'SEL'
        ),

        (
            'SISTEMA DE FRENOS',
            'SF'
        ),

        (
            'SISTEMA DE TRANSMISIÓN',
            'ST'
        )

    ]

    fila_conv = 37

    for nombre, sigla in convenciones:

        rango_nombre = (
            f'A{fila_conv}:C{fila_conv}'
        )

        ws.merge_cells(
            rango_nombre
        )

        ws[
            f'A{fila_conv}'
        ] = nombre

        ws[
            f'D{fila_conv}'
        ] = sigla

        ws[
            f'A{fila_conv}'
        ].fill = fill_gold_light

        ws[
            f'D{fila_conv}'
        ].fill = fill_gray

        ws[
            f'A{fila_conv}'
        ].font = font_label

        ws[
            f'D{fila_conv}'
        ].font = Font(
            name='Aptos',
            bold=True,
            size=9,
            color=NAVY
        )

        ws[
            f'A{fila_conv}'
        ].alignment = center

        ws[
            f'D{fila_conv}'
        ].alignment = center

        aplicar_borde_rango(
            ws,
            rango_nombre,
            border
        )

        ws[
            f'D{fila_conv}'
        ].border = border

        fila_conv += 1

    # ============================================================
    # CONVENCIONES DERECHA
    # ============================================================

    convenciones2 = [

        (
            'SISTEMA DE DIRECCIÓN',
            'SD'
        ),

        (
            'SISTEMA DE MOTOR',
            'SM'
        ),

        (
            'SISTEMA DE SUSPENSIÓN',
            'SS'
        ),

        (
            'SISTEMA DE ESCAPE',
            'SES'
        ),

        (
            'SISTEMA DE LLANTAS',
            'SLL'
        )

    ]

    fila_conv = 37

    for nombre, sigla in convenciones2:

        rango_nombre = (
            f'F{fila_conv}:H{fila_conv}'
        )

        ws.merge_cells(
            rango_nombre
        )

        ws[
            f'F{fila_conv}'
        ] = nombre

        ws[
            f'I{fila_conv}'
        ] = sigla

        ws[
            f'F{fila_conv}'
        ].fill = fill_gold_light

        ws[
            f'I{fila_conv}'
        ].fill = fill_gray

        ws[
            f'F{fila_conv}'
        ].font = font_label

        ws[
            f'I{fila_conv}'
        ].font = Font(
            name='Aptos',
            bold=True,
            size=9,
            color=NAVY
        )

        ws[
            f'F{fila_conv}'
        ].alignment = center

        ws[
            f'I{fila_conv}'
        ].alignment = center

        aplicar_borde_rango(
            ws,
            rango_nombre,
            border
        )

        ws[
            f'I{fila_conv}'
        ].border = border

        fila_conv += 1

    # ============================================================
    # PIE DE PÁGINA
    # ============================================================

    ws.oddFooter.center.text = (
        'IntelliFeet • '
        'Gestión de Mantenimiento Vehicular'
    )

    ws.oddFooter.center.size = 8

    ws.oddFooter.center.font = 'Aptos'

    ws.oddFooter.right.text = (
        'Página &[Page] de &[Pages]'
    )

    ws.oddFooter.right.size = 8

    ws.oddFooter.right.font = 'Aptos'

    # ============================================================
    # IMPORTANTE:
    # NO USAMOS freeze_panes
    #
    # De esta forma el encabezado NO queda fijo al hacer scroll.
    # ============================================================

    # ws.freeze_panes = 'A16'

    # ============================================================
    # CONFIGURACIÓN DE IMPRESIÓN
    # ============================================================

    ws.print_area = 'A1:M42'

    # ============================================================
    # EXPORTAR
    # ============================================================

    output = BytesIO()

    wb.save(
        output
    )

    output.seek(0)

    return send_file(

        output,

        download_name=(
            'FORMATO_MANTENIMIENTO_'
            f'{vehiculo.placa}.xlsx'
        ),

        as_attachment=True,

        mimetype=(
            'application/vnd.openxmlformats-'
            'officedocument.spreadsheetml.sheet'
        )
    )

@reportes_bp.route(
    '/alertas-formato/<int:vehiculo_id>',
    methods=['GET']
)
def exportar_formato_alertas(vehiculo_id):

    from io import BytesIO
    from flask import send_file
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image
    from openpyxl.styles import (
        Font,
        PatternFill,
        Border,
        Side,
        Alignment
    )

    vehiculo = Vehiculo.query.get_or_404(vehiculo_id)

    alertas = (
        Alerta.query
        .filter(Alerta.vehiculo_id == vehiculo_id)
        .order_by(Alerta.created_at.desc())
        .all()
    )

    wb = Workbook()
    ws = wb.active

    ws.title = 'R. ALERTAS'
    ws.sheet_view.showGridLines = False

    # =================================================
    # COLUMNAS
    # =================================================

    columnas = {
        'A': 18,
        'B': 18,
        'C': 18,
        'D': 18,
        'E': 16,
        'F': 16,
        'G': 18,
        'H': 12,
        'I': 12
    }

    for col, width in columnas.items():
        ws.column_dimensions[col].width = width

    # =================================================
    # ESTILOS
    # =================================================

    azul_oscuro = PatternFill(start_color='1F4E78', fill_type='solid')
    azul_claro = PatternFill(start_color='D9EAF7', fill_type='solid')
    azul_header = PatternFill(start_color='8DB4E2', fill_type='solid')
    gris = PatternFill(start_color='F2F2F2', fill_type='solid')

    thin = Side(border_style='thin', color='BFBFBF')

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    center = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )

    left = Alignment(
        horizontal='left',
        vertical='center',
        wrap_text=True
    )

    titulo_font = Font(bold=True, color='FFFFFF', size=14)
    subtitulo_font = Font(bold=True, color='FFFFFF', size=11)
    normal = Font(size=10, color='333333')
    white_bold = Font(bold=True, color='FFFFFF', size=10)

    # =================================================
    # ALTURA FILAS
    # =================================================

    for i in range(1, 60):
        ws.row_dimensions[i].height = 28

    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30

    # =================================================
    # HEADER (NO TOCADO)
    # =================================================

    ws.merge_cells('A1:B3')

    for row in ws['A1:B3']:
        for cell in row:
            cell.fill = azul_oscuro
            cell.border = border

    logo = Image('static/intellifeet.png')
    logo.width = 260
    logo.height = 128
    ws.add_image(logo, 'A1')

    ws.merge_cells('C1:G2')

    ws['C1'] = 'GESTIÓN DIRECCIÓN DE PROYECTOS'
    ws['C1'].fill = azul_oscuro
    ws['C1'].font = titulo_font
    ws['C1'].alignment = center
    ws['C1'].border = border

    ws.merge_cells('C3:G3')

    ws['C3'] = 'REPORTE DE ALERTAS VEHICULARES'
    ws['C3'].fill = azul_claro
    ws['C3'].font = Font(bold=True, size=12, color='1F1F1F')
    ws['C3'].alignment = center
    ws['C3'].border = border

    # =================================================
    # INFO HEADER
    # =================================================

    info_header = [
        ('H1:I1', 'VERSIÓN: 001'),
        ('H2:I2', 'CÓDIGO: ALERT-R-001'),
        ('H3:I3', 'PÁGINA: 1 DE 1')
    ]

    for rango, texto in info_header:
        ws.merge_cells(rango)
        cell = rango.split(':')[0]

        ws[cell] = texto
        ws[cell].fill = azul_oscuro
        ws[cell].font = white_bold
        ws[cell].alignment = center
        ws[cell].border = border

    # =================================================
    # VEHÍCULO
    # =================================================

    ws.merge_cells('A5:I5')

    ws['A5'] = 'INFORMACIÓN GENERAL DEL VEHÍCULO'
    ws['A5'].fill = azul_oscuro
    ws['A5'].font = subtitulo_font
    ws['A5'].alignment = center
    ws['A5'].border = border

    datos = [
        ('CLASE VEHÍCULO', getattr(vehiculo.tipo_vehiculo, 'nombre', '')),
        ('MODELO', vehiculo.modelo or ''),
        ('PLACA', vehiculo.placa or ''),
        ('MARCA', vehiculo.marca or ''),
        ('KM ACTUAL', vehiculo.km_actual or ''),
        ('ESTADO', getattr(vehiculo, 'estado', 'ACTIVO'))
    ]

    posiciones = [
        ('A7:B7', 'C7:E7'),
        ('F7:G7', 'H7:I7'),
        ('A9:B9', 'C9:E9'),
        ('F9:G9', 'H9:I9'),
        ('A11:B11', 'C11:E11'),
        ('F11:G11', 'H11:I11')
    ]

    for i, (label_pos, value_pos) in enumerate(posiciones):

        label, value = datos[i]

        ws.merge_cells(label_pos)
        lc = label_pos.split(':')[0]

        ws[lc] = label
        ws[lc].fill = azul_claro
        ws[lc].font = Font(bold=True, size=10)
        ws[lc].alignment = center
        ws[lc].border = border

        ws.merge_cells(value_pos)
        vc = value_pos.split(':')[0]

        ws[vc] = value
        ws[vc].fill = gris
        ws[vc].alignment = center
        ws[vc].border = border

    # =================================================
    # HEADER TABLA
    # =================================================

    headers = ['TIPO', 'CATEGORÍA', 'PRIORIDAD', 'ESTADO', 'MENSAJE']
    merges = ['A14:B15', 'C14:C15', 'D14:D15', 'E14:E15', 'F14:I15']

    for i, merge in enumerate(merges):

        ws.merge_cells(merge)
        cell = merge.split(':')[0]

        ws[cell] = headers[i]
        ws[cell].fill = azul_header
        ws[cell].font = Font(bold=True, size=10)
        ws[cell].alignment = center
        ws[cell].border = border

    ws.row_dimensions[14].height = 38
    ws.row_dimensions[15].height = 38

    # =================================================
    # COLORES POR ESTADO / PRIORIDAD
    # =================================================

    def color(valor):
        if valor == 'CRITICA':
            return PatternFill(start_color='FCA5A5', fill_type='solid')
        if valor == 'ALTA':
            return PatternFill(start_color='FDBA74', fill_type='solid')
        if valor == 'MEDIA':
            return PatternFill(start_color='FDE68A', fill_type='solid')
        if valor == 'BAJA':
            return PatternFill(start_color='86EFAC', fill_type='solid')
        return None

    # =================================================
    # TABLA
    # =================================================

    fila = 16

    for a in alertas:

        ws.merge_cells(f'A{fila}:B{fila}')
        ws.merge_cells(f'F{fila}:I{fila}')  # MENSAJE BIEN COMBINADO

        ws[f'A{fila}'] = a.tipo
        ws[f'C{fila}'] = a.categoria
        ws[f'D{fila}'] = a.prioridad
        ws[f'E{fila}'] = a.estado
        ws[f'F{fila}'] = a.mensaje

        # base estilo
        for col in range(1, 10):
            c = ws.cell(row=fila, column=col)
            c.border = border
            c.alignment = center
            c.font = normal

        ws[f'A{fila}'].alignment = left
        ws[f'F{fila}'].alignment = left

        # colores
        if color(a.prioridad):
            ws[f'D{fila}'].fill = color(a.prioridad)

        if color(a.estado):
            ws[f'E{fila}'].fill = color(a.estado)

        # zebra
        if fila % 2 == 0:
            for col in range(1, 10):
                ws.cell(row=fila, column=col).fill = gris

        fila += 1

    # =================================================
    # BORDE EXTERIOR (CUADRO FINAL)
    # =================================================

    for r in range(14, fila):
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = border

    # =================================================
    # EXPORTAR
    # =================================================

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name=f'FORMATO_ALERTAS_{vehiculo.placa}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    
    
    
   
# =====================================================
# INDICADOR - CUMPLIMIENTO DE MANTENIMIENTO
# =====================================================

from datetime import timedelta


@reportes_bp.route(
    '/indicador-mantenimiento',
    methods=['GET']
)
def indicador_mantenimiento():

    inicio, fin = obtener_rango_fechas()

    # =================================================
    # NORMALIZAR FECHAS
    # =================================================

    if hasattr(inicio, "date"):
        fecha_inicio = inicio.date()
    else:
        fecha_inicio = inicio

    if hasattr(fin, "date"):
        fecha_fin = fin.date()
    else:
        fecha_fin = fin

    # =================================================
    # FILTROS
    # =================================================

    vehiculo_id = request.args.get(
        'vehiculo_id',
        type=int
    )

    maquinaria_id = request.args.get(
        'maquinaria_id',
        type=int
    )

    # =================================================
    # 1. ALERTAS DE MANTENIMIENTO REPORTADAS
    # =================================================
    #
    # Una alerta MANTENIMIENTO representa una
    # actividad/necesidad reportada.
    #
    # IMPORTANTE:
    # Se utiliza created_at porque representa el
    # momento en que fue creada/reportada la alerta.
    #
    # No se utiliza fecha_evento porque puede contener
    # fechas históricas o cambiar cuando el sistema
    # vuelve a generar/revisar la alerta.
    # =================================================

    query_alertas = Alerta.query.filter(
        Alerta.tipo == 'MANTENIMIENTO',
        Alerta.created_at >= fecha_inicio,
        Alerta.created_at < (
            fecha_fin + timedelta(days=1)
        )
    )

    # =================================================
    # FILTRO VEHÍCULO
    # =================================================

    if vehiculo_id:

        query_alertas = query_alertas.filter(
            Alerta.vehiculo_id == vehiculo_id
        )

    # =================================================
    # FILTRO MAQUINARIA
    # =================================================

    if maquinaria_id:

        query_alertas = query_alertas.filter(
            Alerta.maquinaria_id == maquinaria_id
        )

    alertas_mantenimiento = query_alertas.all()

    # =================================================
    # 2. OBTENER TIPOS DE PLAN ITEM
    # =================================================

    plan_item_ids = {
        alerta.plan_item_id
        for alerta in alertas_mantenimiento
        if alerta.plan_item_id is not None
    }

    tipos_plan_item = {}

    if plan_item_ids:

        registros_plan = PlanItem.query.filter(
            PlanItem.id.in_(plan_item_ids)
        ).all()

        tipos_plan_item = {
            plan.id: (
                plan.tipo_mantenimiento or ''
            ).upper()
            for plan in registros_plan
        }

    # =================================================
    # 3. CLASIFICAR ALERTAS
    # =================================================
    #
    # IMPORTANTE:
    #
    # PREVENTIVO:
    #   Solo alertas PREVENTIVO
    #
    # INSPECCION:
    #   Solo alertas INSPECCION
    #
    # CORRECTIVO:
    #   Solo alertas CORRECTIVO
    #
    # Las inspecciones NO se suman al preventivo
    # en este endpoint.
    #
    # La suma PREVENTIVO + INSPECCION se realizará
    # únicamente en el Excel HSEQ-R-159.
    # =================================================

    alertas_preventivas = []
    alertas_inspecciones = []
    alertas_correctivas = []

    for alerta in alertas_mantenimiento:

        tipo = tipos_plan_item.get(
            alerta.plan_item_id,
            ''
        )

        # -------------------------------------------------
        # PREVENTIVO
        # -------------------------------------------------

        if tipo == 'PREVENTIVO':

            alertas_preventivas.append(
                alerta
            )

        # -------------------------------------------------
        # INSPECCION
        # -------------------------------------------------

        elif tipo == 'INSPECCION':

            alertas_inspecciones.append(
                alerta
            )

        # -------------------------------------------------
        # CORRECTIVO
        # -------------------------------------------------

        elif tipo == 'CORRECTIVO':

            alertas_correctivas.append(
                alerta
            )

    # =================================================
    # 4. MANTENIMIENTOS REALMENTE EJECUTADOS
    # =================================================
    #
    # NO utilizamos Alerta.estado para determinar
    # si el mantenimiento fue ejecutado.
    #
    # Utilizamos los registros reales de:
    #
    #   mantenimientos
    #   maquinaria_mantenimientos
    #
    # con completado = True.
    # =================================================

    # =================================================
    # 4.1 MANTENIMIENTOS DE VEHÍCULOS
    # =================================================

    query_mantenimientos = Mantenimiento.query.filter(
        Mantenimiento.fecha >= fecha_inicio,
        Mantenimiento.fecha <= fecha_fin,
        Mantenimiento.completado == True
    )

    if vehiculo_id:

        query_mantenimientos = (
            query_mantenimientos.filter(
                Mantenimiento.vehiculo_id == vehiculo_id
            )
        )

    mantenimientos_vehiculos = (
        query_mantenimientos.all()
    )

    # =================================================
    # 4.2 MANTENIMIENTOS DE MAQUINARIA
    # =================================================

    query_mantenimientos_maquinaria = (
        MaquinariaMantenimiento.query.filter(
            MaquinariaMantenimiento.fecha >= fecha_inicio,
            MaquinariaMantenimiento.fecha <= fecha_fin,
            MaquinariaMantenimiento.completado == True
        )
    )

    if maquinaria_id:

        query_mantenimientos_maquinaria = (
            query_mantenimientos_maquinaria.filter(
                MaquinariaMantenimiento.maquinaria_id
                == maquinaria_id
            )
        )

    mantenimientos_maquinaria = (
        query_mantenimientos_maquinaria.all()
    )

    # =================================================
    # 5. CLASIFICAR MANTENIMIENTOS DE VEHÍCULOS
    # =================================================

    # -------------------------------------------------
    # PREVENTIVOS
    # -------------------------------------------------

    mantenimientos_preventivos_vehiculos = [
        m
        for m in mantenimientos_vehiculos
        if (m.tipo or '').upper()
        == 'PREVENTIVO'
    ]

    # -------------------------------------------------
    # INSPECCIONES
    # -------------------------------------------------

    mantenimientos_inspecciones_vehiculos = [
        m
        for m in mantenimientos_vehiculos
        if (m.tipo or '').upper()
        == 'INSPECCION'
    ]

    # -------------------------------------------------
    # CORRECTIVOS
    # -------------------------------------------------

    mantenimientos_correctivos_vehiculos = [
        m
        for m in mantenimientos_vehiculos
        if (m.tipo or '').upper()
        == 'CORRECTIVO'
    ]

    # =================================================
    # 6. CLASIFICAR MANTENIMIENTOS DE MAQUINARIA
    # =================================================
    #
    # En la BD actualmente:
    #
    #   I = INSPECCION
    #
    # También soportamos:
    #
    #   INSPECCION
    #
    # por si posteriormente se guarda con ese nombre.
    # =================================================

    # -------------------------------------------------
    # PREVENTIVOS
    # -------------------------------------------------

    mantenimientos_preventivos_maquinaria = [
        m
        for m in mantenimientos_maquinaria
        if (m.tipo or '').upper()
        == 'PREVENTIVO'
    ]

    # -------------------------------------------------
    # INSPECCIONES
    # -------------------------------------------------

    mantenimientos_inspecciones_maquinaria = [
        m
        for m in mantenimientos_maquinaria
        if (m.tipo or '').upper()
        in [
            'INSPECCION',
            'I'
        ]
    ]

    # -------------------------------------------------
    # CORRECTIVOS
    # -------------------------------------------------

    mantenimientos_correctivos_maquinaria = [
        m
        for m in mantenimientos_maquinaria
        if (m.tipo or '').upper()
        == 'CORRECTIVO'
    ]

    # =================================================
    # 7. INDICADOR PREVENTIVO
    # =================================================
    #
    # IMPORTANTE:
    #
    # El indicador de pantalla solamente considera
    # mantenimientos PREVENTIVOS.
    #
    # NO incluye INSPECCIONES.
    #
    # REPORTADOS:
    #   Alertas PREVENTIVO
    #
    # EJECUTADOS:
    #   Mantenimientos PREVENTIVO completados
    # =================================================

    preventivos_reportados = len(
        alertas_preventivas
    )

    preventivos_ejecutados = (
        len(mantenimientos_preventivos_vehiculos)
        +
        len(mantenimientos_preventivos_maquinaria)
    )

    preventivos_pendientes = max(
        preventivos_reportados
        - preventivos_ejecutados,
        0
    )

    if preventivos_reportados == 0:

        porcentaje_preventivo = 100

    else:

        porcentaje_preventivo = (
            preventivos_ejecutados
            / preventivos_reportados
        ) * 100

    # =================================================
    # 8. INDICADOR DE INSPECCIONES
    # =================================================
    #
    # REPORTADAS:
    #   Alertas INSPECCION
    #
    # EJECUTADAS:
    #   Mantenimientos INSPECCION completados
    #
    # Para maquinaria:
    #   I = INSPECCION
    # =================================================

    inspecciones_reportadas = len(
        alertas_inspecciones
    )

    inspecciones_ejecutadas = (
        len(mantenimientos_inspecciones_vehiculos)
        +
        len(mantenimientos_inspecciones_maquinaria)
    )

    if inspecciones_reportadas == 0:

        porcentaje_inspecciones = 100

    else:

        porcentaje_inspecciones = (
            inspecciones_ejecutadas
            / inspecciones_reportadas
        ) * 100

    # =================================================
    # 9. INDICADOR CORRECTIVO
    # =================================================
    #
    # REPORTADOS:
    #   Alertas CORRECTIVO
    #
    # EJECUTADOS:
    #   Mantenimientos CORRECTIVO completados
    # =================================================

    correctivos_reportados = len(
        alertas_correctivas
    )

    correctivos_ejecutados = (
        len(mantenimientos_correctivos_vehiculos)
        +
        len(mantenimientos_correctivos_maquinaria)
    )

    if correctivos_reportados == 0:

        porcentaje_correctivo = 100

    else:

        porcentaje_correctivo = (
            correctivos_ejecutados
            / correctivos_reportados
        ) * 100

    # =================================================
    # 10. DETALLE VEHÍCULOS
    # =================================================

    alertas_preventivas_vehiculos = [
        alerta
        for alerta in alertas_preventivas
        if alerta.vehiculo_id is not None
    ]

    alertas_inspecciones_vehiculos = [
        alerta
        for alerta in alertas_inspecciones
        if alerta.vehiculo_id is not None
    ]

    alertas_correctivas_vehiculos = [
        alerta
        for alerta in alertas_correctivas
        if alerta.vehiculo_id is not None
    ]

    # =================================================
    # 11. DETALLE MAQUINARIA
    # =================================================

    alertas_preventivas_maquinaria = [
        alerta
        for alerta in alertas_preventivas
        if alerta.maquinaria_id is not None
    ]

    alertas_inspecciones_maquinaria = [
        alerta
        for alerta in alertas_inspecciones
        if alerta.maquinaria_id is not None
    ]

    alertas_correctivas_maquinaria = [
        alerta
        for alerta in alertas_correctivas
        if alerta.maquinaria_id is not None
    ]

    # =================================================
    # 12. RESPUESTA
    # =================================================

    return jsonify({

        # =================================================
        # INDICADORES
        # =================================================

        "indicadores": {

            # ---------------------------------------------
            # PREVENTIVO
            # ---------------------------------------------

            "preventivo": {

                "reportados": (
                    preventivos_reportados
                ),

                "ejecutados": (
                    preventivos_ejecutados
                ),

                "pendientes": (
                    preventivos_pendientes
                ),

                "porcentaje": round(
                    porcentaje_preventivo,
                    2
                ),

                "meta": 90,

                "cumple_meta": (
                    porcentaje_preventivo >= 90
                ),

                "metodo_calculo": (
                    "Mantenimientos preventivos "
                    "ejecutados / alertas preventivas "
                    "reportadas × 100"
                )
            },

            # ---------------------------------------------
            # INSPECCIONES
            # ---------------------------------------------

            "inspecciones": {

                "reportadas": (
                    inspecciones_reportadas
                ),

                "resueltas": (
                    inspecciones_ejecutadas
                ),

                "porcentaje": round(
                    porcentaje_inspecciones,
                    2
                ),

                "metodo_calculo": (
                    "Inspecciones ejecutadas / "
                    "alertas de inspección reportadas "
                    "× 100"
                )
            },

            # ---------------------------------------------
            # CORRECTIVO
            # ---------------------------------------------

            "correctivo": {

                "mantenimientos_reportados": (
                    correctivos_reportados
                ),

                "mantenimientos_cerrados": (
                    correctivos_ejecutados
                ),

                "porcentaje": round(
                    porcentaje_correctivo,
                    2
                ),

                "meta": 90,

                "cumple_meta": (
                    porcentaje_correctivo >= 90
                ),

                "metodo_calculo": (
                    "Mantenimientos correctivos "
                    "ejecutados / alertas correctivas "
                    "reportadas × 100"
                )
            }
        },

        # =================================================
        # VEHÍCULOS
        # =================================================

        "vehiculos": {

            "preventivos_programados": (
                len(alertas_preventivas_vehiculos)
            ),

            "preventivos_ejecutados": (
                len(mantenimientos_preventivos_vehiculos)
            ),

            "inspecciones": (
                len(alertas_inspecciones_vehiculos)
            ),

            "inspecciones_ejecutadas": (
                len(mantenimientos_inspecciones_vehiculos)
            ),

            "correctivos": (
                len(alertas_correctivas_vehiculos)
            ),

            "correctivos_cerrados": (
                len(mantenimientos_correctivos_vehiculos)
            )
        },

        # =================================================
        # MAQUINARIA
        # =================================================

        "maquinaria": {

            "preventivos_programados": (
                len(alertas_preventivas_maquinaria)
            ),

            "preventivos_ejecutados": (
                len(mantenimientos_preventivos_maquinaria)
            ),

            "inspecciones": (
                len(alertas_inspecciones_maquinaria)
            ),

            "inspecciones_ejecutadas": (
                len(mantenimientos_inspecciones_maquinaria)
            ),

            "correctivos": (
                len(alertas_correctivas_maquinaria)
            ),

            "correctivos_cerrados": (
                len(mantenimientos_correctivos_maquinaria)
            )
        },

        # =================================================
        # PERIODO
        # =================================================

        "periodo": {

            "inicio": str(fecha_inicio),

            "fin": str(fecha_fin)
        },

        # =================================================
        # FILTROS
        # =================================================

        "filtros": {

            "vehiculo_id": vehiculo_id,

            "maquinaria_id": maquinaria_id
        }
    })
 
@reportes_bp.route(
    '/indicador-mantenimiento/exportar',
    methods=['GET']
)
def exportar_indicador_mantenimiento():

    try:

        from io import BytesIO
        import os

        from flask import (
            request,
            jsonify,
            send_file
        )

        from openpyxl import Workbook
        from openpyxl.drawing.image import Image

        from openpyxl.styles import (
            Font,
            PatternFill,
            Border,
            Side,
            Alignment
        )

        from openpyxl.worksheet.page import PageMargins

        # ==========================================================
        # FILTROS
        # ==========================================================

        tipo = request.args.get(
            'tipo',
            'mensual'
        )

        vehiculo_id = request.args.get(
            'vehiculo_id'
        )

        maquinaria_id = request.args.get(
            'maquinaria_id'
        )

        fecha_inicio = request.args.get(
            'fecha_inicio'
        )

        fecha_fin = request.args.get(
            'fecha_fin'
        )

        # ==========================================================
        # CONVERSIÓN DE IDS
        # ==========================================================

        if vehiculo_id in ['', 'null', 'None']:
            vehiculo_id = None

        elif vehiculo_id:
            vehiculo_id = int(
                vehiculo_id
            )

        if maquinaria_id in ['', 'null', 'None']:
            maquinaria_id = None

        elif maquinaria_id:
            maquinaria_id = int(
                maquinaria_id
            )

        # ==========================================================
        # FECHAS
        # ==========================================================

        hoy = date.today()

        if tipo == 'rango':

            if not fecha_inicio or not fecha_fin:

                return jsonify({
                    'error':
                        'Debe indicar fecha_inicio y fecha_fin'
                }), 400

            fecha_inicio_obj = datetime.strptime(
                fecha_inicio,
                '%Y-%m-%d'
            ).date()

            fecha_fin_obj = datetime.strptime(
                fecha_fin,
                '%Y-%m-%d'
            ).date()

        elif tipo == 'diario':

            fecha_inicio_obj = hoy
            fecha_fin_obj = hoy

        elif tipo == 'semanal':

            fecha_fin_obj = hoy

            fecha_inicio_obj = (
                hoy - timedelta(days=6)
            )

        else:

            # ======================================================
            # MENSUAL
            # ======================================================

            fecha_inicio_obj = hoy.replace(
                day=1
            )

            fecha_fin_obj = hoy

        # ==========================================================
        # VALIDACIÓN DE FECHAS
        # ==========================================================

        if fecha_inicio_obj > fecha_fin_obj:

            return jsonify({
                'error':
                    'La fecha inicial no puede ser mayor que la fecha final'
            }), 400

        # ==========================================================
        # NOMBRE DEL ACTIVO
        # ==========================================================

        activo = 'Todos los activos'

        if vehiculo_id:

            vehiculo = Vehiculo.query.get(
                vehiculo_id
            )

            if vehiculo:

                activo = vehiculo.placa

        elif maquinaria_id:

            maquinaria = Maquinaria.query.get(
                maquinaria_id
            )

            if maquinaria:

                activo = maquinaria.codigo

        # ==========================================================
        # ALERTAS DE MANTENIMIENTO
        # ==========================================================

        query_alertas = Alerta.query.filter(

            Alerta.tipo == 'MANTENIMIENTO',

            Alerta.created_at >= fecha_inicio_obj,

            Alerta.created_at < (
                fecha_fin_obj + timedelta(days=1)
            )
        )

        if vehiculo_id:

            query_alertas = query_alertas.filter(
                Alerta.vehiculo_id == vehiculo_id
            )

        if maquinaria_id:

            query_alertas = query_alertas.filter(
                Alerta.maquinaria_id == maquinaria_id
            )

        alertas_mantenimiento = (
            query_alertas.all()
        )

        # ==========================================================
        # OBTENER PLAN ITEMS
        # ==========================================================

        plan_item_ids = {

            alerta.plan_item_id

            for alerta in alertas_mantenimiento

            if alerta.plan_item_id is not None
        }

        tipos_plan_item = {}

        if plan_item_ids:

            registros_plan = (
                PlanItem.query
                .filter(
                    PlanItem.id.in_(
                        plan_item_ids
                    )
                )
                .all()
            )

            tipos_plan_item = {

                plan.id:
                    (
                        plan.tipo_mantenimiento
                        or ''
                    ).upper()

                for plan in registros_plan
            }

        # ==========================================================
        # CLASIFICAR ALERTAS
        # ==========================================================

        alertas_preventivas = []

        alertas_inspecciones = []

        alertas_correctivas = []

        for alerta in alertas_mantenimiento:

            tipo_plan = tipos_plan_item.get(
                alerta.plan_item_id,
                ''
            )

            if tipo_plan == 'PREVENTIVO':

                alertas_preventivas.append(
                    alerta
                )

            elif tipo_plan == 'INSPECCION':

                alertas_inspecciones.append(
                    alerta
                )

            elif tipo_plan == 'CORRECTIVO':

                alertas_correctivas.append(
                    alerta
                )

        # ==========================================================
        # REPORTADOS
        # ==========================================================

        preventivos_reportados = len(
            alertas_preventivas
        )

        inspecciones_reportadas = len(
            alertas_inspecciones
        )

        correctivos_reportados = len(
            alertas_correctivas
        )

        # ==========================================================
        # PREVENTIVOS VEHÍCULOS
        # ==========================================================

        query_preventivos_vehiculos = (
            Mantenimiento.query.filter(

                Mantenimiento.fecha
                >= fecha_inicio_obj,

                Mantenimiento.fecha
                <= fecha_fin_obj,

                Mantenimiento.tipo
                == 'PREVENTIVO',

                Mantenimiento.completado
                == True
            )
        )

        if vehiculo_id:

            query_preventivos_vehiculos = (
                query_preventivos_vehiculos.filter(
                    Mantenimiento.vehiculo_id
                    == vehiculo_id
                )
            )

        preventivos_vehiculos = (
            query_preventivos_vehiculos.count()
        )

        # ==========================================================
        # INSPECCIONES VEHÍCULOS
        # ==========================================================

        query_inspecciones_vehiculos = (
            Mantenimiento.query.filter(

                Mantenimiento.fecha
                >= fecha_inicio_obj,

                Mantenimiento.fecha
                <= fecha_fin_obj,

                Mantenimiento.tipo
                == 'INSPECCION',

                Mantenimiento.completado
                == True
            )
        )

        if vehiculo_id:

            query_inspecciones_vehiculos = (
                query_inspecciones_vehiculos.filter(
                    Mantenimiento.vehiculo_id
                    == vehiculo_id
                )
            )

        inspecciones_vehiculos = (
            query_inspecciones_vehiculos.count()
        )

        # ==========================================================
        # PREVENTIVOS MAQUINARIA
        # ==========================================================

        query_preventivos_maquinaria = (
            MaquinariaMantenimiento.query.filter(

                MaquinariaMantenimiento.fecha
                >= fecha_inicio_obj,

                MaquinariaMantenimiento.fecha
                <= fecha_fin_obj,

                MaquinariaMantenimiento.tipo
                == 'PREVENTIVO',

                MaquinariaMantenimiento.completado
                == True
            )
        )

        if maquinaria_id:

            query_preventivos_maquinaria = (
                query_preventivos_maquinaria.filter(
                    MaquinariaMantenimiento.maquinaria_id
                    == maquinaria_id
                )
            )

        preventivos_maquinaria = (
            query_preventivos_maquinaria.count()
        )

        # ==========================================================
        # INSPECCIONES MAQUINARIA
        # ==========================================================
        #
        # En maquinaria pueden existir:
        #
        # INSPECCION
        # I
        #
        # ==========================================================

        query_inspecciones_maquinaria = (
            MaquinariaMantenimiento.query.filter(

                MaquinariaMantenimiento.fecha
                >= fecha_inicio_obj,

                MaquinariaMantenimiento.fecha
                <= fecha_fin_obj,

                MaquinariaMantenimiento.completado
                == True
            )
        )

        if maquinaria_id:

            query_inspecciones_maquinaria = (
                query_inspecciones_maquinaria.filter(
                    MaquinariaMantenimiento.maquinaria_id
                    == maquinaria_id
                )
            )

        inspecciones_maquinaria_lista = [

            mantenimiento

            for mantenimiento
            in query_inspecciones_maquinaria.all()

            if (
                (mantenimiento.tipo or '').upper()
                in [
                    'INSPECCION',
                    'I'
                ]
            )
        ]

        inspecciones_maquinaria = len(
            inspecciones_maquinaria_lista
        )

        # ==========================================================
        # PREVENTIVO HSEQ
        # ==========================================================
        #
        # IMPORTANTE:
        #
        # Para el Excel HSEQ-R-159:
        #
        # PREVENTIVO =
        #
        # PREVENTIVOS + INSPECCIONES
        #
        # ==========================================================

        preventivos_programados = (

            preventivos_reportados

            + inspecciones_reportadas
        )

        preventivos_ejecutados = (

            preventivos_vehiculos

            + preventivos_maquinaria

            + inspecciones_vehiculos

            + inspecciones_maquinaria
        )

        # ==========================================================
        # CORRECTIVOS VEHÍCULOS REPORTADOS
        # ==========================================================

        query_correctivos_vehiculos = (
            Mantenimiento.query.filter(

                Mantenimiento.fecha
                >= fecha_inicio_obj,

                Mantenimiento.fecha
                <= fecha_fin_obj,

                Mantenimiento.tipo
                == 'CORRECTIVO'
            )
        )

        if vehiculo_id:

            query_correctivos_vehiculos = (
                query_correctivos_vehiculos.filter(
                    Mantenimiento.vehiculo_id
                    == vehiculo_id
                )
            )

        correctivos_vehiculos_reportados = (
            query_correctivos_vehiculos.count()
        )

        # ==========================================================
        # CORRECTIVOS VEHÍCULOS CERRADOS
        # ==========================================================

        query_correctivos_vehiculos_cerrados = (
            Mantenimiento.query.filter(

                Mantenimiento.fecha
                >= fecha_inicio_obj,

                Mantenimiento.fecha
                <= fecha_fin_obj,

                Mantenimiento.tipo
                == 'CORRECTIVO',

                Mantenimiento.completado
                == True
            )
        )

        if vehiculo_id:

            query_correctivos_vehiculos_cerrados = (
                query_correctivos_vehiculos_cerrados.filter(
                    Mantenimiento.vehiculo_id
                    == vehiculo_id
                )
            )

        correctivos_vehiculos_cerrados = (
            query_correctivos_vehiculos_cerrados.count()
        )

        # ==========================================================
        # CORRECTIVOS MAQUINARIA REPORTADOS
        # ==========================================================

        query_correctivos_maquinaria_reportados = (
            MaquinariaMantenimiento.query.filter(

                MaquinariaMantenimiento.fecha
                >= fecha_inicio_obj,

                MaquinariaMantenimiento.fecha
                <= fecha_fin_obj,

                MaquinariaMantenimiento.tipo
                == 'CORRECTIVO'
            )
        )

        if maquinaria_id:

            query_correctivos_maquinaria_reportados = (
                query_correctivos_maquinaria_reportados.filter(
                    MaquinariaMantenimiento.maquinaria_id
                    == maquinaria_id
                )
            )

        correctivos_maquinaria_reportados = (
            query_correctivos_maquinaria_reportados.count()
        )

        # ==========================================================
        # CORRECTIVOS MAQUINARIA CERRADOS
        # ==========================================================

        query_correctivos_maquinaria_cerrados = (
            MaquinariaMantenimiento.query.filter(

                MaquinariaMantenimiento.fecha
                >= fecha_inicio_obj,

                MaquinariaMantenimiento.fecha
                <= fecha_fin_obj,

                MaquinariaMantenimiento.tipo
                == 'CORRECTIVO',

                MaquinariaMantenimiento.completado
                == True
            )
        )

        if maquinaria_id:

            query_correctivos_maquinaria_cerrados = (
                query_correctivos_maquinaria_cerrados.filter(
                    MaquinariaMantenimiento.maquinaria_id
                    == maquinaria_id
                )
            )

        correctivos_maquinaria_cerrados = (
            query_correctivos_maquinaria_cerrados.count()
        )

        # ==========================================================
        # CORRECTIVOS TOTALES
        # ==========================================================

        correctivos_reportados = (

            correctivos_vehiculos_reportados

            + correctivos_maquinaria_reportados
        )

        correctivos_cerrados = (

            correctivos_vehiculos_cerrados

            + correctivos_maquinaria_cerrados
        )

        # ==========================================================
        # PORCENTAJE PREVENTIVO
        # ==========================================================

        if preventivos_programados > 0:

            porcentaje_preventivo = (

                preventivos_ejecutados
                /
                preventivos_programados

            ) * 100

        else:

            porcentaje_preventivo = 0

        # ==========================================================
        # PORCENTAJE CORRECTIVO
        # ==========================================================

        if correctivos_reportados > 0:

            porcentaje_correctivo = (

                correctivos_cerrados
                /
                correctivos_reportados

            ) * 100

        else:

            porcentaje_correctivo = 0

        # ==========================================================
        # PORCENTAJE INSPECCIONES
        # ==========================================================

        inspecciones_ejecutadas = (

            inspecciones_vehiculos

            + inspecciones_maquinaria
        )

        if inspecciones_reportadas > 0:

            porcentaje_inspecciones = (

                inspecciones_ejecutadas
                /
                inspecciones_reportadas

            ) * 100

        else:

            porcentaje_inspecciones = 0

        # ==========================================================
        # LIMITAR PORCENTAJES
        # ==========================================================

        porcentaje_preventivo = min(
            porcentaje_preventivo,
            100
        )

        porcentaje_correctivo = min(
            porcentaje_correctivo,
            100
        )

        porcentaje_inspecciones = min(
            porcentaje_inspecciones,
            100
        )

        # ==========================================================
        # ESTADOS
        # ==========================================================

        cumple_preventivo = (
            porcentaje_preventivo >= 90
        )

        cumple_correctivo = (
            porcentaje_correctivo >= 90
        )

        cumple_inspecciones = (
            porcentaje_inspecciones >= 90
        )

        # ==========================================================
        # CREAR EXCEL
        # ==========================================================

        wb = Workbook()

        ws = wb.active

        ws.title = 'HSEQ-R-159'

        ws.sheet_view.showGridLines = False

        # ==========================================================
        # COLUMNAS
        # ==========================================================

        columnas = {

            'A': 24,
            'B': 18,
            'C': 18,
            'D': 18,
            'E': 18,
            'F': 18,
            'G': 18,
            'H': 18,
            'I': 18,
            'J': 18

        }

        for col, width in columnas.items():

            ws.column_dimensions[
                col
            ].width = width

        # ==========================================================
        # COLORES
        # ==========================================================

        azul_oscuro = PatternFill(
            start_color='1F4E78',
            end_color='1F4E78',
            fill_type='solid'
        )

        azul_header = PatternFill(
            start_color='8DB4E2',
            end_color='8DB4E2',
            fill_type='solid'
        )

        azul_claro = PatternFill(
            start_color='D9EAF7',
            end_color='D9EAF7',
            fill_type='solid'
        )

        gris = PatternFill(
            start_color='F2F2F2',
            end_color='F2F2F2',
            fill_type='solid'
        )

        verde = PatternFill(
            start_color='E2F0D9',
            end_color='E2F0D9',
            fill_type='solid'
        )

        rojo = PatternFill(
            start_color='F4CCCC',
            end_color='F4CCCC',
            fill_type='solid'
        )

        # ==========================================================
        # BORDES
        # ==========================================================

        thin = Side(
            style='thin',
            color='B7B7B7'
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        # ==========================================================
        # FUENTES
        # ==========================================================

        titulo_font = Font(
            bold=True,
            color='FFFFFF',
            size=14
        )

        subtitulo_font = Font(
            bold=True,
            color='1F1F1F',
            size=12
        )

        blanco_negrita = Font(
            bold=True,
            color='FFFFFF',
            size=10
        )

        negrita = Font(
            bold=True,
            color='1F1F1F',
            size=10
        )

        normal = Font(
            color='333333',
            size=10
        )

        # ==========================================================
        # ALINEACIONES
        # ==========================================================

        center = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True
        )

        left = Alignment(
            horizontal='left',
            vertical='center',
            wrap_text=True
        )

        # ==========================================================
        # ENCABEZADO
        # ==========================================================

        ws.row_dimensions[1].height = 35
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 30

        # ==========================================================
        # LOGO
        # ==========================================================

        ws.merge_cells(
            'A1:B3'
        )

        for row in ws['A1:B3']:

            for cell in row:

                cell.fill = azul_oscuro
                cell.border = border

        ruta_logo = (
            'static/logo.jpg'
        )

        if os.path.exists(ruta_logo):

            logo = Image(
                ruta_logo
            )

            logo.width = 190
            logo.height = 85

            ws.add_image(
                logo,
                'A1'
            )

        else:

            ws['A1'] = 'INTELLIFLEET'

            ws['A1'].font = titulo_font

            ws['A1'].alignment = center

        # ==========================================================
        # TÍTULO PRINCIPAL
        # ==========================================================

        ws.merge_cells(
            'C1:H2'
        )

        ws['C1'] = (
            'MATRIZ DE INDICADORES DE RESOLUCION  40595 DE 2022'
        )

        ws['C1'].fill = azul_oscuro
        ws['C1'].font = titulo_font
        ws['C1'].alignment = center
        ws['C1'].border = border

        # ==========================================================
        # SUBTÍTULO
        # ==========================================================

        ws.merge_cells(
            'C3:H3'
        )

        ws['C3'] = (
            'MATRIZ DE INDICADORES HSEQ'
        )

        ws['C3'].fill = azul_claro
        ws['C3'].font = subtitulo_font
        ws['C3'].alignment = center
        ws['C3'].border = border

        # ==========================================================
        # INFORMACIÓN DOCUMENTO
        # ==========================================================

        info_header = [

            (
                'I1:J1',
                'VERSIÓN: 002'
            ),

            (
                'I2:J2',
                'CÓDIGO: HSEQ-R-159'
            ),

            (
                'I3:J3',
                'PÁGINA: 15 DE 16'
            )

        ]

        for rango, texto in info_header:

            ws.merge_cells(
                rango
            )

            celda = rango.split(':')[0]

            ws[celda] = texto

            ws[celda].fill = azul_oscuro
            ws[celda].font = blanco_negrita
            ws[celda].alignment = center
            ws[celda].border = border

        # ==========================================================
        # INFORMACIÓN DEL PERÍODO
        # ==========================================================

        ws.merge_cells(
            'A5:J5'
        )

        ws['A5'] = (
            'INFORMACIÓN DEL PERÍODO'
        )

        ws['A5'].fill = azul_oscuro
        ws['A5'].font = blanco_negrita
        ws['A5'].alignment = center

        for col in range(1, 11):

            ws.cell(
                5,
                col
            ).border = border

        # ==========================================================
        # PERÍODO
        # ==========================================================

        ws.merge_cells(
            'A6:B6'
        )

        ws['A6'] = (
            'PERÍODO ANALIZADO'
        )

        ws['A6'].fill = gris
        ws['A6'].font = negrita
        ws['A6'].alignment = left

        ws.merge_cells(
            'C6:F6'
        )

        ws['C6'] = (
            f'{fecha_inicio_obj.strftime("%d/%m/%Y")} - '
            f'{fecha_fin_obj.strftime("%d/%m/%Y")}'
        )

        ws['C6'].alignment = left

        ws.merge_cells(
            'G6:H6'
        )

        ws['G6'] = 'ACTIVO'

        ws['G6'].fill = gris
        ws['G6'].font = negrita
        ws['G6'].alignment = left

        ws.merge_cells(
            'I6:J6'
        )

        ws['I6'] = activo
        ws['I6'].alignment = left

        # ==========================================================
        # TIPO DE REPORTE
        # ==========================================================

        ws.merge_cells(
            'A7:B7'
        )

        ws['A7'] = (
            'TIPO DE REPORTE'
        )

        ws['A7'].fill = gris
        ws['A7'].font = negrita
        ws['A7'].alignment = left

        ws.merge_cells(
            'C7:F7'
        )

        nombres_tipo = {

            'diario': 'DIARIO',

            'semanal': 'SEMANAL',

            'mensual': 'MENSUAL',

            'rango': 'RANGO'

        }

        ws['C7'] = nombres_tipo.get(
            tipo,
            tipo.upper()
        )

        ws['C7'].alignment = left

        ws.merge_cells(
            'G7:H7'
        )

        ws['G7'] = (
            'FECHA DE GENERACIÓN'
        )

        ws['G7'].fill = gris
        ws['G7'].font = negrita
        ws['G7'].alignment = left

        ws.merge_cells(
            'I7:J7'
        )

        ws['I7'] = hoy.strftime(
            '%d/%m/%Y'
        )

        ws['I7'].alignment = left

        # ==========================================================
        # BORDES INFORMACIÓN
        # ==========================================================

        for fila in [6, 7]:

            for col in range(1, 11):

                ws.cell(
                    fila,
                    col
                ).border = border

        # ==========================================================
        # FUNCIÓN PARA BORDES COMPLETOS
        # ==========================================================

        def aplicar_bordes(
            fila_inicio,
            fila_fin,
            col_inicio,
            col_fin
        ):

            for fila_borde in range(
                fila_inicio,
                fila_fin + 1
            ):

                for col_borde in range(
                    col_inicio,
                    col_fin + 1
                ):

                    ws.cell(
                        fila_borde,
                        col_borde
                    ).border = border

        # ==========================================================
        # FUNCIÓN PARA COMBINAR CON BORDES
        # ==========================================================

        def combinar_con_borde(
            fila,
            col_inicio,
            col_fin
        ):

            # Aplicar primero
            aplicar_bordes(
                fila,
                fila,
                col_inicio,
                col_fin
            )

            # Combinar
            ws.merge_cells(
                start_row=fila,
                start_column=col_inicio,
                end_row=fila,
                end_column=col_fin
            )

            # Volver a aplicar
            aplicar_bordes(
                fila,
                fila,
                col_inicio,
                col_fin
            )

        # ==========================================================
        # FUNCIÓN INDICADOR
        # ==========================================================

        def escribir_indicador(

            fila,

            titulo,

            definicion,

            interpretacion,

            tipo_indicador,

            fuente,

            proceso,

            sentido,

            meta,

            metodo,

            frecuencia,

            programados,

            ejecutados,

            porcentaje,

            estado

        ):

            # ======================================================
            # TÍTULO
            # ======================================================

            combinar_con_borde(
                fila,
                1,
                10
            )

            celda_titulo = ws.cell(
                fila,
                1
            )

            celda_titulo.value = titulo

            celda_titulo.fill = azul_oscuro

            celda_titulo.font = blanco_negrita

            celda_titulo.alignment = center

            ws.row_dimensions[
                fila
            ].height = 32

            # ======================================================
            # CAMPOS
            # ======================================================

            campos = [

                (
                    'DEFINICIÓN DEL INDICADOR',
                    definicion,
                    60
                ),

                (
                    'INTERPRETACIÓN DEL INDICADOR',
                    interpretacion,
                    70
                ),

                (
                    'TIPO DE INDICADOR',
                    tipo_indicador,
                    32
                ),

                (
                    'FUENTE DE LA INFORMACIÓN',
                    fuente,
                    65
                ),

                (
                    'PROCESO RESPONSABLE',
                    proceso,
                    32
                ),

                (
                    'SENTIDO',
                    sentido,
                    32
                ),

                (
                    'PERSONAS QUE DEBEN CONOCER EL RESULTADO',
                    'Gerencia, HSEQ, Operaciones y Mantenimiento',
                    38
                ),

                (
                    'META',
                    meta,
                    32
                ),

                (
                    'MÉTODO DE CÁLCULO',
                    metodo,
                    60
                ),

                (
                    'FRECUENCIA',
                    frecuencia,
                    38
                )
            ]

            fila_actual = fila + 1

            for nombre, valor, altura in campos:

                # ==================================================
                # A:B - ETIQUETA
                # ==================================================

                combinar_con_borde(
                    fila_actual,
                    1,
                    2
                )

                etiqueta = ws.cell(
                    fila_actual,
                    1
                )

                etiqueta.value = nombre

                etiqueta.fill = gris

                etiqueta.font = negrita

                etiqueta.alignment = left

                # ==================================================
                # C:J - CONTENIDO
                # ==================================================

                combinar_con_borde(
                    fila_actual,
                    3,
                    10
                )

                valor_cell = ws.cell(
                    fila_actual,
                    3
                )

                valor_cell.value = valor

                valor_cell.font = normal

                valor_cell.alignment = left

                # ==================================================
                # ASEGURAR TODOS LOS BORDES
                # ==================================================

                aplicar_bordes(
                    fila_actual,
                    fila_actual,
                    1,
                    10
                )

                ws.row_dimensions[
                    fila_actual
                ].height = altura

                fila_actual += 1

            # ======================================================
            # RESULTADO DEL PERÍODO
            # ======================================================

            combinar_con_borde(
                fila_actual,
                1,
                10
            )

            resultado = ws.cell(
                fila_actual,
                1
            )

            resultado.value = (
                'RESULTADO DEL PERÍODO'
            )

            resultado.fill = azul_header

            resultado.font = negrita

            resultado.alignment = center

            fila_actual += 1

            # ======================================================
            # ENCABEZADOS RESULTADO
            # ======================================================

            encabezados = [

                'PROGRAMADOS / REPORTADOS',

                'EJECUTADOS / CERRADOS',

                'CUMPLIMIENTO',

                'META',

                'ESTADO'

            ]

            rangos = [

                (1, 2),

                (3, 4),

                (5, 6),

                (7, 8),

                (9, 10)

            ]

            valores = [

                programados,

                ejecutados,

                porcentaje / 100,

                0.90,

                estado

            ]

            # ======================================================
            # ENCABEZADOS
            # ======================================================

            for i in range(5):

                col_inicio = rangos[i][0]

                col_fin = rangos[i][1]

                combinar_con_borde(

                    fila_actual,

                    col_inicio,

                    col_fin
                )

                cell = ws.cell(
                    fila_actual,
                    col_inicio
                )

                cell.value = encabezados[i]

                cell.fill = azul_oscuro

                cell.font = blanco_negrita

                cell.alignment = center

                aplicar_bordes(

                    fila_actual,

                    fila_actual,

                    col_inicio,

                    col_fin
                )

            fila_actual += 1

            # ======================================================
            # VALORES
            # ======================================================

            for i in range(5):

                col_inicio = rangos[i][0]

                col_fin = rangos[i][1]

                combinar_con_borde(

                    fila_actual,

                    col_inicio,

                    col_fin
                )

                cell = ws.cell(
                    fila_actual,
                    col_inicio
                )

                cell.value = valores[i]

                cell.alignment = center

                cell.font = Font(
                    bold=True,
                    size=12
                )

                # ==================================================
                # PORCENTAJES
                # ==================================================

                if i in [2, 3]:

                    cell.number_format = (
                        '0.00%'
                    )

                # ==================================================
                # ESTADO
                # ==================================================

                if i == 4:

                    if estado == 'CUMPLE':

                        cell.fill = verde

                    else:

                        cell.fill = rojo

                    cell.font = Font(
                        bold=True,
                        size=11
                    )

                aplicar_bordes(

                    fila_actual,

                    fila_actual,

                    col_inicio,

                    col_fin
                )

            ws.row_dimensions[
                fila_actual
            ].height = 34

            return fila_actual + 2

        # ==========================================================
        # FRECUENCIA
        # ==========================================================

        frecuencia_reporte = (

            f'Fecha desde: '
            f'{fecha_inicio_obj.strftime("%d/%m/%Y")} '
            f'| Fecha hasta: '
            f'{fecha_fin_obj.strftime("%d/%m/%Y")}'
        )

        # ==========================================================
        # INDICADOR PREVENTIVO
        # ==========================================================

        siguiente_fila = escribir_indicador(

            9,

            'INDICADOR DE CUMPLIMIENTO DEL MANTENIMIENTO PREVENTIVO',

            'Mide el cumplimiento de las actividades de mantenimiento '
            'preventivo e inspección reportadas durante el período evaluado.',

            'Permite determinar el porcentaje de actividades preventivas '
            'e inspecciones ejecutadas frente a las actividades reportadas. '
            'Para este indicador HSEQ, las inspecciones hacen parte del '
            'cumplimiento preventivo.',

            'Gestión y cumplimiento.',

            'Alertas de mantenimiento, plan de mantenimiento, órdenes '
            'de trabajo, registros de mantenimiento y hojas de vida '
            'de vehículos y maquinaria.',

            'Gestión de Operaciones y Mantenimiento',

            'Ascendente.',

            '≥ 90 %.',

            '(Mantenimientos preventivos e inspecciones ejecutados / '
            'Mantenimientos preventivos e inspecciones reportados) × 100',

            frecuencia_reporte,

            preventivos_programados,

            preventivos_ejecutados,

            porcentaje_preventivo,

            'CUMPLE'
            if cumple_preventivo
            else 'NO CUMPLE'
        )

        # ==========================================================
        # INDICADOR CORRECTIVO
        # ==========================================================

        siguiente_fila = escribir_indicador(

            siguiente_fila,

            'INDICADOR DE CUMPLIMIENTO DEL MANTENIMIENTO CORRECTIVO',

            'Mide el cumplimiento de las actividades de mantenimiento '
            'correctivo reportadas durante el período evaluado.',

            'Permite evaluar la capacidad de la organización para atender '
            'y cerrar las fallas identificadas en vehículos y maquinaria.',

            'Gestión y eficacia.',

            'Reportes de fallas, órdenes de trabajo, registros de '
            'mantenimiento correctivo y hojas de vida.',

            'Gestión de Operaciones y Mantenimiento',

            'Ascendente.',

            '≥ 90 %.',

            '(Mantenimientos correctivos cerrados / '
            'Mantenimientos correctivos reportados) × 100',

            frecuencia_reporte,

            correctivos_reportados,

            correctivos_cerrados,

            porcentaje_correctivo,

            'CUMPLE'
            if cumple_correctivo
            else 'NO CUMPLE'
        )

        # ==========================================================
        # INDICADOR INSPECCIONES
        # ==========================================================

        siguiente_fila = escribir_indicador(

            siguiente_fila,

            'INDICADOR DE CUMPLIMIENTO DE INSPECCIONES',

            'Mide el cumplimiento de las inspecciones reportadas '
            'durante el período evaluado.',

            'Permite identificar el porcentaje de inspecciones ejecutadas '
            'frente a las inspecciones reportadas.',

            'Gestión y cumplimiento.',

            'Alertas de inspección y registros de inspecciones '
            'de vehículos y maquinaria.',

            'Gestión de Operaciones y Mantenimiento',

            'Ascendente.',

            'Indicador informativo.',

            '(Inspecciones ejecutadas / '
            'Inspecciones reportadas) × 100',

            frecuencia_reporte,

            inspecciones_reportadas,

            inspecciones_ejecutadas,

            porcentaje_inspecciones,

            'CUMPLE'
            if cumple_inspecciones
            else 'NO CUMPLE'
        )

        # ==========================================================
        # PIE DE DOCUMENTO
        # ==========================================================

        fila_pie = siguiente_fila + 1

        combinar_con_borde(
            fila_pie,
            1,
            10
        )

        ws.cell(
            fila_pie,
            1
        ).value = (
            'Documento generado automáticamente por Transmena Smart'
            'para Transmena y Carga.'
        )

        ws.cell(
            fila_pie,
            1
        ).font = Font(
            italic=True,
            size=9,
            color='666666'
        )

        ws.cell(
            fila_pie,
            1
        ).alignment = center

        # ==========================================================
        # CONFIGURACIÓN DE IMPRESIÓN
        # ==========================================================

        ws.freeze_panes = 'A9'

        ws.print_area = (
            f'A1:J{fila_pie}'
        )

        ws.page_setup.orientation = (
            'portrait'
        )

        ws.page_setup.paperSize = (
            ws.PAPERSIZE_A4
        )

        ws.page_setup.fitToWidth = 1

        ws.page_setup.fitToHeight = 0

        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_margins = PageMargins(

            left=0.25,

            right=0.25,

            top=0.35,

            bottom=0.35,

            header=0.15,

            footer=0.15
        )

        # ==========================================================
        # GENERAR ARCHIVO
        # ==========================================================

        output = BytesIO()

        wb.save(
            output
        )

        output.seek(0)

        nombre_archivo = (
            'HSEQ-R-159-Indicadores-Mantenimiento.xlsx'
        )

        return send_file(

            output,

            as_attachment=True,

            download_name=nombre_archivo,

            mimetype=(
                'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.sheet'
            )
        )

    except Exception as e:

        import traceback

        traceback.print_exc()

        return jsonify({

            'error':
                'Error generando el indicador HSEQ',

            'detalle':
                str(e)

        }), 500