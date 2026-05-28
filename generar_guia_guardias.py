from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

AZUL = colors.HexColor('#2962AB')
AZUL_CLARO = colors.HexColor('#EBF2FF')
VERDE = colors.HexColor('#1B7A3C')
NARANJA = colors.HexColor('#DC7814')
GRIS = colors.HexColor('#666666')

doc = SimpleDocTemplate(
    '/home/user/guardias-colegio/guia_guardias_profesores.pdf',
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2.5*cm, bottomMargin=2.5*cm,
    title='Guía Sistema de Guardias - La Asunción Donostia'
)

styles = getSampleStyleSheet()

titulo_principal = ParagraphStyle('TituloPrincipal', fontSize=22, textColor=AZUL,
    spaceAfter=4, alignment=TA_CENTER, fontName='Helvetica-Bold')
subtitulo_doc = ParagraphStyle('SubtituloDoc', fontSize=12, textColor=GRIS,
    spaceAfter=16, alignment=TA_CENTER, fontName='Helvetica')
cuerpo = ParagraphStyle('Cuerpo', fontSize=10, leading=15, spaceAfter=6,
    alignment=TA_JUSTIFY, fontName='Helvetica')
bullet_style = ParagraphStyle('Bullet', fontSize=10, leading=14, spaceAfter=3,
    leftIndent=16, fontName='Helvetica', bulletIndent=6)
subtitulo_sec = ParagraphStyle('SubtituloSec', fontSize=10, textColor=AZUL,
    spaceBefore=8, spaceAfter=4, fontName='Helvetica-Bold')
nota_style = ParagraphStyle('Nota', fontSize=9.5, leading=14, spaceAfter=0,
    alignment=TA_JUSTIFY, fontName='Helvetica', leftIndent=4)

def seccion(titulo, color=AZUL):
    data = [[Paragraph(f'<font color="white"><b>{titulo}</b></font>', ParagraphStyle(
        'Sec', fontSize=11, fontName='Helvetica-Bold', textColor=colors.white))]]
    t = Table(data, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), color),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    return t

def recuadro(titulo, texto, color=AZUL):
    rows = [
        [Paragraph(f'<b>{titulo}</b>', ParagraphStyle('RTitle', fontSize=10,
            fontName='Helvetica-Bold', textColor=color))],
        [Paragraph(texto, nota_style)],
    ]
    t = Table(rows, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), AZUL_CLARO),
        ('LINEABOVE', (0,0), (-1,0), 1.5, color),
        ('LINEBEFORE', (0,0), (-1,-1), 1.5, color),
        ('LINEAFTER', (0,0), (-1,-1), 0.5, colors.HexColor('#C0C0C0')),
        ('LINEBELOW', (0,-1), (-1,-1), 0.5, colors.HexColor('#C0C0C0')),
        ('TOPPADDING', (0,0), (-1,0), 5),
        ('BOTTOMPADDING', (0,0), (-1,0), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,1), (-1,-1), 6),
        ('BOTTOMPADDING', (0,1), (-1,-1), 6),
    ]))
    return t

def bala(texto):
    return Paragraph(f'<bullet>•</bullet>{texto}', bullet_style)

story = []

# ── Cabecera ─────────────────────────────────────────────────────────
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph('Guía del Sistema de Guardias', titulo_principal))
story.append(Paragraph('Curso 2025/2026  ·  La Asunción Donostia', subtitulo_doc))
story.append(HRFlowable(width='100%', thickness=1.5, color=AZUL, spaceAfter=12))
story.append(Paragraph(
    'Esta guía explica cómo funciona la aplicación de gestión de guardias del colegio. '
    'Podréis registrar ausencias, consultar vuestro horario semanal y ver las guardias '
    'asignadas desde cualquier ordenador o móvil con navegador.',
    cuerpo))
story.append(Spacer(1, 0.3*cm))

# ── 1. Acceso ─────────────────────────────────────────────────────────
story.append(seccion('1.  Acceso a la aplicación'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph('La aplicación está disponible en:', cuerpo))
story.append(Paragraph(
    '<font color="#2962AB"><b>https://guardias-colegio.onrender.com</b></font>',
    ParagraphStyle('URL', fontSize=11, alignment=TA_CENTER, spaceAfter=6,
        fontName='Helvetica-Bold')))
story.append(Paragraph(
    'El primer acceso se realiza con el enlace de invitación enviado por la dirección '
    'a vuestro correo electrónico. Ese enlace os permite crear vuestra contraseña. '
    'A partir de entonces, accedéis con vuestro correo y contraseña.',
    cuerpo))
story.append(Spacer(1, 0.4*cm))

# ── 2. Mi horario ─────────────────────────────────────────────────────
story.append(seccion('2.  Mi horario'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'En la sección <b>"Mi horario"</b> (menú lateral izquierdo) podéis consultar '
    'vuestro horario semanal completo: clases, guardias asignadas y horas libres.',
    cuerpo))
story.append(bala('Franjas en <b>azul</b>: clases asignadas.'))
story.append(bala('Franjas en <b>amarillo/naranja</b>: guardias pendientes de realizar.'))
story.append(bala('Franjas en <b>verde</b>: guardias ya completadas.'))
story.append(bala('Franjas <b>vacías</b>: horas libres (disponibles para cubrir guardias).'))
story.append(Spacer(1, 0.4*cm))

# ── 3. Registrar ausencia ─────────────────────────────────────────────
story.append(seccion('3.  Registrar una ausencia'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'Cuando sepáis con antelación que vais a faltar (cita médica, formación, etc.), '
    'debéis registrar la ausencia en <b>"Ausencias" → "Nueva ausencia"</b>.',
    cuerpo))
story.append(Paragraph('<b>Datos a rellenar:</b>', subtitulo_sec))
story.append(bala('<b>Fecha de inicio</b> (y fecha de fin si faltáis más de un día).'))
story.append(bala('<b>Motivo:</b> cita médica, formación, baja médica, etc.'))
story.append(bala(
    '<b>Franjas afectadas:</b> si no faltáis todo el día, indicad solo las horas '
    'concretas (9:00-10:00, 11:30-12:30…). Si no seleccionáis nada, se entiende que '
    'es el día completo.'))
story.append(bala('<b>Baja médica:</b> marcadlo si es una baja indefinida.'))
story.append(Spacer(1, 0.3*cm))
story.append(recuadro(
    'Modo de asignación',
    'Dependiendo de la configuración del centro, las guardias se asignan automáticamente '
    'al registrar la ausencia, o quedan pendientes de aprobación por la dirección. '
    'En cualquier caso, recibiréis un correo electrónico cuando la guardia quede asignada.',
    AZUL))
story.append(Spacer(1, 0.4*cm))

# ── 4. Cómo se asignan ───────────────────────────────────────────────
story.append(seccion('4.  Cómo se asignan las guardias'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'El sistema busca automáticamente el profesor más adecuado para cubrir cada franja. '
    'Los criterios de selección son, por orden de prioridad:',
    cuerpo))
story.append(bala('<b>1.º</b> El/la tutor/a de la clase donde se produce la ausencia.'))
story.append(bala('<b>2.º</b> Un profesor de la misma etapa educativa (Haur Hezkuntza o Lehen Hezkuntza).'))
story.append(bala('<b>3.º</b> Reparto equitativo: se evita sobrecargar a la misma persona, '
    'teniendo en cuenta las guardias realizadas esa semana y en total.'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'Solo se asigna a profesores que estén <b>libres en esa franja horaria</b> '
    '(sin clase y sin otra guardia ya asignada). Los especialistas y PT se utilizan '
    'únicamente si no hay nadie más disponible.',
    cuerpo))
story.append(Spacer(1, 0.4*cm))

# ── 5. Notificaciones ────────────────────────────────────────────────
story.append(seccion('5.  Notificaciones por correo'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'Cuando el sistema os asigna una guardia, recibiréis un correo con:', cuerpo))
story.append(bala('Fecha y franja horaria de la guardia.'))
story.append(bala('Nombre del profesor/a ausente y motivo.'))
story.append(bala('Aula donde debéis ir a hacer la guardia.'))
story.append(Spacer(1, 0.3*cm))
story.append(recuadro(
    'Importante',
    'El correo es un aviso informativo. La guardia ya está asignada a vosotros '
    'en el momento en que se envía. No es necesario confirmar nada para que quede registrada.',
    NARANJA))
story.append(Spacer(1, 0.4*cm))

# ── 6. Ver guardias ──────────────────────────────────────────────────
story.append(seccion('6.  Ver mis guardias asignadas'))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    'En la sección <b>"Guardias"</b> del menú lateral podéis ver todas las guardias del '
    'centro. Las vuestras aparecen destacadas. También podéis consultarlas en "Mi horario".',
    cuerpo))
story.append(bala('<b>Pendiente:</b> guardia asignada, aún no realizada.'))
story.append(bala('<b>Completada:</b> guardia ya realizada y marcada como tal por la dirección.'))
story.append(bala('<b>Sin cobertura:</b> no se encontró profesor disponible; la dirección lo gestionará.'))
story.append(Spacer(1, 0.4*cm))

# ── 7. FAQ ───────────────────────────────────────────────────────────
story.append(seccion('7.  Preguntas frecuentes', VERDE))
story.append(Spacer(1, 0.2*cm))

faqs = [
    ('¿Qué hago si me equivoco al registrar una ausencia?',
     'Contactad con la dirección para que la corrija o elimine desde el panel de administración.'),
    ('¿Puedo ver las ausencias de mis compañeros?',
     'No. Cada profesor solo ve sus propias ausencias y guardias. Solo los administradores tienen acceso al listado completo.'),
    ('¿Qué pasa si no hay nadie disponible para cubrir una guardia?',
     'La guardia queda como "Sin cobertura" y la dirección recibe una alerta para gestionarla manualmente.'),
    ('¿Puedo acceder desde el móvil?',
     'Sí, la aplicación está adaptada para móvil y tablet. Solo necesitáis el navegador web.'),
    ('¿Quién puede registrar ausencias?',
     'Cualquier profesor puede registrar su propia ausencia. Los administradores pueden registrar ausencias de cualquier compañero.'),
]

for pregunta, respuesta in faqs:
    story.append(Paragraph(f'<b>— {pregunta}</b>', subtitulo_sec))
    story.append(Paragraph(respuesta, cuerpo))

# ── Pie ──────────────────────────────────────────────────────────────
story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#C0C0C0'),
    spaceBefore=10, spaceAfter=6))
story.append(Paragraph(
    'Cualquier duda, contactad con la dirección del centro.',
    ParagraphStyle('Pie', fontSize=9, textColor=GRIS, alignment=TA_CENTER,
        fontName='Helvetica-Oblique')))

doc.build(story)
print('PDF generado: guia_guardias_profesores.pdf')
