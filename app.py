from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_mail import Mail, Message
from itsdangerous import URLSafeTimedSerializer, SignatureExpired, BadSignature
from datetime import datetime, date, time
import os
import json
import io
import openpyxl

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'guardias-colegio-secret-2024')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///guardias.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Email config (se configura en settings)
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = ''
app.config['MAIL_PASSWORD'] = ''
app.config['MAIL_DEFAULT_SENDER'] = ''

db = SQLAlchemy(app)
mail = Mail(app)
login_manager = LoginManager(app)

@app.template_filter('from_json')
def from_json_filter(value):
    try:
        return json.loads(value)
    except Exception:
        return []
login_manager.login_view = 'login'
login_manager.login_message = 'Debes iniciar sesión para acceder.'
serializer = URLSafeTimedSerializer(app.config['SECRET_KEY'])

DIAS_SEMANA = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
FRANJAS = ['9:00-10:00', '10:00-11:00', 'Patio', '11:30-12:30', '14:30-15:30', '15:30-16:30']

# ───────────────────────────── MODELOS ─────────────────────────────

ETAPAS = ['1-2 años', '3-4-5 años', 'Primaria', 'Otras']

class Profesor(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200))
    etapa = db.Column(db.String(50))
    aula_tutoria = db.Column(db.String(50))
    aulas_bloqueadas = db.Column(db.Text, default='[]')  # JSON: ["2ºA", "3ºB"]
    es_admin = db.Column(db.Boolean, default=False)
    es_especialista = db.Column(db.Boolean, default=False)
    activo = db.Column(db.Boolean, default=True)
    de_baja = db.Column(db.Boolean, default=False)
    fecha_baja = db.Column(db.Date)
    fecha_vuelta = db.Column(db.Date)
    creado = db.Column(db.DateTime, default=datetime.utcnow)

    horario = db.relationship('HorarioProfesor', backref='profesor', lazy=True, cascade='all, delete-orphan')
    indisponibilidades = db.relationship('Indisponibilidad', backref='profesor', lazy=True, cascade='all, delete-orphan')
    guardias_asignadas = db.relationship('Guardia', foreign_keys='Guardia.profesor_asignado_id', backref='profesor_asignado', lazy=True)
    ausencias = db.relationship('Ausencia', foreign_keys='Ausencia.profesor_id', backref='profesor', lazy=True)

    @property
    def total_guardias(self):
        return Guardia.query.filter_by(profesor_asignado_id=self.id, completada=True).count()

    def guardias_semana(self, fecha_ref=None):
        if fecha_ref is None:
            fecha_ref = date.today()
        # Lunes y domingo de la semana de fecha_ref
        lunes = fecha_ref - __import__('datetime').timedelta(days=fecha_ref.weekday())
        domingo = lunes + __import__('datetime').timedelta(days=6)
        return Guardia.query.filter(
            Guardia.profesor_asignado_id == self.id,
            Guardia.fecha >= lunes,
            Guardia.fecha <= domingo
        ).count()

    @property
    def porcentaje_guardias(self):
        total_profesores_activos = Profesor.query.filter_by(activo=True, de_baja=False).count()
        if total_profesores_activos == 0:
            return 0
        total_guardias_sistema = Guardia.query.filter_by(completada=True).count()
        if total_guardias_sistema == 0:
            return 0
        esperado = total_guardias_sistema / total_profesores_activos
        if esperado == 0:
            return 0
        return round((self.total_guardias / esperado) * 100, 1)

    def esta_libre_en(self, dia, franja):
        horario = HorarioProfesor.query.filter_by(
            profesor_id=self.id, dia=dia, franja=franja, tiene_clase=True
        ).first()
        if horario:
            return False
        indisponible = Indisponibilidad.query.filter_by(
            profesor_id=self.id, dia=dia, franja=franja, activa=True
        ).first()
        if indisponible:
            return False
        return True


class HorarioProfesor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    profesor_id = db.Column(db.Integer, db.ForeignKey('profesor.id'), nullable=False)
    dia = db.Column(db.String(20), nullable=False)
    franja = db.Column(db.String(30), nullable=False)
    tiene_clase = db.Column(db.Boolean, default=False)
    asignatura = db.Column(db.String(100))


class Indisponibilidad(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    profesor_id = db.Column(db.Integer, db.ForeignKey('profesor.id'), nullable=False)
    dia = db.Column(db.String(20), nullable=False)
    franja = db.Column(db.String(30), nullable=False)
    motivo = db.Column(db.String(200))
    recurrente = db.Column(db.Boolean, default=False)
    fecha_especifica = db.Column(db.Date)
    activa = db.Column(db.Boolean, default=True)
    creada = db.Column(db.DateTime, default=datetime.utcnow)


class Ausencia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    profesor_id = db.Column(db.Integer, db.ForeignKey('profesor.id'), nullable=False)
    fecha_inicio = db.Column(db.Date, nullable=False)
    fecha_fin = db.Column(db.Date)
    motivo = db.Column(db.String(200))
    es_baja = db.Column(db.Boolean, default=False)
    notas = db.Column(db.Text)
    reportada_por = db.Column(db.Integer, db.ForeignKey('profesor.id'))
    creada = db.Column(db.DateTime, default=datetime.utcnow)
    # Franjas afectadas (JSON: ["9:00-10:00", "10:00-11:00", ...] o "todas")
    franjas_afectadas = db.Column(db.Text, default='todas')
    # Estado en modo manual: pendiente → el admin aún no ha aprobado
    aprobada = db.Column(db.Boolean, default=True)


class Guardia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False)
    dia_semana = db.Column(db.String(20), nullable=False)
    franja = db.Column(db.String(30), nullable=False)
    profesor_ausente_id = db.Column(db.Integer, db.ForeignKey('profesor.id'), nullable=False)
    profesor_asignado_id = db.Column(db.Integer, db.ForeignKey('profesor.id'))
    motivo_ausencia = db.Column(db.String(200))
    estado = db.Column(db.String(20), default='pendiente')  # pendiente, confirmada, rechazada, sin_cobertura
    completada = db.Column(db.Boolean, default=False)
    notas = db.Column(db.Text)
    creada = db.Column(db.DateTime, default=datetime.utcnow)
    aula = db.Column(db.String(50))

    profesor_ausente = db.relationship('Profesor', foreign_keys=[profesor_ausente_id])


class ConfiguracionEmail(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    mail_server = db.Column(db.String(100), default='smtp.gmail.com')
    mail_port = db.Column(db.Integer, default=587)
    mail_username = db.Column(db.String(120))
    mail_password = db.Column(db.String(200))
    activo = db.Column(db.Boolean, default=False)
    # Modo de gestión de guardias:
    # 'automatico' → el profesor reporta ausencia y se asigna guardia al momento
    # 'manual'     → el admin revisa la ausencia y aprueba la asignación
    modo_guardias = db.Column(db.String(20), default='manual')


# ───────────────────────────── AUTH ─────────────────────────────

@login_manager.user_loader
def load_user(user_id):
    return Profesor.query.get(int(user_id))


def check_password(stored, provided):
    import hashlib
    return hashlib.sha256(provided.encode()).hexdigest() == stored


def hash_password(password):
    import hashlib
    return hashlib.sha256(password.encode()).hexdigest()


# ───────────────────────────── ALGORITMO ─────────────────────────────

def puntuacion_candidato(profesor, aula_guardia, etapa_ausente, fecha):
    """
    Menor puntuación = mayor prioridad.
    Factores (de mayor a menor peso):
      1. Tutor de la clase donde se hace la guardia  → -100
      2. Misma etapa que el profesor ausente          → -25
      3. Guardias realizadas esta semana              → +15 por guardia
      4. Total de guardias acumuladas                 → +1 por guardia
    """
    score = 0

    # Factor 1: tutor de la clase concreta
    if aula_guardia and profesor.aula_tutoria and \
            profesor.aula_tutoria.strip().lower() == aula_guardia.strip().lower():
        score -= 100

    # Factor 2: misma etapa educativa
    if etapa_ausente and profesor.etapa and profesor.etapa == etapa_ausente:
        score -= 25

    # Factor 3: carga semanal (evitar saturar a alguien esta semana)
    score += profesor.guardias_semana(fecha) * 15

    # Factor 4: equidad acumulada a largo plazo
    score += profesor.total_guardias * 1

    return score


def buscar_profesor_para_guardia(dia, franja, excluir_id, aula=None, etapa_ausente=None, fecha=None):
    if fecha is None:
        fecha = date.today()

    def puede_cubrir(p):
        if p.id == excluir_id or not p.esta_libre_en(dia, franja):
            return False
        if aula:
            bloqueadas = json.loads(p.aulas_bloqueadas or '[]')
            if aula.strip().lower() in [b.strip().lower() for b in bloqueadas]:
                return False
        return True

    candidatos = Profesor.query.filter_by(activo=True, de_baja=False, es_especialista=False).all()
    candidatos = [p for p in candidatos if puede_cubrir(p)]

    if not candidatos:
        candidatos = Profesor.query.filter_by(activo=True, de_baja=False, es_especialista=True).all()
        candidatos = [p for p in candidatos if puede_cubrir(p)]

    if not candidatos:
        return None

    candidatos.sort(key=lambda p: puntuacion_candidato(p, aula, etapa_ausente, fecha))
    return candidatos[0]


def get_config():
    cfg = ConfiguracionEmail.query.first()
    if not cfg:
        cfg = ConfiguracionEmail()
        db.session.add(cfg)
        db.session.commit()
    return cfg

def get_mail_config():
    return get_config()


def mail_configurado():
    cfg = get_mail_config()
    return cfg and cfg.activo and cfg.mail_username


def init_mail(cfg):
    app.config['MAIL_USERNAME'] = cfg.mail_username
    app.config['MAIL_PASSWORD'] = cfg.mail_password
    app.config['MAIL_DEFAULT_SENDER'] = cfg.mail_username
    mail.init_app(app)


def enviar_invitacion(profesor):
    cfg = get_mail_config()
    if not cfg or not cfg.activo or not cfg.mail_username:
        return False
    try:
        init_mail(cfg)
        token = serializer.dumps(profesor.email, salt='invitacion-registro')
        url = url_for('registro_invitacion', token=token, _external=True)
        msg = Message(
            subject='Invitación al sistema de guardias — Colegio La Asunción',
            recipients=[profesor.email]
        )
        msg.html = render_template('email_invitacion.html', profesor=profesor, url=url)
        mail.send(msg)
        return True
    except Exception as e:
        print(f'Error enviando invitación: {e}')
        return False


def enviar_email_guardia(guardia):
    cfg = ConfiguracionEmail.query.first()
    if not cfg or not cfg.activo or not cfg.mail_username:
        return False
    try:
        app.config['MAIL_USERNAME'] = cfg.mail_username
        app.config['MAIL_PASSWORD'] = cfg.mail_password
        app.config['MAIL_DEFAULT_SENDER'] = cfg.mail_username
        mail.init_app(app)

        profesor = guardia.profesor_asignado
        ausente = guardia.profesor_ausente
        token_confirmar = serializer.dumps(f'confirmar:{guardia.id}', salt='guardia-action')
        token_rechazar = serializer.dumps(f'rechazar:{guardia.id}', salt='guardia-action')

        url_confirmar = url_for('accion_guardia', token=token_confirmar, _external=True)
        url_rechazar = url_for('accion_guardia', token=token_rechazar, _external=True)

        msg = Message(
            subject=f'Guardia asignada — {guardia.dia_semana} {guardia.franja}',
            recipients=[profesor.email]
        )
        msg.html = render_template('email_guardia.html',
            profesor=profesor, ausente=ausente, guardia=guardia,
            url_confirmar=url_confirmar, url_rechazar=url_rechazar)
        mail.send(msg)
        return True
    except Exception as e:
        print(f'Error email: {e}')
        return False


# ───────────────────────────── RUTAS AUTH ─────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        profesor = Profesor.query.filter_by(email=email).first()
        if profesor and profesor.password_hash and check_password(profesor.password_hash, password):
            login_user(profesor)
            return redirect(url_for('dashboard'))
        flash('Email o contraseña incorrectos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ───────────────────────────── DASHBOARD ─────────────────────────────

@app.route('/')
@login_required
def dashboard():
    hoy = date.today()
    guardias_hoy = Guardia.query.filter_by(fecha=hoy).all()
    mis_guardias_pendientes = Guardia.query.filter_by(
        profesor_asignado_id=current_user.id, completada=False
    ).order_by(Guardia.fecha).limit(5).all()
    total_profesores = Profesor.query.filter_by(activo=True, de_baja=False).count()
    de_baja = Profesor.query.filter_by(de_baja=True).count()
    guardias_mes = Guardia.query.filter(
        Guardia.fecha >= date(hoy.year, hoy.month, 1)
    ).count()
    return render_template('dashboard.html',
        guardias_hoy=guardias_hoy,
        mis_guardias=mis_guardias_pendientes,
        total_profesores=total_profesores,
        de_baja=de_baja,
        guardias_mes=guardias_mes,
        hoy=hoy)


# ───────────────────────────── PROFESORES ─────────────────────────────

@app.route('/profesores')
@login_required
def profesores():
    lista = Profesor.query.order_by(Profesor.nombre).all()
    return render_template('profesores.html', profesores=lista)


@app.route('/profesores/nuevo', methods=['GET', 'POST'])
@login_required
def nuevo_profesor():
    if not current_user.es_admin:
        flash('Solo los administradores pueden añadir profesores.', 'danger')
        return redirect(url_for('profesores'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        if Profesor.query.filter_by(email=email).first():
            flash('Ya existe un profesor con ese email.', 'danger')
            return render_template('form_profesor.html', profesor=None)
        raw = request.form.get('aulas_bloqueadas_raw', '')
        aulas_bloqueadas = json.dumps([a.strip() for a in raw.split(',') if a.strip()])
        p = Profesor(
            nombre=request.form.get('nombre', '').strip(),
            email=email,
            etapa=request.form.get('etapa', '').strip() or None,
            aula_tutoria=request.form.get('aula_tutoria', '').strip() or None,
            aulas_bloqueadas=aulas_bloqueadas,
            es_admin=bool(request.form.get('es_admin')),
            es_especialista=bool(request.form.get('es_especialista')),
        )
        db.session.add(p)
        db.session.commit()
        if mail_configurado():
            ok = enviar_invitacion(p)
            if ok:
                flash(f'Profesor {p.nombre} añadido. Se le ha enviado un email para que cree su contraseña.', 'success')
            else:
                flash(f'Profesor {p.nombre} añadido, pero el email no pudo enviarse. Configura primero el email o usa "Reenviar invitación".', 'warning')
        else:
            flash(f'Profesor {p.nombre} añadido. Configura el email para enviarle la invitación, o usa "Reenviar invitación" más adelante.', 'warning')
        return redirect(url_for('profesores'))
    return render_template('form_profesor.html', profesor=None)


@app.route('/profesores/<int:id>/enlace-invitacion')
@login_required
def enlace_invitacion(id):
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('profesores'))
    p = Profesor.query.get_or_404(id)
    token = serializer.dumps(p.email, salt='invitacion-registro')
    enlace = url_for('registro_invitacion', token=token, _external=True)
    return render_template('enlace_invitacion.html', profesor=p, enlace=enlace)


@app.route('/profesores/<int:id>/reenviar-invitacion', methods=['POST'])
@login_required
def reenviar_invitacion(id):
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('profesores'))
    p = Profesor.query.get_or_404(id)
    if p.password_hash:
        flash(f'{p.nombre} ya tiene contraseña establecida.', 'info')
        return redirect(url_for('profesores'))
    ok = enviar_invitacion(p)
    if ok:
        flash(f'Invitación reenviada a {p.email}.', 'success')
    else:
        flash('No se pudo enviar. Comprueba la configuración de email.', 'danger')
    return redirect(url_for('profesores'))


@app.route('/registro/<token>', methods=['GET', 'POST'])
def registro_invitacion(token):
    try:
        email = serializer.loads(token, salt='invitacion-registro', max_age=172800)  # 48h
    except SignatureExpired:
        flash('El enlace ha caducado. Pide al administrador que te reenvíe la invitación.', 'danger')
        return redirect(url_for('login'))
    except BadSignature:
        flash('Enlace no válido.', 'danger')
        return redirect(url_for('login'))

    profesor = Profesor.query.filter_by(email=email).first_or_404()
    if request.method == 'POST':
        password = request.form.get('password', '')
        confirm = request.form.get('confirm', '')
        if len(password) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'danger')
            return render_template('registro_invitacion.html', profesor=profesor, token=token)
        if password != confirm:
            flash('Las contraseñas no coinciden.', 'danger')
            return render_template('registro_invitacion.html', profesor=profesor, token=token)
        profesor.password_hash = hash_password(password)
        db.session.commit()
        login_user(profesor)
        flash(f'¡Bienvenido/a, {profesor.nombre.split()[0]}! Ya puedes usar la aplicación.', 'success')
        return redirect(url_for('dashboard'))
    return render_template('registro_invitacion.html', profesor=profesor, token=token)


@app.route('/profesores/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def editar_profesor(id):
    p = Profesor.query.get_or_404(id)
    if not current_user.es_admin and current_user.id != id:
        flash('No tienes permiso para editar este perfil.', 'danger')
        return redirect(url_for('profesores'))
    if request.method == 'POST':
        p.nombre = request.form.get('nombre', '').strip()
        p.etapa = request.form.get('etapa', '').strip() or None
        p.aula_tutoria = request.form.get('aula_tutoria', '').strip() or None
        raw = request.form.get('aulas_bloqueadas_raw', '')
        p.aulas_bloqueadas = json.dumps([a.strip() for a in raw.split(',') if a.strip()])
        if current_user.es_admin:
            p.es_admin = bool(request.form.get('es_admin'))
            p.es_especialista = bool(request.form.get('es_especialista'))
        nueva_pass = request.form.get('password', '').strip()
        if nueva_pass:
            p.password_hash = hash_password(nueva_pass)
        db.session.commit()
        flash('Perfil actualizado.', 'success')
        return redirect(url_for('profesores'))
    return render_template('form_profesor.html', profesor=p)


@app.route('/profesores/<int:id>/baja', methods=['POST'])
@login_required
def marcar_baja(id):
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('profesores'))
    p = Profesor.query.get_or_404(id)
    p.de_baja = True
    p.fecha_baja = date.today()
    ausencia = Ausencia(
        profesor_id=p.id,
        fecha_inicio=date.today(),
        motivo=request.form.get('motivo', 'Baja médica'),
        es_baja=True,
        notas=request.form.get('notas', ''),
        reportada_por=current_user.id
    )
    db.session.add(ausencia)
    db.session.commit()
    flash(f'{p.nombre} marcado de baja.', 'warning')
    return redirect(url_for('profesores'))


@app.route('/profesores/<int:id>/alta', methods=['POST'])
@login_required
def marcar_alta(id):
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('profesores'))
    p = Profesor.query.get_or_404(id)
    p.de_baja = False
    p.fecha_vuelta = date.today()
    ausencia = Ausencia.query.filter_by(profesor_id=p.id, es_baja=True, fecha_fin=None).first()
    if ausencia:
        ausencia.fecha_fin = date.today()
    db.session.commit()
    flash(f'{p.nombre} ha vuelto de baja.', 'success')
    return redirect(url_for('profesores'))


# ───────────────────────────── HORARIO PERSONAL ─────────────────────────────

@app.route('/mi-horario')
@login_required
def mi_horario():
    horario = {}
    for dia in DIAS_SEMANA:
        horario[dia] = {}
        for franja in FRANJAS:
            h = HorarioProfesor.query.filter_by(
                profesor_id=current_user.id, dia=dia, franja=franja
            ).first()
            ind = Indisponibilidad.query.filter_by(
                profesor_id=current_user.id, dia=dia, franja=franja, activa=True, recurrente=True
            ).first()
            horario[dia][franja] = {
                'tiene_clase': h.tiene_clase if h else False,
                'asignatura': h.asignatura if h else '',
                'indisponible': ind is not None,
                'motivo_ind': ind.motivo if ind else ''
            }
    return render_template('mi_horario.html', horario=horario, dias=DIAS_SEMANA, franjas=FRANJAS)


@app.route('/mi-horario/guardar', methods=['POST'])
@login_required
def guardar_horario():
    data = request.get_json()
    for dia in DIAS_SEMANA:
        for franja in FRANJAS:
            key = f'{dia}_{franja}'
            celda = data.get(key, {})
            h = HorarioProfesor.query.filter_by(
                profesor_id=current_user.id, dia=dia, franja=franja
            ).first()
            if not h:
                h = HorarioProfesor(profesor_id=current_user.id, dia=dia, franja=franja)
                db.session.add(h)
            h.tiene_clase = celda.get('tiene_clase', False)
            h.asignatura = celda.get('asignatura', '')

            ind = Indisponibilidad.query.filter_by(
                profesor_id=current_user.id, dia=dia, franja=franja, recurrente=True
            ).first()
            if celda.get('indisponible'):
                if not ind:
                    ind = Indisponibilidad(
                        profesor_id=current_user.id, dia=dia, franja=franja, recurrente=True
                    )
                    db.session.add(ind)
                ind.motivo = celda.get('motivo_ind', '')
                ind.activa = True
            elif ind:
                ind.activa = False
    db.session.commit()
    return jsonify({'ok': True})


# ───────────────────────────── AUSENCIAS ─────────────────────────────

@app.route('/ausencias')
@login_required
def ausencias():
    if current_user.es_admin:
        lista = Ausencia.query.order_by(Ausencia.creada.desc()).all()
    else:
        lista = Ausencia.query.filter_by(profesor_id=current_user.id).order_by(Ausencia.creada.desc()).all()
    return render_template('ausencias.html', ausencias=lista)


def crear_guardias_desde_ausencia(ausencia):
    """Crea y asigna guardias para una ausencia. Devuelve lista de guardias creadas."""
    profesor = Profesor.query.get(ausencia.profesor_id)
    fecha = ausencia.fecha_inicio
    guardias_creadas = []

    # Determinar qué franjas cubrir
    if ausencia.franjas_afectadas == 'todas':
        franjas = FRANJAS
    else:
        franjas = json.loads(ausencia.franjas_afectadas)

    dia_semana = DIAS_SEMANA[fecha.weekday()] if fecha.weekday() < 5 else None
    if not dia_semana:
        return []

    for franja in franjas:
        candidato = buscar_profesor_para_guardia(
            dia_semana, franja, profesor.id,
            aula=profesor.aula_tutoria,
            etapa_ausente=profesor.etapa,
            fecha=fecha
        )
        guardia = Guardia(
            fecha=fecha,
            dia_semana=dia_semana,
            franja=franja,
            profesor_ausente_id=profesor.id,
            motivo_ausencia=ausencia.motivo,
            estado='pendiente' if candidato else 'sin_cobertura'
        )
        if candidato:
            guardia.profesor_asignado_id = candidato.id
        db.session.add(guardia)
        guardias_creadas.append((guardia, candidato))

    db.session.commit()

    for guardia, candidato in guardias_creadas:
        if candidato:
            enviar_email_guardia(guardia)

    return guardias_creadas


@app.route('/ausencias/nueva', methods=['GET', 'POST'])
@login_required
def nueva_ausencia():
    cfg = get_config()
    if request.method == 'POST':
        fecha_inicio = datetime.strptime(request.form.get('fecha_inicio'), '%Y-%m-%d').date()
        fecha_fin_str = request.form.get('fecha_fin', '')
        fecha_fin = datetime.strptime(fecha_fin_str, '%Y-%m-%d').date() if fecha_fin_str else None
        es_baja = bool(request.form.get('es_baja'))

        profesor_id = int(request.form.get('profesor_id', current_user.id)) if current_user.es_admin else current_user.id

        franjas_sel = request.form.getlist('franjas')
        franjas_afectadas = json.dumps(franjas_sel) if franjas_sel else 'todas'

        modo = cfg.modo_guardias
        aprobada = (modo == 'automatico') or current_user.es_admin

        ausencia = Ausencia(
            profesor_id=profesor_id,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            motivo=request.form.get('motivo', '').strip(),
            es_baja=es_baja,
            notas=request.form.get('notas', '').strip(),
            reportada_por=current_user.id,
            franjas_afectadas=franjas_afectadas,
            aprobada=aprobada
        )
        db.session.add(ausencia)

        if es_baja:
            p = Profesor.query.get(profesor_id)
            p.de_baja = True
            p.fecha_baja = fecha_inicio

        db.session.commit()

        if aprobada and not es_baja:
            guardias = crear_guardias_desde_ausencia(ausencia)
            asignadas = sum(1 for _, c in guardias if c)
            sin_cobertura = len(guardias) - asignadas
            msg = f'Ausencia registrada. {asignadas} guardia(s) asignada(s) automáticamente.'
            if sin_cobertura:
                msg += f' {sin_cobertura} franja(s) sin cobertura disponible.'
            flash(msg, 'success' if not sin_cobertura else 'warning')
        elif not aprobada:
            flash('Ausencia registrada. El administrador la revisará y asignará la guardia.', 'info')
            # Notificar al admin por email si está configurado
            _notificar_admin_ausencia_pendiente(ausencia)
        else:
            flash('Baja registrada correctamente.', 'success')

        return redirect(url_for('ausencias'))

    profesores_lista = Profesor.query.filter_by(activo=True).order_by(Profesor.nombre).all() if current_user.es_admin else []
    return render_template('form_ausencia.html', profesores=profesores_lista, franjas=FRANJAS, cfg=cfg)


def _notificar_admin_ausencia_pendiente(ausencia):
    if not mail_configurado():
        return
    try:
        cfg = get_mail_config()
        init_mail(cfg)
        admins = Profesor.query.filter_by(es_admin=True, activo=True).all()
        profesor = Profesor.query.get(ausencia.profesor_id)
        for admin in admins:
            url = url_for('aprobar_ausencia', id=ausencia.id, _external=True)
            msg = Message(
                subject=f'Ausencia pendiente de aprobación — {profesor.nombre}',
                recipients=[admin.email]
            )
            msg.html = render_template('email_ausencia_pendiente.html',
                admin=admin, profesor=profesor, ausencia=ausencia, url=url)
            mail.send(msg)
    except Exception as e:
        print(f'Error notificando admin: {e}')


@app.route('/ausencias/<int:id>/aprobar', methods=['POST'])
@login_required
def aprobar_ausencia(id):
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('ausencias'))
    ausencia = Ausencia.query.get_or_404(id)
    ausencia.aprobada = True
    db.session.commit()
    guardias = crear_guardias_desde_ausencia(ausencia)
    asignadas = sum(1 for _, c in guardias if c)
    sin_cobertura = len(guardias) - asignadas
    msg = f'{asignadas} guardia(s) asignada(s).'
    if sin_cobertura:
        msg += f' {sin_cobertura} franja(s) sin cobertura.'
    flash(msg, 'success' if not sin_cobertura else 'warning')
    return redirect(url_for('ausencias'))


# ───────────────────────────── GUARDIAS ─────────────────────────────

@app.route('/guardias')
@login_required
def guardias():
    if current_user.es_admin:
        lista = Guardia.query.order_by(Guardia.fecha.desc(), Guardia.franja).all()
    else:
        lista = Guardia.query.filter(
            (Guardia.profesor_asignado_id == current_user.id) |
            (Guardia.profesor_ausente_id == current_user.id)
        ).order_by(Guardia.fecha.desc()).all()
    return render_template('guardias.html', guardias=lista)


@app.route('/guardias/nueva', methods=['GET', 'POST'])
@login_required
def nueva_guardia():
    if request.method == 'POST':
        fecha = datetime.strptime(request.form.get('fecha'), '%Y-%m-%d').date()
        dia_semana = DIAS_SEMANA[fecha.weekday()] if fecha.weekday() < 5 else 'Lunes'
        franja = request.form.get('franja')
        profesor_ausente_id = int(request.form.get('profesor_ausente_id'))

        aula = request.form.get('aula', '').strip()
        ausente = Profesor.query.get(profesor_ausente_id)
        etapa_ausente = ausente.etapa if ausente else None

        candidato = buscar_profesor_para_guardia(
            dia_semana, franja, profesor_ausente_id,
            aula=aula, etapa_ausente=etapa_ausente, fecha=fecha
        )

        guardia = Guardia(
            fecha=fecha,
            dia_semana=dia_semana,
            franja=franja,
            profesor_ausente_id=profesor_ausente_id,
            motivo_ausencia=request.form.get('motivo_ausencia', '').strip(),
            aula=aula,
            notas=request.form.get('notas', '').strip()
        )

        if candidato:
            guardia.profesor_asignado_id = candidato.id
            guardia.estado = 'pendiente'
            db.session.add(guardia)
            db.session.commit()
            email_ok = enviar_email_guardia(guardia)
            if email_ok:
                flash(f'Guardia asignada a {candidato.nombre} y notificado por email.', 'success')
            else:
                flash(f'Guardia asignada a {candidato.nombre}. (Email no configurado aún)', 'warning')
        else:
            guardia.estado = 'sin_cobertura'
            db.session.add(guardia)
            db.session.commit()
            flash('No hay ningún profesor disponible para esa franja. Guardia sin cobertura.', 'danger')

        return redirect(url_for('guardias'))

    profesores_lista = Profesor.query.filter_by(activo=True, de_baja=False).order_by(Profesor.nombre).all()
    return render_template('form_guardia.html', profesores=profesores_lista, franjas=FRANJAS)


@app.route('/guardia/<int:id>/completar', methods=['POST'])
@login_required
def completar_guardia(id):
    g = Guardia.query.get_or_404(id)
    g.completada = True
    g.estado = 'completada'
    db.session.commit()
    flash('Guardia marcada como completada.', 'success')
    return redirect(url_for('guardias'))


@app.route('/accion-guardia/<token>')
def accion_guardia(token):
    try:
        data = serializer.loads(token, salt='guardia-action', max_age=86400)
        accion, guardia_id = data.split(':')
        guardia = Guardia.query.get_or_404(int(guardia_id))
        if accion == 'confirmar':
            guardia.estado = 'confirmada'
            db.session.commit()
            return render_template('accion_guardia.html', accion='confirmada', guardia=guardia)
        elif accion == 'rechazar':
            guardia.estado = 'rechazada'
            candidato_anterior_id = guardia.profesor_asignado_id
            nuevo = buscar_profesor_para_guardia(
                guardia.dia_semana, guardia.franja, candidato_anterior_id,
                aula=guardia.aula,
                etapa_ausente=guardia.profesor_ausente.etapa if guardia.profesor_ausente else None,
                fecha=guardia.fecha
            )
            if nuevo and nuevo.id != candidato_anterior_id:
                guardia.profesor_asignado_id = nuevo.id
                guardia.estado = 'pendiente'
                db.session.commit()
                enviar_email_guardia(guardia)
                return render_template('accion_guardia.html', accion='rechazada_reasignada', guardia=guardia, nuevo=nuevo)
            else:
                guardia.estado = 'sin_cobertura'
                db.session.commit()
                return render_template('accion_guardia.html', accion='sin_cobertura', guardia=guardia)
    except Exception:
        return render_template('accion_guardia.html', accion='error', guardia=None)


# ───────────────────────────── IMPORTAR HORARIOS ─────────────────────────────

@app.route('/horarios/importar', methods=['GET', 'POST'])
@login_required
def importar_horarios():
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        f = request.files.get('archivo')
        if not f or not f.filename.endswith(('.xlsx', '.xls')):
            flash('Sube un fichero Excel (.xlsx).', 'danger')
            return redirect(url_for('importar_horarios'))

        try:
            wb = openpyxl.load_workbook(io.BytesIO(f.read()))
            ws = wb.active
            filas = list(ws.iter_rows(values_only=True))
            if not filas:
                flash('El fichero está vacío.', 'danger')
                return redirect(url_for('importar_horarios'))

            # Cabecera esperada: Profesor | Dia | Franja | Tipo | Asignatura
            errores = []
            importados = 0
            for i, fila in enumerate(filas[1:], start=2):  # salta cabecera
                if not any(fila):
                    continue
                nombre_prof = str(fila[0]).strip() if fila[0] else ''
                dia = str(fila[1]).strip() if fila[1] else ''
                franja = str(fila[2]).strip() if fila[2] else ''
                tipo = str(fila[3]).strip().lower() if fila[3] else ''
                asignatura = str(fila[4]).strip() if len(fila) > 4 and fila[4] else ''

                if not nombre_prof or not dia or not franja:
                    errores.append(f'Fila {i}: datos incompletos')
                    continue
                if dia not in DIAS_SEMANA:
                    errores.append(f'Fila {i}: día "{dia}" no reconocido')
                    continue
                if franja not in FRANJAS:
                    errores.append(f'Fila {i}: franja "{franja}" no reconocida')
                    continue

                # Buscar profesor por nombre (parcial, sin distinción mayúsculas)
                profesor = Profesor.query.filter(
                    Profesor.nombre.ilike(f'%{nombre_prof}%')
                ).first()
                if not profesor:
                    errores.append(f'Fila {i}: profesor "{nombre_prof}" no encontrado')
                    continue

                tiene_clase = tipo in ('clase', 'sí', 'si', 'yes', '1', 'true')
                es_indisponible = tipo in ('reunión', 'reunion', 'no disponible', 'ocupado', 'ind')

                h = HorarioProfesor.query.filter_by(
                    profesor_id=profesor.id, dia=dia, franja=franja
                ).first()
                if not h:
                    h = HorarioProfesor(profesor_id=profesor.id, dia=dia, franja=franja)
                    db.session.add(h)
                h.tiene_clase = tiene_clase
                h.asignatura = asignatura if tiene_clase else ''

                if es_indisponible:
                    ind = Indisponibilidad.query.filter_by(
                        profesor_id=profesor.id, dia=dia, franja=franja, recurrente=True
                    ).first()
                    if not ind:
                        ind = Indisponibilidad(
                            profesor_id=profesor.id, dia=dia, franja=franja, recurrente=True
                        )
                        db.session.add(ind)
                    ind.motivo = asignatura or 'Ocupado'
                    ind.activa = True

                importados += 1

            db.session.commit()
            msg = f'Importados {importados} registros correctamente.'
            if errores:
                msg += f' {len(errores)} filas con errores: ' + '; '.join(errores[:5])
                if len(errores) > 5:
                    msg += f' y {len(errores)-5} más.'
                flash(msg, 'warning')
            else:
                flash(msg, 'success')

        except Exception as e:
            flash(f'Error procesando el fichero: {e}', 'danger')

        return redirect(url_for('importar_horarios'))

    profesores_lista = Profesor.query.filter_by(activo=True).order_by(Profesor.nombre).all()
    return render_template('importar_horarios.html', profesores=profesores_lista, dias=DIAS_SEMANA, franjas=FRANJAS)


@app.route('/horarios/plantilla')
@login_required
def descargar_plantilla():
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('dashboard'))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Horarios'

    # Cabecera
    cabecera = ['Profesor', 'Dia', 'Franja', 'Tipo', 'Asignatura']
    for col, titulo in enumerate(cabecera, 1):
        ws.cell(row=1, column=col, value=titulo)
        ws.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)

    # Filas de ejemplo
    ejemplos = [
        ['García López, María', 'Lunes', '1ª hora', 'clase', 'Matemáticas'],
        ['García López, María', 'Lunes', '2ª hora', 'libre', ''],
        ['García López, María', 'Martes', 'Recreo', 'reunion', 'Reunión de departamento'],
        ['Fernández Ruiz, Juan', 'Lunes', '1ª hora', 'libre', ''],
    ]
    for row, fila in enumerate(ejemplos, 2):
        for col, val in enumerate(fila, 1):
            ws.cell(row=row, column=col, value=val)

    # Hoja de ayuda
    ws2 = wb.create_sheet('Instrucciones')
    instrucciones = [
        ['INSTRUCCIONES DE USO'],
        [''],
        ['Columna "Tipo" — valores válidos:'],
        ['  clase       → el profesor tiene clase en esa franja (no puede hacer guardia)'],
        ['  libre        → franja libre (puede hacer guardia)'],
        ['  reunion     → ocupado por reunión u otro motivo (no puede hacer guardia)'],
        [''],
        ['Columna "Dia" — valores exactos:'],
        ['  Lunes, Martes, Miércoles, Jueves, Viernes'],
        [''],
        ['Columna "Franja" — valores exactos:'],
        ['  1ª hora, 2ª hora, 3ª hora, Recreo, 4ª hora, 5ª hora, 6ª hora'],
        [''],
        ['El nombre del profesor debe coincidir (al menos en parte) con el nombre en la aplicación.'],
    ]
    for row, linea in enumerate(instrucciones, 1):
        ws2.cell(row=row, column=1, value=linea[0])

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 30

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    from flask import send_file
    return send_file(output, download_name='plantilla_horarios.xlsx',
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ───────────────────────────── ESTADÍSTICAS ─────────────────────────────

@app.route('/estadisticas')
@login_required
def estadisticas():
    profesores_lista = Profesor.query.filter_by(activo=True).order_by(Profesor.nombre).all()
    hoy = date.today()
    guardias_por_mes = []
    for mes in range(1, hoy.month + 1):
        count = Guardia.query.filter(
            db.extract('month', Guardia.fecha) == mes,
            db.extract('year', Guardia.fecha) == hoy.year
        ).count()
        guardias_por_mes.append({'mes': mes, 'total': count})
    return render_template('estadisticas.html',
        profesores=profesores_lista,
        guardias_por_mes=guardias_por_mes)


# ───────────────────────────── CONFIGURACIÓN EMAIL ─────────────────────────────

@app.route('/configuracion', methods=['GET', 'POST'])
@login_required
def configuracion():
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('dashboard'))
    cfg = get_config()
    if request.method == 'POST':
        cfg.mail_username = request.form.get('mail_username', '').strip()
        cfg.mail_password = request.form.get('mail_password', '').strip()
        cfg.activo = bool(request.form.get('activo'))
        db.session.commit()
        flash('Configuración de email guardada.', 'success')
    return render_template('configuracion.html', cfg=cfg)


@app.route('/configuracion/modo', methods=['POST'])
@login_required
def cambiar_modo():
    if not current_user.es_admin:
        flash('Solo administradores.', 'danger')
        return redirect(url_for('configuracion'))
    cfg = get_config()
    cfg.modo_guardias = request.form.get('modo_guardias', 'manual')
    db.session.commit()
    modo_texto = 'Automático' if cfg.modo_guardias == 'automatico' else 'Manual'
    flash(f'Modo cambiado a: {modo_texto}.', 'success')
    return redirect(url_for('configuracion'))


# ───────────────────────────── INIT DB ─────────────────────────────

def init_db():
    db.create_all()
    if not Profesor.query.filter_by(es_admin=True).first():
        admin = Profesor(
            nombre='Administrador',
            email='admin@colegio.es',
            es_admin=True,
            password_hash=hash_password('admin123')
        )
        db.session.add(admin)
        db.session.commit()
        print('Admin creado: admin@colegio.es / admin123')


# Se ejecuta tanto con gunicorn como con python app.py
with app.app_context():
    init_db()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
