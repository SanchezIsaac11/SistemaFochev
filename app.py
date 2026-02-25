import os
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
)
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from sqlalchemy import text

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "fochev.db")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# APP & CONFIG
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "fochev-secret-2024")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# --- Conexión de Base de Datos (Supabase / PostgreSQL en Render, o SQLite en local) ---
database_url = os.environ.get("DATABASE_URL", "")

if database_url:
    # Supabase/Render a veces entrega "postgres://", SQLAlchemy necesita "postgresql://"
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    # Asegurar SSL para Supabase
    if "sslmode" not in database_url:
        sep = "&" if "?" in database_url else "?"
        database_url += f"{sep}sslmode=require"
    app.config["SQLALCHEMY_DATABASE_URI"] = database_url
    # Pool settings para evitar conexiones caídas en instancias gratuitas de Render
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "pool_pre_ping": True,       # verifica la conexión antes de usarla
        "pool_recycle": 280,         # recicla conexiones antes de que expire el timeout del servidor
        "pool_timeout": 20,
        "pool_size": 5,
        "max_overflow": 10,
        "connect_args": {
            "sslmode": "require",
            "connect_timeout": 10,
        },
    }
else:
    # Desarrollo local con SQLite
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

db = SQLAlchemy(app)

# Flag para inicializar la DB solo una vez por proceso
_db_initialized = False


# ---------------------------------------------------------------------------
# MODELOS
# ---------------------------------------------------------------------------
class Zone(db.Model):
    __tablename__ = "zones"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    manager_name = db.Column(db.String(120), nullable=False)
    manager_email = db.Column(db.String(120), nullable=False)
    distributors = db.relationship("User", back_populates="zone", lazy=True)
    orders = db.relationship("Order", back_populates="zone", lazy=True)


class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # admin | distributor | customer
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(50), nullable=True)
    address = db.Column(db.String(255), nullable=True)
    city = db.Column(db.String(100), nullable=True)
    state = db.Column(db.String(100), nullable=True)
    postal_code = db.Column(db.String(20), nullable=True)
    zone_id = db.Column(db.Integer, db.ForeignKey("zones.id"), nullable=True)
    active = db.Column(db.Boolean, default=True, nullable=False)
    zone = db.relationship("Zone", back_populates="distributors", lazy=True)
    customer_orders = db.relationship(
        "Order", back_populates="customer",
        foreign_keys="Order.customer_id", lazy=True
    )
    distributor_orders = db.relationship(
        "Order", back_populates="distributor",
        foreign_keys="Order.distributor_id", lazy=True
    )

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


class Product(db.Model):
    __tablename__ = "products"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    description = db.Column(db.Text, nullable=True)
    price = db.Column(db.Float, nullable=False)
    stock = db.Column(db.Integer, nullable=False, default=0)
    active = db.Column(db.Boolean, default=True, nullable=False)
    order_items = db.relationship("OrderItem", back_populates="product", lazy=True)


class Order(db.Model):
    __tablename__ = "orders"
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    status = db.Column(db.String(20), nullable=False, default="pendiente")
    customer_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    distributor_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    zone_id = db.Column(db.Integer, db.ForeignKey("zones.id"), nullable=True)
    total = db.Column(db.Float, nullable=False, default=0.0)
    excel_path = db.Column(db.String(255), nullable=True)
    customer = db.relationship("User", foreign_keys=[customer_id], back_populates="customer_orders")
    distributor = db.relationship("User", foreign_keys=[distributor_id], back_populates="distributor_orders")
    zone = db.relationship("Zone", back_populates="orders")
    items = db.relationship("OrderItem", back_populates="order", lazy=True)


class OrderItem(db.Model):
    __tablename__ = "order_items"
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey("orders.id"), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    subtotal = db.Column(db.Float, nullable=False)
    order = db.relationship("Order", back_populates="items")
    product = db.relationship("Product", back_populates="order_items")


# ---------------------------------------------------------------------------
# SEED DATA
# ---------------------------------------------------------------------------
def init_db_with_sample_data():
    """Crea datos de ejemplo solo si la base aún no tiene un admin."""
    try:
        if User.query.filter_by(role="admin").first():
            return
    except Exception:
        # Las tablas aún no existen — db.create_all() se encargó de eso antes
        return

    # Zonas
    z_centro = Zone(name="Centro", manager_name="Carlos López", manager_email="carlos@fochev.com")
    z_norte = Zone(name="Norte", manager_name="Ana Martínez", manager_email="ana@fochev.com")
    db.session.add_all([z_centro, z_norte])
    db.session.flush()

    # Admin
    admin = User(
        email="admin@fochev.com",
        role="admin",
        first_name="Admin",
        last_name="Fochev",
        phone="8001234567",
        address="Av. Principal 100",
        city="CDMX",
        state="CDMX",
        postal_code="06600",
    )
    admin.set_password("admin123")

    # Distribuidores
    dist1 = User(
        email="dist.centro@fochev.com",
        role="distributor",
        first_name="Roberto",
        last_name="Sánchez",
        phone="8001112222",
        address="Calle 5 de Mayo 22",
        city="CDMX",
        state="CDMX",
        postal_code="06700",
        zone_id=z_centro.id,
    )
    dist1.set_password("dist123")

    dist2 = User(
        email="dist.norte@fochev.com",
        role="distributor",
        first_name="Laura",
        last_name="García",
        phone="8003334444",
        address="Blvd. Norte 88",
        city="Monterrey",
        state="Nuevo León",
        postal_code="64000",
        zone_id=z_norte.id,
    )
    dist2.set_password("dist123")

    # Cliente demo
    cliente = User(
        email="cliente.demo@fochev.com",
        role="customer",
        first_name="María",
        last_name="González",
        phone="8005556666",
        address="Calzada Sur 33",
        city="Guadalajara",
        state="Jalisco",
        postal_code="44100",
        zone_id=z_centro.id,
    )
    cliente.set_password("cliente123")

    db.session.add_all([admin, dist1, dist2, cliente])

    # Productos
    products = [
        Product(name="Shampoo Hidratante Pro", description="Limpieza profunda con hidratación intensiva 500 ml.",
                price=285.00, stock=50),
        Product(name="Acondicionador Reparador", description="Reconstruye fibra capilar dañada 400 ml.",
                price=265.00, stock=40),
        Product(name="Mascarilla Nutrición Total", description="Mascarilla profesional de keratina 250 ml.",
                price=320.00, stock=30),
        Product(name="Aceite Argan Premium", description="Serum antifrizz de argán marroquí 60 ml.",
                price=350.00, stock=25),
        Product(name="Spray Protector Térmico", description="Protección hasta 230 °C para planchas y rizadores.",
                price=230.00, stock=35),
        Product(name="Ampolleta Vitaminas C+E", description="Tratamiento intensivo en ampolleta 12 ml.",
                price=180.00, stock=60),
    ]
    db.session.add_all(products)
    try:
        db.session.commit()
        print("[INFO] Datos de ejemplo creados correctamente.")
    except Exception as e:
        db.session.rollback()
        print(f"[WARN] No se pudieron crear los datos de ejemplo: {e}")


def _init_db():
    """Crea tablas e inserta datos de ejemplo. Se llama una vez al arrancar."""
    global _db_initialized
    if _db_initialized:
        return
    try:
        # Verifica que la conexión a Supabase funcione
        db.session.execute(text("SELECT 1"))
        db.create_all()
        init_db_with_sample_data()
        _db_initialized = True
        print("[INFO] Base de datos inicializada correctamente.")
    except Exception as e:
        print(f"[ERROR] No se pudo inicializar la DB: {e}")
        # No matamos el proceso — gunicorn seguirá corriendo
        # El error aparecerá en los logs de Render


@app.before_request
def ensure_db_initialized():
    """Inicializa la DB en el primer request real (lazy init para Render)."""
    global _db_initialized
    if not _db_initialized:
        _init_db()


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    return db.session.get(User, user_id)


def login_required(role=None):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            user = get_current_user()
            if not user:
                flash("Debes iniciar sesión primero.", "warning")
                return redirect(url_for("login"))
            if role and user.role != role:
                flash("No tienes permiso para acceder a esta sección.", "warning")
                return redirect(url_for("index"))
            return fn(*args, **kwargs)
        return wrapper
    return decorator


@app.context_processor
def inject_now():
    return {"now": datetime.utcnow()}


# ---------------------------------------------------------------------------
# RUTAS GENERALES
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    user = get_current_user()
    if not user:
        return redirect(url_for("login"))
    if user.role == "admin":
        return redirect(url_for("admin_dashboard"))
    if user.role == "distributor":
        return redirect(url_for("distributor_dashboard"))
    return redirect(url_for("customer_dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if get_current_user():
        return redirect(url_for("index"))
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        user = User.query.filter_by(email=email, active=True).first()
        if user and user.check_password(password):
            session.clear()
            session["user_id"] = user.id
            session["role"] = user.role
            flash(f"Bienvenido, {user.first_name}!", "success")
            return redirect(url_for("index"))
        flash("Correo o contraseña incorrectos.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sesión cerrada correctamente.", "info")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# RUTAS ADMIN
# ---------------------------------------------------------------------------
@app.route("/admin")
@login_required(role="admin")
def admin_dashboard():
    user = get_current_user()
    total_orders = Order.query.count()
    total_sales = db.session.query(db.func.sum(Order.total)).scalar() or 0.0
    products_low_stock = Product.query.filter(Product.stock <= 5, Product.active == True).all()
    recent_orders = Order.query.order_by(Order.created_at.desc()).limit(10).all()
    return render_template(
        "admin_dashboard.html",
        user=user,
        total_orders=total_orders,
        total_sales=total_sales,
        products_low_stock=products_low_stock,
        recent_orders=recent_orders,
    )


# --- Productos ---
@app.route("/admin/products")
@login_required(role="admin")
def admin_products():
    products = Product.query.order_by(Product.name).all()
    return render_template("admin_products.html", products=products)


@app.route("/admin/products/new", methods=["GET", "POST"])
@login_required(role="admin")
def admin_product_new():
    if request.method == "POST":
        p = Product(
            name=request.form["name"].strip(),
            description=request.form.get("description", "").strip() or None,
            price=float(request.form["price"]),
            stock=int(request.form["stock"]),
            active="active" in request.form,
        )
        db.session.add(p)
        db.session.commit()
        flash("Producto creado correctamente.", "success")
        return redirect(url_for("admin_products"))
    return render_template("admin_product_form.html", product=None)


@app.route("/admin/products/<int:product_id>/edit", methods=["GET", "POST"])
@login_required(role="admin")
def admin_product_edit(product_id):
    p = db.session.get(Product, product_id)
    if not p:
        flash("Producto no encontrado.", "danger")
        return redirect(url_for("admin_products"))
    if request.method == "POST":
        p.name = request.form["name"].strip()
        p.description = request.form.get("description", "").strip() or None
        p.price = float(request.form["price"])
        p.stock = int(request.form["stock"])
        p.active = "active" in request.form
        db.session.commit()
        flash("Producto actualizado.", "success")
        return redirect(url_for("admin_products"))
    return render_template("admin_product_form.html", product=p)


@app.route("/admin/products/<int:product_id>/delete", methods=["POST"])
@login_required(role="admin")
def admin_product_delete(product_id):
    p = db.session.get(Product, product_id)
    if p:
        db.session.delete(p)
        db.session.commit()
        flash("Producto eliminado.", "info")
    return redirect(url_for("admin_products"))


# --- Distribuidores ---
@app.route("/admin/distributors")
@login_required(role="admin")
def admin_distributors():
    distributors = User.query.filter_by(role="distributor").order_by(User.last_name).all()
    return render_template("admin_distributors.html", distributors=distributors)


@app.route("/admin/distributors/new", methods=["GET", "POST"])
@login_required(role="admin")
def admin_distributor_new():
    zones = Zone.query.order_by(Zone.name).all()
    if request.method == "POST":
        d = User(
            email=request.form["email"].strip().lower(),
            role="distributor",
            first_name=request.form["first_name"].strip(),
            last_name=request.form["last_name"].strip(),
            phone=request.form.get("phone", "").strip() or None,
            address=request.form.get("address", "").strip() or None,
            city=request.form.get("city", "").strip() or None,
            state=request.form.get("state", "").strip() or None,
            postal_code=request.form.get("postal_code", "").strip() or None,
            zone_id=int(request.form["zone_id"]) if request.form.get("zone_id") else None,
            active=True,
        )
        d.set_password(request.form["password"])
        db.session.add(d)
        db.session.commit()
        flash("Distribuidor creado correctamente.", "success")
        return redirect(url_for("admin_distributors"))
    return render_template("admin_distributor_form.html", distributor=None, zones=zones)


@app.route("/admin/distributors/<int:dist_id>/edit", methods=["GET", "POST"])
@login_required(role="admin")
def admin_distributor_edit(dist_id):
    d = db.session.get(User, dist_id)
    if not d or d.role != "distributor":
        flash("Distribuidor no encontrado.", "danger")
        return redirect(url_for("admin_distributors"))
    zones = Zone.query.order_by(Zone.name).all()
    if request.method == "POST":
        d.email = request.form["email"].strip().lower()
        d.first_name = request.form["first_name"].strip()
        d.last_name = request.form["last_name"].strip()
        d.phone = request.form.get("phone", "").strip() or None
        d.address = request.form.get("address", "").strip() or None
        d.city = request.form.get("city", "").strip() or None
        d.state = request.form.get("state", "").strip() or None
        d.postal_code = request.form.get("postal_code", "").strip() or None
        d.zone_id = int(request.form["zone_id"]) if request.form.get("zone_id") else None
        d.active = "active" in request.form
        if request.form.get("password"):
            d.set_password(request.form["password"])
        db.session.commit()
        flash("Distribuidor actualizado.", "success")
        return redirect(url_for("admin_distributors"))
    return render_template("admin_distributor_form.html", distributor=d, zones=zones)


@app.route("/admin/distributors/<int:dist_id>/delete", methods=["POST"])
@login_required(role="admin")
def admin_distributor_delete(dist_id):
    d = db.session.get(User, dist_id)
    if d and d.role == "distributor":
        db.session.delete(d)
        db.session.commit()
        flash("Distribuidor eliminado.", "info")
    return redirect(url_for("admin_distributors"))


# --- Zonas ---
@app.route("/admin/zones")
@login_required(role="admin")
def admin_zones():
    zones = Zone.query.order_by(Zone.name).all()
    return render_template("admin_zones.html", zones=zones)


@app.route("/admin/zones/new", methods=["GET", "POST"])
@login_required(role="admin")
def admin_zone_new():
    if request.method == "POST":
        z = Zone(
            name=request.form["name"].strip(),
            manager_name=request.form["manager_name"].strip(),
            manager_email=request.form["manager_email"].strip().lower(),
        )
        db.session.add(z)
        db.session.commit()
        flash("Zona creada correctamente.", "success")
        return redirect(url_for("admin_zones"))
    return render_template("admin_zone_form.html", zone=None)


@app.route("/admin/zones/<int:zone_id>/edit", methods=["GET", "POST"])
@login_required(role="admin")
def admin_zone_edit(zone_id):
    z = db.session.get(Zone, zone_id)
    if not z:
        flash("Zona no encontrada.", "danger")
        return redirect(url_for("admin_zones"))
    if request.method == "POST":
        z.name = request.form["name"].strip()
        z.manager_name = request.form["manager_name"].strip()
        z.manager_email = request.form["manager_email"].strip().lower()
        db.session.commit()
        flash("Zona actualizada.", "success")
        return redirect(url_for("admin_zones"))
    return render_template("admin_zone_form.html", zone=z)


@app.route("/admin/zones/<int:zone_id>/delete", methods=["POST"])
@login_required(role="admin")
def admin_zone_delete(zone_id):
    z = db.session.get(Zone, zone_id)
    if z:
        db.session.delete(z)
        db.session.commit()
        flash("Zona eliminada.", "info")
    return redirect(url_for("admin_zones"))


# --- Ventas (admin) ---
@app.route("/admin/sales")
@login_required(role="admin")
def admin_sales():
    orders = Order.query.order_by(Order.created_at.desc()).all()
    return render_template("admin_sales.html", orders=orders)


# ---------------------------------------------------------------------------
# RUTAS DISTRIBUIDOR
# ---------------------------------------------------------------------------
@app.route("/distributor")
@login_required(role="distributor")
def distributor_dashboard():
    user = get_current_user()
    if user.zone_id:
        pending_orders = (
            Order.query
            .filter_by(zone_id=user.zone_id)
            .filter(Order.status.in_(["pendiente", "asignado", "en_ruta"]))
            .order_by(Order.created_at.desc())
            .all()
        )
    else:
        pending_orders = []
    return render_template("distributor_dashboard.html", user=user, pending_orders=pending_orders)


@app.route("/distributor/order/<int:order_id>", methods=["GET", "POST"])
@login_required(role="distributor")
def distributor_order_detail(order_id):
    distributor = get_current_user()
    order = db.session.get(Order, order_id)
    if not order or order.zone_id != distributor.zone_id:
        flash("Pedido no encontrado o no pertenece a tu zona.", "danger")
        return redirect(url_for("distributor_dashboard"))
    if request.method == "POST":
        new_status = request.form.get("status")
        valid_statuses = ["pendiente", "asignado", "en_ruta", "entregado", "cancelado"]
        if new_status in valid_statuses:
            order.status = new_status
            if new_status == "asignado" and not order.distributor_id:
                order.distributor_id = distributor.id
            db.session.commit()
            flash("Estatus actualizado correctamente.", "success")
        return redirect(url_for("distributor_order_detail", order_id=order.id))
    return render_template("distributor_order_detail.html", order=order)


# ---------------------------------------------------------------------------
# RUTAS CLIENTE
# ---------------------------------------------------------------------------
@app.route("/shop")
@login_required(role="customer")
def customer_dashboard():
    user = get_current_user()
    products = Product.query.filter_by(active=True).order_by(Product.name).all()
    return render_template("customer_dashboard.html", user=user, products=products)


@app.route("/shop/cart/add/<int:product_id>", methods=["POST"])
@login_required(role="customer")
def customer_add_to_cart(product_id):
    product = db.session.get(Product, product_id)
    if not product or not product.active:
        flash("Producto no disponible.", "warning")
        return redirect(url_for("customer_dashboard"))

    quantity = int(request.form.get("quantity", 1))
    if quantity < 1:
        quantity = 1
    if quantity > product.stock:
        flash(f"Stock insuficiente. Disponible: {product.stock}.", "warning")
        return redirect(url_for("customer_dashboard"))

    # Carrito en sesión: dict { str(product_id): quantity }
    cart = session.get("cart", {})
    pid_str = str(product_id)
    cart[pid_str] = cart.get(pid_str, 0) + quantity
    session["cart"] = cart
    flash(f"'{product.name}' agregado al carrito.", "success")
    return redirect(url_for("customer_dashboard"))


@app.route("/shop/cart")
@login_required(role="customer")
def customer_cart():
    user = get_current_user()
    cart = session.get("cart", {})
    items = []
    total = 0.0
    for pid_str, qty in cart.items():
        product = db.session.get(Product, int(pid_str))
        if product:
            subtotal = product.price * qty
            total += subtotal
            items.append(type("CartItem", (), {"product": product, "quantity": qty, "subtotal": subtotal})())
    return render_template("customer_cart.html", user=user, items=items, total=total)


@app.route("/shop/cart/confirm", methods=["POST"])
@login_required(role="customer")
def customer_confirm_order():
    user = get_current_user()
    cart = session.get("cart", {})
    if not cart:
        flash("Tu carrito está vacío.", "warning")
        return redirect(url_for("customer_dashboard"))

    # Calcular total y validar stock
    total = 0.0
    order_items_data = []
    for pid_str, qty in cart.items():
        product = db.session.get(Product, int(pid_str))
        if not product or not product.active:
            flash(f"El producto ya no está disponible. Por favor revisa tu carrito.", "warning")
            return redirect(url_for("customer_cart"))
        if product.stock < qty:
            flash(f"Stock insuficiente para '{product.name}'. Disponible: {product.stock}.", "warning")
            return redirect(url_for("customer_cart"))
        subtotal = product.price * qty
        total += subtotal
        order_items_data.append((product, qty, subtotal))

    # Crear orden
    order = Order(
        customer_id=user.id,
        zone_id=user.zone_id,
        total=total,
        status="pendiente",
    )
    db.session.add(order)
    db.session.flush()  # Para obtener order.id

    for product, qty, subtotal in order_items_data:
        item = OrderItem(
            order_id=order.id,
            product_id=product.id,
            quantity=qty,
            unit_price=product.price,
            subtotal=subtotal,
        )
        db.session.add(item)
        product.stock -= qty  # Descontar stock

    # Generar Excel
    try:
        excel_filename = f"pedido_{order.id}.xlsx"
        excel_path = os.path.join(EXPORT_DIR, excel_filename)
        wb = Workbook()
        ws = wb.active
        ws.title = f"Pedido #{order.id}"
        ws.append(["FOCHEV — Detalle de Pedido"])
        ws.append([])
        ws.append(["Pedido #", order.id])
        ws.append(["Fecha", datetime.utcnow().strftime("%Y-%m-%d %H:%M")])
        ws.append(["Estado", "Pendiente"])
        ws.append([])
        ws.append(["--- CLIENTE ---"])
        ws.append(["Nombre", f"{user.first_name} {user.last_name}"])
        ws.append(["Correo", user.email])
        ws.append(["Teléfono", user.phone or "-"])
        ws.append(["Dirección", user.address or "-"])
        ws.append(["Ciudad", user.city or "-"])
        ws.append(["Estado", user.state or "-"])
        ws.append(["C.P.", user.postal_code or "-"])
        ws.append([])
        if user.zone:
            ws.append(["--- ZONA ---"])
            ws.append(["Zona", user.zone.name])
            ws.append(["Responsable", user.zone.manager_name])
            ws.append(["Correo responsable", user.zone.manager_email])
            ws.append([])
        ws.append(["--- PRODUCTOS ---"])
        ws.append(["Producto", "Cantidad", "Precio unitario", "Subtotal"])
        for product, qty, subtotal in order_items_data:
            ws.append([product.name, qty, product.price, subtotal])
        ws.append([])
        ws.append(["TOTAL", "", "", total])
        wb.save(excel_path)
        order.excel_path = excel_path
    except Exception as e:
        print(f"[WARN] No se pudo generar el Excel: {e}")

    db.session.commit()
    session.pop("cart", None)
    flash(f"¡Pedido #{order.id} confirmado exitosamente! Se ha generado el Excel del pedido.", "success")
    return redirect(url_for("customer_orders"))


@app.route("/shop/orders")
@login_required(role="customer")
def customer_orders():
    user = get_current_user()
    orders = (
        Order.query
        .filter_by(customer_id=user.id)
        .order_by(Order.created_at.desc())
        .all()
    )
    return render_template("customer_orders.html", user=user, orders=orders)


# ---------------------------------------------------------------------------
# ARRANQUE
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    with app.app_context():
        _init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)