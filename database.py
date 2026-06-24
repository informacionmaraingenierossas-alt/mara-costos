from sqlalchemy import create_engine, Column, Integer, String, Float, Date, ForeignKey, Text, DateTime, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
import datetime

DATABASE_URL = "sqlite:///mara_ingenieros.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class Usuario(Base):
    __tablename__ = "usuarios"
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, unique=True, index=True, nullable=False)
    password_hash = Column(String, nullable=False)
    rol = Column(String, nullable=False)  # Gerencia, Operario, Auxiliar Contable
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    is_active = Column(Boolean, default=True)

class Proyecto(Base):
    __tablename__ = "proyectos"
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, nullable=False)
    ubicacion = Column(String, default="Colombia")
    latitud = Column(Float, nullable=True)
    longitud = Column(Float, nullable=True)
    estado = Column(String, default="Activo")
    cliente = Column(String, default="WOM")
    n_requerimiento = Column(String, default="")
    acta_conciliacion = Column(String, default="SIN CONCILIAR")
    created_by = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    partidas = relationship("PartidaPresupuesto", back_populates="proyecto")
    gastos = relationship("Gasto", back_populates="proyecto")

class Proveedor(Base):
    __tablename__ = "proveedores"
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, nullable=False)
    nit = Column(String)
    contacto = Column(String)
    telefono = Column(String)
    created_by = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    gastos = relationship("Gasto", back_populates="proveedor")

class PartidaPresupuesto(Base):
    __tablename__ = "partidas_presupuesto"
    id = Column(Integer, primary_key=True, index=True)
    proyecto_id = Column(Integer, ForeignKey("proyectos.id"), nullable=False)
    categoria = Column(String, nullable=False)
    descripcion = Column(String, nullable=False)
    cantidad = Column(Float, default=1.0)
    valor_unitario = Column(Float, nullable=False)
    total = Column(Float, nullable=False)
    # Campos para gestión de cobros (ingresos)
    cobrado = Column(Boolean, default=False)
    conciliado_ingreso = Column(Boolean, default=False)
    acta_conciliacion_ingreso = Column(String, nullable=True)
    archivo_evidencia_ingreso = Column(String, nullable=True)
    created_by = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    proyecto = relationship("Proyecto", back_populates="partidas")
    cobros = relationship("CobroCliente", back_populates="partida")

class Gasto(Base):
    __tablename__ = "gastos"
    id = Column(Integer, primary_key=True, index=True)
    proyecto_id = Column(Integer, ForeignKey("proyectos.id"), nullable=False)
    concepto = Column(String, nullable=False)
    categoria = Column(String, nullable=False)
    unidad = Column(String, default="U")
    cantidad = Column(Float, default=1.0)
    valor_unitario = Column(Float, nullable=False)
    valor_total = Column(Float, nullable=False)
    proveedor_id = Column(Integer, ForeignKey("proveedores.id"), nullable=True)
    # Campos para gestión de gastos (pagos a proveedores)
    estado_pago = Column(String, default="Pendiente")          # Pendiente, Parcial, Pagado
    conciliado = Column(Boolean, default=False)                # Conciliado contablemente
    acta_conciliacion = Column(String, nullable=True)
    archivo_evidencia = Column(String, nullable=True)
    created_by = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    proyecto = relationship("Proyecto", back_populates="gastos")
    proveedor = relationship("Proveedor", back_populates="gastos")
    pagos = relationship("Pago", back_populates="gasto")

class Pago(Base):
    __tablename__ = "pagos"
    id = Column(Integer, primary_key=True, index=True)
    gasto_id = Column(Integer, ForeignKey("gastos.id"), nullable=False)
    tipo = Column(String, nullable=False)  # Factura, Anticipo
    numero_factura = Column(String)
    fecha = Column(Date, default=datetime.date.today)
    monto = Column(Float, nullable=False)
    concepto = Column(String(50), nullable=True)
    observaciones = Column(Text)
    created_by = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    gasto = relationship("Gasto", back_populates="pagos")

class CobroCliente(Base):
    __tablename__ = "cobros_clientes"
    id = Column(Integer, primary_key=True, index=True)
    partida_id = Column(Integer, ForeignKey("partidas_presupuesto.id"), nullable=False)
    proyecto_id = Column(Integer, ForeignKey("proyectos.id"), nullable=False)
    monto = Column(Float, nullable=False)
    fecha = Column(Date, default=datetime.date.today)
    numero_factura = Column(String, nullable=True)
    concepto = Column(String(50), nullable=True)
    observaciones = Column(Text, nullable=True)
    creado_por = Column(Integer, ForeignKey("usuarios.id"))
    created_at = Column(DateTime, default=datetime.datetime.now)
    updated_at = Column(DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)
    updated_by = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    partida = relationship("PartidaPresupuesto", back_populates="cobros")

class Auditoria(Base):
    __tablename__ = "auditoria"
    id = Column(Integer, primary_key=True, index=True)
    tabla = Column(String, nullable=False)
    registro_id = Column(Integer, nullable=False)
    accion = Column(String, nullable=False)  # insert, update, delete
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    datos_anteriores = Column(Text, nullable=True)
    datos_nuevos = Column(Text, nullable=True)
    fecha = Column(DateTime, default=datetime.datetime.now)

Base.metadata.create_all(bind=engine)