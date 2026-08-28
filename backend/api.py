#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mas Campo Inventory System - Backend API REST
Flask + SQLite (via SQLAlchemy)
Puerto: 5000
Base de datos: backend/mascampo.db (archivo fisico en disco)
"""

import os
import json
import uuid
from datetime import datetime
from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'mascampo.db')

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

CORS(app, resources={r"/api/*": {"origins": "*"}})
db = SQLAlchemy(app)

_db_initialized = False

@app.before_request
def ensure_db():
    global _db_initialized
    if not _db_initialized:
        init_db()
        _db_initialized = True

class User(db.Model):
    __tablename__ = 'users'
    id            = db.Column(db.String(64), primary_key=True)
    firstName     = db.Column(db.String(120), nullable=False)
    lastName      = db.Column(db.String(120), nullable=False)
    email         = db.Column(db.String(200), unique=True, nullable=False)
    password      = db.Column(db.Text, nullable=False)
    role          = db.Column(db.String(30), nullable=False, default='LOGISTICA')
    phone         = db.Column(db.String(40))
    address       = db.Column(db.String(300))
    vinculacion   = db.Column(db.String(100))
    document      = db.Column(db.String(30))
    active        = db.Column(db.Boolean, default=True)
    isSuperuser   = db.Column(db.Boolean, default=False)
    mustChangePassword = db.Column(db.Boolean, default=False)
    createdAt     = db.Column(db.String(30))
    def to_dict(self):
        return {'id':self.id,'firstName':self.firstName,'lastName':self.lastName,
                'name':f"{self.firstName} {self.lastName}",'email':self.email,
                'password':self.password,'role':self.role,'phone':self.phone or '',
                'address':self.address or '','vinculacion':self.vinculacion or 'Planta (Nomina)',
                'document':self.document or '','active':self.active,
                'isSuperuser':self.isSuperuser,'mustChangePassword':self.mustChangePassword,
                'createdAt':self.createdAt or ''}

class Product(db.Model):
    __tablename__ = 'products'
    id            = db.Column(db.String(64), primary_key=True)
    sku           = db.Column(db.String(100), unique=True, nullable=False)
    barcode       = db.Column(db.String(100))
    name          = db.Column(db.String(300), nullable=False)
    reference     = db.Column(db.String(200))
    description   = db.Column(db.Text)
    category      = db.Column(db.String(100))
    brand         = db.Column(db.String(100))
    supplier      = db.Column(db.String(100))
    requiresSerial= db.Column(db.Boolean, default=False)
    unitOfMeasure = db.Column(db.String(50))
    minStockAlert = db.Column(db.Integer, default=5)
    baseCost      = db.Column(db.Float, default=0)
    salePrice     = db.Column(db.Float, default=0)
    salePrice2    = db.Column(db.Float, default=0)
    salePrice3    = db.Column(db.Float, default=0)
    physicalStock = db.Column(db.Integer, default=0)
    reservedStock = db.Column(db.Integer, default=0)
    locationStock = db.Column(db.Text)
    createdAt     = db.Column(db.String(30))
    updatedAt     = db.Column(db.String(30))
    def to_dict(self):
        loc={}
        if self.locationStock:
            try: loc=json.loads(self.locationStock)
            except: loc={}
        return {'id':self.id,'sku':self.sku,'barcode':self.barcode or '','name':self.name,
                'reference':self.reference or '','description':self.description or '',
                'category':self.category or '','brand':self.brand or '',
                'supplier':self.supplier or '','requiresSerial':self.requiresSerial,
                'unitOfMeasure':self.unitOfMeasure or 'Unidad','minStockAlert':self.minStockAlert,
                'baseCost':self.baseCost,'salePrice':self.salePrice,
                'salePrice2':self.salePrice2 or 0.0,'salePrice3':self.salePrice3 or 0.0,
                'physicalStock':self.physicalStock,
                'reservedStock':self.reservedStock,'locationStock':loc,'stockByLocation':loc,
                'createdAt':self.createdAt or '','updatedAt':self.updatedAt or ''}

class Movement(db.Model):
    __tablename__ = 'movements'
    id            = db.Column(db.String(64), primary_key=True)
    type          = db.Column(db.String(50), nullable=False)
    productId     = db.Column(db.String(64), nullable=False)
    quantity      = db.Column(db.Integer, nullable=False)
    invoiceNumber = db.Column(db.String(100))
    customerId    = db.Column(db.String(64))
    customerName  = db.Column(db.String(300))
    serialNumber  = db.Column(db.String(200))
    notes         = db.Column(db.Text)
    userId        = db.Column(db.String(64))
    locationId    = db.Column(db.String(64))
    fromLocationId= db.Column(db.String(64))
    toLocationId  = db.Column(db.String(64))
    attachments   = db.Column(db.Text)
    date          = db.Column(db.String(30))
    def to_dict(self):
        att=[]
        if self.attachments:
            try: att=json.loads(self.attachments)
            except: att=[]
        return {'id':self.id,'type':self.type,'productId':self.productId,'quantity':self.quantity,
                'invoiceNumber':self.invoiceNumber or '','customerId':self.customerId or '',
                'customerName':self.customerName or '','serialNumber':self.serialNumber or '',
                'notes':self.notes or '','user':self.userId or '','userId':self.userId or '',
                'locationId':self.locationId or '','fromLocationId':self.fromLocationId or '',
                'toLocationId':self.toLocationId or '','attachments':att,'date':self.date or ''}

class Location(db.Model):
    __tablename__ = 'locations'
    id          = db.Column(db.String(64), primary_key=True)
    name        = db.Column(db.String(200), nullable=False)
    address     = db.Column(db.String(300))
    phone       = db.Column(db.String(50))
    manager     = db.Column(db.String(200))
    active      = db.Column(db.Boolean, default=True)
    createdAt   = db.Column(db.String(30))
    def to_dict(self):
        return {'id':self.id,'name':self.name,'address':self.address or '','phone':self.phone or '',
                'manager':self.manager or '','active':self.active,'createdAt':self.createdAt or ''}

class Customer(db.Model):
    __tablename__ = 'customers'
    id               = db.Column(db.String(64), primary_key=True)
    documentType     = db.Column(db.String(100))
    documentTypeAbbr = db.Column(db.String(10))
    documentNum      = db.Column(db.String(50))
    name             = db.Column(db.String(300), nullable=False)
    phone            = db.Column(db.String(50))
    email            = db.Column(db.String(200))
    city             = db.Column(db.String(100))
    address          = db.Column(db.String(300))
    regimenIva       = db.Column(db.String(100))
    extraData        = db.Column(db.Text)
    createdAt        = db.Column(db.String(30))
    def to_dict(self):
        extra={}
        if self.extraData:
            try: extra=json.loads(self.extraData)
            except: extra={}
        base={'id':self.id,'documentType':self.documentType or '','documentTypeAbbr':self.documentTypeAbbr or '',
              'documentNum':self.documentNum or '','name':self.name,'phone':self.phone or '',
              'email':self.email or '','city':self.city or '','address':self.address or '',
              'regimenIva':self.regimenIva or '','createdAt':self.createdAt or ''}
        base.update(extra)
        return base

class Reservation(db.Model):
    __tablename__ = 'reservations'
    id              = db.Column(db.String(64), primary_key=True)
    productId       = db.Column(db.String(64), nullable=False)
    customerId      = db.Column(db.String(64), nullable=False)
    locationId      = db.Column(db.String(64))
    quantity        = db.Column(db.Integer, nullable=False)
    status          = db.Column(db.String(30), default='ACTIVA')
    reason          = db.Column(db.Text)
    serialIds       = db.Column(db.Text)
    createdByUserId = db.Column(db.String(64))
    reservationDate = db.Column(db.String(30))
    expiryDate      = db.Column(db.String(30))
    def to_dict(self):
        sids=[]
        if self.serialIds:
            try: sids=json.loads(self.serialIds)
            except: sids=[]
        return {'id':self.id,'productId':self.productId,'customerId':self.customerId,
                'locationId':self.locationId or '','quantity':self.quantity,'status':self.status,
                'reason':self.reason or '','serialIds':sids,'createdByUserId':self.createdByUserId or '',
                'reservationDate':self.reservationDate or '','expiryDate':self.expiryDate or ''}

class SerializedItem(db.Model):
    __tablename__ = 'serialized_items'
    id                = db.Column(db.String(64), primary_key=True)
    productId         = db.Column(db.String(64), nullable=False)
    serialNumber      = db.Column(db.String(200), unique=True, nullable=False)
    status            = db.Column(db.String(30), default='EN_STOCK')
    currentCustomerId = db.Column(db.String(64))
    locationId        = db.Column(db.String(64))
    entryDate         = db.Column(db.String(30))
    saleDate          = db.Column(db.String(30))
    warrantyExpiry    = db.Column(db.String(30))
    invoiceNumber     = db.Column(db.String(100))
    events            = db.Column(db.Text)
    updatedAt         = db.Column(db.String(30))
    def to_dict(self):
        evts=[]
        if self.events:
            try: evts=json.loads(self.events)
            except: evts=[]
        return {'id':self.id,'productId':self.productId,'serialNumber':self.serialNumber,
                'status':self.status,'currentCustomerId':self.currentCustomerId or '',
                'locationId':self.locationId or '','entryDate':self.entryDate or '',
                'saleDate':self.saleDate or '','warrantyExpiry':self.warrantyExpiry or '',
                'invoiceNumber':self.invoiceNumber or '','events':evts,'updatedAt':self.updatedAt or ''}

class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    id          = db.Column(db.String(64), primary_key=True)
    action      = db.Column(db.String(50))
    module      = db.Column(db.String(100))
    description = db.Column(db.Text)
    userId      = db.Column(db.String(64))
    userEmail   = db.Column(db.String(200))
    userRole    = db.Column(db.String(30))
    ipAddress   = db.Column(db.String(50))
    timestamp   = db.Column(db.String(30))
    def to_dict(self):
        return {'id':self.id,'action':self.action or '','module':self.module or '',
                'description':self.description or '','userId':self.userId or '',
                'userEmail':self.userEmail or '','userRole':self.userRole or '',
                'ipAddress':self.ipAddress or '','timestamp':self.timestamp or ''}

class Category(db.Model):
    __tablename__ = 'categories'
    id          = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name        = db.Column(db.String(200), unique=True, nullable=False)
    description = db.Column(db.Text)
    color       = db.Column(db.String(50))
    createdAt   = db.Column(db.String(30))
    def to_dict(self):
        return {'id':self.id,'name':self.name,'description':self.description or '',
                'color':self.color or '','createdAt':self.createdAt or ''}

class PendingIntake(db.Model):
    __tablename__ = 'pending_intakes'
    id            = db.Column(db.String(64), primary_key=True)
    invoiceNumber = db.Column(db.String(100))
    productId     = db.Column(db.String(64))
    sku           = db.Column(db.String(100))
    productName   = db.Column(db.String(300))
    quantity      = db.Column(db.Integer, default=1)
    locationId    = db.Column(db.String(64))
    supplier      = db.Column(db.String(300))
    requester     = db.Column(db.String(200))
    requesterEmail= db.Column(db.String(200))
    status        = db.Column(db.String(50), default='PENDIENTE_APROBACION')
    dataJson      = db.Column(db.Text)
    createdAt     = db.Column(db.String(30))
    def to_dict(self):
        obj = {}
        if self.dataJson:
            try: obj = json.loads(self.dataJson)
            except: obj = {}
        base = {
            'id': self.id,
            'invoiceNumber': self.invoiceNumber or '',
            'productId': self.productId or '',
            'sku': self.sku or '',
            'productName': self.productName or '',
            'quantity': self.quantity or 1,
            'locationId': self.locationId or '',
            'supplier': self.supplier or '',
            'requester': self.requester or '',
            'requesterEmail': self.requesterEmail or '',
            'status': self.status or 'PENDIENTE_APROBACION',
            'createdAt': self.createdAt or ''
        }
        base.update(obj)
        return base

class SystemSetting(db.Model):
    __tablename__ = 'system_settings'
    key       = db.Column(db.String(100), primary_key=True)
    value     = db.Column(db.Text, nullable=False)
    updatedAt = db.Column(db.String(30))
    def to_dict(self):
        return {'key': self.key, 'value': self.value, 'updatedAt': self.updatedAt or ''}

SALT = '$'
INITIAL_USERS = [
    {'id':'usr-super','firstName':'Superusuario','lastName':'Gerencia','email':'gerencia@softproductiva.com',
     'role':'SUPERADMINISTRADOR','vinculacion':'Planta (Nomina)','address':'Sede Principal SoftProductiva',
     'phone':'+573000000000','document':'0000000000',
     'password':'$MAS_CAMPO_SECURE_SALT_2026$170cf7594de1a72cf96e7b2dc03b7ac2990d8320909c71f5faf4be2caa99862b',
     'active':True,'isSuperuser':True,'mustChangePassword':False},
]
INITIAL_CATEGORIES = [
    {'name':'EQUIPO_ELECTRONICO','description':'Equipos Serializados con Hoja de Vida','color':'#10b981'},
    {'name':'CONSUMIBLE','description':'Consumibles por Cantidad / Lote','color':'#f59e0b'},
    {'name':'ACCESORIO','description':'Accesorios y complementos','color':'#6366f1'},
    {'name':'REPUESTO','description':'Repuestos y partes','color':'#ef4444'},
    {'name':'SERVICIO','description':'Servicios tecnicos y mantenimiento','color':'#3b82f6'},
]

def now_iso(): return datetime.now().isoformat()
def gen_id(): return str(uuid.uuid4())

def init_db():
    db.create_all()
    # Migraciones automáticas de columnas SQLite para tablas existentes
    try:
        engine = db.engine
        with engine.connect() as conn:
            # Columnas nuevas de movements
            existing_mov_cols = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(movements)").fetchall()}
            mov_new_cols = {
                'customerName': 'VARCHAR(300)',
                'serialNumber': 'VARCHAR(200)',
                'locationId': 'VARCHAR(64)',
                'fromLocationId': 'VARCHAR(64)',
                'toLocationId': 'VARCHAR(64)',
                'notes': 'TEXT',
                'attachments': 'TEXT'
            }
            for col, col_type in mov_new_cols.items():
                if col not in existing_mov_cols:
                    conn.exec_driver_sql(f"ALTER TABLE movements ADD COLUMN {col} {col_type}")

            # Columnas nuevas de reservations
            existing_res_cols = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(reservations)").fetchall()}
            if 'locationId' not in existing_res_cols:
                conn.exec_driver_sql("ALTER TABLE reservations ADD COLUMN locationId VARCHAR(64)")

            # Columnas nuevas de products
            existing_prod_cols = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(products)").fetchall()}
            if 'reference' not in existing_prod_cols:
                conn.exec_driver_sql("ALTER TABLE products ADD COLUMN reference VARCHAR(200)")
            if 'supplier' not in existing_prod_cols:
                conn.exec_driver_sql("ALTER TABLE products ADD COLUMN supplier VARCHAR(100)")
            if 'salePrice2' not in existing_prod_cols:
                conn.exec_driver_sql("ALTER TABLE products ADD COLUMN salePrice2 FLOAT DEFAULT 0")
            if 'salePrice3' not in existing_prod_cols:
                conn.exec_driver_sql("ALTER TABLE products ADD COLUMN salePrice3 FLOAT DEFAULT 0")

            conn.commit()
    except Exception as e:
        print(f"[WARN] Error en migracion de columnas: {e}")

    for u in INITIAL_USERS:
        if not db.session.get(User, u['id']):
            db.session.add(User(id=u['id'],firstName=u['firstName'],lastName=u['lastName'],
                email=u['email'],password=u['password'],role=u['role'],phone=u['phone'],
                address=u['address'],vinculacion=u['vinculacion'],document=u['document'],
                active=u['active'],isSuperuser=u['isSuperuser'],
                mustChangePassword=u['mustChangePassword'],createdAt=now_iso()[:10]))
    for c in INITIAL_CATEGORIES:
        if not Category.query.filter_by(name=c['name']).first():
            db.session.add(Category(name=c['name'],description=c['description'],
                color=c['color'],createdAt=now_iso()[:10]))

    # Cargar clientes iniciales a SQLite si la tabla de clientes está vacía
    if Customer.query.count() == 0:
        cust_json_path = os.path.join(os.path.dirname(BASE_DIR), 'all_excel_customers.json')
        if not os.path.exists(cust_json_path):
            cust_json_path = os.path.join(BASE_DIR, 'all_excel_customers.json')
        if os.path.exists(cust_json_path):
            try:
                with open(cust_json_path, 'r', encoding='utf-8') as f:
                    raw_custs = json.load(f)
                known = {'id','documentType','documentTypeAbbr','documentNumber','documentNum','name','phone','email','city','address','regimenIva','createdAt'}
                for cdata in raw_custs:
                    cid = cdata.get('id') or gen_id()
                    doc_num = cdata.get('documentNumber') or cdata.get('documentNum') or ''
                    extra = {k:v for k,v in cdata.items() if k not in known}
                    db.session.add(Customer(
                        id=cid,
                        documentType=cdata.get('documentType', ''),
                        documentTypeAbbr=cdata.get('documentTypeAbbr', ''),
                        documentNum=doc_num,
                        name=cdata.get('name', 'Sin Nombre'),
                        phone=cdata.get('phone', ''),
                        email=cdata.get('email', ''),
                        city=cdata.get('city', ''),
                        address=cdata.get('address', ''),
                        regimenIva=cdata.get('regimenIva', ''),
                        extraData=json.dumps(extra, ensure_ascii=False),
                        createdAt=cdata.get('createdAt', now_iso()[:10])
                    ))
                print(f"[*] Se migraron {len(raw_custs)} clientes a la BD SQLite.")
            except Exception as ce:
                print(f"[WARN] Error al sembrar clientes en BD SQLite: {ce}")

    db.session.commit()
    print(f"BD SQLite lista: {DB_PATH}")

@app.route('/api/health')
def health():
    return jsonify({'status':'ok','users':User.query.count(),'products':Product.query.count(),
                    'movements':Movement.query.count(),'customers':Customer.query.count(),
                    'db':DB_PATH,'timestamp':now_iso()})

@app.route('/api/users', methods=['GET'])
def get_users(): return jsonify([u.to_dict() for u in User.query.all()])

@app.route('/api/users', methods=['POST'])
def save_users():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    for udata in data:
        uid=udata.get('id') or gen_id()
        email=(udata.get('email') or '').strip().lower()
        # Buscar por ID o por Email para evitar violaciones de unicidad en SQLite
        user = User.query.get(uid) or (User.query.filter_by(email=email).first() if email else None)
        
        fName = udata.get('firstName') or ''
        lName = udata.get('lastName') or ''
        if not fName and udata.get('name'):
            parts = str(udata.get('name')).strip().split(' ', 1)
            fName = parts[0]
            lName = parts[1] if len(parts) > 1 else ''

        if user:
            if fName: user.firstName = fName
            if lName: user.lastName = lName
            if email: user.email = email
            if udata.get('role'): user.role = udata.get('role')
            user.phone = udata.get('phone', user.phone)
            user.address = udata.get('address', user.address)
            user.vinculacion = udata.get('vinculacion', user.vinculacion)
            user.document = udata.get('document', user.document)
            if 'active' in udata: user.active = bool(udata.get('active'))
            if 'mustChangePassword' in udata: user.mustChangePassword = bool(udata.get('mustChangePassword'))
            if udata.get('password') and udata['password'] != user.password:
                user.password = udata['password']
        else:
            db.session.add(User(
                id=uid,
                firstName=fName or 'Usuario',
                lastName=lName or '',
                email=email,
                password=udata.get('password',''),
                role=udata.get('role','LOGISTICA'),
                phone=udata.get('phone',''),
                address=udata.get('address',''),
                vinculacion=udata.get('vinculacion','Planta (Nomina)'),
                document=udata.get('document',''),
                active=bool(udata.get('active', True)),
                isSuperuser=bool(udata.get('isSuperuser', False)),
                mustChangePassword=bool(udata.get('mustChangePassword', False)),
                createdAt=udata.get('createdAt', now_iso()[:10])
            ))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/users/<uid>', methods=['DELETE'])
def delete_user(uid):
    user = User.query.get(uid) or User.query.filter_by(email=uid).first()
    if not user:
        return jsonify({'error':'Usuario no encontrado'}), 404
    if user.isSuperuser or user.email == 'gerencia@softproductiva.com':
        return jsonify({'error':'No se puede eliminar el superadministrador'}), 403
    db.session.delete(user)
    db.session.commit()
    return jsonify({'ok':True})

@app.route('/api/products', methods=['GET'])
def get_products(): return jsonify([p.to_dict() for p in Product.query.all()])

@app.route('/api/products', methods=['POST'])
def save_products():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    for pdata in data:
        pid=pdata.get('id') or gen_id()
        loc_data = pdata.get('stockByLocation') or pdata.get('locationStock') or {}
        loc_str=json.dumps(loc_data)
        p=Product.query.get(pid)
        if p:
            p.sku=pdata.get('sku',p.sku); p.barcode=pdata.get('barcode',p.barcode)
            p.name=pdata.get('name',p.name); p.reference=pdata.get('reference',p.reference)
            p.description=pdata.get('description',p.description)
            p.category=pdata.get('category',p.category); p.brand=pdata.get('brand',p.brand)
            p.supplier=pdata.get('supplier',p.supplier)
            p.requiresSerial=pdata.get('requiresSerial',p.requiresSerial)
            p.unitOfMeasure=pdata.get('unitOfMeasure',p.unitOfMeasure)
            p.minStockAlert=pdata.get('minStockAlert',p.minStockAlert)
            p.baseCost=pdata.get('baseCost',p.baseCost); p.salePrice=pdata.get('salePrice',p.salePrice)
            p.salePrice2=pdata.get('salePrice2',p.salePrice2); p.salePrice3=pdata.get('salePrice3',p.salePrice3)
            p.physicalStock=pdata.get('physicalStock',p.physicalStock)
            p.reservedStock=pdata.get('reservedStock',p.reservedStock)
            p.locationStock=loc_str; p.updatedAt=now_iso()
        else:
            db.session.add(Product(id=pid,sku=pdata.get('sku',pid),barcode=pdata.get('barcode',''),
                name=pdata.get('name',''),reference=pdata.get('reference',''),
                description=pdata.get('description',''),
                category=pdata.get('category',''),brand=pdata.get('brand',''),
                supplier=pdata.get('supplier',''),
                requiresSerial=pdata.get('requiresSerial',False),
                unitOfMeasure=pdata.get('unitOfMeasure','Unidad'),
                minStockAlert=pdata.get('minStockAlert',0),baseCost=pdata.get('baseCost',0),
                salePrice=pdata.get('salePrice',0),salePrice2=pdata.get('salePrice2',0),
                salePrice3=pdata.get('salePrice3',0),physicalStock=pdata.get('physicalStock',0),
                reservedStock=pdata.get('reservedStock',0),locationStock=loc_str,
                createdAt=pdata.get('createdAt',now_iso()[:10]),updatedAt=now_iso()))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/movements', methods=['GET'])
def get_movements(): return jsonify([m.to_dict() for m in Movement.query.order_by(Movement.date.desc()).all()])

@app.route('/api/movements', methods=['POST'])
def save_movements():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    existing={m.id for m in Movement.query.all()}
    for mdata in data:
        mid=mdata.get('id') or gen_id()
        if mid not in existing:
            db.session.add(Movement(id=mid,type=mdata.get('type',''),productId=mdata.get('productId',''),
                quantity=mdata.get('quantity',0),invoiceNumber=mdata.get('invoiceNumber',''),
                customerId=mdata.get('customerId',''),customerName=mdata.get('customerName',''),
                serialNumber=mdata.get('serialNumber',''),
                notes=mdata.get('notes',''),userId=mdata.get('user',mdata.get('userId',mdata.get('registeredBy',''))),
                locationId=mdata.get('locationId',''),
                fromLocationId=mdata.get('fromLocationId',''),
                toLocationId=mdata.get('toLocationId',''),
                attachments=json.dumps(mdata.get('attachments',[])),date=mdata.get('date',now_iso())))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/locations', methods=['GET'])
def get_locations(): return jsonify([l.to_dict() for l in Location.query.all()])

@app.route('/api/locations', methods=['POST'])
def save_locations():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    for ldata in data:
        lid=ldata.get('id') or gen_id()
        loc=Location.query.get(lid)
        if loc:
            loc.name=ldata.get('name',loc.name)
            loc.address=ldata.get('address',loc.address)
            loc.phone=ldata.get('phone',loc.phone)
            loc.manager=ldata.get('manager',loc.manager)
            loc.active=ldata.get('active',loc.active)
        else:
            db.session.add(Location(id=lid,name=ldata.get('name','Sede'),address=ldata.get('address',''),
                phone=ldata.get('phone',''),manager=ldata.get('manager',''),
                active=ldata.get('active',True),createdAt=ldata.get('createdAt',now_iso()[:10])))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/customers', methods=['GET'])
def get_customers(): return jsonify([c.to_dict() for c in Customer.query.all()])

@app.route('/api/customers', methods=['POST'])
def save_customers():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    known={'id','documentType','documentTypeAbbr','documentNum','name','phone','email','city','address','regimenIva','createdAt'}
    for cdata in data:
        cid=cdata.get('id') or gen_id()
        extra={k:v for k,v in cdata.items() if k not in known}
        c=Customer.query.get(cid)
        if c:
            c.documentType=cdata.get('documentType',c.documentType)
            c.documentTypeAbbr=cdata.get('documentTypeAbbr',c.documentTypeAbbr)
            c.documentNum=cdata.get('documentNum',c.documentNum)
            c.name=cdata.get('name',c.name); c.phone=cdata.get('phone',c.phone)
            c.email=cdata.get('email',c.email); c.city=cdata.get('city',c.city)
            c.address=cdata.get('address',c.address); c.regimenIva=cdata.get('regimenIva',c.regimenIva)
            c.extraData=json.dumps(extra,ensure_ascii=False)
        else:
            db.session.add(Customer(id=cid,documentType=cdata.get('documentType',''),
                documentTypeAbbr=cdata.get('documentTypeAbbr',''),documentNum=cdata.get('documentNum',''),
                name=cdata.get('name','Sin Nombre'),phone=cdata.get('phone',''),email=cdata.get('email',''),
                city=cdata.get('city',''),address=cdata.get('address',''),regimenIva=cdata.get('regimenIva',''),
                extraData=json.dumps(extra,ensure_ascii=False),createdAt=cdata.get('createdAt',now_iso()[:10])))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/reservations', methods=['GET'])
def get_reservations(): return jsonify([r.to_dict() for r in Reservation.query.all()])

@app.route('/api/reservations', methods=['POST'])
def save_reservations():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    existing={r.id for r in Reservation.query.all()}
    for rdata in data:
        rid=rdata.get('id') or gen_id()
        if rid not in existing:
            db.session.add(Reservation(id=rid,productId=rdata.get('productId',''),
                customerId=rdata.get('customerId',''),locationId=rdata.get('locationId',''),
                quantity=rdata.get('quantity',1),
                status=rdata.get('status','ACTIVA'),reason=rdata.get('reason',''),
                serialIds=json.dumps(rdata.get('serialIds',[])),
                createdByUserId=rdata.get('createdByUserId',''),
                reservationDate=rdata.get('reservationDate',now_iso()),expiryDate=rdata.get('expiryDate','')))
        else:
            r=Reservation.query.get(rid)
            if r:
                r.status=rdata.get('status',r.status)
                if rdata.get('locationId'): r.locationId=rdata.get('locationId')
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/serialized-items', methods=['GET'])
def get_serialized_items(): return jsonify([i.to_dict() for i in SerializedItem.query.all()])

@app.route('/api/serialized-items', methods=['POST'])
def save_serialized_items():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    for idata in data:
        iid=idata.get('id') or gen_id()
        item=SerializedItem.query.get(iid)
        evts_str=json.dumps(idata.get('events',[]),ensure_ascii=False)
        if item:
            item.status=idata.get('status',item.status)
            item.currentCustomerId=idata.get('currentCustomerId',item.currentCustomerId)
            item.locationId=idata.get('locationId',item.locationId)
            item.saleDate=idata.get('saleDate',item.saleDate)
            item.warrantyExpiry=idata.get('warrantyExpiry',item.warrantyExpiry)
            item.invoiceNumber=idata.get('invoiceNumber',item.invoiceNumber)
            item.events=evts_str; item.updatedAt=now_iso()
        else:
            db.session.add(SerializedItem(id=iid,productId=idata.get('productId',''),
                serialNumber=idata.get('serialNumber',iid),status=idata.get('status','EN_STOCK'),
                currentCustomerId=idata.get('currentCustomerId',''),locationId=idata.get('locationId',''),
                entryDate=idata.get('entryDate',now_iso()[:10]),saleDate=idata.get('saleDate',''),
                warrantyExpiry=idata.get('warrantyExpiry',''),invoiceNumber=idata.get('invoiceNumber',''),
                events=evts_str,updatedAt=now_iso()))
    db.session.commit()
    return jsonify({'ok':True,'count':len(data)})

@app.route('/api/categories', methods=['GET'])
def get_categories(): return jsonify([c.to_dict() for c in Category.query.all()])

@app.route('/api/categories', methods=['POST'])
def save_categories():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    for cdata in data:
        name=cdata.get('name','').strip()
        if not name: continue
        cat=Category.query.filter_by(name=name).first()
        if cat: cat.description=cdata.get('description',cat.description); cat.color=cdata.get('color',cat.color)
        else: db.session.add(Category(name=name,description=cdata.get('description',''),
                color=cdata.get('color',''),createdAt=cdata.get('createdAt',now_iso()[:10])))
    db.session.commit()
    return jsonify({'ok':True})

@app.route('/api/audit-logs', methods=['GET'])
def get_audit_logs(): return jsonify([l.to_dict() for l in AuditLog.query.order_by(AuditLog.timestamp.desc()).limit(500).all()])

@app.route('/api/audit-logs', methods=['POST'])
def save_audit_logs():
    data=request.get_json()
    if not isinstance(data,list): return jsonify({'error':'Lista esperada'}),400
    existing={l.id for l in AuditLog.query.all()}
    for ldata in data:
        lid=ldata.get('id') or gen_id()
        if lid not in existing:
            db.session.add(AuditLog(id=lid,action=ldata.get('action',''),
                module=ldata.get('module',ldata.get('entityName','')),description=ldata.get('description',''),
                userId=ldata.get('userId',''),userEmail=ldata.get('userEmail',''),
                userRole=ldata.get('userRole',''),ipAddress=ldata.get('ipAddress',''),
                timestamp=ldata.get('timestamp',now_iso())))
    db.session.commit()
    return jsonify({'ok':True})

@app.route('/api/pending-intakes', methods=['GET'])
def get_pending_intakes():
    return jsonify([pi.to_dict() for pi in PendingIntake.query.order_by(PendingIntake.createdAt.desc()).all()])

@app.route('/api/pending-intakes', methods=['POST'])
def save_pending_intakes():
    data = request.get_json()
    if not isinstance(data, list): return jsonify({'error': 'Lista esperada'}), 400
    known = {'id','invoiceNumber','productId','sku','productName','quantity','locationId','supplier','requester','requesterEmail','status','createdAt'}
    for pdata in data:
        pid = pdata.get('id') or gen_id()
        extra = {k: v for k, v in pdata.items() if k not in known}
        pi = PendingIntake.query.get(pid)
        if pi:
            pi.invoiceNumber = pdata.get('invoiceNumber', pi.invoiceNumber)
            pi.productId = pdata.get('productId', pi.productId)
            pi.sku = pdata.get('sku', pi.sku)
            pi.productName = pdata.get('productName', pi.productName)
            pi.quantity = pdata.get('quantity', pi.quantity)
            pi.locationId = pdata.get('locationId', pi.locationId)
            pi.supplier = pdata.get('supplier', pi.supplier)
            pi.requester = pdata.get('requester', pi.requester)
            pi.requesterEmail = pdata.get('requesterEmail', pi.requesterEmail)
            pi.status = pdata.get('status', pi.status)
            pi.dataJson = json.dumps(extra, ensure_ascii=False)
        else:
            db.session.add(PendingIntake(
                id=pid,
                invoiceNumber=pdata.get('invoiceNumber', ''),
                productId=pdata.get('productId', ''),
                sku=pdata.get('sku', ''),
                productName=pdata.get('productName', ''),
                quantity=pdata.get('quantity', 1),
                locationId=pdata.get('locationId', ''),
                supplier=pdata.get('supplier', ''),
                requester=pdata.get('requester', ''),
                requesterEmail=pdata.get('requesterEmail', ''),
                status=pdata.get('status', 'PENDIENTE_APROBACION'),
                dataJson=json.dumps(extra, ensure_ascii=False),
                createdAt=pdata.get('createdAt', now_iso()[:10])
            ))
    db.session.commit()
    return jsonify({'ok': True, 'count': len(data)})

@app.route('/api/settings', methods=['GET'])
def get_settings():
    settings = {s.key: s.value for s in SystemSetting.query.all()}
    return jsonify(settings)

@app.route('/api/settings', methods=['POST'])
def save_settings():
    data = request.get_json()
    if not isinstance(data, dict): return jsonify({'error': 'Objeto esperado'}), 400
    for k, v in data.items():
        val_str = json.dumps(v, ensure_ascii=False) if not isinstance(v, str) else v
        s = SystemSetting.query.get(k)
        if s:
            s.value = val_str
            s.updatedAt = now_iso()
        else:
            db.session.add(SystemSetting(key=k, value=val_str, updatedAt=now_iso()))
    db.session.commit()
    return jsonify({'ok': True})

from flask import send_from_directory

@app.route('/')
def serve_index():
    project_root = os.path.dirname(BASE_DIR)
    return send_from_directory(project_root, 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    project_root = os.path.dirname(BASE_DIR)
    if os.path.exists(os.path.join(project_root, path)):
        return send_from_directory(project_root, path)
    return jsonify({'error': 'Not found'}), 404

if __name__ == '__main__':
    with app.app_context(): init_db()
    print('[OK] Servidor Mas Campo API listo en: http://localhost:8080')
    app.run(host='0.0.0.0', port=8080, debug=False)
