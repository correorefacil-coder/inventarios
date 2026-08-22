const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');

const app = express();
const PORT = process.env.PORT || 4000;
const JWT_SECRET = process.env.JWT_SECRET || 'mas-campo-super-secret-key-2026';

app.use(cors());
app.use(express.json());

const fs = require('fs');
const path = require('path');

const LOGS_DIR = path.join(__dirname, '../logs');
const LOGS_FILE = path.join(LOGS_DIR, 'audit_activity.json');

if (!fs.existsSync(LOGS_DIR)) {
  fs.mkdirSync(LOGS_DIR, { recursive: true });
}

function loadAuditLogsFromDisk() {
  try {
    if (fs.existsSync(LOGS_FILE)) {
      const data = fs.readFileSync(LOGS_FILE, 'utf8');
      if (data.trim()) return JSON.parse(data);
    }
  } catch (err) {
    console.error('Error leyendo logs de auditoría desde archivo físico:', err);
  }
  return [];
}

function appendAuditLogToFile(logEntry) {
  try {
    const logs = loadAuditLogsFromDisk();
    logs.unshift(logEntry);
    fs.writeFileSync(LOGS_FILE, JSON.stringify(logs, null, 2), 'utf8');
    console.log(`💾 [DISCO FISICO] Log guardado en archivo '${LOGS_FILE}'`);
  } catch (err) {
    console.error('Error escribiendo log de auditoría en archivo físico:', err);
  }
}

// In-Memory Database (Synced with Disk Logs)
const db = {
  roles: ['ADMINISTRADOR', 'SUPERADMINISTRADOR', 'LOGISTICA', 'VENTAS'],
  auditLogs: loadAuditLogsFromDisk(),
  products: [],
  movements: [],
  serializedItems: [],
  reservations: []
};

// Middlewares
function authMiddleware(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Acceso no autorizado: Token JWT requerido' });
  }

  const token = authHeader.split(' ')[1];
  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded;
    next();
  } catch (err) {
    return res.status(401).json({ error: 'Token JWT inválido o expirado' });
  }
}

// RBAC Middleware Guard
function rbacGuard(allowedRoles) {
  return (req, res, next) => {
    const userRole = req.user ? req.user.role : 'Desconocido';
    const isAdminRole = userRole === 'ADMINISTRADOR' || userRole === 'SUPERADMINISTRADOR';
    const isAllowed = allowedRoles.includes(userRole) || (allowedRoles.includes('ADMINISTRADOR') && isAdminRole);
    
    if (!req.user || !isAllowed) {
      return res.status(403).json({ 
        error: `Acceso Denegado (RBAC): Su rol '${userRole}' no tiene permisos para esta acción.` 
      });
    }
    next();
  };
}

// Audit Logging Interceptor Middleware (Writes to physical file on disk)
function auditMiddleware(action, entityName) {
  return (req, res, next) => {
    const userId = req.user ? req.user.id : 'ANONYMOUS';
    const logEntry = {
      id: `audit-${Date.now()}`,
      userId,
      action,
      entityName,
      payload: req.body,
      ip: req.ip || req.connection.remoteAddress,
      timestamp: new Date().toISOString()
    };
    db.auditLogs.unshift(logEntry);
    appendAuditLogToFile(logEntry);
    next();
  };
}

// Endpoints

// 1. Auth Endpoint
app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body;
  
  // Mock login check
  let role = 'ADMINISTRADOR';
  if (email.includes('superadmin') || email.includes('gerencia')) role = 'SUPERADMINISTRADOR';
  else if (email.includes('logistica')) role = 'LOGISTICA';
  else if (email.includes('ventas')) role = 'VENTAS';

  const user = { id: 'usr-101', name: 'Carlos Mendoza', email, role };
  const token = jwt.sign(user, JWT_SECRET, { expiresIn: '8h' });

  res.json({ message: 'Autenticación exitosa', token, user });
});

// 2. Products Catalog (RBAC Sensitive: BaseCost masked for non-Admins)
app.get('/api/products', authMiddleware, (req, res) => {
  const userRole = req.user.role;
  const isAdmin = userRole === 'ADMINISTRADOR' || userRole === 'SUPERADMINISTRADOR';

  // Filter financial details if not Admin
  const safeProducts = db.products.map(p => {
    if (!isAdmin) {
      const { baseCost, ...safeProduct } = p;
      return safeProduct;
    }
    return p;
  });

  res.json(safeProducts);
});

// 3. Create Product (Admin Only)
app.post('/api/products', authMiddleware, rbacGuard(['ADMINISTRADOR', 'SUPERADMINISTRADOR']), auditMiddleware('CREATE', 'Product'), (req, res) => {
  const newProduct = { id: `prod-${Date.now()}`, ...req.body, physicalStock: req.body.physicalStock || 0, reservedStock: req.body.reservedStock || 0 };
  db.products.push(newProduct);
  res.status(201).json(newProduct);
});

// 3b. Update Product & Inventory Values (Admin & Superadmin Only)
app.put('/api/products/:id', authMiddleware, rbacGuard(['ADMINISTRADOR', 'SUPERADMINISTRADOR']), auditMiddleware('UPDATE', 'Product'), (req, res) => {
  const { id } = req.params;
  const index = db.products.findIndex(p => p.id === id);
  if (index === -1) {
    return res.status(404).json({ error: 'Producto no encontrado' });
  }

  db.products[index] = {
    ...db.products[index],
    ...req.body,
    updatedAt: new Date().toISOString()
  };

  res.json({ message: 'Producto e inventario actualizado exitosamente', product: db.products[index] });
});

// 4. Linear Regression Projections Endpoint
app.get('/api/analytics/forecast/:productId', authMiddleware, (req, res) => {
  const { productId } = req.params;
  const product = db.products.find(p => p.id === productId);

  if (!product) {
    return res.status(404).json({ error: 'Producto no encontrado' });
  }

  // Linear Regression Calculation
  const X = [1, 5, 10, 15, 20, 25, 30];
  const Y = [2, 5, 8, 12, 16, 22, 28]; // Mock sales velocity

  const n = X.length;
  const sumX = X.reduce((a, b) => a + b, 0);
  const sumY = Y.reduce((a, b) => a + b, 0);
  const sumXY = X.reduce((acc, x, i) => acc + x * Y[i], 0);
  const sumX2 = X.reduce((acc, x) => acc + x * x, 0);

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const dailyRate = Math.max(0.1, parseFloat(slope.toFixed(2)));
  const daysRemaining = Math.round(product.physicalStock / dailyRate);

  res.json({
    productId,
    sku: product.sku,
    physicalStock: product.physicalStock,
    dailyConsumptionRate: dailyRate,
    estimatedDaysUntilDepletion: daysRemaining,
    purchaseAlertSuggested: daysRemaining < 15
  });
});

// 5. Export Sigo Flat File (CSV)
app.get('/api/reports/sigo-export', authMiddleware, (req, res) => {
  let csv = 'Fecha;TipoMovimiento;SKU;Producto;Cantidad;Factura;Cliente;CostoBase;PrecioVenta\n';
  db.movements.forEach(m => {
    csv += `"${m.date}";"${m.type}";"${m.sku}";"${m.productName}";${m.quantity};"${m.invoiceNumber || ''}";"${m.customerName || ''}";${m.baseCost};${m.salePrice}\n`;
  });

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="sigo_export_mascampo.csv"');
  res.send(csv);
});

// Server Listen
app.listen(PORT, () => {
  console.log(`🚀 Servidor SaaS Backend Más Campo ejecutándose en puerto ${PORT}`);
});
