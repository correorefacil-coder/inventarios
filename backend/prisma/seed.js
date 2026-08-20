const { PrismaClient } = require('@prisma/client');
const fs = require('fs');
const zipfile = require('zipfile'); // or parse JSON extracted seed

const prisma = new PrismaClient();

async function main() {
  console.log("🌱 Iniciando carga de datos iniciales (Seeding) desde el Excel de Más Campo...");

  // Load extracted processed seed
  const rawData = fs.readFileSync('/tmp/processed_seed.json', 'utf-8');
  const seedData = JSON.parse(rawData);

  // 1. Roles
  const adminRole = await prisma.role.upsert({
    where: { name: 'ADMINISTRADOR' },
    update: {},
    create: { name: 'ADMINISTRADOR', description: 'Acceso Total Financiero y Configuración' }
  });

  const logRole = await prisma.role.upsert({
    where: { name: 'LOGISTICA' },
    update: {},
    create: { name: 'LOGISTICA', description: 'Acceso Bodega y Escaneo de Códigos' }
  });

  const ventRole = await prisma.role.upsert({
    where: { name: 'VENTAS' },
    update: {},
    create: { name: 'VENTAS', description: 'Consulta Stock y Reservas de Clientes' }
  });

  // 2. Categories
  const catEquipo = await prisma.category.upsert({
    where: { name: 'EQUIPO_ELECTRONICO' },
    update: {},
    create: { name: 'EQUIPO_ELECTRONICO', description: 'Equipos Serializados con Hoja de Vida' }
  });

  const catConsumible = await prisma.category.upsert({
    where: { name: 'CONSUMIBLE' },
    update: {},
    create: { name: 'CONSUMIBLE', description: 'Consumibles por Cantidad / Lote' }
  });

  // 3. Populate Customers
  for (const c of seedData.customers) {
    await prisma.customer.upsert({
      where: { documentId: c.documentId },
      update: {},
      create: {
        documentId: c.documentId,
        companyOrName: c.name,
        phone: c.phone,
        email: c.email,
        address: c.address
      }
    });
  }

  // 4. Populate Products (First 100 for seed performance)
  for (const p of seedData.products.slice(0, 100)) {
    const categoryId = p.category === 'EQUIPO_ELECTRONICO' ? catEquipo.id : catConsumible.id;
    
    await prisma.product.upsert({
      where: { sku: p.sku },
      update: {},
      create: {
        sku: p.sku,
        name: p.name,
        description: p.description,
        categoryId: categoryId,
        requiresSerial: p.requiresSerial,
        unitOfMeasure: p.unitOfMeasure,
        minStockAlert: p.minStockAlert,
        baseCost: p.baseCost,
        salePrice: p.salePrice
      }
    });
  }

  console.log(`✅ ¡Seeding completado con éxito! Se cargaron ${seedData.customers.length} clientes y los productos iniciales del Excel.`);
}

main()
  .catch((e) => {
    console.error("❌ Error en seeding:", e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
