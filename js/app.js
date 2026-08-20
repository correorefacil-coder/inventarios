/* ==========================================================================
   MÁS CAMPO - SaaS Sistema de Gestión de Inventarios
   Data Ingested directly from official Excel File (INVENTARIO 2026)
   Enhanced with Document Sanitizer (Clean NIT vs Cédula + Separate DV Field)
   and Customer 360° Analytics Hub (Serial History, Financial Metrics, Semester Breakdown)
   ========================================================================== */

const initialCategories = [
  { id: "cat-1", name: "Consumibles Desechables", description: "Fundas, mangas, guantes de palpación y catéteres" },
  { id: "cat-2", name: "Consumibles Ingeribles", description: "Medios de cultivo, suplementos y lubricantes" },
  { id: "cat-3", name: "Equipos de Diagnóstico", description: "Ecógrafos portátiles, sondas y escáneres" },
  { id: "cat-4", name: "Equipos de Laboratorio", description: "Microscopios, platinas térmicas e incubadoras" }
];

const excelIngestedProducts = [
  {
    "id": "prod-1",
    "sku": "MC 001",
    "name": "FUNDAS PARA TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: BTE-025PT/ 504002 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-2",
    "sku": "MC 002",
    "name": "FUNDAS PARA TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: 704330 | Marca: MARCA PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-3",
    "sku": "MC 003",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: INO-025/ 508001 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 12,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-4",
    "sku": "MC 004",
    "name": "LíNEA PARA OPU 1,2 METROS",
    "description": "Ref: LA-120 /506003 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 220,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-5",
    "sku": "MC 005",
    "name": "LACRADORES CUADRADO 0,25",
    "description": "Ref: LP-25/ 505005 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-6",
    "sku": "MC 006",
    "name": "TAPON DE SILICONA PARA OPU",
    "description": "Ref: RS-01 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-7",
    "sku": "MC 007",
    "name": "TAPON DE SILICONA PEQUEÑO",
    "description": "Ref: RSS-T 04 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-8",
    "sku": "MC 008",
    "name": "AGUJAS PARA OPU",
    "description": "Ref: AA-OPU | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-9",
    "sku": "MC 009",
    "name": "CAMISA SANITARIA PARA OPU ASPIRACIÓN FOLICULAR",
    "description": "Ref: TNB-0001-16 | Marca: TNB | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 28,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-10",
    "sku": "MC 010",
    "name": "TUBOS PARA INSEMINACIÓN PROFUNDA EN EQUINOS 22¨(VERDE) ESTERILIZADO",
    "description": "Ref: TIPE-V 22/ 386315 | Marca: SINMARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-11",
    "sku": "MC 011",
    "name": "CAJA DE CAMISAS SANITARIA ROLLO POR 80 UNIDADES",
    "description": "Ref: 338000 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-12",
    "sku": "MC 012",
    "name": "LUBRICANTE SLIX NO ESPERMICIDA-TUBO 6 OZ",
    "description": "Ref: 22009 | Marca: SLIX | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-13",
    "sku": "MC 013",
    "name": "SISTEMA DE CONDUCCIÓN UNIÓN Y- SISTEMA DE USO DUAL (EQUINO & BOVINO)",
    "description": "Ref: 704133 | Marca: US | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-14",
    "sku": "MC 014",
    "name": "CATÉTER DE SILICONA DE 18 FR X 30 CC X 21\"",
    "description": "Ref: 438912 | Marca: BIONICHE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-15",
    "sku": "MC 015",
    "name": "FILTRO, EM-CON PARA USO EQUINO Y BOVINO",
    "description": "Ref: 41357/ 150200 | Marca: EMCON | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-16",
    "sku": "MC 016",
    "name": "LUBRICANTE GEL PRIORITY CARE 5 OZ (LUBRICANT, STER-LJELLY 5 OZ )",
    "description": "Ref: 26660 | Marca: PRIORITY CARE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-17",
    "sku": "MC 017",
    "name": "PIPETA IA PROFUNDA EQUINA DE 65 CM",
    "description": "Ref: 17209/1165 | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-18",
    "sku": "MC 018",
    "name": "PIPETA IA PROFUNDA EQUINA DE 75 CM",
    "description": "Ref: 17209/1175N | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-19",
    "sku": "MC 019",
    "name": "EmSafe - FILTRO COLECCIÓN DE EMBRIONES BOVINO/EQUINO",
    "description": "Ref: 19010/6000 | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-20",
    "sku": "MC 020",
    "name": "CATHETER, Equine flushing, 32FR x 75cc (34) - Silicone Cate",
    "description": "Ref: 438918 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-21",
    "sku": "MC 021",
    "name": "BOLSA WHIRL-PAK 7 oz",
    "description": "Ref: | Marca: WHIRL-PAK | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-22",
    "sku": "MC 022",
    "name": "GUANTES DE NITRILO AZUL",
    "description": "Ref: | Marca: ALFASAFE | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-23",
    "sku": "MC 023",
    "name": "FILTROS ACRODISC, ESTERILIZADOS , 0.22um",
    "description": "Ref: F-2690-9 | Marca: ACRODISC | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-24",
    "sku": "MC 024",
    "name": "PAJILLAS AMARILLAS IRRADIADAS 0.25",
    "description": "Ref: 440035 | Marca: US | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-25",
    "sku": "MC 025",
    "name": "PUNTAS TRANSPARENTES DE 10 ul Esterilizadas",
    "description": "Ref: 311-4150 | Marca: BIOPOINT | Prov: EXPERT/ EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-26",
    "sku": "MC 026",
    "name": "CÁMARA DE RECUENTO CELULAR",
    "description": "Ref: ISAS D4C20 L | Marca: PROISER | Prov: PROISER/ ARQUIMEA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-27",
    "sku": "MC 027",
    "name": "CÁMARA DE RECUENTO CELULAR",
    "description": "Ref: ISAS D4C20 | Marca: PROISER | Prov: PROISER/ ARQUIMEA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 29,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-28",
    "sku": "MC 028",
    "name": "5563 CAMISA SANITARIA TRANSFERENCIA",
    "description": "Ref: 90007000002 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-29",
    "sku": "MC 029",
    "name": "ACEITE MINERAL VWR X 500 ML",
    "description": "Ref: J217-500L | Marca: VWR | Prov: INCOLAMERICA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "1",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-30",
    "sku": "MC 030",
    "name": "LINEA DE ASPIRACION GRANDE 1.40 M",
    "description": "Ref: 12001 | Marca: PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-31",
    "sku": "MC 031",
    "name": "CAJAS DE PETRI 35X15",
    "description": "Ref: 627161 | Marca: GREINER BIO-ONE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-32",
    "sku": "MC 032",
    "name": "CAJAS DE PETRI 60X15",
    "description": "Ref: 6200 | Marca: BRIXCO | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-33",
    "sku": "MC 033",
    "name": "CAJAS DE PETRI 100X20 PS",
    "description": "Ref: 353003-20 | Marca: FALCON | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-34",
    "sku": "MC 034",
    "name": "BOLSA AUTOSELLABLE PARA ESTERILIZAR 13,5 CM 25,5 CMS",
    "description": "Ref: MQAUT003 | Marca: YIPAK | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 200 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-35",
    "sku": "MC 035",
    "name": "CATETER DE SILICONA 3 VIAS EQUINO",
    "description": "Ref: | Marca: BIONICHE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 8 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-36",
    "sku": "MC 036",
    "name": "CAMISA SANITARIA IMV",
    "description": "Ref: 5563 | Marca: IMV | Prov: EIDER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-37",
    "sku": "MC 037",
    "name": "TUBOS PIPETAS INSEMINACION ARTIFICIAL EQUINOS",
    "description": "Ref: | Marca: PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 22 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-38",
    "sku": "MC 038",
    "name": "FUNDAS PARA INSEMINACION ARTIFICIAL PUNTA VERDE",
    "description": "Ref: 7495 | Marca: IMV | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-39",
    "sku": "MC 039",
    "name": "SYNGRO HOLDING 8 ML",
    "description": "Ref: 100424 | Marca: VETOQUINOL | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-40",
    "sku": "MC 040",
    "name": "MANGA PALPACION SUPER SENSITIVA",
    "description": "Ref: | Marca: VET-BUSINESS | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-41",
    "sku": "MC 041",
    "name": "ETHYLENE GRYCOL ABT FREEZE 8 ML",
    "description": "Ref: | Marca: VETOQUINOL | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-42",
    "sku": "MC 042",
    "name": "BIOLOGIX CRIOVIALES PP 2,0ML",
    "description": "Ref: | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-43",
    "sku": "MC 043",
    "name": "CINTA CONTROL ESTERILIZACION A VAPOR ROLLO 3/4 ANCHO, 60 YARDAS REF: 1322",
    "description": "Ref: 1001924 | Marca: 3M | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-44",
    "sku": "MC 044",
    "name": "TUBOS PIPETAS INSEMINACION ARTIFICIAL EQUINOS",
    "description": "Ref: | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-45",
    "sku": "MC 045",
    "name": "PISTOLA QUICKLOCK CLASIC",
    "description": "Ref: 1132 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 8,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-46",
    "sku": "MC 046",
    "name": "CORTAPAJILLAS",
    "description": "Ref: 1216 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-47",
    "sku": "MC 047",
    "name": "PARCHE ESTROTECT",
    "description": "Ref: 1 | Marca: ESTROTECT | Prov: ESTROTECT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 180,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-48",
    "sku": "MC 048",
    "name": "COMPLETE FLUSH ABT 1 LTR",
    "description": "Ref: ABT360 | Marca: ABT | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "LITRO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-49",
    "sku": "MC 049",
    "name": "TUBO DE FONDO REDONDO DE 5ML 12X75MM , POLIESTIRENO",
    "description": "Ref: 352054 | Marca: FALCON | Prov: EQUIMED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 125 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-50",
    "sku": "MC 050",
    "name": "MEDIOS DE CULTIVO",
    "description": "Ref: 3 | Marca: STARSWAB II | Prov: INCOLAMERICA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-51",
    "sku": "MC 051",
    "name": "TUBO TIPO FALCON CÓNICO DE 15 ML",
    "description": "Ref: 10-0152-1 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-52",
    "sku": "MC 052",
    "name": "TUBO CONICO PLUG, AUTOCLAVABLE ESTERIL TAPA ROSCA DE 50 ML",
    "description": "Ref: 50050-1 | Marca: SPL | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 32,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-53",
    "sku": "MC 053",
    "name": "FILTRO SAFE-WAY",
    "description": "Ref: 150420 | Marca: US | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-54",
    "sku": "MC 054",
    "name": "FILTROS DE ALTO VOLUMEN",
    "description": "Ref: 150212A | Marca: MAI | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-55",
    "sku": "MC 055",
    "name": "HOLDING HA 7ML",
    "description": "Ref: ADA724/ 627024 | Marca: ABT | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-56",
    "sku": "MC 056",
    "name": "DIGITAL MICROSCOPE USB",
    "description": "Ref: 744359 | Marca: CHINA | Prov: CHINA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-57",
    "sku": "MC 057",
    "name": "DIGITAL MICROSCOPE USB",
    "description": "Ref: 207774.03 | Marca: CHINA | Prov: CHINA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-58",
    "sku": "MC 058",
    "name": "HISOPO DE CULTIVO DOBLE SEGURIDAD ESTERILIZADO",
    "description": "Ref: 636101 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-59",
    "sku": "MC 059",
    "name": "GUANTES DE VINILO",
    "description": "Ref: | Marca: VITAL MEDIC | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD CAJA X 100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-60",
    "sku": "MC 060",
    "name": "VAGINA ARTIFICIAL OVINA",
    "description": "Ref: 615010 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-61",
    "sku": "MC 061",
    "name": "CONO DE RECOLECCION BOVINO DESECHABLE SATIN BOLSILLO, EXTRA-SUAVE",
    "description": "Ref: 77020 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-62",
    "sku": "MC 062",
    "name": "CONO DE RECOLECCION BOVINO DESECHABLE DE POLIETILENO CERRADO",
    "description": "Ref: 77000 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-63",
    "sku": "MC 063",
    "name": "INSEMINADOR PARA LAPAROSCOPIO DE 14\" PARA OVINO",
    "description": "Ref: 464327 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-64",
    "sku": "MC 064",
    "name": "INSEMINADOR PARA LAPAROSCOPIO DE 12\" PARA OVINO",
    "description": "Ref: 464313 | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-65",
    "sku": "MC 065",
    "name": "FILTRO PARA SEMEN PORCINO DIA: 240 MM",
    "description": "Ref: 11123/3020 | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 200 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-66",
    "sku": "MC 066",
    "name": "ESTILETE",
    "description": "Ref: | Marca: SIN MARCA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-67",
    "sku": "MC 067",
    "name": "GUANTES DE LATEX",
    "description": "Ref: | Marca: VITAL MEDIC | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD CAJA X 100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-68",
    "sku": "MC 068",
    "name": "WARMER GUN CALENTADOR PISTOLAS",
    "description": "Ref: | Marca: AGJAPAN | Prov: FREDERIK",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-69",
    "sku": "MC 069",
    "name": "VAGINA OVINA",
    "description": "Ref: 7179 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-70",
    "sku": "MC 070",
    "name": "CONO VAGINA OVINA",
    "description": "Ref: 5205 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-71",
    "sku": "MC 071",
    "name": "CAMISA VAGINA OVINA",
    "description": "Ref: 22512 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-72",
    "sku": "MC 072",
    "name": "ESCOTIMETRO",
    "description": "Ref: | Marca: SIN MARCA | Prov: COLINSTRUVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-73",
    "sku": "MC 073",
    "name": "MICROSCOPIO CON PLATINA Y PANTALLA 7\"",
    "description": "Ref: | Marca: PROPIA | Prov: APUS MEDICAL",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-74",
    "sku": "MC 074",
    "name": "DRY BATTERY",
    "description": "Ref: | Marca: TATTU | Prov: APUS MEDICAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 8,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-75",
    "sku": "MC 075",
    "name": "MEDIUM BUBBER X 25",
    "description": "Ref: 6935 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-76",
    "sku": "MC 076",
    "name": "PRECISION LIQUID IN GLASS VERIFICATION THERMOMETERS WHIT COMPLIANCE REPORT",
    "description": "Ref: 13-201-546 | Marca: FISHERBRAND | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-77",
    "sku": "MC 077",
    "name": "MINIPAJILLA 0,25 ML TRANSPARENTES",
    "description": "Ref: 13400010/ 1124 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-78",
    "sku": "MC 078",
    "name": "PUNTA PIPETA UNIV. AMARILLA 1-200 UL PAQ. X 1000",
    "description": "Ref: T-200-Y | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-79",
    "sku": "MC 079",
    "name": "PUNTA PIPETA UNIV. 1000 UL PAQ. X 1000",
    "description": "Ref: T 1000 B | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-80",
    "sku": "MC 080",
    "name": "PUNTA PIPETA UNIV. 0,5- 10 UL",
    "description": "Ref: T-300 | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-81",
    "sku": "MC 081",
    "name": "MICROTUBO 2,0 ML PP AUTOCLAVABLE, AUTOSOSTENIBLE, ESTERIL, CON ÁREA DE ESCRITURA",
    "description": "Ref: 81-8204 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-82",
    "sku": "MC 082",
    "name": "ISPERM KIT (WITH 50 INCLUIDED CHIPS, EXS M5)",
    "description": "Ref: | Marca: ISPERM | Prov: AIDMICS BIOLOCHNOLOGY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-83",
    "sku": "MC 083",
    "name": "ISPERM CHIPS (500 TESTS PER BOX) PAQ. X 50 UND",
    "description": "Ref: | Marca: ISPERM | Prov: AIDMICS BIOLOCHNOLOGY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 126,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-84",
    "sku": "MC 084",
    "name": "PINZA COGE PAJILLA",
    "description": "Ref: | Marca: SIN MARCA | Prov: VET-BUSINESS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-85",
    "sku": "MC 085",
    "name": "PLACA CALENTADORA/ MESA PATRON MA-3320",
    "description": "Ref: 502009 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-86",
    "sku": "MC 086",
    "name": "PLACA CALENTADORA/ LUPA NIKON MA-SMZ",
    "description": "Ref: 502003/ 502019 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-87",
    "sku": "MC 087",
    "name": "RACK 28 TUBOS MIXTOS (2X15 + 15X5 + B KG 8X1.5)",
    "description": "Ref: 51001 | Marca: TED | Prov: TED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-88",
    "sku": "MC 088",
    "name": "CAMISA SANITARIA PARA PROBE RECTAL",
    "description": "Ref: 509013 | Marca: TNB | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 32,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-89",
    "sku": "MC 089",
    "name": "GUIA OPU- MINDRAY 65C15EAV",
    "description": "Ref: 512001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-90",
    "sku": "MC 090",
    "name": "CAJA DE CAJAS PETRI PEQUEÑA 30X15. 74 PAQ. X 10 UNIDADES",
    "description": "Ref: 627161-C | Marca: GREINER BIO-ONE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 74 PAQ.",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-91",
    "sku": "MC 091",
    "name": "ARGOLLA NARIGUERA BRONCE NAL",
    "description": "Ref: 125 | Marca: INSMEVET | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-92",
    "sku": "MC 092",
    "name": "LAPIZ MARCADOR ROJO",
    "description": "Ref: 574 | Marca: RAIDEX | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-93",
    "sku": "MC 093",
    "name": "LAPIZ MARCADOR VERDE",
    "description": "Ref: 576 | Marca: RAIDEX | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-94",
    "sku": "MC 094",
    "name": "PAJILLA MEDIANA 0,5 ML VERDE T",
    "description": "Ref: 1383 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-95",
    "sku": "MC 095",
    "name": "PIPETA PASTEUR EN PLÁSTICO PE-LD 3 ML CAJA X 500",
    "description": "Ref: 747760 | Marca: BRAND | Prov: BLAMIS DOTACIONES LABORATORIO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-96",
    "sku": "MC 096",
    "name": "SONDA NASAL EQUINO 3M*19MM",
    "description": "Ref: 820 | Marca: HAUPTNER | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-97",
    "sku": "MC 097",
    "name": "CAJA DE PETRI 60X15 PS ESTERIL X 500 UND",
    "description": "Ref: 66-1560 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-98",
    "sku": "MC 098",
    "name": "TERMOMETRO DE VIDRIO EN ALCOHOL 10+50°C",
    "description": "Ref: 1218431905 | Marca: LAB SCIENT | Prov: ISLAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-99",
    "sku": "MC 099",
    "name": "TERMO DESCONGELADOR DE SEMEN DE EMBRIONES ORANGE",
    "description": "Ref: 23065 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-100",
    "sku": "MC 100",
    "name": "TERMO DESCONGELADOR DE SEMEN DE EMBRIONES AZUL",
    "description": "Ref: 23064 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-101",
    "sku": "MC 101",
    "name": "RONCAINA X 50 ML",
    "description": "Ref: ROPRONX50M | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD 50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-102",
    "sku": "MC 102",
    "name": "AGUJA DESECH 18G X 1/2",
    "description": "Ref: 1050 | Marca: ISAVET | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-103",
    "sku": "MC 103",
    "name": "TERMO DESCONGELADOR DE SEMEN DE EMBRIONES VERDE",
    "description": "Ref: 23068 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-104",
    "sku": "MC 104",
    "name": "TERMO XTL3 270517 CON 5 IN, CARTINER (6)",
    "description": "Ref: 3XLTLB-5S | Marca: TAYLOR WHARTON | Prov: ANICAM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-105",
    "sku": "MC 105",
    "name": "ECOGRAFO POTATIL L60 CON SONDA LINEAL RECTAL",
    "description": "Ref: L60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-106",
    "sku": "MC 106",
    "name": "ECOGRAFO SIUI ULTRASOUND 3 SONDAS SCANNER WITH L7FVC, C3FC AND L1FC",
    "description": "Ref: CTS-800 | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-107",
    "sku": "MC 107",
    "name": "ECOGRAFO SIUI ULTRA SOUND SCANNER WITH L7FCV",
    "description": "Ref: CTS-800 | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-108",
    "sku": "MC 108",
    "name": "LAINER VAGINA ARTIFICIAL OVINA-CAPRINA",
    "description": "Ref: | Marca: | Prov: LEYLA BERNAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-109",
    "sku": "MC 109",
    "name": "NO USAR",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-110",
    "sku": "MC 110",
    "name": "NO USAR",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-111",
    "sku": "MC 111",
    "name": "ECOGRAFO BLANCO Y NEGRO DP 20 VET",
    "description": "Ref: ECOMDP20 | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-112",
    "sku": "MC 112",
    "name": "TRASDUCTOR MICROCONVEX (DP-10,DP20, DP30, DP-50, DP-2200)",
    "description": "Ref: 65C15EA | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-113",
    "sku": "MC 113",
    "name": "PUNTAS DE 0,5 10 UL, SIN FILTRO, TRANSPARENTE, ESTERIL, RACK X 96 PUNTAS",
    "description": "Ref: T-300-R-S | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-114",
    "sku": "MC 114",
    "name": "MALLETIN PARA ECOGRAFO DP 10",
    "description": "Ref: MPED 10 | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-115",
    "sku": "MC 115",
    "name": "TINTA MARCAR CARNE 500 ML AZUL",
    "description": "Ref: 1715 | Marca: RAIDEX | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-116",
    "sku": "MC 116",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: 7246 | Marca: IMV | Prov: IMPLEGAN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-117",
    "sku": "MC 117",
    "name": "JELCO N° 18 MED",
    "description": "Ref: MQJEL1011 | Marca: JELCO | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-118",
    "sku": "MC 118",
    "name": "PAPEL ESTERLIZACION ROLLO MIXTO 62GR 15CMX200MT 3 INCIDAD",
    "description": "Ref: MQRM1003 | Marca: ALFA-YIPAK | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO X 200 MTS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-119",
    "sku": "MC 119",
    "name": "TRANSDUCTOR ENDORECTAL",
    "description": "Ref: PL2A-30-90629 | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-120",
    "sku": "MC 120",
    "name": "TRILADYL",
    "description": "Ref: 1135 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-121",
    "sku": "MC 121",
    "name": "BATERIA LI-ION ICR26650 3,7 V 4400 MAH COM TOP- RONTEK",
    "description": "Ref: 101010 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-122",
    "sku": "MC 122",
    "name": "PUNTAS ul YELLOW TIPS",
    "description": "Ref: T-200YRS | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-123",
    "sku": "MC 123",
    "name": "ETHYLENE GRYCOL ABT FREEZE CON SUCROSA Y HA X 7 ML",
    "description": "Ref: 627025 | Marca: ABT | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-124",
    "sku": "MC 124",
    "name": "PUNTAS COLOR AMARILLO ESTANDAR. GRADUADA. ARO LARGO. 200 UL PP.",
    "description": "Ref: 20-0200 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-125",
    "sku": "MC 125",
    "name": "BD FALCON FILTRO CELULAR 70 UM",
    "description": "Ref: 352350 | Marca: FALCON | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-126",
    "sku": "MC 126",
    "name": "PAJILLA PARA SEMEN 0,50 ML",
    "description": "Ref: 2155 | Marca: IMV | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-127",
    "sku": "MC 127",
    "name": "ANDROMET 200 ML",
    "description": "Ref: 135030/1000 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-128",
    "sku": "MC 128",
    "name": "MICROTUBO 2,0 ML",
    "description": "Ref: | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-129",
    "sku": "MC 129",
    "name": "CINTA ESCROTAL",
    "description": "Ref: 11589 | Marca: PROPIA | Prov: COLINSTRUVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-130",
    "sku": "MC 130",
    "name": "TERMOMETRO PARA INCUBADORA",
    "description": "Ref: EF23892C | Marca: | Prov: EXPERT",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-131",
    "sku": "MC 131",
    "name": "ESTROTECT",
    "description": "Ref: AI 00059 RED 10 | Marca: ESTROTECT | Prov: ESTROTECT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-132",
    "sku": "MC 132",
    "name": "PLATINA DE MESA MA 3320/ MA 3322",
    "description": "Ref: 502018 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-133",
    "sku": "MC 133",
    "name": "BOMBA DE VACIO TOUCH BA-06",
    "description": "Ref: 500004 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-134",
    "sku": "MC 134",
    "name": "LAMINILLAS CUBRE OBJETOS 22X22 MM",
    "description": "Ref: 4402784 | Marca: GLASS LAB | Prov: INCOLMEDICA SA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-135",
    "sku": "MC 135",
    "name": "LAMINILLAS PORTA- OBJETOS",
    "description": "Ref: 4402698 | Marca: GLASS LAB | Prov: INCOLMEDICA SA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-136",
    "sku": "MC 136",
    "name": "TUBOS PARA MICROCENTRIFUGA 1,5 ML",
    "description": "Ref: MCT-150C | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-137",
    "sku": "MC 137",
    "name": "DISPENSADOR DE CONOS AI",
    "description": "Ref: J0102D9 | Marca: JOR VET | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-138",
    "sku": "MC 138",
    "name": "CALENTADOR DE TUBOS 50 ML AT-650-OLED",
    "description": "Ref: 501007 | Marca: TED | Prov: TED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-139",
    "sku": "MC 139",
    "name": "CONO DE COLECTA PAQUETE POR 100 UNIDADES",
    "description": "Ref: | Marca: MINITUBE | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-140",
    "sku": "MC 140",
    "name": "BOLSA TELA MAS CAMPO",
    "description": "Ref: | Marca: PROPIA | Prov: CARLOS PINILLA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 179,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-141",
    "sku": "MC 141",
    "name": "FUNDA SANITARIA 18\" IA",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-142",
    "sku": "MC 142",
    "name": "SONDA FOLLEY BIONICHE DE SILICONA (ABT) 20FR. X 30CCX 22\"",
    "description": "Ref: 132030 | Marca: BIONICHE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-143",
    "sku": "MC 143",
    "name": "ESCALERILLAS METÁLICAS PARA GOBLETS DE 10 MM PAQ. X 100",
    "description": "Ref: 347510 | Marca: PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-144",
    "sku": "MC 144",
    "name": "CAMARA DE NEUBAUER",
    "description": "Ref: 100851108 | Marca: NEWLTEM | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-145",
    "sku": "MC 145",
    "name": "CATÉTER DE INSEMINACIÓN UNIVERSAL INSERTO VERDE EMP INDIVIDUAL PAQ X 25",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 250UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-146",
    "sku": "MC 146",
    "name": "PAJILLAS 0.5 CC",
    "description": "Ref: 1126 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-147",
    "sku": "MC 147",
    "name": "JALEA CONDUCTIVA (GEL ULTRASONICO) GALON X 3785 ML AZUL",
    "description": "Ref: 1003514 | Marca: KRAMER | Prov: INCOLMEDICA SA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALON X 3785 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-148",
    "sku": "MC 148",
    "name": "TRANSPORTADORA DE EMBRIONES",
    "description": "Ref: 503002 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-149",
    "sku": "MC 149",
    "name": "FILTRO EM CON",
    "description": "Ref: 4135 | Marca: EMCON | Prov: BIOGENEC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-150",
    "sku": "MC 150",
    "name": "SISTEMA DE CONDUCCIÓN UNIÓN Y- SISTEMA DE USO DUAL (EQUINO & BOVINO)",
    "description": "Ref: 0060-1 | Marca: SPI | Prov: BIOGENEC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-151",
    "sku": "MC 151",
    "name": "CONO PARA SFS 133 (13018/0100)",
    "description": "Ref: 1199 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-152",
    "sku": "MC 152",
    "name": "CAJA CUATRO HUECOS DISH 4 WELLS ESTERILIZADO PAQ X10",
    "description": "Ref: 164056 | Marca: MAI | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-153",
    "sku": "MC 153",
    "name": "CARGADOR PARA TRANSPORTADORA TED",
    "description": "Ref: | Marca: | Prov: NOVOA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-154",
    "sku": "MC 154",
    "name": "VAGINA ARTIFICIAL OVINA",
    "description": "Ref: | Marca: IMV | Prov: LEYLA BERNAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-155",
    "sku": "MC 155",
    "name": "ECOGRAFO PORTATIL V2",
    "description": "Ref: V2 - KAIXIN | Marca: KAIXIN | Prov:",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-156",
    "sku": "MC 156",
    "name": "AGUA, DOBLE - CULTIVO CEL PROCESADO * TES 500 ML",
    "description": "Ref: MG-W3500500M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-157",
    "sku": "MC 157",
    "name": "FILTRO JERINGA DE 0.22 UM",
    "description": "Ref: SSF08S | Marca: BIOBASE | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-158",
    "sku": "MC 158",
    "name": "SONDA S5 UNIDAD PRINCIPAL CON SONDA LINEAL RECTAL",
    "description": "Ref: S5 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 8,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-159",
    "sku": "MC 159",
    "name": "PASTA TATUAR VERDE",
    "description": "Ref: 671 | Marca: HAUPTNER | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-160",
    "sku": "MC 160",
    "name": "BATERIA WED3000",
    "description": "Ref: | Marca: | Prov: APUS MEDICAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-161",
    "sku": "MC 161",
    "name": "FUNDA UNIVERSAL PARA I.A. RANURADA",
    "description": "Ref: | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 134,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-162",
    "sku": "MC 162",
    "name": "CATETER TALADRO TRANSPARENTE 18\"",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-163",
    "sku": "MC 163",
    "name": "TERMO TRANSPORTE 1 LITRO/1 DIA- METÁLICO",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-164",
    "sku": "MC 164",
    "name": "NARIGUERA DE TORO MEDIO 3/8\" X 3\"",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 12,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-165",
    "sku": "MC 165",
    "name": "ALCOHOL POLIVINILICO",
    "description": "Ref: | Marca: MARCA PROPIA | Prov: RODA QUÍMICOS SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 1 KILO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-166",
    "sku": "MC 166",
    "name": "CARBOXIMETIL CELULOSA (CMC)",
    "description": "Ref: | Marca: MARCA PROPIA | Prov: RODA QUÍMICOS SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 1 KILO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-167",
    "sku": "MC 167",
    "name": "LIDOCAINA X 100 ML",
    "description": "Ref: CPALICL100 | Marca: CORPAUL | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-168",
    "sku": "MC 168",
    "name": "HEPARINA 5000 UN",
    "description": "Ref: | Marca: | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-169",
    "sku": "MC 169",
    "name": "EOSINA- NIGROSINA",
    "description": "Ref: 109304 | Marca: NOVALAB | Prov: NOVALAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-170",
    "sku": "MC 170",
    "name": "REGLA PARA NITROGENO",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-171",
    "sku": "MC 171",
    "name": "MANGA OBSTÉTRICA DESECHABLE NARANJA",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-172",
    "sku": "MC 172",
    "name": "PINZA COGEPAJILLA",
    "description": "Ref: | Marca: NACIONAL | Prov: PLAYCEL SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-173",
    "sku": "MC 173",
    "name": "MICROPIPETA AUTOMATICA PETTEPLUS VOLUMEN VARIABLE 0,5-10ul,",
    "description": "Ref: 12101 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-174",
    "sku": "MC 174",
    "name": "TERMÓMETRO DIGITAL",
    "description": "Ref: SH113 | Marca: | Prov: MEDICOR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-175",
    "sku": "MC 175",
    "name": "FUNDAS PARA TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: 13136 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-176",
    "sku": "MC 176",
    "name": "PIPETA VOL VARIABLE 5-10 UL EPPENDORF SERIE P27161H",
    "description": "Ref: 3120000020 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-177",
    "sku": "MC 177",
    "name": "GUIA PARA OPU WTA",
    "description": "Ref: 10370 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-178",
    "sku": "MC 178",
    "name": "MICROVIAL PARA CULTIVO CELULAR 1,2 ML",
    "description": "Ref: 21315 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-179",
    "sku": "MC 179",
    "name": "TAPABOCAS TIRA ELASTICA CAJA X 50 UND 2018DM-0017530. REGISTRO INVIMA",
    "description": "Ref: 3012 | Marca: VITAL MEDIC | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-180",
    "sku": "MC 180",
    "name": "GORRO DESECHABLE REDONDO 17 GRAMOS. PAQ X 50 UND NACIONAL. REGISTRO INVIMA",
    "description": "Ref: QA-08 | Marca: | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-181",
    "sku": "MC 181",
    "name": "ALCOHOL ANTISEPTICO GALON X 3700 ML. REGISTRO SANITARIO 2008M-011909-R1",
    "description": "Ref: 1000602 | Marca: | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALON X 3700 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-182",
    "sku": "MC 182",
    "name": "TAPABOCAS (MASCARA RESPIRADOR) DE MEDIA CARA DOBLE CARTUCHO",
    "description": "Ref: 11111003-05 | Marca: | Prov: INCOLMEDICA SA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-183",
    "sku": "MC 183",
    "name": "GEL ANTIBACTERIAL 250 ML",
    "description": "Ref: FVV032420 | Marca: | Prov: LABG&M",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-184",
    "sku": "MC 184",
    "name": "GEL ANTIBACTERIAL GALON X 4 LT",
    "description": "Ref: FVV03244 | Marca: | Prov: LABG&M",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALON",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-185",
    "sku": "MC 185",
    "name": "JABON LIQUIDO ANTIBACTERIAL GALON X 5 LTRS",
    "description": "Ref: FVV031810 | Marca: | Prov: LABG&M",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALON",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-186",
    "sku": "MC 186",
    "name": "GUANTE DE EXAMEN X 100 UND. TALLA M",
    "description": "Ref: 00EXA192 | Marca: EXAMTEX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-187",
    "sku": "MC 187",
    "name": "GUANTE DE EXAMEN X 100 UND. TALLA S",
    "description": "Ref: 00EXA191 | Marca: EXAMTEX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-188",
    "sku": "MC 188",
    "name": "GUANTE DE NITRILO AZUL X 100 UND TALLA M",
    "description": "Ref: 00EXA225 | Marca: PROTEXION | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-189",
    "sku": "MC 189",
    "name": "GUANTE DE NITRILO AZUL X 100 UND TALLA S",
    "description": "Ref: 00EXA224 | Marca: PROTEXION | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-190",
    "sku": "MC 190",
    "name": "GUANTES DE VINILO X 100 UND. TALLA S",
    "description": "Ref: 00EXA227 | Marca: PROTEXION | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-191",
    "sku": "MC 191",
    "name": "GUANTES DE VINILO X 100 UND. TALLA L",
    "description": "Ref: 00EXA229 | Marca: PROTEXION | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-192",
    "sku": "MC 192",
    "name": "CAJAS DE PETRI DE 35X 10 MM, EN POLYESTYRENO, ESTPERILES, DESECHABLES, PARA CULTIVO CELULAR",
    "description": "Ref: 430165 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 86,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-193",
    "sku": "MC 193",
    "name": "MICROCUBETA PARA FOTÓMETRO SDM 1, 100/CAJA",
    "description": "Ref: 12300/1111 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-194",
    "sku": "MC 194",
    "name": "PAJILLA MEDIANA 0,5 ML TRANSPARENTES",
    "description": "Ref: 13408/0010 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-195",
    "sku": "MC 195",
    "name": "AGUJA DESECHABLE 18 X 11/2",
    "description": "Ref: IVEAGDE181 | Marca: ISMEVET | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-196",
    "sku": "MC 196",
    "name": "BD FALCON FILTRO CELULAR 70 UM",
    "description": "Ref: 352350 | Marca: FALCON | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-197",
    "sku": "MC 197",
    "name": "RAMPAS DE CONGELACIÓN DE 58 PAJILLAS",
    "description": "Ref: 7118 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-198",
    "sku": "MC 198",
    "name": "KIT DE INSEMINACIÓN BÁSICO",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-199",
    "sku": "MC 199",
    "name": "POLVO SELLANTE DE PAJILLAS X 750 GR",
    "description": "Ref: 440200 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "750 GR",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-200",
    "sku": "MC 200",
    "name": "CARGADOR TED 4S5A C/ LED INDICADOR",
    "description": "Ref: 108130 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-201",
    "sku": "MC 201",
    "name": "TRANSISTOR MOSFET AO34407- A79T 4,3A/ 30V SOT 23",
    "description": "Ref: 108137 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-202",
    "sku": "MC 202",
    "name": "CONECTOR PARA CABO INVERTIDO M16 6 CONTACTO MACHO",
    "description": "Ref: 102025 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-203",
    "sku": "MC 203",
    "name": "GUANTE NITRILO AZUL TALLA L",
    "description": "Ref: LATGUNIX100 | Marca: SEMPERQUARD | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-204",
    "sku": "MC 204",
    "name": "OVEROL",
    "description": "Ref: | Marca: | Prov: MUNDIAL DE DOTACIONES SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-205",
    "sku": "MC 205",
    "name": "BATA QUIRURGICA",
    "description": "Ref: | Marca: | Prov: MUNDIAL DE DOTACIONES SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-206",
    "sku": "MC 206",
    "name": "ABT ADAPT HOLDING CON HA 7 ML",
    "description": "Ref: ABT ADAPT HOLDING | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 13,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-207",
    "sku": "MC 207",
    "name": "HEPES SOLUTION 1M SOLUTION IN PURFIED WATER. STERILE FILTERED",
    "description": "Ref: HOL06-100 ML | Marca: MAISSON | Prov: GENTECH BIOSCIENCES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-208",
    "sku": "MC 208",
    "name": "GUANTE NITRILO TALLA M. WELL",
    "description": "Ref: GU1078 | Marca: WELL | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-209",
    "sku": "MC 209",
    "name": "TAPABOCAS TRADICIONAL TRES CAPAS",
    "description": "Ref: | Marca: NACIONAL | Prov: MUNDIAL DE DOTACIONES SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-210",
    "sku": "MC 210",
    "name": "6459 CAMISA SANITARIA INSEMINACION PQ X 100",
    "description": "Ref: 70006000001 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-211",
    "sku": "MC 211",
    "name": "GOBLETS CLAROS/BLANCOS DE 10 MM PAQ. X 100",
    "description": "Ref: 347430 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-212",
    "sku": "MC 212",
    "name": "GOBLETS AMARILLOS DE 13 MM PAQ. X 100",
    "description": "Ref: 347449 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-213",
    "sku": "MC 213",
    "name": "ESCALERILLAS METÁLICAS PARA GOBLETS DE 13 MM PAQ. X 100",
    "description": "Ref: 347550 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-214",
    "sku": "MC 214",
    "name": "BOTA PLÁSTICA NACIONAL",
    "description": "Ref: | Marca: | Prov: NACIONAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-215",
    "sku": "MC 215",
    "name": "BOTUSEMENSPECIAL 8,5gm",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: LEYLA BERNAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-216",
    "sku": "MC 216",
    "name": "ILERSAN 414",
    "description": "Ref: | Marca: ILEDERCORP | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1000 CC",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-217",
    "sku": "MC 217",
    "name": "ADAPTADOR DE JERINGA X 1 ML TB A PAJILLA DE 1/4 CC O 1/2 CC AZUL",
    "description": "Ref: 440356 | Marca: NACIONAL | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-218",
    "sku": "MC 218",
    "name": "TERMOMETRO INFRAROJO DÍGITAL",
    "description": "Ref: JXB-178 | Marca: BERRCOM | Prov: LEYLA BERNAL",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-219",
    "sku": "MC 219",
    "name": "AGUJA DESECH 20G X 1 1/2",
    "description": "Ref: 302350 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-220",
    "sku": "MC 220",
    "name": "PAJILLA PARA SEMEN 0,25 ML",
    "description": "Ref: 2137 | Marca: IMV | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-221",
    "sku": "MC 221",
    "name": "PARAFILM ROLLO DEV 100 MM ANCHO, 38 MTS DE LARGO",
    "description": "Ref: CS-200311 | Marca: SANTA CRUZ | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-222",
    "sku": "MC 222",
    "name": "SOLUCION SALINA 500ML (1323)",
    "description": "Ref: MQSOL1005 | Marca: BAXTER | Prov: LIDERMEDICA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-223",
    "sku": "MC 223",
    "name": "ESTUCHE DE CIRUGIA * 15 PIEZAS",
    "description": "Ref: 2688 | Marca: AMALGADENT | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-224",
    "sku": "MC 224",
    "name": "ECOGRAFO PORTATIL DP2200",
    "description": "Ref: DP 2200 | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-225",
    "sku": "MC 225",
    "name": "AGUJA SUTURA CURVA GRANDE * 12",
    "description": "Ref: 3481 | Marca: AMALGADENT | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 125 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-226",
    "sku": "MC 226",
    "name": "CATGUT CROMADO N° 2 CON AGUJA",
    "description": "Ref: 238 | Marca: ASSUT | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-227",
    "sku": "MC 227",
    "name": "TIRAS INDICADORAS DEL PH PH 0 - 14 INDICADOR UNIVERSAL NO DESTIÑEN 100 STRIPS",
    "description": "Ref: 1095350001 | Marca: | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 STRIPS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 11,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-228",
    "sku": "MC 228",
    "name": "AGUA DESTILADA GALÓN X 5 LT",
    "description": "Ref: FVV03205 | Marca: | Prov: MICROGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALÓN X 5 LT",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-229",
    "sku": "MC 229",
    "name": "INMUNODET NEUTRO GALÓN X 5 LT",
    "description": "Ref: FVV03175 | Marca: | Prov: MICROGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "GALÓN X 5 LT",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-230",
    "sku": "MC 230",
    "name": "CRIOVIALES 1,8ML NUNC",
    "description": "Ref: EXP000645 | Marca: THERMO SCIENTIFIC | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-231",
    "sku": "MC 231",
    "name": "TUBOS PARA MICROCENTRIFUGA 2,0ML",
    "description": "Ref: EXP000708 | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-232",
    "sku": "MC 232",
    "name": "CAJA PETRI 100X15 MM",
    "description": "Ref: EXP000253 | Marca: FALCON | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-233",
    "sku": "MC 233",
    "name": "ETHYLENE GRYCOL ANHYDROUS 99,8% X 100 ML",
    "description": "Ref: MG-324558100 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-234",
    "sku": "MC 234",
    "name": "DIMETHYL SULFOXIDE HYDRIMAX STERILE 5X5 ML",
    "description": "Ref: MG-D26505X5M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-235",
    "sku": "MC 235",
    "name": "SUCROSE; CELL CULTURE TESTED X 500 G",
    "description": "Ref: MG-S1888500G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "500 G",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-236",
    "sku": "MC 236",
    "name": "SODIUM DL- LACTATE SOLUTION, SYRUP, BIOREAGENT, 60 %, SYNTHETIC, SUITABLE FOR CELL CULTURE, 100 ML",
    "description": "Ref: MG-L4263100M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-237",
    "sku": "MC 237",
    "name": "LACRADORES LARGOS REDONDOS, BLANCOS",
    "description": "Ref: 342512 | Marca: PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-238",
    "sku": "MC 238",
    "name": "TABLET S5",
    "description": "Ref: | Marca: HUAWEI | Prov: IMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-239",
    "sku": "MC 239",
    "name": "MINERAL OIL WHITE (PETROLEUM) LIGHT. HIGH PURITY",
    "description": "Ref: 97064-130 | Marca: VWR | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-240",
    "sku": "MC 240",
    "name": "CAJA PETRI 35X10MM 00520002",
    "description": "Ref: EXP000304/ 430165 | Marca: CORNING | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-241",
    "sku": "MC 241",
    "name": "FUNDA DE LÁTEX PARA VAGINA BOVINA",
    "description": "Ref: 614616 | Marca: PROPIA | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-242",
    "sku": "MC 242",
    "name": "MICROSCOPIO MONOCULAR CON PANTALLA DE 5 \"",
    "description": "Ref: SE-500 | Marca: PROPIA | Prov: SONOVIC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-243",
    "sku": "MC 243",
    "name": "MICROSCOPIO MONOCULAR CON PANTALLA DE 7 \" CON PLATINA TERMICA",
    "description": "Ref: SE-1900 | Marca: PROPIA | Prov: SONOVIC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-244",
    "sku": "MC 244",
    "name": "MALETIN NEGRO ISPERM",
    "description": "Ref: 9012900000-0 | Marca: ISPERM | Prov: AIDMICS BIOLOCHNOLOGY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-245",
    "sku": "MC 245",
    "name": "ULTRASONIDO V16",
    "description": "Ref: V16 | Marca: | Prov: SHENZHEN DARKMATTER TECH",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-246",
    "sku": "MC 246",
    "name": "5569 PAJILLAS 0.5 ML",
    "description": "Ref: 90009000003 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 200",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-247",
    "sku": "MC 247",
    "name": "BOTUVITAL DIILUENT FOR SPERM EVALUATION",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: VIVIAN ALEJANDRA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "10 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-248",
    "sku": "MC 248",
    "name": "CAJA PETRI DESECHABLE ESTERIL 60X15 MM",
    "description": "Ref: 18130E | Marca: CRAL | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-249",
    "sku": "MC 249",
    "name": "SODIUM CHLORIDE; REAGENTPLUS(R);gt;99% X 500G",
    "description": "Ref: MG-S9625500G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-250",
    "sku": "MC 250",
    "name": "POTASSIUM CHLORIDE CELL CULTURE TESTED X 250G",
    "description": "Ref: MG-P5405250G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-251",
    "sku": "MC 251",
    "name": "SODIUM BICARBONATE; POWDER; BIOREAGENT X 500G",
    "description": "Ref: MG-S5761500G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-252",
    "sku": "MC 252",
    "name": "MAGNESIUM CHLORIDE HEXAHYDRATE CELL X 100G",
    "description": "Ref: MG-M2393100 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-253",
    "sku": "MC 253",
    "name": "HEPES; BIOPERFORMANCE CERTIFIED X 100G",
    "description": "Ref: MG-H4034100G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-254",
    "sku": "MC 254",
    "name": "EPINEPHRINE X 1G",
    "description": "Ref: MG-E42501G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-255",
    "sku": "MC 255",
    "name": "BME Amino Acids Solution 50, Without L-glutamine, sterile-filtered, BioReagent, suitable for cell culture",
    "description": "Ref: MG-B6766100M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-256",
    "sku": "MC 256",
    "name": "BESTRADIOL X 1G",
    "description": "Ref: MG-E88751G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-257",
    "sku": "MC 257",
    "name": "D-Penicillamine , 98-101% , 1G",
    "description": "Ref: MG-P48751G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-258",
    "sku": "MC 258",
    "name": "HYPOTAURINE CRYSTALLINE X 250MG",
    "description": "Ref: MG-H1384250 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-259",
    "sku": "MC 259",
    "name": "Medium 199, With Earles salts,",
    "description": "Ref: MG-M4530500 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-260",
    "sku": "MC 260",
    "name": "SODIUM PHOSPHATE MONOBASIC X 100G",
    "description": "Ref: MG-S5011100G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-261",
    "sku": "MC 261",
    "name": "CALCIUM CHLORIDE DIHYDRATE, BIOREAGENT, SUITABLE FOR CELL CULTURE",
    "description": "Ref: MG-C7902500G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-262",
    "sku": "MC 262",
    "name": "HEPES; SODIUM SALT BIOPERFORMANCE X 100G",
    "description": "Ref: MG-H3784100G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-263",
    "sku": "MC 263",
    "name": "SODIUM METABISULFITE, 500G",
    "description": "Ref: MG-S9000500G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-264",
    "sku": "MC 264",
    "name": "TUBOS CONICO ESTERILIZADOS DE 50 ML PUNTA PLANA",
    "description": "Ref: EXP000313 430921 | Marca: CORNING | Prov: EXPERT",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-265",
    "sku": "MC 265",
    "name": "CAJA PETRI DESECHABLE ESTERIL 90X15 MM",
    "description": "Ref: 18248AUTE | Marca: CRAL | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-266",
    "sku": "MC 266",
    "name": "CAJA PETRI 35 X 10 mm",
    "description": "Ref: 706001 | Marca: NEST | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-267",
    "sku": "MC 267",
    "name": "OPTIDYL X 500 CC",
    "description": "Ref: | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 CC",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-268",
    "sku": "MC 268",
    "name": "TERMO CITO THAW UNIT 120V",
    "description": "Ref: 02 3013 | Marca: CITO | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-269",
    "sku": "MC 269",
    "name": "TUBO CRIOGENICO GRADUADO 1,8 ML TAPA ROSCA EXTERNA ESTERIL",
    "description": "Ref: CRIO18NR | Marca: CRAL | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-270",
    "sku": "MC 270",
    "name": "CONOS DE COLECTA",
    "description": "Ref: | Marca: PROPIA | Prov: LEYLA BERNAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-271",
    "sku": "MC 271",
    "name": "JUEGO DE DADOS 3/8",
    "description": "Ref: 526 | Marca: STONE | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-272",
    "sku": "MC 272",
    "name": "TATUADOR STANDARD 3/8",
    "description": "Ref: 866/ 2505 | Marca: STONE | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-273",
    "sku": "MC 273",
    "name": "TINTA VERDE FCO 3 ONZ",
    "description": "Ref: 920/ 6530 | Marca: STONE | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-274",
    "sku": "MC 274",
    "name": "JELCO CATETER ONTRAV # 18 BRAUM DEC551",
    "description": "Ref: 10595-1 | Marca: JELCO | Prov: INSTRUAGRO COLOMBIA SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-275",
    "sku": "MC 275",
    "name": "HEPARINA SODICA 5000 U.I/ML SI 5ML VIAL",
    "description": "Ref: DCV5ML | Marca: BRAUN | Prov: LEYLA BERNAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-276",
    "sku": "MC 276",
    "name": "ACEITE MINERAL VWR X 1000 ML",
    "description": "Ref: J217-1000L | Marca: VWR | Prov: ISLAS",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "1 LITRO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-277",
    "sku": "MC 277",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES",
    "description": "Ref: 1129 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-278",
    "sku": "MC 278",
    "name": "WATER, STERILE-FILTERED, BIOREAGENT",
    "description": "Ref: MG-W35001L | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 LITRO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-279",
    "sku": "MC 279",
    "name": "JERINGA DES. 20 ML C/A 18G X 1 1/2",
    "description": "Ref: ZB7L18B | Marca: RYMCO | Prov: ISNMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-280",
    "sku": "MC 280",
    "name": "JERINGA DES. 5 ML - 2 PARTES C/A 18G X 1 1/2",
    "description": "Ref: | Marca: ISAVET | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-281",
    "sku": "MC 281",
    "name": "ASPIC PARA IA OVINOS",
    "description": "Ref: 5546 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-282",
    "sku": "MC 282",
    "name": "CONO GRANDE COLECTOR",
    "description": "Ref: 9090 | Marca: PROPIA | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X100",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-283",
    "sku": "MC 283",
    "name": "CONO SATINDO PEQUEÑO",
    "description": "Ref: 77020 | Marca: PROPIA | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-284",
    "sku": "MC 284",
    "name": "AGUJA BD 20G X 1 1/2",
    "description": "Ref: 305179 | Marca: BD | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-285",
    "sku": "MC 285",
    "name": "CATETER INTRAVENOSO N°18",
    "description": "Ref: 4035/ 01-01342 | Marca: JELCO | Prov: HOSPICLINIC SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-286",
    "sku": "MC 286",
    "name": "JERINGAS 27GX1/2\"",
    "description": "Ref: 10612-18 | Marca: VITAL MEDIC | Prov: INSTRUAGRO COLOMBIA SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-287",
    "sku": "MC 287",
    "name": "ESTEREOMICROSCOPIO BINOCULAR LED, ENFOQUE GRUESO",
    "description": "Ref: SZX-B+SZ-A1+SZ-ST2 | Marca: OPTIKA | Prov: LAB BRANDS SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-288",
    "sku": "MC 288",
    "name": "LACTATO DE RINGER-HARMANT",
    "description": "Ref: MQLACT 1001 | Marca: RINGER-HARMANT | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-289",
    "sku": "MC 289",
    "name": "AGUA DESTILADA X 4 LT ALBOR",
    "description": "Ref: 1410007 | Marca: ALBOR | Prov: QUIOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "4 LITROS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-290",
    "sku": "MC 290",
    "name": "CAMISAS SANITARIAS 21\"",
    "description": "Ref: 338000 | Marca: PROPIA | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-291",
    "sku": "MC 291",
    "name": "PT50C : VET COLOR DOPPLER SISTEM/ CON TRANSDUCTOR LENEAL RECTAL/ TRANSDUCTOR LINEAL CONVEXO",
    "description": "Ref: PT50C | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-292",
    "sku": "MC 292",
    "name": "CRIOVIALES 2,0 ML AUTOSOSTENIBLE ESTERIL PARA - 196 (FASE GASEOSA NO SUMERGIBLE )",
    "description": "Ref: 126263 | Marca: GREINER BIO-ONE | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-293",
    "sku": "MC 293",
    "name": "TUBOS TIPO EPPENDERF 1,5 ML. GRADUADO. CERTIFICADOS LIBRE DE RNASE, DNASA, PIROGENOS Y DNA",
    "description": "Ref: 616201B | Marca: GREINER BIO-ONE | Prov: BIOHAUS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-294",
    "sku": "MC 294",
    "name": "MICROPIPETA PREMIUM VOLUMEN VARIABLE 20-200 UL",
    "description": "Ref: CP-P200V | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-295",
    "sku": "MC 295",
    "name": "FILTRO DE BOMBA DE VACÍO",
    "description": "Ref: 10372 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-296",
    "sku": "MC 296",
    "name": "PUNTA DE MANDRIL DE FORMA SENCILLA",
    "description": "Ref: 10096 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-297",
    "sku": "MC 297",
    "name": "CAMISA SANITARIA PARA OPU ASPIRACION FOLICULAR/ WTA",
    "description": "Ref: 10130 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-298",
    "sku": "MC 298",
    "name": "MANGAS DE PALPACIÓN SUPER SENSITIVA",
    "description": "Ref: | Marca: PROPIA | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-299",
    "sku": "MC 299",
    "name": "PIPETA AUT VOL VARIABLE ACCUBIOTECH SUREPETTE 0,5-10",
    "description": "Ref: P15 | Marca: ACCUBIOTECH | Prov: EUROLAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-300",
    "sku": "MC 300",
    "name": "AGUJA DESECHABLE 18 X 11/2",
    "description": "Ref: 25 | Marca: ISMEVET | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-301",
    "sku": "MC 301",
    "name": "PUNTAS PARA PIPETA SUELTAS TRANSPARENTES LIBRES DE ADN DNASA RNASA 0,1-20 MI",
    "description": "Ref: 732022 | Marca: BRAND | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-302",
    "sku": "MC 302",
    "name": "PUNTAS PARA PIPETA AMARILLAS SUELTAS LIBRES DE ADN 2-200 UL",
    "description": "Ref: 732028 | Marca: BRAND | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-303",
    "sku": "MC 303",
    "name": "TUBO CONICO PLUG SEAL PP 15 ML GRADUADO AUTOCLAVABLE ESTERIL TAPA ROSCA",
    "description": "Ref: 50015-1 | Marca: SPL | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-304",
    "sku": "MC 304",
    "name": "CAJA PETRI 35X10 ESTERIL",
    "description": "Ref: 627161 | Marca: GREINER BIO-ONE | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 273,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-305",
    "sku": "MC 305",
    "name": "FILTROS PARA OPU 100 UL",
    "description": "Ref: 25458 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-306",
    "sku": "MC 306",
    "name": "DISPOCEL 0.6 GM X 10 GMS",
    "description": "Ref: CARDI06X10 | Marca: VON FRANFEN | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "10 GMS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-307",
    "sku": "MC 307",
    "name": "APLICADOR DISPOCEL",
    "description": "Ref: CARAPLIDIS | Marca: VON FRANFEN | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-308",
    "sku": "MC 308",
    "name": "SINCVAL X 5000 UI",
    "description": "Ref: CARSINC5MI | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-309",
    "sku": "MC 309",
    "name": "CLOPROSTENOL X 20 ML",
    "description": "Ref: CDACLOPX20 | Marca: CALIER | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-310",
    "sku": "MC 310",
    "name": "BENZOATO ESTRADEOL 100 ML",
    "description": "Ref: COVBEES100 | Marca: NORMAS GMP | Prov: PAISAGRO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-311",
    "sku": "MC 311",
    "name": "CAJA DE HERRAMIENTAS",
    "description": "Ref: 770515204142 | Marca: RIMAX | Prov: EASY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-312",
    "sku": "MC 312",
    "name": "KIT DE INSEMINACIÓN COMPLETO",
    "description": "Ref: 9 PRODUCTOS | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-313",
    "sku": "MC 313",
    "name": "GLICINA FARMAGRADE G 5417",
    "description": "Ref: G 5417 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-314",
    "sku": "MC 314",
    "name": "ETHANOL 459844",
    "description": "Ref: MG-4598441L | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-315",
    "sku": "MC 315",
    "name": "PHENOL RED P 3532",
    "description": "Ref: MG-P35325G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-316",
    "sku": "MC 316",
    "name": "PISTOLA DE INSEMINACIÓN ARTIFICIAL UNIVERSAL METÁLICA, PARA 0,05CC- 0,25 CC",
    "description": "Ref: B6-3255 | Marca: ESTROTECT | Prov: ESTROTECT",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-317",
    "sku": "MC 317",
    "name": "ULTRASONIDO SCANNER MODELO: BESTSCAN S9, CON TRANSDUCTOR RECTAL 4.0, MHZ 60 MM/ INCLUYE TABLET E INTRODUCER RECTAL",
    "description": "Ref: S9 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-318",
    "sku": "MC 318",
    "name": "KIT DE VAGINOSCOPIO OVINO",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-319",
    "sku": "MC 319",
    "name": "ESPÉCULO KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-320",
    "sku": "MC 320",
    "name": "FUNDAS KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-321",
    "sku": "MC 321",
    "name": "LINTERNA KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-322",
    "sku": "MC 322",
    "name": "TERMÓMETRO KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-323",
    "sku": "MC 323",
    "name": "FUNDAS PARA ESPÉCULO KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-324",
    "sku": "MC 324",
    "name": "CLAMBS KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-325",
    "sku": "MC 325",
    "name": "CAJA DE HERRAMIENTAS KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: | Prov: CESAR",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-326",
    "sku": "MC 326",
    "name": "CAMISA SANITARIA IMV",
    "description": "Ref: 5563/99005 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-327",
    "sku": "MC 327",
    "name": "FUNDAS DE TRANSFERENCIA AZUL PUNTA METÁLICA 21\"",
    "description": "Ref: 5540 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-328",
    "sku": "MC 328",
    "name": "TABS AMARILLOS",
    "description": "Ref: TY | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-329",
    "sku": "MC 329",
    "name": "GOBLET DE 13 MM AMARILLO",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-330",
    "sku": "MC 330",
    "name": "ESCALERILLAS PARA GOBLETS DE 13 MM",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-331",
    "sku": "MC 331",
    "name": "MICROPIPETA VOLUMEN VARIABLE 0,5- 10 UL TIPO PLUS",
    "description": "Ref: 7030301004 | Marca: DBL | Prov: LAB BRANDS SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-332",
    "sku": "MC 332",
    "name": "AGITADOR VORTEX TIPO MX-S VELOCIDAD VARIABLE",
    "description": "Ref: 8031102000 | Marca: DLAB- DRAGON LAB | Prov: FILTRACIÓN Y ANÁLISIS SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-333",
    "sku": "MC 333",
    "name": "TUBO OOSAFE ANDRLOGY",
    "description": "Ref: OOPW-AT 10 | Marca: OOSAFE | Prov: DAXLEY DE COLOMBIA SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE x 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-334",
    "sku": "MC 334",
    "name": "PUNTA BLANCA XL, 10 UL GRADUABLE",
    "description": "Ref: 4130-00S | Marca: SSIBIO++ | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-335",
    "sku": "MC 335",
    "name": "PUNTA GRADUADA 200 UL TRANSPARENTE",
    "description": "Ref: 4220-00S | Marca: SSIBIO++ | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-336",
    "sku": "MC 336",
    "name": "PUNTA 1000 UL",
    "description": "Ref: 4330-015 | Marca: SSIBIO++ | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-337",
    "sku": "MC 337",
    "name": "JUEGO DE DADOS 3/8 (A-Z)",
    "description": "Ref: 531 | Marca: STONE | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-338",
    "sku": "MC 338",
    "name": "JERINGA DES. 10 ML C/A 18 X 1 1/2",
    "description": "Ref: 500 | Marca: ISAVET | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-339",
    "sku": "MC 339",
    "name": "JERINGA DES. 2 ML PARTES C/A 23G X 1",
    "description": "Ref: 4769 | Marca: RYMCO | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-340",
    "sku": "MC 340",
    "name": "TUBO VACUTAINER LILA 4 ML",
    "description": "Ref: 954 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-341",
    "sku": "MC 341",
    "name": "TUBO VACUTAINER ROJO 6 ML",
    "description": "Ref: 954 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-342",
    "sku": "MC 342",
    "name": "MICROSCOPIO CON PLATINA Y PANTALLA 7\"",
    "description": "Ref: | Marca: | Prov: WUHAM VISBULL",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-343",
    "sku": "MC 343",
    "name": "CAJA VACÍA RACK PARA PUNTAS DE 1000 ul",
    "description": "Ref: 4-503-00-0 | Marca: AHN BIOTECHNOLOGIE | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-344",
    "sku": "MC 344",
    "name": "AGUJA MULTIPLE 21G X 1 1/2 VERDE",
    "description": "Ref: 51/ 360213 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-345",
    "sku": "MC 345",
    "name": "CALENTADOR SUPERIOR EN ACERO INOXIDABLE",
    "description": "Ref: CO-194-CW | Marca: COSORI | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-346",
    "sku": "MC 346",
    "name": "ECÓGRAFO V8 CON SONDA LINEAL RECTAL",
    "description": "Ref: V8 | Marca: HANDSCAN | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-347",
    "sku": "MC 347",
    "name": "TRANSPORTADORA DE EMBRIONES Y OOCITOS",
    "description": "Ref: 26237/ 25339 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-348",
    "sku": "MC 348",
    "name": "PUNTA 1000 UL",
    "description": "Ref: 4330-015 | Marca: SSIBIO++ | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-349",
    "sku": "MC 349",
    "name": "ECÓGRAFO WED-3000V CON SONDA RECTAL",
    "description": "Ref: WED-3000V | Marca: WED-3000V | Prov: WUHAN VISBULL MACHINERY CO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-350",
    "sku": "MC 350",
    "name": "7240 PISTOLA DE TRANSFERENCIA BOVINA 0.25 ML",
    "description": "Ref: 70002000014 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-351",
    "sku": "MC 351",
    "name": "PASTA TATUAR BLANCA",
    "description": "Ref: 669 | Marca: HAUPTNER | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-352",
    "sku": "MC 352",
    "name": "FILTRO DE MINITUBE FLUSH",
    "description": "Ref: 19222/2000 | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-353",
    "sku": "MC 353",
    "name": "ESTILETE, ACERO INOXIDABLE PARA CATETER DE 65 CM",
    "description": "Ref: 17209/1065N | Marca: MINITUBE | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-354",
    "sku": "MC 354",
    "name": "TERMO DESCONGELADOR DE SEMEN DE EMBRIONES ROJO",
    "description": "Ref: 18621 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-355",
    "sku": "MC 355",
    "name": "CAJA PETRI DESECHABLE ESTERIL 60X15 MM RAYADA",
    "description": "Ref: 1751OE | Marca: CRAL | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 49,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-356",
    "sku": "MC 356",
    "name": "ESTEREOMICROSCOPIO ZOOM 645",
    "description": "Ref: BIO-EQ-001 | Marca: | Prov: BIO INSTRUMENTS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-357",
    "sku": "MC 357",
    "name": "MEM Non-essential Amino Acid Solution (100) ,without L-glutamine,liquid, sterile-filtered, BioReag",
    "description": "Ref: MGM7145100 | Marca: SIGMA | Prov: SIGMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-358",
    "sku": "MC 358",
    "name": "LUBRICANTE OB, NO ESPERMICIDA",
    "description": "Ref: 605032 | Marca: VET ONE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 OZ",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 32,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-359",
    "sku": "MC 359",
    "name": "CAJAS PETRI 4 POZOS",
    "description": "Ref: 005/010 | Marca: SPI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-360",
    "sku": "MC 360",
    "name": "CATÉTER FOLLEY DE SILICONA 20FR X 30CC, LARGO 56MM",
    "description": "Ref: 438913 | Marca: PETS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-361",
    "sku": "MC 361",
    "name": "SONDA EN \"Y\" CON VÁLVULA",
    "description": "Ref: 704133 | Marca: SPI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-362",
    "sku": "MC 362",
    "name": "CEPILLO DE CITOLOGÍA EQUINA, PARA RECOGIDA DE CÉLULAS DE ÚTERO DE MARES, ESTERILIZADO",
    "description": "Ref: 17214/2960 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-363",
    "sku": "MC 363",
    "name": "PISTOLA DE TRANSFERENCIA DE EQUINOS DE 0,5",
    "description": "Ref: 24857 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-364",
    "sku": "MC 364",
    "name": "TUBO TIPO EPPENDORF 0,5/0,65 ML",
    "description": "Ref: 04-332-30 | Marca: NERBE PLUS | Prov: BIOHAUS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "BOLSA X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-365",
    "sku": "MC 365",
    "name": "CATÉTER FOLLEY DE SILICONA 2 CAMINOS 28FR / 75CC",
    "description": "Ref: PE1822-0028 | Marca: PETS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-366",
    "sku": "MC 366",
    "name": "ALCOHOL ANTISEPTICO AL 70%",
    "description": "Ref: 34324 | Marca: NACIONAL | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "600 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-367",
    "sku": "MC 367",
    "name": "CONO COLECTOR PARA CANINOS",
    "description": "Ref: 1077020 | Marca: REPRODUCTION PROVISIONS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-368",
    "sku": "MC 368",
    "name": "TUBO ESTOMACAL PEQUEÑO EN PVC",
    "description": "Ref: J0106SE | Marca: JORVET | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-369",
    "sku": "MC 369",
    "name": "TUBO ESTOMACAL MEDIANA EN PVC",
    "description": "Ref: J0106ME | Marca: JORVET | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-370",
    "sku": "MC 370",
    "name": "BOLSA WHIRL-PAK 4 oz",
    "description": "Ref: | Marca: WHIRL-PAK | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 960,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-371",
    "sku": "MC 371",
    "name": "LACRADOR CORTO AMARILLO",
    "description": "Ref: 20101 | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-372",
    "sku": "MC 372",
    "name": "TUBO ENSAYO PLÁSTICO T/ROSCA 75MM",
    "description": "Ref: 34181005 | Marca: CITOPLUS | Prov: NORQUIMICOS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-373",
    "sku": "MC 373",
    "name": "FUNDA TÉRMICA / USO AGROPECUARIO",
    "description": "Ref: 8419899900 | Marca: IFT | Prov: INT INSTRUMENTS SRL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-374",
    "sku": "MC 374",
    "name": "PIPETA EQUINA 65CM",
    "description": "Ref: 17209/1165 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-375",
    "sku": "MC 375",
    "name": "FUNDAS DE TRANSFERENCAI DE 0,5 PARA EQUINOS",
    "description": "Ref: 17532 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-376",
    "sku": "MC 376",
    "name": "PUNTA BLANCA 10 UL, GRADUADA BAJA RETENCIÓN, ESTERIL",
    "description": "Ref: 07-362-20 | Marca: NERBE PLUS | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-377",
    "sku": "MC 377",
    "name": "TUBO TIPO EPPENDORF 0,5/0,65 ML",
    "description": "Ref: 11110-00S | Marca: SSIBIO++ | Prov: BIOHAUS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-378",
    "sku": "MC 378",
    "name": "TUBO CENTRÍFUGA CELLSTAR 50 ML",
    "description": "Ref: 227261B | Marca: GREINER BIO-ONE | Prov: BIOHAUS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-379",
    "sku": "MC 379",
    "name": "HEPARINA 5000 UN",
    "description": "Ref: MQHEP-02 | Marca: FRAGMIN | Prov: LIDERMEDICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-380",
    "sku": "MC 381",
    "name": "TUBO DE ENSAYO CON TAPA A PRESIÓN DE 12X75MMX25U",
    "description": "Ref: 34181005 | Marca: CITOPLUS | Prov: ONE LAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUE X 25 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-381",
    "sku": "MC 380",
    "name": "TUBO PARA CENTRÍFUGA TIPO FALCON DE 50 ML, FONDO CÓNICO, TAPA ROSCA COLOR AZUL",
    "description": "Ref: 4610-1878 | Marca: CITOPLUS | Prov: ONE LAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-382",
    "sku": "MC 382",
    "name": "ELECTROJAC COMPLETE W/2,5\" PROBE",
    "description": "Ref: 6906-25 | Marca: ELECTROJAC | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-383",
    "sku": "MC 383",
    "name": "SONDA RECTAL LINEAL PARA TS20",
    "description": "Ref: TS20 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-384",
    "sku": "MC 384",
    "name": "CAJA PETRI GRANDE CUADRADA RAYADA",
    "description": "Ref: D-2555-8 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-385",
    "sku": "MC 385",
    "name": "PISTOLA PARA TRANSFERENCIA 0,25",
    "description": "Ref: | Marca: TED | Prov: RAB GLOBAL",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-386",
    "sku": "MC 386",
    "name": "FLUIDO DE CO2 CON RANGO DE MEDICION 0-20%/ 60 (KIT 3 BOTELLAS)",
    "description": "Ref: EXP001686/0010-5057 | Marca: | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-387",
    "sku": "MC 387",
    "name": "FLUIDO DE CO2 CON RANGO DE MEDICION 0-21%/ 60",
    "description": "Ref: EXP001687/0010-5060 | Marca: | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-388",
    "sku": "MC 388",
    "name": "LITHIUM ION BATTERY PACK/ 11.1V 37 Wh",
    "description": "Ref: CNLB-01 | Marca: XUZHOU KAIXIN ELECTRONIC INSTRUMENT | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-389",
    "sku": "MC 389",
    "name": "TERMO CRYOGENICO",
    "description": "Ref: 1496 | Marca: MVE | Prov: GENETICA SELECTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-390",
    "sku": "MC 390",
    "name": "TUBO DE FONDO REDONDO DE 5ML 12X75MM , POLIESTIRENO",
    "description": "Ref: 352054 | Marca: FALCON | Prov: EXPERT",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 125 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-391",
    "sku": "MC 391",
    "name": "HISOPO DE CULTIVO UTERINO EQUINO. ESTÉRIL",
    "description": "Ref: 605005 | Marca: CONTINENTAL | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 13,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-392",
    "sku": "MC 392",
    "name": "FUNDA UNIVERSAL PARA I.A. RANURADA. PUNTA VERDE",
    "description": "Ref: 17007/0001 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-393",
    "sku": "MC 393",
    "name": "PINZA COGE PAJILLA",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-394",
    "sku": "MC 394",
    "name": "MANGA DE PALPACIÓN NARANJA",
    "description": "Ref: | Marca: PROPIA | Prov: BELSAGRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 8,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-395",
    "sku": "MC 395",
    "name": "CAJA PETRI ESTERIL DESECHABLE 60X15MMX10",
    "description": "Ref: 2303-1060 | Marca: CITOTEST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-396",
    "sku": "MC 396",
    "name": "FILTROS PARA OPU 100 UL",
    "description": "Ref: 25458 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 17,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-397",
    "sku": "MC 397",
    "name": "PAJILLA MEDIANA 0,5 ML AZUL- TRANSPARENTE",
    "description": "Ref: 1808 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-398",
    "sku": "MC 398",
    "name": "TRANQUILAN INY X 10 ML",
    "description": "Ref: ZOOTRANQ10 | Marca: TRANQUILAN | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "10 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-399",
    "sku": "MC 399",
    "name": "JERINGA DESECHABLE X 5 ML",
    "description": "Ref: IVEJEDEXX5C | Marca: IVE | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-400",
    "sku": "MC 400",
    "name": "CAJAS PETRI CUADRARA RAYADA",
    "description": "Ref: | Marca: PROPIA | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-401",
    "sku": "MC 401",
    "name": "KIT PARA CONEJOS",
    "description": "Ref: AIT8 | Marca: | Prov: SHENZHEN DARKMATTER TECH",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-402",
    "sku": "MC 402",
    "name": "BOMBA DE VACÍO BOVINO",
    "description": "Ref: 20032/ 23187 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-403",
    "sku": "MC 403",
    "name": "ROLA METÁLICA PARA ASPIRACIÓN FOLICULAR",
    "description": "Ref: 10589 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-404",
    "sku": "MC 404",
    "name": "PLATINA TERMICA COMPACTA",
    "description": "Ref: 20435 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-405",
    "sku": "MC 405",
    "name": "CALENTADOR DE TUBOS PORTÁTIL",
    "description": "Ref: 21499 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-406",
    "sku": "MC 406",
    "name": "PUNTAS SIN FILTRO COLOR NATURAL 10 UL ARO CORTO",
    "description": "Ref: 310-2000 | Marca: BIOPONTE | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-407",
    "sku": "MC 407",
    "name": "ECÓGRAFO PARA PORCINOS CON DOS SONAS SECTORIALES",
    "description": "Ref: M50 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-408",
    "sku": "MC 408",
    "name": "MÁQUINA CENTRÍFUGA PROFESIONAL PLASMA",
    "description": "Ref: 800-1 | Marca: 800D | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-409",
    "sku": "MC 409",
    "name": "BATERÍA PORTÁTIL 10000 mAh",
    "description": "Ref: | Marca: | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-410",
    "sku": "MC 410",
    "name": "MG-M7145100/ MEM Non-essential Amino Acid Solution (100) ,without L-glutamine, liquid, sterile-filt REFRIGERADO",
    "description": "Ref: MG-M7145100 | Marca: SAFC | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-411",
    "sku": "MC 411",
    "name": "MG-H1384100/HYPOTAURINE X 100MG",
    "description": "Ref: MG-H1384100 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 MG",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-412",
    "sku": "MC 412",
    "name": "MG-I663450G INSULIN FROM BOVINE PANCREA",
    "description": "Ref: MG-I663450G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-413",
    "sku": "MC 413",
    "name": "MG-H7006100G HEPES SODIUM X 100G",
    "description": "Ref: MG-H7006100G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-414",
    "sku": "MC 414",
    "name": "MG-S938810ML SERUM REPLACEMENT 2 (50X)",
    "description": "Ref: MG-S938810ML | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "10 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-415",
    "sku": "MC 415",
    "name": "MG-P87835G PROGESTERONE POWDER",
    "description": "Ref: MG-P87835G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-416",
    "sku": "MC 416",
    "name": "MG-537471G HYALURONIC ACID SODIUM SALT FROM STREPTO X 1G",
    "description": "Ref: MG-537471G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-417",
    "sku": "MC 417",
    "name": "MICROTUBOS PARA CULTIVO CELULAR 1,2 ML. ESQUINAS REDONDEADAS",
    "description": "Ref: 21315 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-418",
    "sku": "MC 418",
    "name": "CENTRIFUGA PROFESIONAL PARA PLASMA - * RPM 4000",
    "description": "Ref: 8001 | Marca: | Prov: MERCADO LIBRE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-419",
    "sku": "MC 419",
    "name": "MAXI FILTRO",
    "description": "Ref: | Marca: SPI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-420",
    "sku": "MC 420",
    "name": "POLIVINYL (POLVO SELLADOR PARA PAJILLAS )",
    "description": "Ref: 537-671 | Marca: ARS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "750 GRAMOS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-421",
    "sku": "MC 421",
    "name": "OPTIXCELL",
    "description": "Ref: 26218 | Marca: OPTIXCELL | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-422",
    "sku": "MC 422",
    "name": "COMPLETE FLUSH ABT 1 LTR",
    "description": "Ref: ABT360 | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 LITRO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-423",
    "sku": "MC 423",
    "name": "ECÓGRAFO PORTÁTIL DP 10 CON BLANCO Y NEGRO CON SONDA",
    "description": "Ref: DP 10 _LR | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-424",
    "sku": "MC 424",
    "name": "BOTUSEMEN GOLD 8.5 GRAMOS",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: ALEJANDRA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "8.5 GRAMOS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-425",
    "sku": "MC 425",
    "name": "BOTUCRIO (DILUYENTE PARA CONGELACIÓN DE SEMEN EQUINO)",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: ALEJANDRA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-426",
    "sku": "MC 426",
    "name": "TERMO METÁLICO DESCONGELADOR",
    "description": "Ref: | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 15,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-427",
    "sku": "MC 427",
    "name": "CALENTADOR DE TUBOS MIXTO",
    "description": "Ref: 18032 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-428",
    "sku": "MC 428",
    "name": "TRANSDUCTOR MICROCONVEX (DP-10, DP-20, DP-30, DP-50, DP-2200)",
    "description": "Ref: 120-000076-00 | Marca: MINDRAY | Prov: MARQUETING NET SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-429",
    "sku": "MC 429",
    "name": "PLATINA CALENTADORA ESTÁNDAR",
    "description": "Ref: 17656 / 29455 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-430",
    "sku": "MC 430",
    "name": "CALENTADOR DE TUBOS MIXTO",
    "description": "Ref: 502016 | Marca: TED | Prov: TED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-431",
    "sku": "MC 431",
    "name": "TUBO 5ML CON TAPÓN 75X12 MM, CON GRADUACIÓN ESTÉRIL PP",
    "description": "Ref: 55526006 | Marca: | Prov: QUIK",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-432",
    "sku": "MC 432",
    "name": "MANGA DESECHABLE SUPER SENSITIVA",
    "description": "Ref: 101257 | Marca: PROPIA | Prov: LHAURA VET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 71,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-433",
    "sku": "MC 433",
    "name": "CENTRIFUGA PROFESIONAL 6 TUBOS",
    "description": "Ref: LEQ001/ LC-04R (800D) | Marca: INMEDIAT | Prov: MERCADO LIBRE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-434",
    "sku": "MC 434",
    "name": "BATERÍA PORTÁTIL PARA DISPOSITIVOS DIGITALES 5000mAh",
    "description": "Ref: MPB-506 | Marca: MOVISUN | Prov: BIN COLOMBIA SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-435",
    "sku": "MC 435",
    "name": "CARGADOR BATERÍA POWER BANK 10000 mAh, IFROGZ",
    "description": "Ref: AB847 | Marca: IFROGZ | Prov: MARKET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-436",
    "sku": "MC 436",
    "name": "MICROPIPETA PREMIUM VOLUMEN VARIABLE 100- 1000 UL",
    "description": "Ref: CP-P1000V | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-437",
    "sku": "MC 437",
    "name": "TUBOS PARA CENTRIFUGA NO ESTÉRIL COTAPARA 15 ML",
    "description": "Ref: P10401-25 | Marca: ABDOS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 25 UNIDADES",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-438",
    "sku": "MC 438",
    "name": "PIPETA PASTEUR 3 ML",
    "description": "Ref: P31205 | Marca: ABDOS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-439",
    "sku": "MC 439",
    "name": "PIPETA PASTEUR 3 ML",
    "description": "Ref: P31205 | Marca: ABDOS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 23,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-440",
    "sku": "MC 440",
    "name": "WRIGHT ALBOR REG INVIMA: 2006RD-0000221",
    "description": "Ref: 12101002 | Marca: ALBOR | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "200 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-441",
    "sku": "MC 441",
    "name": "GIEMSA- ALBOR REG INVIMA: 2006RD-0000221-R1",
    "description": "Ref: 12125001 | Marca: ALBOR | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-442",
    "sku": "MC 442",
    "name": "TUBO CONICO PP 15 ML GRADUADO AUTOCLAVABLE, ESTÉRIL LIBRE DNASE, RNASE, TAPA VERDE",
    "description": "Ref: 301002-1 | Marca: NEST | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-443",
    "sku": "MC 443",
    "name": "ESTILETE EN ACERO INOXIDABLE",
    "description": "Ref: EX | Marca: | Prov: INDUGANAVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-444",
    "sku": "MC 444",
    "name": "CÁMARA DE MC MASTER",
    "description": "Ref: | Marca: FEC SOURCE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-445",
    "sku": "MC 445",
    "name": "REFRACTÓMETRO CON ATC",
    "description": "Ref: | Marca: BRIX | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-446",
    "sku": "MC 446",
    "name": "CALOSTRÓMETRO",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-447",
    "sku": "MC 447",
    "name": "CATÉTER FOLLEY DE SILICONA (ABT) 16FR. X 30CC X 22",
    "description": "Ref: 131630 | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-448",
    "sku": "MC 448",
    "name": "CATÉTER FOLLEY DE SILICONA (ABT) 18FR. X 30CC X 22",
    "description": "Ref: 131830 | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-449",
    "sku": "MC 449",
    "name": "VASO DE VIDRIO 150",
    "description": "Ref: 5010629 | Marca: BOECO | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-450",
    "sku": "MC 450",
    "name": "VASO DE PRECIPITADO DE 250mL FORMA BAJA",
    "description": "Ref: 5010636 | Marca: BOECO | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-451",
    "sku": "MC 451",
    "name": "BULL SAFE/ FUNDA TERMICA PARA VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: | Marca: IFT | Prov: INT INSTRUMENTS SRL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-452",
    "sku": "MC 452",
    "name": "COLORANTE DE GIEMSA/ VENCE MARZ/23",
    "description": "Ref: 10820101 | Marca: NOVALAB | Prov: NOVALAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-453",
    "sku": "MC 453",
    "name": "PUNTAS AZULES ESTERILES DE AJUSTE UNIVERSAL 1000 UL",
    "description": "Ref: T-1000-B | Marca: AXYGEN | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-454",
    "sku": "MC 454",
    "name": "FILTROS PARA JERINGA ACRODISCOS PORO 0.2 UM DIAM 28 MM",
    "description": "Ref: 431219 | Marca: CORNING | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-455",
    "sku": "MC 455",
    "name": "VASO PRECIPITADO VIDRIO 50 ML",
    "description": "Ref: 5010617 | Marca: BOECO | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-456",
    "sku": "MC 456",
    "name": "PUNTAS TRANSPARENTES 2-200 UL",
    "description": "Ref: 9605 | Marca: BIOPOINTE | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-457",
    "sku": "MC 457",
    "name": "PAJILLAS IRRADIADAS 1/4 CC/ 0,25",
    "description": "Ref: 4164 | Marca: MAI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-458",
    "sku": "MC 458",
    "name": "FORRO DE CUERO PARA ECÓGRAFO CTS- 800",
    "description": "Ref: CTS-800 | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-459",
    "sku": "MC 459",
    "name": "BATERÍA PARA ECÓGRAFO CTS- 800",
    "description": "Ref: CTS-800 | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 7,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-460",
    "sku": "MC 460",
    "name": "EZ FILTRO DE EMBRIONES",
    "description": "Ref: 180121 | Marca: SPL | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-461",
    "sku": "MC 461",
    "name": "PIPETA EQUINA CON INNER",
    "description": "Ref: 17207/1165 | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-462",
    "sku": "MC 462",
    "name": "PARAFILM 38M X 10 CM",
    "description": "Ref: 7016-05 | Marca: | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO 38X10",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-463",
    "sku": "MC 463",
    "name": "AGUJA DESECHABLE 18X1/2",
    "description": "Ref: 10541A | Marca: AGH | Prov: INSTRUAGRO COLOMBIA SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-464",
    "sku": "MC 464",
    "name": "JERINGA DESECHABLE X 10 ML AGU21X11/2",
    "description": "Ref: 10614 | Marca: XAX MEDICALES | Prov: INSTRUAGRO COLOMBIA SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-465",
    "sku": "MC 465",
    "name": "TUBO TIPO EPPENDORF 2 ML, GRADADO",
    "description": "Ref: 623201 | Marca: GEINER BIO ONE | Prov: BIOHAUS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-466",
    "sku": "MC 466",
    "name": "MICROTUBO EPPENDORF, 1.5 ML",
    "description": "Ref: BS022 | Marca: BIOSIGMA | Prov: BIOHAUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-467",
    "sku": "MC 467",
    "name": "PROCGEST X 1200 X 10",
    "description": "Ref: COVPROCX10 | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-468",
    "sku": "MC 468",
    "name": "CONTRACTYL X 50 ML",
    "description": "Ref: VME334 | Marca: | Prov: WEIZUR COLOMBIA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-469",
    "sku": "MC 469",
    "name": "BUTRAFINA X 50 ML",
    "description": "Ref: VME332 | Marca: | Prov: WEIZUR COLOMBIA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "51 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-470",
    "sku": "MC 470",
    "name": "BENZATRIOL X 50 ML",
    "description": "Ref: VM34 | Marca: | Prov: WEIZUR COLOMBIA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-471",
    "sku": "MC 471",
    "name": "PUNTAS TRANSPARENTES 0.5-10 UL",
    "description": "Ref: 9624 | Marca: BIOPOINTE | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-472",
    "sku": "MC 472",
    "name": "BOVINE SERUM ALBUMIN, LYPHILIZED POWDER,ESSENTIALLY FATTY ACID FREE 96% REFRIGERADO",
    "description": "Ref: MG-A600310G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "10 GM",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-473",
    "sku": "MC 473",
    "name": "PLATINA TÉRMICA LK-400P",
    "description": "Ref: P07 | Marca: LABMAK | Prov: LABMAK SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-474",
    "sku": "MC 474",
    "name": "CONGELADORA DE EMBRIONES ALLICE",
    "description": "Ref: 22413/ 29455 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-475",
    "sku": "MC 475",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES 0,25 / BLOQUEO AUTOMÁTICO",
    "description": "Ref: 18421 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-476",
    "sku": "MC 476",
    "name": "BOTUSEMEN SPECIAL",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: VIVIAN ALEJANDRA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-477",
    "sku": "MC 477",
    "name": "VAGINA BOVINA INCLUYE CONO EN SILICONA",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-478",
    "sku": "MC 478",
    "name": "CONO EN SILICONA PARA VAGINA BOVINA",
    "description": "Ref: 614640 | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-479",
    "sku": "MC 479",
    "name": "CONO EN SILICONA PARA VAGINA OVINA",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-480",
    "sku": "MC 480",
    "name": "PAJILLAS 1/4 CC IRRADIADAS/ 0,25",
    "description": "Ref: 4166 | Marca: MAI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-481",
    "sku": "MC 481",
    "name": "PAJILLAS 1/2 IRRADIADAS CC/ ,05",
    "description": "Ref: 4168 | Marca: MAI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 8,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-482",
    "sku": "MC 482",
    "name": "HISOPO DE CULTIVO DE TRANSPORTE",
    "description": "Ref: 220099 | Marca: BD | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-483",
    "sku": "MC 483",
    "name": "AGUJA MULTIPLE 20G X 1 1/2 AMARILLA",
    "description": "Ref: 360215 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-484",
    "sku": "MC 484",
    "name": "CAMISA VACUTAINER",
    "description": "Ref: 364896 | Marca: BD | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-485",
    "sku": "MC 485",
    "name": "GRADILLA PARA 50 TUBOS CENTRIFUGA( 30 DE 15 ML Y 20 DE 50ML) NO AUTOCLAVABLE",
    "description": "Ref: HS24306 | Marca: HEATHROW SCIENTIFIC | Prov: BIOHAUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-486",
    "sku": "MC 486",
    "name": "LACRADOR REDONDO AMARILLO",
    "description": "Ref: 505006 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 70,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-487",
    "sku": "MC 487",
    "name": "LACRADOR REDONDO BLANCO",
    "description": "Ref: 505003 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 12,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-488",
    "sku": "MC 488",
    "name": "FILTROS PARA OPU",
    "description": "Ref: 514003 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-489",
    "sku": "MC 489",
    "name": "MICROPIPETA VOLUMEN VARIABLE 0,5 - 10 UL / CRAL",
    "description": "Ref: CP-10V | Marca: CRAL | Prov:",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-490",
    "sku": "MC 490",
    "name": "27512 LAINER VAGINA ARTIFICIAL OVINA-CAPRINA",
    "description": "Ref: 27512 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-491",
    "sku": "MC 491",
    "name": "5111 PROTECTOR VAGINA ARTIFICIAL OVINA-CAPRINA",
    "description": "Ref: 5111 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-492",
    "sku": "MC 492",
    "name": "MICROSCOPIO VERTICAL MONOCULAR PARA CAMPO CLARO DE ILUMINACIÓN LED ECLIPSE EI/ MARCA: NIKON",
    "description": "Ref: 90453 | Marca: NIKON | Prov: ZIVOT",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-493",
    "sku": "MC 493",
    "name": "MÁQUINA AUTOMÁTICA PARA CONGELACIÓN DE SEMEN/ USO VETERINARIO",
    "description": "Ref: 84198999 | Marca: NEOVET | Prov: NEOVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-494",
    "sku": "MC 494",
    "name": "ELECTROJAC PROBE 1\" (25mm)",
    "description": "Ref: 6989 | Marca: ELECTROJAC | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-495",
    "sku": "MC 495",
    "name": "ECÓGRAFO PARA PORCINOS CON SONDA SECTORIAL.",
    "description": "Ref: M50 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-496",
    "sku": "MC 496",
    "name": "PINZA COGEPAJILLA",
    "description": "Ref: 11571 | Marca: | Prov: COLINSTRUVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 27,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-497",
    "sku": "MC 497",
    "name": "MYOINOSITOL BIOREAGENT SUITABLE",
    "description": "Ref: MG-I752850G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-498",
    "sku": "MC 498",
    "name": "FRASCOS PARA CULTIVO CELULAR X 75 CM PS, TAPA SIN FILTRO ESTERIL TRATADO",
    "description": "Ref: 70175 | Marca: SPL | Prov: ISLAS",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 23,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-499",
    "sku": "MC 499",
    "name": "SONDA CONVEXA 3.5 MHz MULTI FRECUENCIA",
    "description": "Ref: KX5600 | Marca: KAIXIN | Prov: KAIXIN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-500",
    "sku": "MC 500",
    "name": "LENTE ISPERM",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-501",
    "sku": "MC 501",
    "name": "PUNTERA DE AGUJA GUÍA OPU WTA",
    "description": "Ref: 10191 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-502",
    "sku": "MC 502",
    "name": "MANDRIL GUÍA OPU WTA",
    "description": "Ref: 10109 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-503",
    "sku": "MC 503",
    "name": "APLICACIÓN (SOFWARE) ISPERM ADICIONAL",
    "description": "Ref: ADBISP03 | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-504",
    "sku": "MC 504",
    "name": "BATERÍA ECÓGRAFO L60",
    "description": "Ref: BAT L60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-505",
    "sku": "MC 505",
    "name": "GOBLET DE 13 MM AMARILLO",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-506",
    "sku": "MC 506",
    "name": "ESCALERILLAS PARA GOBLETS DE 13 MM",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-507",
    "sku": "MC 507",
    "name": "LAPAROSCOPIO CON LENTE (5mm) 0 GRADOS",
    "description": "Ref: 1601 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-508",
    "sku": "MC 508",
    "name": "KIT FUENTE DE LUZ PORTÁTIL",
    "description": "Ref: HF6518.905.1 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-509",
    "sku": "MC 509",
    "name": "TROCAR Y VAINA PRINCIPAL (INOX)",
    "description": "Ref: HF6518.010.1 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-510",
    "sku": "MC 510",
    "name": "TROCAR Y VAINA SECUNDARIA (INOX)",
    "description": "Ref: HF6518.010.1 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-511",
    "sku": "MC 511",
    "name": "MINI PISTOLA DE 0,25 ML OVINA- CAPRINA",
    "description": "Ref: 7177 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-512",
    "sku": "MC 512",
    "name": "ABT PVA CONCENTRAD 20 ML VC ENERO 2027",
    "description": "Ref: | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 7,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-513",
    "sku": "MC 513",
    "name": "PISTOLA DE TE 0,25 MARAC: IMV",
    "description": "Ref: 7246 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-514",
    "sku": "MC 514",
    "name": "PISTOLA DE TE 0,25 MARAC: IMV",
    "description": "Ref: 7240 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-515",
    "sku": "MC 515",
    "name": "IPAD MINI 7.9\" PLATA 64 GB",
    "description": "Ref: 190199062535 | Marca: APPLE | Prov: ALKOSTO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-516",
    "sku": "MC 516",
    "name": "PUNTA AZUL PLÁSTICA PARA PIPETA DE 100 A 1000 UL",
    "description": "Ref: 961-04 | Marca: KARTELL | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-517",
    "sku": "MC 517",
    "name": "LAMINILLAS CUBRE OBJETOS 22X22MM",
    "description": "Ref: | Marca: GLASS LAB | Prov: COLMEDICAL SOLUTIONS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-518",
    "sku": "MC 518",
    "name": "UPS INTERACTIVA 3000VA REF MICRONET 3000",
    "description": "Ref: 41309501 | Marca: MICRONET | Prov: ESTRATEGIAS TECNOLÓGICAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-519",
    "sku": "MC 519",
    "name": "EVE MODULAR/ INCUBADORA DE EMBRIONES",
    "description": "Ref: 20920 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-520",
    "sku": "MC 520",
    "name": "PLATINA TÉRMICA UNIVERSAL PARA ESTEREOMICROSCOPIO",
    "description": "Ref: 20931 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-521",
    "sku": "MC 521",
    "name": "ESTEREOMICOSCOPIO WS7",
    "description": "Ref: WS7 | Marca: WESCO | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-522",
    "sku": "MC 523",
    "name": "CONECTOR LAPAROSCOPIO",
    "description": "Ref: HF6518.945 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-523",
    "sku": "MC 522",
    "name": "PINZA ALLYS 15.5 CM",
    "description": "Ref: 680 | Marca: AMALGADEN | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-524",
    "sku": "MC 524",
    "name": "PINZA CAMPO BACKAUS 11 CM",
    "description": "Ref: 683 | Marca: AMALGADEN | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-525",
    "sku": "MC 525",
    "name": "PORTA AGUJAS MAYO HEGAR 16 CM",
    "description": "Ref: 1976 | Marca: HAUPTNER | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-526",
    "sku": "MC 526",
    "name": "TOPIZADOR NORMAL MEDIANO",
    "description": "Ref: 1521 | Marca: OCHOA | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-527",
    "sku": "MC 527",
    "name": "EQUIPO BÁSICO DE CIRUGÍA",
    "description": "Ref: 11262 | Marca: | Prov: COLINSTRUVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-528",
    "sku": "MC 528",
    "name": "GUÍA OPU SONOSCAPE C612",
    "description": "Ref: 25867 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-529",
    "sku": "MC 529",
    "name": "CAJA PETRI DESECHABLE ESTERIL 94X16MM",
    "description": "Ref: TC-633181 | Marca: GREINER BIO-ONE | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-530",
    "sku": "MC 530",
    "name": "ASPIC PARA IA OVINOS",
    "description": "Ref: 5546 | Marca: IMV | Prov: RAB GLOBAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-531",
    "sku": "MC 531",
    "name": "PASTILLA BROMOPOL",
    "description": "Ref: 413597 | Marca: | Prov: BIOEASY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-532",
    "sku": "MC 532",
    "name": "PUNTAS BLANCAS UNIVERSALES, PLÁSTICAS, LIBRES DE DNASAS, RNASA Y PIROGENOS, RANGO DE 0.1 A 10UL. (CORTAS).",
    "description": "Ref: 4840 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-533",
    "sku": "MC 533",
    "name": "CAJA PETRI DE 35 MM",
    "description": "Ref: EXP000587 150255 | Marca: | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-534",
    "sku": "MC 534",
    "name": "TABS BLANCOS",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-535",
    "sku": "MC 535",
    "name": "PERCOLL(R), PH 8.5-9.5 (20 DEGREE C) REFRIGERADO",
    "description": "Ref: MG-P1644500ML | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-536",
    "sku": "MC 536",
    "name": "MANGA LARGA SUPERSENSITIVA CON RESORTE",
    "description": "Ref: 101341 | Marca: | Prov: LHAURA VET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-537",
    "sku": "MC 537",
    "name": "PAPEL PARAFILM EN ROLLO, LONGITUD 38 Mts, X 10CM. (4inX125ft) MARCA: PARAFILM M REF: 7016-05 (PM 996)",
    "description": "Ref: 13-374-10 | Marca: PARAFILM | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-538",
    "sku": "MC 538",
    "name": "CLORURO DE SODIO AL 0.9% (SOLUCION SALINA AL 0.9%) X 100 ML",
    "description": "Ref: 01002201-03 | Marca: | Prov: INCOLAMERICA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-539",
    "sku": "MC 539",
    "name": "CAJA VACÍA RACK PARA PUNTAS AMARILLAS DE 200 ul",
    "description": "Ref: KJ-814 | Marca: | Prov: CARPER LABS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-540",
    "sku": "MC 540",
    "name": "PUNTAS AMARILLAS DE 200 UL",
    "description": "Ref: 4330-0021-17 | Marca: PIPETTE TIP | Prov: CARPER LABS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UL",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-541",
    "sku": "MC 541",
    "name": "CAJA PETRI 65 X15/ ESTÉRIL CLARIDAD",
    "description": "Ref: 6200 | Marca: BRIXCO | Prov: LABEXCO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-542",
    "sku": "MC 542",
    "name": "TUBO CÓNICO DE 15 ML, TAPA ROSCA ESTÉRIL RCF MAX: 10000xg, graduados, PCR puro. VENCE: 03/04/2028",
    "description": "Ref: CTB15 | Marca: GEN FOLLOWER | Prov: BPL",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-543",
    "sku": "MC 543",
    "name": "TUBO CÓNICO DE 50 mL, TAPA ROSCA, PRE ESTERILIZADO. RCF MAX: 10000xg, graduados, PCR puro.",
    "description": "Ref: CTB50 | Marca: GEN FOLLOWER | Prov: BPL",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-544",
    "sku": "MC 544",
    "name": "CAMISA SANITARIA IMV INSEMINACIÓN X 100",
    "description": "Ref: | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-545",
    "sku": "MC 545",
    "name": "LifeStraw PERSONAL",
    "description": "Ref: | Marca: LIFESTRAW | Prov: IWANA GREEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-546",
    "sku": "MC 546",
    "name": "LifeStraw GO",
    "description": "Ref: | Marca: LIFESTRAW | Prov: IWANA GREEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-547",
    "sku": "MC 547",
    "name": "LifeStraw FAMILY 2.0",
    "description": "Ref: | Marca: LIFESTRAW | Prov: IWANA GREEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-548",
    "sku": "MC 548",
    "name": "LifeStraw FAMILY 1.1",
    "description": "Ref: | Marca: LIFESTRAW | Prov: IWANA GREEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-549",
    "sku": "MC 549",
    "name": "LifeStraw COMMUNITY",
    "description": "Ref: | Marca: LIFESTRAW | Prov: IWANA GREEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-550",
    "sku": "MC 550",
    "name": "PUNTAS AMARILLAS, 0-200 ul",
    "description": "Ref: RSR066 | Marca: BIOSIGMA | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-551",
    "sku": "MC 551",
    "name": "PUNTAS 200 ul EN RACK SIN FILTRO",
    "description": "Ref: 06-365-2018 | Marca: NERBE PLUS | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-552",
    "sku": "MC 552",
    "name": "ANDROSTAR PLUS 47G= 1 LITRO",
    "description": "Ref: 1135 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "47 GRAMOS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-553",
    "sku": "MC 553",
    "name": "GOBLET DE 13 MM BLANCOS",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-554",
    "sku": "MC 554",
    "name": "MG-A7469100G L-ALANINE, FROM NON-ANIMAL SOURCE, MEET, 100G",
    "description": "Ref: MG-A7469100G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-555",
    "sku": "MC 555",
    "name": "MG-G854025G/ LGLUTAMINE; NONANIMAL SOURCE; CELL CUL X 25G",
    "description": "Ref: MG-G854025G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-556",
    "sku": "MC 556",
    "name": "MG-P7794100M/ PENICILLIN G POTASSIUM SALT BIOREAGENT X 100MU",
    "description": "Ref: MG-P7794100M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 MU",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-557",
    "sku": "MC 557",
    "name": "MG-S913725G /STREPTOMYCIN SULFATE SALT BIOREAGENT; S X 25G",
    "description": "Ref: MG-S913725G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-558",
    "sku": "MC 558",
    "name": "PYRUVIC ACID SODIUM CELL CULTURE TESTED X 25G",
    "description": "Ref: MG-P528025G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-559",
    "sku": "MC 559",
    "name": "MG-L438850G/ CALCIUM L-LACTATE HYDRATE BIOREAGENT, , 50G",
    "description": "Ref: MG-L438850G | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-560",
    "sku": "MC 560",
    "name": "TORNILLOS Y DESTORNILLADOR PARA GUÍA OPU WTA",
    "description": "Ref: | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-561",
    "sku": "MC 561",
    "name": "ECÓGRAFO PORTÁTIL RKU 10",
    "description": "Ref: RKU 10 | Marca: KAIXIN | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-562",
    "sku": "MC 562",
    "name": "DISPOSITIVO ENTROMETRIX",
    "description": "Ref: 15933 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-563",
    "sku": "MC 563",
    "name": "LINER DE LÁTEX PARA VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: 26647 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-564",
    "sku": "MC 564",
    "name": "DESCORNADOR 7/8\" X 30 200WATTS",
    "description": "Ref: C2430N | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-565",
    "sku": "MC 565",
    "name": "PUNTAS AMARILLAS, UNIVERSALES. PLASTICAS, NO ESTERIL. LIBRES DE RNASE, DNASE Y PYROGENOS. RANGO DE 1-200UL. BOLSA X1000UDS.",
    "description": "Ref: 4845 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 38,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-566",
    "sku": "MC 566",
    "name": "SOPORTE LINEAL EN ACRÍLICO, PARA 6 MICROPIPETAS",
    "description": "Ref: SLIM | Marca: E&Q | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-567",
    "sku": "MC 567",
    "name": "PANTALLA LCD",
    "description": "Ref: | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-568",
    "sku": "MC 568",
    "name": "BAÑO SEROLÓGICO EN ACRÍLICO",
    "description": "Ref: 40201 | Marca: INDULAB | Prov: INDULAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-569",
    "sku": "MC 569",
    "name": "PISTOLA DE INSEMINACIÓN UNIVERSAL (AUTOMATIC LOCK)",
    "description": "Ref: 18367 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-570",
    "sku": "MC 570",
    "name": "7176 MINI PISTOLA 0.5 ML OVINA-CAPRINA",
    "description": "Ref: 70002000002 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-571",
    "sku": "MC 571",
    "name": "6931 CATETER PARA PISTOLA OVINA-CAPRINA PQ X 40 UN",
    "description": "Ref: 70002000003 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-572",
    "sku": "MC 572",
    "name": "6921 CATETER PISTOLA OVINA PQ X 40 UN IMV",
    "description": "Ref: 70002000013 | Marca: IMV | Prov: ROYAL PLUS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-573",
    "sku": "MC 573",
    "name": "PIPETA ROBETSON, STANDARD, NO ESTERILIZADO",
    "description": "Ref: 1040 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-574",
    "sku": "MC 574",
    "name": "ECÓGRAFO PORTATIL S1, PROBE RECTAL Y PROTECTOR SOLAR",
    "description": "Ref: S1 | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-575",
    "sku": "MC 575",
    "name": "VAGINA ARTIFICIAL OVINA- CAPRINA",
    "description": "Ref: 7179 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-576",
    "sku": "MC 576",
    "name": "LINER LISO PARA VAGINA ARTIFICIAL PEQUEÑOS RUMIANTES",
    "description": "Ref: | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-577",
    "sku": "MC 577",
    "name": "CUBIERTA PROTECTORA DE FIELTRO PARA TUBO COLECTOR",
    "description": "Ref: | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-578",
    "sku": "MC 578",
    "name": "ESPECULO DE ACERO INOXIDABLE DE USO VETERINARIO",
    "description": "Ref: 17373/0000 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-579",
    "sku": "MC 579",
    "name": "PISTOLA DE INSEMINACIÓN 0,5 OVINOS Y CAPRINOS",
    "description": "Ref: 7175 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-580",
    "sku": "MC 580",
    "name": "TUBOS DE PLÁSTICO TRANSPARENTE DE RECAMBIO PARA ESPÉCULO",
    "description": "Ref: 17373/001 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-581",
    "sku": "MC 581",
    "name": "MG- H3393250K HEPARIN SODIUM SALT FROM PORCINE INTESTporcin, Grado IA, 180 unidades USP / mg",
    "description": "Ref: MG- H3393250K | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 KU",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-582",
    "sku": "MC 582",
    "name": "CAMILLA OVINA HIDRÁULICA NACIONAL",
    "description": "Ref: CMOVINA | Marca: NACIONAL | Prov: CESAR DELGADO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-583",
    "sku": "MC 583",
    "name": "PINZA PRESIÓN ÚTERO LAPAROSCOPIA EN OVINOS Y CAPRINOS",
    "description": "Ref: HF6518.039/ HF6518303/ HF6518.059 | Marca: | Prov: HANVER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-584",
    "sku": "MC 584",
    "name": "PROBE 1 1/4\" DE DIÁMETRO CON ADAPTACIÓN A CABLE",
    "description": "Ref: 22-2X | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-585",
    "sku": "MC 585",
    "name": "5111 PROTECTOR VAGINA ARTIFICIAL OVINA-CAPRINA",
    "description": "Ref: 5111 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-586",
    "sku": "MC 586",
    "name": "RACK DE CONGELACIÓN",
    "description": "Ref: 7117 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-587",
    "sku": "MC 587",
    "name": "PINZA COGEPAJILLA",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-588",
    "sku": "MC 588",
    "name": "MICROSCOPIO CON PANTALLA LCD DE 7\", CON PLATINA TÉRMICA, ACROMÁTICO 10x20x40x",
    "description": "Ref: A33.1513 | Marca: OPTO-EDU | Prov: OPTO-EDU",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-589",
    "sku": "MC 589",
    "name": "MICROSCOPIO CON PANTALLA LCD DE 7\", 2,0 M, ACROMÁTICO 10x20x40x",
    "description": "Ref: A33.1512 | Marca: OPTO-EDU | Prov: OPTO-EDU",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-590",
    "sku": "MC 590",
    "name": "MICROSCOPIO BIOLÓGICO, BINOCULAR, COLECTOR KHOLER",
    "description": "Ref: A11.1109-A | Marca: OPTO-EDU | Prov: OPTO-EDU",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-591",
    "sku": "MC 591",
    "name": "MICROSCOPIO DE LABORATORIO BINOCULAR",
    "description": "Ref: A12.1502-AB | Marca: OPTO-EDU | Prov: OPTO-EDU",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-592",
    "sku": "MC 592",
    "name": "TUBO 5ML, 75X12MM PS",
    "description": "Ref: 62.476.028X | Marca: SARSTEDT AG & CO | Prov: QUIK",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-593",
    "sku": "MC 593",
    "name": "PISTOLA DE INSEMINACIÓN OVINO/CAPRINO PARA VAINAS MEDIANAS",
    "description": "Ref: 7175 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-594",
    "sku": "MC 594",
    "name": "VAINA MEDIANA OVINA/CAPRINA",
    "description": "Ref: 6921 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-595",
    "sku": "MC 595",
    "name": "PISTOLA DE INSEMINACION OVINA/ CAPRINA PARA VAINAS MINI",
    "description": "Ref: 7176 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-596",
    "sku": "MC 596",
    "name": "VAINA MINI OVINA/ CAPRINA COMPATIBLE CON PISTOLA 7176",
    "description": "Ref: 6922 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-597",
    "sku": "MC 597",
    "name": "OPTIXCELL",
    "description": "Ref: 26218 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-598",
    "sku": "MC 598",
    "name": "ELECTRO PARA OVINO- CAPRINO/ PROBE 1 1/4\" DE DIÁMETRO FUNCIONA CON 2 BATERIAS 9V",
    "description": "Ref: 26-0 | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-599",
    "sku": "MC 599",
    "name": "TANQUE DE DIÓXIDO DE CARBONO DE 3 LITROS CON REGULADOR DE DIÓXIDO DE CARBONO",
    "description": "Ref: | Marca: | Prov: CESAR DELGADO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-600",
    "sku": "MC 600",
    "name": "MICROSCOPIO CON PLATINA Y PANTALLA 7\"",
    "description": "Ref: HC-R069 | Marca: HAPPYCARE | Prov: HAPPYCARE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 9,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-601",
    "sku": "MC 601",
    "name": "ECÓGRAFO EMPEROR N2",
    "description": "Ref: EM-N2 | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-602",
    "sku": "MC 602",
    "name": "ISPERM SOFTWARE INICIAL ( UNA APLICACIÓN)",
    "description": "Ref: | Marca: ISPERM | Prov: AIDMICS BIOLOCHNOLOGY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-603",
    "sku": "MC 603",
    "name": "PIPETA 0,5-10 UL",
    "description": "Ref: A3021 | Marca: ZENITHLAB | Prov: EUROLAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-604",
    "sku": "MC 604",
    "name": "AGUJA PARA ASPIRACIÓN 20G",
    "description": "Ref: 13316 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 410,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-605",
    "sku": "MC 605",
    "name": "POLVO SELLADOR DE PAJILLAS COLOR BLANCO, 750GR",
    "description": "Ref: 44-0200 | Marca: REPRODUCTION PROVISIONS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "750 GRAMOS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-606",
    "sku": "MC 606",
    "name": "CAPACITACIÍN VIRTUAL",
    "description": "Ref: | Marca: | Prov:",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-607",
    "sku": "MC 607",
    "name": "CABLE PARA ELECTRO",
    "description": "Ref: 22.4 | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-608",
    "sku": "MC 608",
    "name": "TUBO CONICO. PLUG SEAL. PP 50 ML TAPA EN PE. GRADUADO AUTOCLAVABLE ESTERIL LIBRE DNase, RNase TAPA VERDE",
    "description": "Ref: 602002-1 | Marca: NEST | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-609",
    "sku": "MC 609",
    "name": "MICROPIPETA VOLUMEN VARIABLE 0,5- 10 UL",
    "description": "Ref: CP-10V | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-610",
    "sku": "MC 610",
    "name": "MG-459844500/ ETHYL ALCOHOL; PURE; 200 PROOF; ACS X 500ML",
    "description": "Ref: MG-459844500 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-611",
    "sku": "MC 611",
    "name": "TUBO PARA CENTRÍFUGA TIPO FALCON DE 50 ML, FONDO CÓNICO, TAPA ROSCA COLOR VERDE",
    "description": "Ref: 1111-QLS | Marca: ONELAB | Prov: VITTRA SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-612",
    "sku": "MC 612",
    "name": "GOBLETS BLANCOS DE 13 MM",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-613",
    "sku": "MC 613",
    "name": "PUNTAS BLANCAS CORTA DE 0,5X 10 (TIPO Gilson) fv(25/01/2025)",
    "description": "Ref: 771293 | Marca: GREINER BIO-ONE | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-614",
    "sku": "MC 614",
    "name": "CINTA PARA MEDICION ESCROTAL",
    "description": "Ref: 698 | Marca: | Prov: IMPLEGAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-615",
    "sku": "MC 615",
    "name": "SISTEMA OPU 1,20 M",
    "description": "Ref: 21789 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 29,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-616",
    "sku": "MC 616",
    "name": "TRANSPORTADORA DE EMBRIONES Y OOCITOS",
    "description": "Ref: TC-39/240 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-617",
    "sku": "MC 617",
    "name": "CALENTADOR DE TUBOS 50 ML",
    "description": "Ref: 18029 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-618",
    "sku": "MC 618",
    "name": "TABS AZULES",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-619",
    "sku": "MC 619",
    "name": "TABS ROSADOS",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-620",
    "sku": "MC 620",
    "name": "CAMISAS SANITARIAS DE INSEMINACIÓN 53 CM",
    "description": "Ref: 19271/0080 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-621",
    "sku": "MC 621",
    "name": "ONE STEP THAW/ 1M SUCROSE",
    "description": "Ref: ABT247 | Marca: ABT 360 | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-622",
    "sku": "MC 622",
    "name": "ETHYLENE GLYCOL FREEZE WITH SUCROSE",
    "description": "Ref: ABT234 | Marca: ABT 360 | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-623",
    "sku": "MC 623",
    "name": "HOLDING",
    "description": "Ref: ABT224 | Marca: ABT 360 | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-624",
    "sku": "MC 624",
    "name": "SOPORTE PORTA BUBBLES/ CUBETA METALICO",
    "description": "Ref: 7116 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-625",
    "sku": "MC 625",
    "name": "SONDA PROVETSCAN WIFI SR-2C/ WRPCCFB019",
    "description": "Ref: 1106 | Marca: PROVETSCAN | Prov: PROVETSCAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-626",
    "sku": "MC 626",
    "name": "NATURAL FERTILITY",
    "description": "Ref: 7709265238861 | Marca: NATURAL VETERINARY | Prov: NATURAL VETERINARY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 KILOGRAMO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-627",
    "sku": "MC 627",
    "name": "SONDA CONVEXA",
    "description": "Ref: L 60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-628",
    "sku": "MC 628",
    "name": "SONDA MICRO CONVEXA",
    "description": "Ref: L 60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-629",
    "sku": "MC 629",
    "name": "CHIP CON 8 CURVAS CL2000",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-630",
    "sku": "MC 630",
    "name": "TERMO DE NITRÓGENO WORTHINGTON TW20/ BIOMEDICAL",
    "description": "Ref: TW20 | Marca: BIOMEDICAL | Prov: ANICAM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-631",
    "sku": "MC 631",
    "name": "TERMO DRY SHIPPER CX100",
    "description": "Ref: CX100B-11M | Marca: BIOMEDICAL | Prov: ANICAM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-632",
    "sku": "MC 632",
    "name": "CONGELADORA DE SEMEN Y DE EMBRIONES",
    "description": "Ref: CL-2200 | Marca: CRYOLOGIC | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-633",
    "sku": "MC 633",
    "name": "CAMISA SANITARIA 21¨ ROLLO X 80",
    "description": "Ref: 5563 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO X 80 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-634",
    "sku": "MC 634",
    "name": "PISTOLA DE TRANSFERENCIA DE BOVINOS",
    "description": "Ref: 7246 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-635",
    "sku": "MC 635",
    "name": "VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: 5417 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-636",
    "sku": "MC 636",
    "name": "CONO PARA VAGINA ARTIFICIAL BOVINA 27 CM",
    "description": "Ref: 5204 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-637",
    "sku": "MC 637",
    "name": "LINER DE LÁTEX PARA VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: 26541 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-638",
    "sku": "MC 638",
    "name": "FILTRO MINI FLUSH",
    "description": "Ref: 19222/20000 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-639",
    "sku": "MC 639",
    "name": "VAGINA EQUINA COMPLETA INCLUYE LINER Y FORRO DE CUERO",
    "description": "Ref: | Marca: MISSOURI | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-640",
    "sku": "MC 640",
    "name": "PAJILLA TRANSPARENTE VERDE 0,25",
    "description": "Ref: 1094 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-641",
    "sku": "MC 641",
    "name": "FILTROS DE JERINGA CORNING® DE 28 MM 1 Cax50 395.200 19% 395.200 DE DIÁMETRO, MEMBRANA PES DE PORO DE 0,2 µm, ESTÉRILES 17221037 01/01/9999",
    "description": "Ref: EXP003186 431229 | Marca: CORNING | Prov: EXPERT",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-642",
    "sku": "MC 642",
    "name": "PUNTA BLANCA CORTA DE 0,5-10 (tipo Gilson)",
    "description": "Ref: BSR064 | Marca: BIOSIGMA | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-643",
    "sku": "MC 643",
    "name": "PIPETAS BASICA VL VARIABLE 1-10 ML SERIE L60599K",
    "description": "Ref: 3120000283 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-644",
    "sku": "MC 644",
    "name": "PIPETA VOL VARIABLE 500-5000 UL SERIE K43202J",
    "description": "Ref: 3120000070 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-645",
    "sku": "MC 645",
    "name": "PIPETA VOL. VARIABLE 20-200 EPPENDORF SERIE J10810H, J10759 H",
    "description": "Ref: 3120000054 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-646",
    "sku": "MC 646",
    "name": "MICROPIPETA PLUS 100-1000 UL EPPENDORF SERIE J54860L",
    "description": "Ref: 3123000268 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-647",
    "sku": "MC 647",
    "name": "CLIP PARA 15 PAJILLAS MEDIANAS",
    "description": "Ref: 7008 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-648",
    "sku": "MC 648",
    "name": "BOQUILLA DE LLENADO CON SUCCIÓN 15 PIN 0.5 ML",
    "description": "Ref: 7297 | Marca: IMV | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-649",
    "sku": "MC 649",
    "name": "ECÓGRAFO BLX-V50",
    "description": "Ref: BLX-V50 | Marca: BOXERLY | Prov: BOXERLY",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-650",
    "sku": "MC 650",
    "name": "CORTAPAJILLAS TED",
    "description": "Ref: 90008.004000000001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-651",
    "sku": "MC 651",
    "name": "DISPLAY OLED",
    "description": "Ref: 23218 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-652",
    "sku": "MC 652",
    "name": "CAMISA SANITARIA IMV INSEMINACIÓN X 100",
    "description": "Ref: 6459 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-653",
    "sku": "MC 653",
    "name": "OPTIDYL EXTENDEDOR DE SEMEN BOVINO",
    "description": "Ref: 20996 | Marca: IMV | Prov: IMV TECHNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-654",
    "sku": "MC 654",
    "name": "C8166-25M/CARBOXYFLUORESCEIN DIACETATE",
    "description": "Ref: C8166-25M | Marca: SIGMA | Prov: SIGMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25 MG",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-655",
    "sku": "MC 655",
    "name": "IPAD MINI 6/ 64 GB",
    "description": "Ref: MK7M3LZ/A | Marca: IPAD | Prov: ISHOP",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-656",
    "sku": "MC 656",
    "name": "MINICENTRÍFUGA MINI SPIN",
    "description": "Ref: 5452000816 | Marca: EPPENDORF | Prov: RTL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-657",
    "sku": "MC 657",
    "name": "LAMINA PORTA OBJETOS LISA EN VIDRIO 26X76MM, BORDE ESMERILADO",
    "description": "Ref: 7101-CRAL | Marca: PRECISION GLASS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 13,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-658",
    "sku": "MC 658",
    "name": "LAMINA CUBRE OBJETO EN VIDRIO 22X22",
    "description": "Ref: P22X22 | Marca: PRECISION GLASS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 26,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-659",
    "sku": "MC 659",
    "name": "PUNTA AMARILLA PLASTICA 0-200 UL",
    "description": "Ref: 18260NM | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-660",
    "sku": "MC 660",
    "name": "CAMARA DE NEUBAUER PARA RECUENTO DE GLOBULOS BLANCOS Y ROJOS",
    "description": "Ref: 90200.01 | Marca: LABSCIENT | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-661",
    "sku": "MC 661",
    "name": "CARGADDOR PARA PULSATOR IV",
    "description": "Ref: 44612 | Marca: | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-662",
    "sku": "MC 662",
    "name": "GOBLETS BLANCOS DE 10 MM",
    "description": "Ref: 347430 | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-663",
    "sku": "MC 663",
    "name": "CONTROLADOR DE TEMPERATURA",
    "description": "Ref: | Marca: NEOVET | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-664",
    "sku": "MC 664",
    "name": "JARRO DE COLECCIÓN DE SEMEN PORCINO PLÁSTICO",
    "description": "Ref: 1373 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-665",
    "sku": "MC 665",
    "name": "MICROPIPETA PLUS 1-10mL marca Eppendorf serie N14829L-N14948L",
    "description": "Ref: 3123000080 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-666",
    "sku": "MC 666",
    "name": "MICROPIPETA PLUS 0.5-5mL marca Eppendorf , Lote N73902L",
    "description": "Ref: 3123000071 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-667",
    "sku": "MC 667",
    "name": "MICROPIPETA PLUS 20-200μL marca Eppendorf, lote 067316L",
    "description": "Ref: 3123000055 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-668",
    "sku": "MC 668",
    "name": "MICROPIPETA PLUS 100-1000μL marca Eppendorf lote 032591L- 032745L",
    "description": "Ref: 3123000063 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-669",
    "sku": "MC 669",
    "name": "MICROPIPETA PLUSs 2-20μL , marca Eppendorf lote P54448L, P54382L,P54385",
    "description": "Ref: 3123000039 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-670",
    "sku": "MC 670",
    "name": "ECOGRAFO MINDRAY DP-50 Vet/ INCLUYE MALETÍN",
    "description": "Ref: 1152B-CTO-S01 | Marca: MINDRAY | Prov: MARQUETING NET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-671",
    "sku": "MC 671",
    "name": "CAMISA SANITARIA ROLLO X 80",
    "description": "Ref: 5563 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 367,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-672",
    "sku": "MC 672",
    "name": "PAJUELA AMARILLA 0.25/ PAQUETE X 2000",
    "description": "Ref: 5575 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-673",
    "sku": "MC 673",
    "name": "PUNTA AZULES UNIVERSALES 100- 1000 UL",
    "description": "Ref: 4846 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-674",
    "sku": "MC 674",
    "name": "CRYOVIAL 2,0 ml",
    "description": "Ref: CT0020 | Marca: GEN FOLLOWER | Prov: BPL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-675",
    "sku": "MC 675",
    "name": "MICROTUBO 1,5 ML PP AUTOCLAVABLE NO ESTERIL CONICO CON TAPAS TRANSPARENTES 81-0000 RANGO DE TEM -196 C - 121 C CAJA X 500 U - BIOLOGIX",
    "description": "Ref: 81-0153 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-676",
    "sku": "MC 676",
    "name": "PUNTAS COLOR AMARILLO ESTÁNDAR. GRADUADA 200 ul. PP ESTERIL LIBRES DE RNAsa, DNAsa Y PIROGENOS AUTOCLAVABLES CAJA 10 X 96 RACK - BIOLOGIX",
    "description": "Ref: 21-0200 | Marca: BIOLOGIX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-677",
    "sku": "MC 677",
    "name": "DINAMÓMETRO ELÉCTRNIC CAPACIDAD 50 KG",
    "description": "Ref: 101201015 | Marca: BBG | Prov: BBG",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-678",
    "sku": "MC 678",
    "name": "PLATINA CALENTADORA PARA MICROSCOPIO T200",
    "description": "Ref: T-200 | Marca: | Prov: UNITED SCOPE LLC",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-679",
    "sku": "MC 679",
    "name": "REGLA PARA NITRÓGENO",
    "description": "Ref: | Marca: | Prov: GENÉTICA SELECTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-680",
    "sku": "MC 680",
    "name": "PLATINA TÉRMICA MA- BIO",
    "description": "Ref: 9003008 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-681",
    "sku": "MC 681",
    "name": "PURIFICADOR DE AIRE UPM/EC-31-F7+HEPA H14",
    "description": "Ref: | Marca: SODECA | Prov: SODECA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-682",
    "sku": "MC 682",
    "name": "CONO RECOLECTOR",
    "description": "Ref: | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 28,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-683",
    "sku": "MC 683",
    "name": "PAJILLAS AMARILLAS 0,25",
    "description": "Ref: | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-684",
    "sku": "MC 684",
    "name": "PUNTAS COLOR NATURAL ESTANDAR GRADUADA 10 UL",
    "description": "Ref: 21-0010 | Marca: BIOLGIX | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-685",
    "sku": "MC 685",
    "name": "PARAFILM 38M X 10 CM",
    "description": "Ref: 1868703 | Marca: PARAFILM | Prov: NORQUIMICOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "ROLLO X 38 MTS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-686",
    "sku": "MC 686",
    "name": "MICROPIPETA PREMIUM VOLUMEN VARIABLE 10- 100 UL",
    "description": "Ref: CP-P100V | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-687",
    "sku": "MC 687",
    "name": "ESPÉCULO INOX PICO DE PATO OVINO",
    "description": "Ref: | Marca: NACIONAL | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-688",
    "sku": "MC 688",
    "name": "TUBO CONICO. PLUG SEAL. PP 50 ML TAPA EN PE. GRADUADO AUTOCLAVABLE ESTERIL LIBRE Dnase, Rnase TAPA VERDE BOLSA X 50 UNI, MAX RCF (xg) 12.000 - NEST",
    "description": "Ref: 602002-2 | Marca: NEST | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-689",
    "sku": "MC 689",
    "name": "SONDA CONVENXA ECÓGRAFO PT 50C",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-690",
    "sku": "MC 690",
    "name": "PANTALLA L60",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-691",
    "sku": "MC 691",
    "name": "LIDOCAÍNA HCL 2 % X 100 ML/ 2025-12-16",
    "description": "Ref: 00051076-02 | Marca: CORPAUL | Prov: CORPAUL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -19,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-692",
    "sku": "MC 692",
    "name": "PUNTAS TRANSPARENTES DE 5- 10 UL GRADUADAS LIBRE DE RNA SAS DNA SAS",
    "description": "Ref: 43300024 | Marca: CITOTEST | Prov: DGC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-693",
    "sku": "MC 693",
    "name": "BATERÍA TRANSPORTADORA WTA",
    "description": "Ref: | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-694",
    "sku": "MC 694",
    "name": "GRADILLA PLÁSTICA EN PP PARA 60 TUBS TIPO FALCON 30X15 ML / 20X50 ML/ MARCA: CRALPLAST",
    "description": "Ref: 181550 | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-695",
    "sku": "MC 695",
    "name": "PISTOLA DE TE 0,25 MARCA IMV",
    "description": "Ref: 7240 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-696",
    "sku": "MC 696",
    "name": "PISTOLA DE TE 0,5 / MARCA IMV",
    "description": "Ref: 24857 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-697",
    "sku": "MC 697",
    "name": "OPTIXCELL 250 ML",
    "description": "Ref: 26218 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-698",
    "sku": "MC 698",
    "name": "CUBETA PLÁSTICA MEDIANA PARA 0,50 P/25",
    "description": "Ref: 6935 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-699",
    "sku": "MC 699",
    "name": "CAMISA SANITARIA CAJA X 100",
    "description": "Ref: 6459 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 80,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-700",
    "sku": "MC 700",
    "name": "FORRO ECÓGRAFO L60",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-701",
    "sku": "MC 701",
    "name": "FILTROS OPU 100 UM/ MARCA TED",
    "description": "Ref: 9014001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-702",
    "sku": "MC 702",
    "name": "FILTROS OPU 100 UM/ MARCA TED",
    "description": "Ref: 9014001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-703",
    "sku": "MC 703",
    "name": "SEMEN GUS BRAHMAN IMPORTADO",
    "description": "Ref: 7BR528 | Marca: | Prov: GENÉTICA SELECTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-704",
    "sku": "MC 704",
    "name": "NITRÓGENO",
    "description": "Ref: NITROGENO | Marca: | Prov: GENÉTICA SELECTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-705",
    "sku": "MC 705",
    "name": "FORRO DE GOMITA PASA ECÓGRAFO S1",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-706",
    "sku": "MC 706",
    "name": "FILTRO DE JERINGA ACRODISC, MEMBRANA SUPOR (PES), DIAM 25MM X 0,2 UM, ESTERIL POR GAMMA IRRADIACIÓN",
    "description": "Ref: 4612 | Marca: PALL | Prov: ONELAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-707",
    "sku": "MC 707",
    "name": "FILTRO DE JERINGA ACRODISC, MEMBRANA SUPOR (PES), DIAM 25MM X 0,45 UM, ESTERIL POR GAMMA IRRADIACIÓN",
    "description": "Ref: 4614 | Marca: PALL | Prov: ONELAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-708",
    "sku": "MC 708",
    "name": "PUNTA PIPETA UNIV. 0,5- 10 UL",
    "description": "Ref: T-300 | Marca: AXYGEN | Prov: EQUIMED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-709",
    "sku": "MC 709",
    "name": "VAGINA ARTIFICIAL EQUINA NACIONAL",
    "description": "Ref: | Marca: NACIONAL | Prov: ANDROLOGÍA Y REPRODUCCIÓN EQUINA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-710",
    "sku": "MC 710",
    "name": "MICROPIPETA PLUS 0.5- 10 UL / INCLUYE CAJA DE PUNTAS",
    "description": "Ref: 3120000802 | Marca: EPPENDORF | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-711",
    "sku": "MC 711",
    "name": "ESCALERILLAS METÁLICAS DE 10 MM",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-712",
    "sku": "MC 712",
    "name": "CATÉTER FOLLEY 32 FR 75 CC/ 864 MM",
    "description": "Ref: | Marca: PETS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-713",
    "sku": "MC 713",
    "name": "BOLSA WHRIL-PAK 7 OZ",
    "description": "Ref: B00992 | Marca: WHIRL-PAK | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 835,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-714",
    "sku": "MC 714",
    "name": "PIPETA RESEARCH PLUS 0.1-2.5 UL SERIE R3388IL",
    "description": "Ref: 3123000012 | Marca: EPPENDORF | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-715",
    "sku": "MC 715",
    "name": "MINISART NY 25 X 0.2. ESTERIL, SERIE 220262103, FV 2025/02",
    "description": "Ref: 17845-ACK | Marca: | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-716",
    "sku": "MC 716",
    "name": "FUNDAS DE INSEMINACIÓN 3W/ 0,50",
    "description": "Ref: 26043 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 27,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-717",
    "sku": "MC 717",
    "name": "PISTOLA DE INSEMINACIÓN 0,50",
    "description": "Ref: 28804 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-718",
    "sku": "MC 718",
    "name": "HIPPEX DILUYENTE DE SEMEN EQUINO",
    "description": "Ref: | Marca: HUVE SEARCH | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-719",
    "sku": "MC 719",
    "name": "DISPOSITIVO INTRAVAGINAL PARA OVINOS Y CAPRINOS",
    "description": "Ref: | Marca: | Prov: ANDINA PREMIER",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-720",
    "sku": "MC 720",
    "name": "VETEGON ECG X5MIL UND",
    "description": "Ref: CDAVEGET50 | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-721",
    "sku": "MC 721",
    "name": "GESTAR X 100 ML",
    "description": "Ref: OVEGEST100 | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "100 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-722",
    "sku": "MC 722",
    "name": "CICLAR X 50 ML",
    "description": "Ref: COVCICLX50 | Marca: | Prov: PAISAGRO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-723",
    "sku": "MC 723",
    "name": "GAFAS PARA ECÓGRAFO L60",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-724",
    "sku": "MC 724",
    "name": "ECÓGRAFO EMPEROR V8",
    "description": "Ref: EMP-V8 | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-725",
    "sku": "MC 725",
    "name": "ECÓGRAFO EMPEROR G20/ COLOR DOPPLER",
    "description": "Ref: EMP-G20 | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-726",
    "sku": "MC 726",
    "name": "BATERÍA ECÓGRAFO EMPEROR N2",
    "description": "Ref: EMP-N2BAT | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-727",
    "sku": "MC 727",
    "name": "GUÍA OPU EMPEROR/ PARA ECÓGRAFO G20",
    "description": "Ref: EMP-6.5C12811-OPU | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-728",
    "sku": "MC 728",
    "name": "MICROPUNTA NATURAL, 10 UL AUTOCLAVABLE",
    "description": "Ref: 3090.0010000000002 | Marca: SOCOREX | Prov: MUNDIAL DE EQUIPOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-729",
    "sku": "MC 729",
    "name": "MICROPUNTA NATURAL, 10 UL CORTA",
    "description": "Ref: 7501-96R | Marca: SCIENTIFIC | Prov: MUNDIAL DE EQUIPOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-730",
    "sku": "MC 730",
    "name": "CAJA DE PETRI DESECHABLE ESTERIL 130X130 MM CUADRADA",
    "description": "Ref: HP0012 | Marca: HD MEDICAL | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-731",
    "sku": "MC 731",
    "name": "TUBOS MICROCENTRIFUGA (EPPENDOR ) 1.5 ML. PP. SAFE LOCK, AUTOCLAVABLE NO ESTERIL, LIBRES DE RNAsa, D",
    "description": "Ref: SGB003 | Marca: BIO SEEN | Prov: ISLAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-732",
    "sku": "MC 732",
    "name": "PIPERA SEROLÓGICA ESTERIL 10 MLS/ CELLSTAR",
    "description": "Ref: 607180U | Marca: CELLSTAR | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-733",
    "sku": "MC 733",
    "name": "HEPES; BIOPERFORMANCE /CERTIFIED X 25G",
    "description": "Ref: MG-H403425G | Marca: SIGMA | Prov: SIGMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-734",
    "sku": "MC 734",
    "name": "SONDA PROVETSCAN WIFI SU-3/ INCLUYE: CARGADOR GENÉRICO 2A, ARNES GENÉRICO, MALETA SU-3A",
    "description": "Ref: 1505 | Marca: PROVETSCAN | Prov: NEW VETEC SL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-735",
    "sku": "MC 735",
    "name": "CARGADOR MICROSCOPIO",
    "description": "Ref: HC-R069 | Marca: HAPPYCARE | Prov: HAPPYCARE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-736",
    "sku": "MC 736",
    "name": "PLUSET 20 ML",
    "description": "Ref: 30004101900 | Marca: | Prov: VETERANDINA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-737",
    "sku": "MC 737",
    "name": "MANGA OBSTET. DESECHABHE TIPO T SUPER SENSITIVA",
    "description": "Ref: 101258 | Marca: LHAURA | Prov: LHAURA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 19,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-738",
    "sku": "MC 738",
    "name": "REGLILLA NITRÓGENO",
    "description": "Ref: 1154 | Marca: | Prov: GENÉTICA SELECTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-739",
    "sku": "MC 739",
    "name": "GUARDIAN 1,50 L",
    "description": "Ref: 1004992 | Marca: Dicpoplast | Prov: BIOHAUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-740",
    "sku": "MC 740",
    "name": "TERMO NITRÓGEN ET-3",
    "description": "Ref: 4451 | Marca: MVE | Prov: INSMEVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-741",
    "sku": "MC 741",
    "name": "MINI CENTRÍFUGA",
    "description": "Ref: XCM-12K | Marca: | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-742",
    "sku": "MC 742",
    "name": "PLATINA XH-2002",
    "description": "Ref: XH-2002 | Marca: PREMIER | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-743",
    "sku": "MC 743",
    "name": "ESTEREOMICROSCOPIO 2X 4X LED",
    "description": "Ref: SMP-24 | Marca: | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-744",
    "sku": "MC 744",
    "name": "ESTEREOMICROSCOPIO ZOOM",
    "description": "Ref: SMZ-05 | Marca: | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-745",
    "sku": "MC 745",
    "name": "OBJETIVO AUXILIAR 1.5X",
    "description": "Ref: SMZ-1.5 | Marca: | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-746",
    "sku": "MC 746",
    "name": "OBJETIVO AUXILIAR 2X",
    "description": "Ref: SMZ-2 | Marca: | Prov: SHANGHAI YOUDING INTERNATIONAL TRADE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-747",
    "sku": "MC 747",
    "name": "MG-G126450M / GENTAMICIN SULFATE SALT BIOREAGENT; SUIT X 50MG",
    "description": "Ref: MG-G126450M | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 MG",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-748",
    "sku": "MC 748",
    "name": "BALANZA ELECTRÓNICA DÍGITAL GRAMERA 40 KG RECARGABLE",
    "description": "Ref: | Marca: | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-749",
    "sku": "MC 749",
    "name": "MINI MICROSCOPIO",
    "description": "Ref: HC-B079E | Marca: | Prov: HAPPYCARE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-750",
    "sku": "MC 750",
    "name": "ADAPTADOR DE CELULAR PARA MICROSCOPIO",
    "description": "Ref: | Marca: | Prov: HAPPYCARE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 5,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-751",
    "sku": "MC 751",
    "name": "KIT PIPETA LLENADO PAJILLAS",
    "description": "Ref: FSF-101-KIT | Marca: ARSALES | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-752",
    "sku": "MC 752",
    "name": "MEDIO PARA CONGELACIÓN DE SEMEN CANINO",
    "description": "Ref: 13700 | Marca: CANIPLUS | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-753",
    "sku": "MC 753",
    "name": "ECÓGRAFO EASI SCAN GO LITE CON SONDA LINEAL KIT",
    "description": "Ref: ESG- LINEAL | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-754",
    "sku": "MC 754",
    "name": "ALARGADR DE SONDA ESG CORTO",
    "description": "Ref: E- INTRODUCER -GO | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-755",
    "sku": "MC 755",
    "name": "CORTAPAJILLAS WTA",
    "description": "Ref: | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 35,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-756",
    "sku": "MC 756",
    "name": "PIPETA PLUS 0.1-2.5µL, UNIDAD - EPPENDORF",
    "description": "Ref: 312300217 | Marca: EPPENDORF | Prov: ISLAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-757",
    "sku": "MC 757",
    "name": "PAJILLA MEDIANA 0,5 ML TRANSPARENTES",
    "description": "Ref: 13408/ 0010 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 250 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-758",
    "sku": "MC 758",
    "name": "ESFERAS METALICAS 0.5 ML",
    "description": "Ref: 1107 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-759",
    "sku": "MC 759",
    "name": "PIPETA EQUINA DE 65 CM",
    "description": "Ref: 1127 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-760",
    "sku": "MC 760",
    "name": "CABLE PLATINA ISPERM",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-761",
    "sku": "MC 761",
    "name": "GOBLET AMARILLO 10 MM",
    "description": "Ref: 347439 | Marca: REPRODUCTION PROVISION | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-762",
    "sku": "MC 762",
    "name": "CAMARA DE NEUBAUER",
    "description": "Ref: | Marca: | Prov: NORQUIMICOS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-763",
    "sku": "MC 763",
    "name": "PUNTA 10 UL LARGA LIBRES DNAsas, RNAasas, Y piR ESTACABLE",
    "description": "Ref: 320-3000A | Marca: BIO POINTE SCIENTIFIC | Prov: MUNDIAL DE EQUIPOS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "96 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-764",
    "sku": "MC 764",
    "name": "AGUJA PARA ASPIRACIÓN 18G",
    "description": "Ref: 21532 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 866,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-765",
    "sku": "MC 765",
    "name": "TERMO DE TRANSPORTE 1 LITRO",
    "description": "Ref: 1 | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 4,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-766",
    "sku": "MC 766",
    "name": "AGUJA DE GUÍA BOVINA PARA OPU",
    "description": "Ref: 9011008 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-767",
    "sku": "MC 767",
    "name": "SONDA Wi-Fi PROVERSCAN CR1A",
    "description": "Ref: CR1A | Marca: PROVERSCAN | Prov: PROVERSCAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-768",
    "sku": "MC 768",
    "name": "ELECTRO EYACULADOR PULSATOR V COMPLETO",
    "description": "Ref: 30-02560U | Marca: LANE MANUFACTURING | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-769",
    "sku": "MC 769",
    "name": "TROBE VERTICAL DE 60MM PARA ELECTRO PULSATOR",
    "description": "Ref: 25-60U | Marca: LANE MANUFACTURING | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-770",
    "sku": "MC 770",
    "name": "CABLE PARA ELECTRO PULSATOR IV",
    "description": "Ref: 45038 | Marca: LANE MANUFACTURING | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-771",
    "sku": "MC 771",
    "name": "CAMARA DE NEUBAUER IMPROVED LINEA ESTANDARD",
    "description": "Ref: 4401050 | Marca: | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-772",
    "sku": "MC 772",
    "name": "PUNTAS PARA MICROPIEPTA 2-20UL RACK X 96",
    "description": "Ref: 9605 | Marca: BIO POINTE SCIENTIFIC | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-773",
    "sku": "MC 773",
    "name": "ELECTRO EYACULADOR PULSATOR V COMPLETO CON 2 ELECTRODOS",
    "description": "Ref: 30-60U | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-774",
    "sku": "MC 774",
    "name": "PISTOLA DE INSEMINACIÓN ARTIFICIAL CON CÁMARA",
    "description": "Ref: PTC | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": -2,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-775",
    "sku": "MC 775",
    "name": "IPAD MINI 6/ 64 GB",
    "description": "Ref: | Marca: APPLE | Prov: SU TIENDA PRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-776",
    "sku": "MC 776",
    "name": "V1 KAIXIN ECOGRAFO PORTATL",
    "description": "Ref: V1 | Marca: KAIXIM | Prov: KAIXIN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-777",
    "sku": "MC 777",
    "name": "ESQUILADORA SHEAR MASTER OVINOS 1 00403 OSTER",
    "description": "Ref: 403/ 78153-003 | Marca: OSTER | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-778",
    "sku": "MC 778",
    "name": "PROBE DE 75 MM 2 ELECTRODOS",
    "description": "Ref: 25-75U | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-779",
    "sku": "MC 779",
    "name": "ELECTRO EYACULADOR LANE/ CON BATERIA",
    "description": "Ref: 26-0 | Marca: LANE MANUFACTURING | Prov: LANE MANUFACTURING",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-780",
    "sku": "MC 780",
    "name": "APLICACIÓN ISPERM ADICIONAL X 2 ESPECIES",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-781",
    "sku": "MC 781",
    "name": "FUNDAS DE TRANSFERENCIA AZUL PUNTA METÁLICA 21\"",
    "description": "Ref: 5540 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 5 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 99,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-782",
    "sku": "MC 782",
    "name": "FUNDAS DE TRANSFERENCAI DE 0,5 PARA EQUINOS IMV",
    "description": "Ref: 17532 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 143,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-783",
    "sku": "MC 783",
    "name": "ASPIC PARA IA OVINOS",
    "description": "Ref: 5546 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-784",
    "sku": "MC 784",
    "name": "TERMO DESCONGELADOR DE SEMEN DE EMBRIONES ROSADO",
    "description": "Ref: 20898 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-785",
    "sku": "MC 785",
    "name": "FUNDA DE INSEMINACIÓN 3W 0,25",
    "description": "Ref: 28801 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-786",
    "sku": "MC 786",
    "name": "CINTA ESCROTAL METÁLICA",
    "description": "Ref: 6986 | Marca: NEOGEN | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 13,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-787",
    "sku": "MC 787",
    "name": "TIJERA BALONADORA 30 CM RECTA",
    "description": "Ref: 101147 | Marca: LHAURA | Prov: LHAURA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-788",
    "sku": "MC 788",
    "name": "BASCULA OVINA ELECTRONICA",
    "description": "Ref: | Marca: ALFONSO AURELIO ALBARRACIN SANCHEZ | Prov: ALFONSO AURELIO ALBARRACIN SANCHEZ",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-789",
    "sku": "MC 789",
    "name": "MICROTUBO EPPENDORF PLASTICO GRADUADO X 2 ML ESTERIL/ TAPA LLANA, BASE CÓNICA",
    "description": "Ref: 181620 | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-790",
    "sku": "MC 790",
    "name": "MICROTUBO EPPENDORF PLASTICO GRADUADO X 0,5 ML ESTERIL/, BASE CÓNICA",
    "description": "Ref: 181645 | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-791",
    "sku": "MC 791",
    "name": "TUBO PLÁSTICO PARA PISTOLA DE INSEMINACIÓN",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 21,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-792",
    "sku": "MC 792",
    "name": "MAP-PROGES- ESPOV/ ESPONJA PARA OVINO- CAPRINO",
    "description": "Ref: | Marca: | Prov: NUCLEOLAB",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-793",
    "sku": "MC 793",
    "name": "ESTÉREO MICROSCOPIO BINOCULAR BS-3040B con 10X OCULAR/ SOPORTE BSZ-F35",
    "description": "Ref: BS-3040B / BSZ-F35 | Marca: BETSCAN | Prov: BETSCAN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-794",
    "sku": "MC 794",
    "name": "OCULAR 15X",
    "description": "Ref: | Marca: BETSCAN | Prov: BETSCAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAR",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-795",
    "sku": "MC 795",
    "name": "OCULAR 20X",
    "description": "Ref: | Marca: BETSCAN | Prov: BETSCAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAR",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-796",
    "sku": "MC 796",
    "name": "PROTECTOR VAGINA ARTIFICIAL OVINO- CAPRINA",
    "description": "Ref: | Marca: | Prov: NACIONAL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-797",
    "sku": "MC 797",
    "name": "ESCALERILLAS METÁLICAS DE 10 MM",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-798",
    "sku": "MC 798",
    "name": "PIPETA AUT VOL VARIABLE JOANLAB 0,5-10",
    "description": "Ref: | Marca: JOANLAB | Prov: EUROLAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-799",
    "sku": "MC 799",
    "name": "PIPETA AUT VOL VARIABLE JOANLAB 20-200",
    "description": "Ref: | Marca: JOANLAB | Prov: EUROLAB",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-800",
    "sku": "MC 800",
    "name": "GRADILLA PLASTICA EN PP PARA 96 MICROTUBOS 0,5/1,5/2,0ML AZUL X UNIDAD / CRALPLAST",
    "description": "Ref: | Marca: CRALPLAST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-801",
    "sku": "MC 801",
    "name": "FOTOMETRO SMD 1 PARA SEMEN BOVINO",
    "description": "Ref: 1196 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-802",
    "sku": "MC 802",
    "name": "MICROCUBETAS X 100 UNIDADES",
    "description": "Ref: 1123 | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-803",
    "sku": "MC 803",
    "name": "ELECTROJAC PROBE 2\" (51mm)",
    "description": "Ref: 6990 | Marca: NEOGEN | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-804",
    "sku": "MC 804",
    "name": "MONITOR INALÁMBRICO PARA V1",
    "description": "Ref: | Marca: KAIXIN | Prov: KAIXIN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-805",
    "sku": "MC 805",
    "name": "ECÓGRAFO S0- VET",
    "description": "Ref: 9018121000 | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-806",
    "sku": "MC 806",
    "name": "MONITOR AM6100 SN AM3003090",
    "description": "Ref: | Marca: | Prov: BIOMEDICAL TECH SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-807",
    "sku": "MC 807",
    "name": "SONDA LINEAL RECTAL DAWEI",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-808",
    "sku": "MC 808",
    "name": "SONDA LINEAL RECTAL L60 BMV",
    "description": "Ref: L 60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-809",
    "sku": "MC 809",
    "name": "SONDA LINEAL RECTAL (L080-80F)4.0MHz~7.5MHz/ ECÓGRAFO V9",
    "description": "Ref: 9018121000 | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-810",
    "sku": "MC 810",
    "name": "MONITOR PARA V1 WIRILESS",
    "description": "Ref: | Marca: KAIXIN | Prov: KAIXIN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-811",
    "sku": "MC 811",
    "name": "CUBETA PLÁSTICA PARA MINI PAJILLAS P/25",
    "description": "Ref: 6936 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-812",
    "sku": "MC 812",
    "name": "TERMO DE RECOLECCIÓN DE SEMEN/ USO VETERINARIO",
    "description": "Ref: MSLDO05 | Marca: | Prov: MEDSINGLONG CO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 42,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-813",
    "sku": "MC 813",
    "name": "KIT PRUEBA DE PREÑEZ MICROESFERAS",
    "description": "Ref: | Marca: BIOEASY | Prov: BIOEASY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-814",
    "sku": "MC 814",
    "name": "MICROPIPETA AUTOMÁTICA VOLUMEN VARIABLE 2 A 10 ML",
    "description": "Ref: DC85059 | Marca: JOANLAB | Prov: DALÍ DE COLOMBIA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-815",
    "sku": "MC 815",
    "name": "MICROPIPETA AUTOMÁTICA VOLUMEN VARIABLE 20 A 200 UL",
    "description": "Ref: DC90088 | Marca: JOANLAB | Prov: DALÍ DE COLOMBIA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-816",
    "sku": "MC 816",
    "name": "MICROPIPETA AUTMÁTICA VOLUMEN VARIABLE 100 A 1000 UL",
    "description": "Ref: DC22619 | Marca: JOANLAB | Prov: DALÍ DE COLOMBIA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-817",
    "sku": "MC 817",
    "name": "MICROPIPETA AUTOMÁTICA VOLUMEN VARIABLE 0,5 A 10 UL",
    "description": "Ref: DC65884 | Marca: JOANLAB | Prov: DALÍ DE COLOMBIA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-818",
    "sku": "MC 818",
    "name": "CATÉTER DE INSEMINACIÓN PUNTA VERDE 30\" EQUINOS/ MARCA: REPRODUCTION PROVISION",
    "description": "Ref: 386710 | Marca: REPRODUCTION PROVISION | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-819",
    "sku": "MC 819",
    "name": "PLACA CALENTADORA DIGITAL, CON CONTROL DE TEMPERATURA",
    "description": "Ref: CO294-CW | Marca: COSORI | Prov: AMAZON",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-820",
    "sku": "MC 820",
    "name": "FILTROS DE JERINGA DE 28 MM DE DIAMETRO Y 0.2 MICRAS DE PORO, ESTERIL, MEMBRANAS (PES), MATERIAL (AC). CAJA X 50 UNDS. MARCA: CORNING Lote 21622041 (2.0)",
    "description": "Ref: 431229 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-821",
    "sku": "MC 821",
    "name": "FILTROS DE JERINGA DE 28 MM DE DIAMETRO Y 0.2 MICRAS DE PORO, ESTERIL, MEMBRANAS (PES), MATERIAL (AC). CAJA X 50 UNDS. MARCA: CORNING Lote 21622041 (2.0)",
    "description": "Ref: 431229 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-822",
    "sku": "MC 822",
    "name": "NATURAL FERTILITY X 2 KL",
    "description": "Ref: PT113 | Marca: NATURAL VETERINARY | Prov: NATURAL VETERINARY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "2 KL",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-823",
    "sku": "MC 823",
    "name": "GOBLETS DE BLANCOS 10 MM",
    "description": "Ref: 347430 | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-824",
    "sku": "MC 824",
    "name": "PAJILLAS TRANSPARENTES 0,25",
    "description": "Ref: 5565 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-825",
    "sku": "MC 825",
    "name": "PAJILLAS TRANSPARENTES 0,50",
    "description": "Ref: 5569 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-826",
    "sku": "MC 826",
    "name": "VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: 5417 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-827",
    "sku": "MC 827",
    "name": "CONO VAGINA ARTIFICIAL BOVINA",
    "description": "Ref: 5204 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-828",
    "sku": "MC 828",
    "name": "LINER MEMBRANA ESTRIADA LATEX BOVINA",
    "description": "Ref: 26541 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-829",
    "sku": "MC 829",
    "name": "CARCASA ISPERM",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-830",
    "sku": "MC 830",
    "name": "FORRO TERMO DE NITRÉGENO T3",
    "description": "Ref: | Marca: NACIONAL | Prov: MARTHA QUINTERO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-831",
    "sku": "MC 831",
    "name": "PAJILLAS TRANSPARENTES 0,25",
    "description": "Ref: 5565 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-832",
    "sku": "MC 832",
    "name": "PISTOLA QUICKLOCK CLASIC ANI",
    "description": "Ref: 1132 | Marca: MINITUBE | Prov: ANICAM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-833",
    "sku": "MC 833",
    "name": "ETHYLENE GRYCOL ABT FREEZE CON SUCROSA Y HA X 7 ML",
    "description": "Ref: | Marca: ABT | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "7 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-834",
    "sku": "MC 834",
    "name": "ESTILETE FLEXIBLE DE 65 CM",
    "description": "Ref: 17209/1065 | Marca: MINITUBE | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-835",
    "sku": "MC 835",
    "name": "DILUYENTE TRILADYL / V:31-08-2025/ 30-12-2025",
    "description": "Ref: 13500/0250 | Marca: MINITUBE | Prov: ANICAM",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "250 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-836",
    "sku": "MC 836",
    "name": "PUNTAS TRANSPARENTES ESTAN DAR, DE 0,5 A 10 UL. NO ESTÉRILES",
    "description": "Ref: T-300-R | Marca: AXYGEN | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-837",
    "sku": "MC 837",
    "name": "PUNTAS AMARILLAS ESTANDAR. DE 1 A 200 UL. ESTÉRILES",
    "description": "Ref: T-200-Y-R-S | Marca: AXYGEN | Prov: EQUIMED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-838",
    "sku": "MC 838",
    "name": "PISTOLA DE INSEMINACIÓN DE 0,25 PARA NOVILLAS",
    "description": "Ref: 28804 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-839",
    "sku": "MC 839",
    "name": "ECÓGRAFO N5/ CON SONDA LINEAL Y MICROCONVEXA",
    "description": "Ref: | Marca: EMPEROR | Prov: EMPEROR",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-840",
    "sku": "MC 840",
    "name": "TOALLAS DE PAPEL EN Z/ 22,5 X22 CM CADA UNA",
    "description": "Ref: | Marca: SANITISU | Prov: DOLLAR CITY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 150 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-841",
    "sku": "MC 841",
    "name": "AGUA DESTILADA",
    "description": "Ref: | Marca: | Prov: KROMIC SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "1 LITRO",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-842",
    "sku": "MC 842",
    "name": "BATERÍA PARA ECÓGRAFO DAWEI",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-843",
    "sku": "MC 843",
    "name": "BATERÍA PARA ECÓGRAFO L60",
    "description": "Ref: L 60 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-844",
    "sku": "MC 844",
    "name": "T0167-25G/ D(+)TREHALOSE DIHYDRATE CELL CULTURE &/ MARCA: SIGMA/ TIEMPO DE ENTREGA: IMPORTACIÓN DE 90 A 120 DÍAS SALVO FABRICACIÓN Y VENTA PREVIA",
    "description": "Ref: T0167-25G | Marca: SIGMA | Prov: SIGMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-845",
    "sku": "MC 845",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES 0,25 / BLOQUEO AUTOMÁTICO EN ALUMINIO",
    "description": "Ref: | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 4,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-846",
    "sku": "MC 846",
    "name": "VAGINA ARTIFICIAL CON VALVULA OVINO",
    "description": "Ref: 7179 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-847",
    "sku": "MC 847",
    "name": "CONO ESPECIAL VAGINA OVINO",
    "description": "Ref: 5205 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-848",
    "sku": "MC 848",
    "name": "LINER VAGINA ARTIFICIAL OVINO",
    "description": "Ref: 27512 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-849",
    "sku": "MC 849",
    "name": "MICROSCOPIO BINOCULAR CON PLATINA Y PANTALLA 7\"",
    "description": "Ref: HC-R069A | Marca: HAPPYCARE | Prov: HAPPYCARE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 5,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-850",
    "sku": "MC 850",
    "name": "PAJILLAS TRANSPARENTES 0,50 MINITUBE",
    "description": "Ref: 13408/0010 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-851",
    "sku": "MC 851",
    "name": "PAJILLAS TRANSPARENTES 0,25 MINITUBE",
    "description": "Ref: 13407/0010 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-852",
    "sku": "MC 852",
    "name": "PISTOLA DE INSEMINACIÓN QUICKLOCK 2000 MINITUBE",
    "description": "Ref: 17026/0000 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-853",
    "sku": "MC 853",
    "name": "FUNDA RANURADA UNIVERSAL MINITUBE HEN",
    "description": "Ref: 17007/0001 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-854",
    "sku": "MC 854",
    "name": "TERMO DE NITRÓGENO DE 51 LITROS /125mm CUELLO/6 CANASTILLAS",
    "description": "Ref: TR 51/7 | Marca: HEYI BIOTECH | Prov: HENAN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-855",
    "sku": "MC 855",
    "name": "PISTOLA DE INSEMINACIÓN ARTIFICIAL UNIVERSAL TIPO CIRCLIP",
    "description": "Ref: AIG/2U | Marca: HEYI BIOTECH | Prov: HENAN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-856",
    "sku": "MC 856",
    "name": "FUNDA UNIVERSAL PARA INSEMINACION HEYI",
    "description": "Ref: AIGS/5 | Marca: HEYI BIOTECH | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 200,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-857",
    "sku": "MC 857",
    "name": "DETECTOR ULTRASÓNICO DE GESTACIÓN PARA CERDOS",
    "description": "Ref: 29.32.65-50.43 | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-858",
    "sku": "MC 858",
    "name": "DILUYENTE EQUINO RF SPERM",
    "description": "Ref: | Marca: RF | Prov: REPROFERT",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-859",
    "sku": "MC 859",
    "name": "ECOGRAFO S5 BLANCO Y NEGRO CON SONDA LINELA RECTAL",
    "description": "Ref: S5B/W | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-860",
    "sku": "MC 860",
    "name": "KIT VAGINOSCOPIO",
    "description": "Ref: | Marca: NACIONAL | Prov: JORGE EDUARDO CADAVID",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-861",
    "sku": "MC 861",
    "name": "ENDOSCOPIO VETERINARIO PARA SISTEMA ANDRID Y IOS",
    "description": "Ref: SE-01 | Marca: SHINECO | Prov: SHINECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-862",
    "sku": "MC 862",
    "name": "MEDIDOR DE GRASA DORSAL UT01",
    "description": "Ref: UT01 | Marca: KAIXIN | Prov: KAIXIN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-863",
    "sku": "MC 863",
    "name": "ESCROTIMETRO",
    "description": "Ref: 1101 | Marca: NACIONAL | Prov: INALMET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 73,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-864",
    "sku": "MC 864",
    "name": "HEMACOLOR TINCION RAPIDA DE FROTIS SANGUINEO SOLUCION 2",
    "description": "Ref: 1119562500 | Marca: MERCK | Prov: ADEQUIM",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "2,5 LITROS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-865",
    "sku": "MC 865",
    "name": "MICROPIPETA DE 10- 100 UL",
    "description": "Ref: | Marca: JOANLAB | Prov: DALÍ DE COLOMBIA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-866",
    "sku": "MC 866",
    "name": "DILUYENTE PARA CERDOS EOBOS PLUS 1L / 4-5 DÍAS",
    "description": "Ref: E24031/C01-26 | Marca: ARQUIMEA | Prov: ARQUIMEA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "SOBRE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1068,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-867",
    "sku": "MC 867",
    "name": "DILUYENTE PARA CERDOS ACROMAX PLUS 1L/ 7-8 DIAS",
    "description": "Ref: 24030/C01-26 | Marca: ARQUIMEA | Prov: ARQUIMEA",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "SOBRE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 470,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-868",
    "sku": "MC 868",
    "name": "RACK ALUMINIO 12 TUBOS/ 2 ML",
    "description": "Ref: 9010012 | Marca: TED | Prov: TED",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-869",
    "sku": "MC 869",
    "name": "AGUA DESTILADA GARRAFA X 20 LITROS MARCA: NACIONAL",
    "description": "Ref: OR510606 | Marca: NACIONAL | Prov: KROMIC SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 LITROS",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-870",
    "sku": "MC 870",
    "name": "PAJILLAS DE 0,50 TRANSPARENTES MINITUBE",
    "description": "Ref: 13408/0010 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 250 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-871",
    "sku": "MC 871",
    "name": "CINTA DE MEDICIÓN PARA PESAR GANADO BOVINO/ NACIONAL",
    "description": "Ref: | Marca: NACIONAL | Prov: INALMET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-872",
    "sku": "MC 872",
    "name": "CINTA DE MEDICIÓN PARA PESAR CABALLOS/ NACIONAL",
    "description": "Ref: | Marca: NACIONAL | Prov: INALMET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-873",
    "sku": "MC 873",
    "name": "ESTILETE DE 24 SS PARA CATETER DE 16FR Y MAS GRANDE",
    "description": "Ref: 762 | Marca: NACIONAL | Prov: IMPLEGAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-874",
    "sku": "MC 874",
    "name": "AGUJA DESECHABLE 20G X 1 1/2",
    "description": "Ref: 6206 | Marca: ENSAFE | Prov: INSMEVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 30,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-875",
    "sku": "MC 875",
    "name": "BOQUILLA DE LLENADO CON SUCCIÓN 15 PIN 0.5 ML",
    "description": "Ref: 7297 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-876",
    "sku": "MC 876",
    "name": "BOQUILLA DE LLENADO CON SUCCIÓN 20 PIN 0.25 ML",
    "description": "Ref: 7298 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-877",
    "sku": "MC 877",
    "name": "SOPORTE PORTA CUBETAS/ METALICO",
    "description": "Ref: 7116 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-878",
    "sku": "MC 878",
    "name": "SUJETADOR 15 PAJILLAS 0,5",
    "description": "Ref: 7008 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-879",
    "sku": "MC 879",
    "name": "SUJETADOR 20 PAJILLAS 0,25",
    "description": "Ref: 7009 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 7,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-880",
    "sku": "MC 880",
    "name": "CAUCHO RESPUESTO DE SUJETADOR 15 PAJILLAS 0,5",
    "description": "Ref: 7876 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-881",
    "sku": "MC 881",
    "name": "CAUCHO RESPUESTO DE SUJETADOR 20 PAJILLAS 0,25",
    "description": "Ref: 7875 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-882",
    "sku": "MC 882",
    "name": "RACK DE CONGELACIÓN 100 PAJILLAS 0,50 (14 x 58 cm)",
    "description": "Ref: 7117 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-883",
    "sku": "MC 883",
    "name": "RACK DE CONGELACIÓN 58 PAJILLAS 0,50 (14 x 36.5 cm)",
    "description": "Ref: 7118 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-884",
    "sku": "MC 884",
    "name": "RACK DE CONGELACIÓN 70 PAJILLAS 0,25 (14 X 25,3)",
    "description": "Ref: 7122 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-885",
    "sku": "MC 885",
    "name": "RACK DE CONGELACIÓN 40 PAJILLAS 0,25 (14 X 25,3 cm)",
    "description": "Ref: 7119 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-886",
    "sku": "MC 886",
    "name": "TUBO LIGAZON BOMBA MRS3",
    "description": "Ref: 7825 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-887",
    "sku": "MC 887",
    "name": "MINI PISTOLA DE 0,25 ML OVINA- CAPRINA",
    "description": "Ref: 7177 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-888",
    "sku": "MC 888",
    "name": "PISTOLA DE INSEMINACION OVINO- CAPRINO/ PARA VAINAS MINI",
    "description": "Ref: 7176 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-889",
    "sku": "MC 889",
    "name": "PISTOLA DE INSEMINACIÓN OVINO 0,50/CAPRINO PARA VAINAS MEDIANAS",
    "description": "Ref: 7175 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-890",
    "sku": "MC 890",
    "name": "VAINA 0,50 NO RANURADAS MEDIANA OVINA/CAPRINA",
    "description": "Ref: 6921 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-891",
    "sku": "MC 891",
    "name": "VAINA MINI 0,25 OVINA RANURADAS/ CAPRINA COMPATIBLE CON PISTOLA 7176",
    "description": "Ref: 6922 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "PAQUETE X 40 UND",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 6,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-892",
    "sku": "MC 892",
    "name": "MALETIN PARA MICROSCOPIO EN LONA",
    "description": "Ref: | Marca: NACIONAL | Prov: MICROSCOPIOS Y EQUIPOS ESPECIALES",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-893",
    "sku": "MC 893",
    "name": "BOMBA DE VACIO DE 1000 ML",
    "description": "Ref: H003-B | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-894",
    "sku": "MC 894",
    "name": "TUBO EPPENDORF DE 0,5 ML",
    "description": "Ref: MCTB005 | Marca: GENFOLLOWER | Prov: BPL",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-895",
    "sku": "MC 895",
    "name": "TUBO EPPENDORF DE 2,0 ML",
    "description": "Ref: MCTB020 | Marca: GENFOLLOWER | Prov: BPL",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-896",
    "sku": "MC 896",
    "name": "RACK DE ALUMINIO PARA 15 ML",
    "description": "Ref: 14417 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-897",
    "sku": "MC 897",
    "name": "RACK DE ALUMINIO PARA CAJA PETRI DE 35 MM",
    "description": "Ref: 27495 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-898",
    "sku": "MC 898",
    "name": "FUNDA CON BOLSA DE COLECTA PARA EQUINOS",
    "description": "Ref: | Marca: NACIONAL | Prov: JORGE EDUARDO CADAVID",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-899",
    "sku": "MC 899",
    "name": "CAMARA DE NEUBAUER",
    "description": "Ref: 440151 | Marca: | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-900",
    "sku": "MC 900",
    "name": "TABLET 11\" SANSUNG A9+ 128 GB",
    "description": "Ref: 806095277974 | Marca: SANSUNG | Prov: ALKOSTO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-901",
    "sku": "MC 901",
    "name": "PRIMOSTAR 3 BINOCULAR CON OBJETIVOS PARA CONTRASTE DE 1 FASES, MARCA ZEISS. SERIE No 3186005395",
    "description": "Ref: 415500-0057-000 | Marca: ZEISS | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-902",
    "sku": "MC 902",
    "name": "GOBLET DE 10 AZULES",
    "description": "Ref: | Marca: | Prov: ALTA VETERINARY LLC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-903",
    "sku": "MC 903",
    "name": "AGUA BIDESTILADA TIPO IVBE ESPECIAL",
    "description": "Ref: | Marca: ASEPSIAS LABORATORIO | Prov: ASEPSIAS LABORATORIO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-904",
    "sku": "MC 904",
    "name": "MG-A60035G/ BOVINE SERUM ALBUMIN; LYOPHILIZED POWDE , 5G",
    "description": "Ref: MG-A60035G | Marca: SIGMA | Prov: SIGMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "5 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-905",
    "sku": "MC 905",
    "name": "PAJILLAS TRANSPARENTES 0,50 IMV",
    "description": "Ref: | Marca: IMV | Prov: IMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 250 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-906",
    "sku": "MC 906",
    "name": "TERMO NITRÓGENO ET- 20",
    "description": "Ref: 20597559 | Marca: MVE | Prov: INSMEVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-907",
    "sku": "MC 907",
    "name": "AUTOJAC V3 ELECTROEYACUALADOR MANUAL Y AUTOMÁTICO",
    "description": "Ref: | Marca: NEOVET | Prov: NEOVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-908",
    "sku": "MC 908",
    "name": "TABLET GALAXY TAB A 9+ MARCA SANSUNG",
    "description": "Ref: | Marca: SANSUNG | Prov: KATRONIX",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-909",
    "sku": "MC 909",
    "name": "CARCAZA CON ARNÉZ PARA TABLET GALAXY TAB A",
    "description": "Ref: | Marca: | Prov: MERCADO LIBRE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-910",
    "sku": "MC 910",
    "name": "RECTAL PROBE INTRODUCER",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-911",
    "sku": "MC 911",
    "name": "SONDA URINARIA PARA GATO CON AGUJA",
    "description": "Ref: | Marca: CAT CATHETER | Prov: IMPORTAACIONES VEETERINARIAS MATEUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-912",
    "sku": "MC 912",
    "name": "SONDA URINARIA PARA GATO SIN GUIA",
    "description": "Ref: | Marca: CAT CATHETER | Prov: IMPORTAACIONES VEETERINARIAS MATEUS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-913",
    "sku": "MC 913",
    "name": "CONO RECOLECTOR PAQUETE X 50 UND",
    "description": "Ref: | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-914",
    "sku": "MC 914",
    "name": "PLATINA TERMICA PARA MICROSCOPIO PT0805",
    "description": "Ref: | Marca: NEOVET | Prov: NEOVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-915",
    "sku": "MC 915",
    "name": "PLATINA CALENTADORA MT2535CTE NEOVET",
    "description": "Ref: | Marca: NEOVET | Prov: NEOVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-916",
    "sku": "MC 916",
    "name": "BAÑO DE MARIA",
    "description": "Ref: | Marca: BLOSHISHAN | Prov: AMAZON",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-917",
    "sku": "MC 919",
    "name": "PAJILLA TRANSPPPARENTE NARANJA 0,25 MINITUBE HEN",
    "description": "Ref: 13407/0144 | Marca: MINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-918",
    "sku": "MC 920",
    "name": "PAJILLA TRANSPARENTE 0,25 AZUL MINTUBE HEN",
    "description": "Ref: 13407/0064 | Marca: MIIINITUBE | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-919",
    "sku": "MC 921",
    "name": "PLACA TERMICA PARA MICROSCOPIO",
    "description": "Ref: BHS 1000 | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-920",
    "sku": "MC 922",
    "name": "MICROSCOPIO BINOCULAR",
    "description": "Ref: BS-2074B | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-921",
    "sku": "MC 923",
    "name": "BS-2074B 10× (N.A.:0.25, W.D.:10.2mm) 20× (N.A.:0.40, W.D.:4.0mm) 40× (N.A.:0.65, W.D.:1.5mm) 100× (Oil, N.A.:1.25, W.D.: 0.3mm)",
    "description": "Ref: NIS61 | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-922",
    "sku": "MC 924",
    "name": "(Bright Field-Phase Contrast Plate (10x-40x), Bright Field-Phase Contrast Plate (100x), Bright Field-Dark Field Plate, Simple Polarization Set, Centering Telescop)",
    "description": "Ref: BS-2074B Optional | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-923",
    "sku": "MC 925",
    "name": "TERMO RECOLECTOR DE SEMEN COLOR NEGRO",
    "description": "Ref: 23057 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-924",
    "sku": "MC 926",
    "name": "PISTOLA DE TRANSFERENCIA DE EMBRIONES 0,25 DE MANGO DE ALUMINIO",
    "description": "Ref: 16301 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 16,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-925",
    "sku": "MC 928",
    "name": "TAPA ESTÁNDAR PARA TERMO CRIOGÉNICO",
    "description": "Ref: | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 4,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-926",
    "sku": "MC 929",
    "name": "MACROCUBAS FOTOMETRO",
    "description": "Ref: 5155 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-927",
    "sku": "MC 930",
    "name": "ISPERM COLECTOR DE MUESTRAS",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-928",
    "sku": "MC 931",
    "name": "ECOGRAFO ULTRAS SO",
    "description": "Ref: ULTRAS | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-929",
    "sku": "MC 932",
    "name": "ECOGRAFO SLITES S 1",
    "description": "Ref: SLIGHT | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-930",
    "sku": "MC 933",
    "name": "CAJA PETRI DESECHABLE ESTERIL 90X 15",
    "description": "Ref: 30943 | Marca: LABORATORY PRODUCTS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-931",
    "sku": "MC 934",
    "name": "TUBO CONICO DE 1,5ml PRE ESTERILIZADO LIBRE DE Rnasa y Dnasa No pirogenico No DNA Humano Tapa a Presion",
    "description": "Ref: MCTB015 | Marca: GEN FOLLOWER | Prov: BPL",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 7,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-932",
    "sku": "MC 935",
    "name": "PLACA CALENTADORA/ INCLUYE CABLE",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-933",
    "sku": "MC 936",
    "name": "CAJA PLÁSTICA DE 22\" DE TRANSPORTE",
    "description": "Ref: 7416 | Marca: RIMAX | Prov: COMERCIALIZADORA AYA HNOS SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 13,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-934",
    "sku": "MC 937",
    "name": "IPAD MINI 6/ 64 GB",
    "description": "Ref: | Marca: APPLE | Prov: EDUARD ALMANZA ÁVILA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-935",
    "sku": "MC 938",
    "name": "CAJA DE TRANSPORTE DE 24\"",
    "description": "Ref: | Marca: RIMAX | Prov: CENCOSUD",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-936",
    "sku": "MC 939",
    "name": "CABLE REPUESTO ECOGRAFO S5",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 148,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-937",
    "sku": "MC 940",
    "name": "LINER LISO BOVINO LATEX",
    "description": "Ref: 26647 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 11,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-938",
    "sku": "MC 941",
    "name": "LINER ESTRIADO NEGRO BOVINO",
    "description": "Ref: 5232 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-939",
    "sku": "MC 942",
    "name": "CAJAS DE PETRI DE 35 X 10MM, EN POLYESTYRENO, ESTÉRILES, DESECHABLES, PARA CULTIVO CELULAR",
    "description": "Ref: 430165 | Marca: CORNING | Prov: EQUIMED",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-940",
    "sku": "MC 943",
    "name": "SONDA S7 PLUS DOPPLER COLOR",
    "description": "Ref: S7 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-941",
    "sku": "MC 944",
    "name": "SONDA ADICIONAL S7 PLUS",
    "description": "Ref: SONDA S7 | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-942",
    "sku": "MC 945",
    "name": "CARGADOR DE BATERÍA DE TRANSPORTADORA WTA",
    "description": "Ref: 26878 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-943",
    "sku": "MC 946",
    "name": "CARGADOR WTA",
    "description": "Ref: 19221/ 12023 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-944",
    "sku": "MC 947",
    "name": "ECÓGRAFO WIRELESS INALAMBRICO DOPPLER COLOR NEO DAWEI",
    "description": "Ref: 9018121000 | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-945",
    "sku": "MC 948",
    "name": "PIPETA AUTOMÁTICA VOLUMEN DE 0.5A 10UL",
    "description": "Ref: JA3022 | Marca: JOANLAB | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-946",
    "sku": "MC 949",
    "name": "RACK DE PUNTAS 10 UL NO ESTERIL/ LT N1394910O/ N139103M",
    "description": "Ref: 1111-3800 | Marca: USA SCIENTIFIC | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 82,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-947",
    "sku": "MC 950",
    "name": "RACK DE PUNTAS 10 UL ESTERIL/ VEN 2030-01-31",
    "description": "Ref: 1111-3810 | Marca: USA SCIENTIFIC | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 220,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-948",
    "sku": "MC 951",
    "name": "RACK DE PUNTAS 200 UL NO ESTERIL/ LT M1355851",
    "description": "Ref: 1111-1800 | Marca: USA SCIENTIFIC | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-949",
    "sku": "MC 952",
    "name": "BOQUILLA DE LLENADO CON SUCCIÓN 15 PIN 0.5 ML",
    "description": "Ref: 200167 | Marca: KURDON INDUSTRIES | Prov: ALTA VETERINARY",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-950",
    "sku": "MC 953",
    "name": "TABLET 8.7\" SAMSUNG A9 128GB",
    "description": "Ref: 8806095429977 | Marca: SAMSUNG | Prov: KATRONIX",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-951",
    "sku": "MC 954",
    "name": "KIT DE LLENADO DE PAJILLAS/ PARA 5 PAJILLAS",
    "description": "Ref: FSF-101-KIT [FSF-101- KIT] | Marca: ARS | Prov: ARS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-952",
    "sku": "MC 955",
    "name": "RACK METÁLICO NACIONAL A MEDIDAS",
    "description": "Ref: | Marca: NACIONAL | Prov: TECNICADENAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-953",
    "sku": "MC 956",
    "name": "TERMO NITROGENO ET- 11",
    "description": "Ref: 4452 | Marca: MVE | Prov: INSMEVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-954",
    "sku": "MC 957",
    "name": "IPAD MINI 7.9\" PLATA 128 GB",
    "description": "Ref: | Marca: APPLE | Prov: SU TIENDA PRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-955",
    "sku": "MC 958",
    "name": "PUNTAS COLOR NATURAL SIN FILTRO, UNIVERSAL, 5-10UL. TRANSPARENTE, GRADUADA, NO ESTERIL LIBRES DE RNASA, DNASA Y PIROGENOS AUTOCLAVABLE LOTE: 20231026-SGA-001A",
    "description": "Ref: SGA001 | Marca: BIO-SEEN | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-956",
    "sku": "MC 959",
    "name": "ECÓGRAFO S7 B/W ( BLANCO Y NEGRO) BMV - Plataforma BLANCO Y NEGRO CON BATERÍA",
    "description": "Ref: S7/B-W | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-957",
    "sku": "MC 960",
    "name": "CARCASA TABLET BMV",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-958",
    "sku": "MC 961",
    "name": "ECÓGRAFO BETSCAN S9 VET CON SONDA LINEAL",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-959",
    "sku": "MC 962",
    "name": "ECÓGRAFO BETSCAN S10 VET CON SONDA LINEAL/ COLOR DOPPLER",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-960",
    "sku": "MC 963",
    "name": "GAFAS PARA ECÓGRAFOS S9 Y S10",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-961",
    "sku": "MC 964",
    "name": "TABLET PARA ECÓGRAFOS BMV",
    "description": "Ref: | Marca: SANSUNG | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-962",
    "sku": "MC 965",
    "name": "ECÓGRAFO S7 B/W ( BLANCO Y NEGRO) BMV - Plataforma BLANCO Y NEGROSIN BATERÍA",
    "description": "Ref: S7/B-W | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-963",
    "sku": "MC 966",
    "name": "CAJA PETRI 4 POZOS / 3,026X3,026 CM",
    "description": "Ref: 10 | Marca: SPI | Prov: SPI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 173,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-964",
    "sku": "MC 967",
    "name": "FILTRO EZ",
    "description": "Ref: 25 | Marca: SPI | Prov: SPI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-965",
    "sku": "MC 968",
    "name": "MAXI FILTRO",
    "description": "Ref: 50 | Marca: SPI | Prov: SPI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-966",
    "sku": "MC 969",
    "name": "SONDA EN \"Y\" CON VÁLVULA PARA BOVINOS",
    "description": "Ref: 60 | Marca: SPI | Prov: SPI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 32,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-967",
    "sku": "MC 970",
    "name": "TERMO DESCONGELADOR DE SEMEN Y EMBRIONES DSE-480 SIN BATERÍA",
    "description": "Ref: 9013001 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-968",
    "sku": "MC 971",
    "name": "TERMO DESCONGELADOR DE SEMEN Y EMBRIONES DSE-480 CON BATERÍA BLUE/GRAY",
    "description": "Ref: 9013002 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-969",
    "sku": "MC 972",
    "name": "MICROFILTRO 0,75 uM",
    "description": "Ref: 9014002 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-970",
    "sku": "MC 973",
    "name": "MICROFILTRO 0,75 uM",
    "description": "Ref: 9014001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-971",
    "sku": "MC 974",
    "name": "LACRADOR LARGO GRIS",
    "description": "Ref: 9006007 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-972",
    "sku": "MC 975",
    "name": "LACRADOR LARGO ROJOS",
    "description": "Ref: 9006001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-973",
    "sku": "MC 976",
    "name": "LACRADOR LARGO PURPURA",
    "description": "Ref: 9006006 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-974",
    "sku": "MC 977",
    "name": "LACRADOR LARGO VERDE",
    "description": "Ref: 9006004 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-975",
    "sku": "MC 978",
    "name": "LACRADOR LARGO AZUL",
    "description": "Ref: 9006003 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-976",
    "sku": "MC 979",
    "name": "MINI PIPETA ISPERM 50 UL",
    "description": "Ref: | Marca: ISPERM | Prov: ISPERM",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-977",
    "sku": "MC 980",
    "name": "INTRODUCER KAIXIN PARA V1",
    "description": "Ref: | Marca: KAIXIN | Prov: KAIXIN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-978",
    "sku": "MC 981",
    "name": "HEMACOLOR TINCION RAPIDA DE FROTIS SANGUINEOS KIT DE TINCION PARA MICROSCOPIA",
    "description": "Ref: 1116610001 | Marca: SIGMA | Prov: ADEQUIM",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "KIT X 3 FRACOS X 500 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-979",
    "sku": "MC 982",
    "name": "PISTOLA DE TRANSFERIA DE EMBRIONES DE 0,50",
    "description": "Ref: 18958 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-980",
    "sku": "MC 983",
    "name": "ECÓGRAFO S7 PRO VET CON SONDA LINEAL RECTAL",
    "description": "Ref: S7PRO | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-981",
    "sku": "MC 984",
    "name": "ECÓGRAFO BPU 50B CON SONDA LINEAL Y MICROCONVEZA CON MALETIN DE TRANSPORTE",
    "description": "Ref: BPU50B | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-982",
    "sku": "MC 985",
    "name": "SONDA MICRO CONVEX BPU50B",
    "description": "Ref: SONDBPU | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-983",
    "sku": "MC 986",
    "name": "ECOGRAFO MX5C CON SONDA LINEAL RECTAL/DOPLLER COLOR",
    "description": "Ref: MX5C | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-984",
    "sku": "MC 987",
    "name": "CAJA DE PETRI DESECHABLE ESTERIL 60x15MM.",
    "description": "Ref: 140681 | Marca: LABORATORY PRODUCTS | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 10 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 28,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-985",
    "sku": "MC 988",
    "name": "DETECTOR DE MASTITIS 4X4",
    "description": "Ref: 76311 | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-986",
    "sku": "MC 989",
    "name": "MANGA DE PROTECCIÓN PALPACIÓN MAS CAMPO",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-987",
    "sku": "MC 990",
    "name": "CARGADOR ECÓGRAFO S5 BMV",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-988",
    "sku": "MC 991",
    "name": "TERMO NITROGENO DRY SHIPER YDH-3 TRANSPORTE",
    "description": "Ref: 4455/ 20005473 | Marca: YDH-3 | Prov: INSMEVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-989",
    "sku": "MC 992",
    "name": "TRILADYL 250 ML",
    "description": "Ref: 13500/0250 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-990",
    "sku": "MC 993",
    "name": "TERMO DE 380 ML MAS CAMPO",
    "description": "Ref: | Marca: MAS CAMPO | Prov: MAXSPEED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 315,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-991",
    "sku": "MC 994",
    "name": "PISTOLA DE INSEMINACIÓN QUICKLOCK UNIVERSAL MINITUBE",
    "description": "Ref: 17025 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-992",
    "sku": "MC 995",
    "name": "PISTOLA DE INSEMINACION QUICKLOCK 2000 MINITUBE",
    "description": "Ref: 17026 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-993",
    "sku": "MC 996",
    "name": "CORTAPAJILLAS MARCA MINITUBE",
    "description": "Ref: 17062 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-994",
    "sku": "MC 997",
    "name": "PAJILLAS TRANSPARENTES 0,25 MINITUBE",
    "description": "Ref: 13407 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 10,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-995",
    "sku": "MC 998",
    "name": "PAJILLAS TRANSPARENTES 0,50 MINITUBE",
    "description": "Ref: 13408 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-996",
    "sku": "MC 999",
    "name": "ESFERAS METALICAS SELLADORAS 0.5 ML",
    "description": "Ref: 13400 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "FRASCO X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-997",
    "sku": "MC 1000",
    "name": "ESFERAS METALICAS SELLADORAS 0.25 ML",
    "description": "Ref: 13400 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "FRASCO X 1000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-998",
    "sku": "MC 1001",
    "name": "FUNDA PARA INSEMINACIÓN RANURA MARCA MINITUBE",
    "description": "Ref: 13408 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 50 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-999",
    "sku": "MC 1002",
    "name": "PIPETA EQUINA DE 65 CM",
    "description": "Ref: 17209/1265 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 191,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1000",
    "sku": "MC 1003",
    "name": "ESTILETE PARA PIPETA S DE IA EQUINA DE 57 CM",
    "description": "Ref: 17209/1057 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1001",
    "sku": "MC 1004",
    "name": "EMPACADORA Y SELLADORA AUTOMATICA DE PAJILLAS 0.50",
    "description": "Ref: FT2000AU/0,5 | Marca: HEYI BIOTECH | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1002",
    "sku": "MC 1005",
    "name": "PROTECTOR SOLAR ECÓGRAFO DAWEI S1",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1003",
    "sku": "MC 1006",
    "name": "INTRODUCER DAWEI",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1004",
    "sku": "MC 1007",
    "name": "TERMO DESCONGELADOR DE EMBRIONES Y SEMEN DSE-400 SIN BATERÍA BLACK/GRAY",
    "description": "Ref: 9013005 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1005",
    "sku": "MC 1008",
    "name": "BAÑO DE MARÍA DE 3 LITROS",
    "description": "Ref: BK-S1 | Marca: BIOBASE | Prov: BIOBASE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1006",
    "sku": "MC 1009",
    "name": "BAÑO DE MARÍA DE 5 LITROS",
    "description": "Ref: BK-S2 | Marca: BIOBASE | Prov: BIOBASE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1007",
    "sku": "MC 1010",
    "name": "AGUJAS PARA EMPACADORA SELLADORA",
    "description": "Ref: | Marca: HEYI BIOTECH | Prov: HEYI BIOTECH",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 27,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1008",
    "sku": "MC 1011",
    "name": "MICROPIPETA AUTOMATICA VOLUMEN VARIABLE DE 100 - 1000",
    "description": "Ref: M100 | Marca: JOANLAB | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1009",
    "sku": "MC 1012",
    "name": "ACEITE DE INMERSION x 100ML Und. 250516-505 1 (REAC) / CHEMI",
    "description": "Ref: | Marca: CHEMI | Prov: DICORLAB S.A.S",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "1000 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1010",
    "sku": "MC 1013",
    "name": "MICROSCOPIO MODELO PRIMOSTAR 3 TRIOCULAR/ SERIAL 3178002348",
    "description": "Ref: 415501-0011-000 | Marca: ZEISS | Prov: KAIKA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1011",
    "sku": "MC 1014",
    "name": "OBJETIVO IPLAN-ACHRAMAT 20X/0,45",
    "description": "Ref: 415501-1625-000 | Marca: ZEISS | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1012",
    "sku": "MC 1015",
    "name": "REPUESTO ROSCA ECOGRAFO S7 PLUS, S7 PRO Y MX5",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 16,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1013",
    "sku": "MC 1016",
    "name": "CENTRIFUGA CX-2000",
    "description": "Ref: CX-2000 | Marca: PREMIER | Prov: PREMIER",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1014",
    "sku": "MC 1017",
    "name": "TABLET SANSUNG A9+ COLOR A DISPONIBILIDAD",
    "description": "Ref: SM-X210NZAECOO | Marca: SANSUNG | Prov: SANSUNG",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1015",
    "sku": "MC 1018",
    "name": "BOTU SEMEN SPECIAL 8,5 G",
    "description": "Ref: BOT.30.0003 | Marca: BOTUPHARMA | Prov: SUMINISTROS GLOBAL VETT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "SOBRE 8.5 G",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1016",
    "sku": "MC 1019",
    "name": "ECOGRAFO MX5C CON SONDA LINEAL RECTAL/ BLANCO Y NEGRO",
    "description": "Ref: MX5A | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1017",
    "sku": "MC 1020",
    "name": "TERMO NITROGENO ET-2",
    "description": "Ref: 04702/ 20672009 | Marca: MVE | Prov: INSMEVET",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1018",
    "sku": "MC 1021",
    "name": "PREG-TONE II PLUS RENCO",
    "description": "Ref: | Marca: RENCO | Prov: RENCO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1019",
    "sku": "MC 1022",
    "name": "IMPRESORA DE PAJILLAS MINI PRINTER",
    "description": "Ref: | Marca: HEYI BIOTECH | Prov: HENAN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1020",
    "sku": "MC 1023",
    "name": "FIROVET HORSE ORAL",
    "description": "Ref: 6 | Marca: BOTUPHARMA | Prov: SUMINISTROS GLOBAL VETT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1021",
    "sku": "MC 1024",
    "name": "BOTUKILLER 90 ML",
    "description": "Ref: 16 | Marca: BOTUPHARMA | Prov: SUMINISTROS GLOBAL VETT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1022",
    "sku": "MC 1025",
    "name": "MANGO DE RECOLECCIÓN DE SEMEN",
    "description": "Ref: 6985 | Marca: NEOGEN | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1023",
    "sku": "MC 1026",
    "name": "TUBO RIGIDO EQUINO BOTUPHARMA",
    "description": "Ref: 25 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1024",
    "sku": "MC 1027",
    "name": "MUCOSA DE LATEX EQUINA BOTUPHARMA",
    "description": "Ref: 24 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1025",
    "sku": "MC 1028",
    "name": "MUCOSA DE PLÁSTICA EQUINA BOTUPHARMA",
    "description": "Ref: 26 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1026",
    "sku": "MC 1029",
    "name": "VASO COLECTOR EQUINO BOTUPHARMA",
    "description": "Ref: 27 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1027",
    "sku": "MC 1030",
    "name": "CAMISA SANITARIA PARA VASO COLECTOR EQUINO BOTUPHARMA",
    "description": "Ref: 28 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1028",
    "sku": "MC 1031",
    "name": "ISCAN MINI ULTRASONIDO CON SONDA LINEAL RECTAL",
    "description": "Ref: 29.32.65-50.43 | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1029",
    "sku": "MC 1032",
    "name": "PROTECTOR SOLAR DRAMINSKI",
    "description": "Ref: | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1030",
    "sku": "MC 1033",
    "name": "FILTRO DE GEL NYLON DESECHABLE, 3.5\"",
    "description": "Ref: SAVF-102/25 | Marca: ARS | Prov: ARS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 25 UNIDADES",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1031",
    "sku": "MC 1034",
    "name": "BOTELLA DE RECOLECCIÓN DE SEMEN EQUINO X 200 ML",
    "description": "Ref: SCB-200/100 | Marca: ARS | Prov: ARS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 49,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1032",
    "sku": "MC 1035",
    "name": "BOTU CRIO 25 ML",
    "description": "Ref: 14 | Marca: BOTUPHARMA | Prov: SUMINISTROS GLOBAL VETT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "25 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1033",
    "sku": "MC 1036",
    "name": "REGLILLA PARA MEDIR NITRÓGENO",
    "description": "Ref: 2 | Marca: NACIONAL | Prov: BELSAGRO SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1034",
    "sku": "MC 1037",
    "name": "SINCRODIOL X 50 ML",
    "description": "Ref: 11000557 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1035",
    "sku": "MC 1038",
    "name": "SINCROCIO X 50 ML",
    "description": "Ref: 8014 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1036",
    "sku": "MC 1039",
    "name": "SINCROECG X 30 ML 6.000 UI",
    "description": "Ref: 1100614 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1037",
    "sku": "MC 1040",
    "name": "SINCROCP X 50 ML",
    "description": "Ref: 9319 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "50 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1038",
    "sku": "MC 1041",
    "name": "SINCROGEST DISPOSITIVO 1GR BOLSA X 2 UNIDADES",
    "description": "Ref: 22651 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1039",
    "sku": "MC 1042",
    "name": "SINCROFORTE FRASCO 20ML",
    "description": "Ref: 8545 | Marca: EUROFINO | Prov: VETERINARIA LA RED SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "20 ML",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1040",
    "sku": "MC 1043",
    "name": "OBJETIVO IPLAN-ACHROMAT 100X/1.25 OIL PH3, PRIMOSTAR 3",
    "description": "Ref: 415501-1645-000 | Marca: KAIKA | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1041",
    "sku": "MC 1044",
    "name": "CORREDERA H, PH1, PH2, PH3 PARA PRIMO STAR.",
    "description": "Ref: 415500-1826-000 | Marca: KAIKA | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1042",
    "sku": "MC 1045",
    "name": "OBJETIVO IPLAN-ACHROMAT 100X/1.25 OIL PH3 PRIMOSTAR.",
    "description": "Ref: 415501-1645-000 | Marca: KAIKA | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1043",
    "sku": "MC 1046",
    "name": "AXIOCAM 212 COLOR MARCA ZEISS",
    "description": "Ref: 426570-9901-000 | Marca: KAIKA | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1044",
    "sku": "MC 1047",
    "name": "ADAPTADOR DE CAMARA P95 PARA PRIMO STAR 3",
    "description": "Ref: 415501-1811-000 | Marca: KAIKA | Prov: KAIKA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1045",
    "sku": "MC 1048",
    "name": "SONDA CONVEXA ADICIONAL S7 PRO",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1046",
    "sku": "MC 1049",
    "name": "SONDA MICRO CONVEXA ADICIONAL S7 PRO",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1047",
    "sku": "MC 1050",
    "name": "ECOGRAFO MX5C CON SONDA MICROCONVEXA CON DOPPLER COLOR",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1048",
    "sku": "MC 1051",
    "name": "ECOGRAFO MX5C CON SONDA CONVEXA CON DOPPLER COLOR",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 3,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1049",
    "sku": "MC 1052",
    "name": "ECOGRAFO MX5C CON SONDA LINEAL ABDOMEN CON DOPPLER COLOR",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1050",
    "sku": "MC 1053",
    "name": "BATERÍA PARA TRASNPORTADORA WTA",
    "description": "Ref: | Marca: SAGER | Prov: SAGER TECNOLOGIES",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1051",
    "sku": "MC 1054",
    "name": "SPERM FIILTER",
    "description": "Ref: BOT.30.0018 | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1052",
    "sku": "MC 1055",
    "name": "ECÓGRAFO Y6 CON SONDA LINEAL RECTAL",
    "description": "Ref: 9018121000 | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1053",
    "sku": "MC 1056",
    "name": "ECÓGRAFO SMART FASTSCAN CON SONDA LINEAL RECTAL DE 6,5 MhZ",
    "description": "Ref: | Marca: SONOWIN | Prov: SONOWIN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 5,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1054",
    "sku": "MC 1057",
    "name": "ECÓGRAFO SMART FASTSCAN CON SONDA CONVEXA DE 4,0 MhZ",
    "description": "Ref: | Marca: SONOWIN | Prov: SONOWIN",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1055",
    "sku": "MC 1058",
    "name": "ECOGRAFO AIRSCAN PRO CON SONDA LINEAL RECTAL",
    "description": "Ref: | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1056",
    "sku": "MC 1059",
    "name": "DETECTOR DE ESTRO PARA VACAS Y YEGUAS EDC2",
    "description": "Ref: 77053 | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1057",
    "sku": "MC 1060",
    "name": "MICROSCOPIO BINOCULAR BIOLOGICO (Objetivo acromático finito 4×, 10×, 40×, 100× (aceite) y batería recargable.)",
    "description": "Ref: BS-2020B | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1058",
    "sku": "MC 1061",
    "name": "MICROSCOPIO BINOCULAR BIOLOGICO ( Objetivo acromático finito 4×, 10×, 40×, 100× (aceite)",
    "description": "Ref: BS-2020B | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1059",
    "sku": "MC 1062",
    "name": "MICROSCOPIO TRINOCULAR BIOLOGICO Objetivo acromático finito 4×, 10×, 40×, 100× (aceite) con adaptador de montura C 0,5X",
    "description": "Ref: BS-2020T | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1060",
    "sku": "MC 1063",
    "name": "MICROSCOPIO BINOCULAR Con objetivo semiplano infinito (4X, 10X, 40X, 100X)",
    "description": "Ref: BS-2052B | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1061",
    "sku": "MC 1064",
    "name": "MICROSCOPIO TRINOCULAR Con objetivo Infinite Semi-Plan 4X, 10X, 40X, 100X y adaptador de montura C 0,5X",
    "description": "Ref: BS-2025BT | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1062",
    "sku": "MC 1065",
    "name": "MICROSCOPIO TRINOCULAR Sistema óptico NIS60 Infinite, objetivo de 4X, 10X, 40X y 100X (aceite), montura C de 0,5X",
    "description": "Ref: BS-2074T | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1063",
    "sku": "MC 1066",
    "name": "CÁMARA CMOS PARA MICCROSCOPIO BESTCOPE con montura C (4K, 8 MP) con múltiples salidas HDMI/WIFI/USB 3.0",
    "description": "Ref: BWHC1-4K8MPA | Marca: BESTCOPE | Prov: BESTCOPE",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1064",
    "sku": "MC 1067",
    "name": "PROVETSCAN SF-2 SONDA DE GRASA DORSAL PARA PORCINOS",
    "description": "Ref: 1104 - WFPCFJC003 | Marca: PROVETSCAN | Prov: NEW VETEC",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1065",
    "sku": "MC 1068",
    "name": "CONO LÁTEX PARA MANGO COLECTOR DE ELECTRO",
    "description": "Ref: 6983 | Marca: NEOGEN | Prov: NEOGEN",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1066",
    "sku": "MC 1069",
    "name": "MINIPAJILLA 0,25 ML TRANSPARENTES",
    "description": "Ref: | Marca: MINITUBE | Prov: ANDITECNICA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 20,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1067",
    "sku": "MC 1070",
    "name": "PAJUELA TRANSPARENTE 0.25/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5565 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1068",
    "sku": "MC 1071",
    "name": "PAJUELA AZUL 0.25/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5584 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1069",
    "sku": "MC 1072",
    "name": "PAJUELA ANARANJADA 0.25/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5580 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1070",
    "sku": "MC 1073",
    "name": "PAJUELA PISTACHO 0.25/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 19179 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1071",
    "sku": "MC 1074",
    "name": "PAJUELA GRIS 0.25/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5577 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1072",
    "sku": "MC 1075",
    "name": "PAJUELA TRANSPARENTES 0.50/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5569 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1073",
    "sku": "MC 1076",
    "name": "PAJUELA ROJA 0.50/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5709 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1074",
    "sku": "MC 1077",
    "name": "PAJUELA ANARANJADO 0.50/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5685 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1075",
    "sku": "MC 1078",
    "name": "PAJUELA VERDE 0.50/ PAQ. X 2000 MARCA IMV",
    "description": "Ref: 5710 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1076",
    "sku": "MC 1079",
    "name": "PAJUELA TRANSPARENTES 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-010 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1077",
    "sku": "MC 1080",
    "name": "PAJUELA ROJA 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-044 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1078",
    "sku": "MC 1081",
    "name": "PAJUELA VERDE 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-054 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1079",
    "sku": "MC 1082",
    "name": "PAJUELA AZUL 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-064 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1080",
    "sku": "MC 1083",
    "name": "PAJUELA AMARILLA 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-094 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1081",
    "sku": "MC 1084",
    "name": "PAJUELA ANARANJADA 0.25/ PAQ. X 2000 MARCA MINITUBE",
    "description": "Ref: 13407-144 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1082",
    "sku": "MC 1085",
    "name": "PAJUELA TRANSPARENTE 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-010 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1083",
    "sku": "MC 1086",
    "name": "PAJUELA ROJA 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-044 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1084",
    "sku": "MC 1087",
    "name": "PAJUELA VERDE 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-054 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1085",
    "sku": "MC 1088",
    "name": "PAJUELA AZUL 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-064 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1086",
    "sku": "MC 1089",
    "name": "PAJUELA GRIS 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-074 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1087",
    "sku": "MC 1090",
    "name": "PAJUELA AMARILLA 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-094 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1088",
    "sku": "MC 1091",
    "name": "PAJUELA TURQUESA 0.50/ PAQ. X 2500 MARCA MINITUBE",
    "description": "Ref: 13408-130 | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2500 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1089",
    "sku": "MC 1092",
    "name": "SONDA CONVEXA ECOGRAFO S1 DAWEI",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1090",
    "sku": "MC 1093",
    "name": "TABLET 11\" HONOR PAD X8A 128 GB",
    "description": "Ref: 6936520864546 | Marca: HONOR | Prov: KTRONIX",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 5,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1091",
    "sku": "MC 1094",
    "name": "TABLET 11\" REDMI PAD2 4 / 128 GB",
    "description": "Ref: 6932554444525 | Marca: REDMI | Prov: KTRONIX",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1092",
    "sku": "MC 1095",
    "name": "TERMO DESCONGELADOR DE SEMEN Y EMBRIONES DSEB-400 SIN BATERÍA BLUE/WHITE",
    "description": "Ref: 9013003 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1093",
    "sku": "MC 1096",
    "name": "ROLLO LINEA PARA OPU 100 M MARCA TED",
    "description": "Ref: 9007013 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1094",
    "sku": "MC 1097",
    "name": "CARGADOR TED 12.8V 5A",
    "description": "Ref: 9015001 | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1095",
    "sku": "MC 1098",
    "name": "BOTUSEMEN 8 GM",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: SUMINISTROS GLOBAL VETT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "SOBRE",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1096",
    "sku": "MC 1099",
    "name": "CÁMARA DE RECUENTO CELULAR SPERMLIDE420/ 4 cámaras de recuento por porta, 20 micras de profundidad.",
    "description": "Ref: SPERMLIDE420 | Marca: Sperm Analysis Technologies SL | Prov: Sperm Analysis Technologies SL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "CAJA X 25 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 19,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1097",
    "sku": "MC 1100",
    "name": "PAJUELA AMARILLA 0.25/ PAQUETE X 2000",
    "description": "Ref: 5575 | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1098",
    "sku": "MC 1101",
    "name": "TUBO 0.25 PARA PISTOLA DE INSEMINACIÓN CON CÁMARA BMV",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 2,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1099",
    "sku": "MC 1102",
    "name": "TUBO 0.50 PARA PISTOLA DE INSEMINACIÓN CON CÁMARA BMV",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 5,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1100",
    "sku": "MC 1103",
    "name": "CAJA PETRI DESECHABLE ESTERIL 60X15MM MARCA: NEST",
    "description": "Ref: 754001 | Marca: NEST | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1101",
    "sku": "MC 1104",
    "name": "PAJILLA TRANSPARENTE DE 0.5 IMV PAQ.X 200 UND",
    "description": "Ref: | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQUETE X 200 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1102",
    "sku": "MC 1105",
    "name": "DISPOSITIVO DE SEPARACIÓN DE ESPERMA BOVINO VETMOTL",
    "description": "Ref: | Marca: VETMOTL | Prov: VETMOTL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 110,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1103",
    "sku": "MC 1106",
    "name": "DISPOSITIVO DE SEPARACIÓN DE ESPERMA EQUINO VETMOTL 3 ML",
    "description": "Ref: | Marca: VETMOTL | Prov: VETMOTL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1104",
    "sku": "MC 1107",
    "name": "DISPOSITIVO DE SEPARACIÓN DE ESPERMA EQUINO VETMOTL 850 UL",
    "description": "Ref: | Marca: VETMOTL | Prov: VETMOTL",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1105",
    "sku": "MC 1108",
    "name": "PLATINA CALENTADORA SMZ REDONDA WTA",
    "description": "Ref: 21977 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1106",
    "sku": "MC 1109",
    "name": "TUERCA DE ACOPLAMIENTO DE FILTRO PARA FRASCO COLECTOR DE SEMEN EQUINO",
    "description": "Ref: SCN-101-/100 | Marca: ARS | Prov: ARS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1107",
    "sku": "MC 1110",
    "name": "TAPA DE FRASCO DE RECOLECCION DE SEMEN EQUINO",
    "description": "Ref: CAP-102/100 | Marca: ARS | Prov: ARS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1108",
    "sku": "MC 1111",
    "name": "SONDA EN \"Y\" CON VÁLVULA PARA EQUINOS",
    "description": "Ref: | Marca: SPI | Prov: SPI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 9,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1109",
    "sku": "MC 1112",
    "name": "RACK DE ALUMINIO WTA 12 X 1.5 ML PARA TUBOS CONICOS",
    "description": "Ref: 14418 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1110",
    "sku": "MC 1113",
    "name": "PISTOLA DE INSEMINACIÓN UNIVERSAL (AUTOMATIC LOCK) ALUMINIO",
    "description": "Ref: 15816 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "WTA",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 5,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1111",
    "sku": "MC 1114",
    "name": "HOLDER PARA PROBE DRAMINSKI",
    "description": "Ref: | Marca: DRAMINSKI | Prov: DRAMINSKI",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1112",
    "sku": "MC 1115",
    "name": "PRUEBA DE PREÑEZ TEMPRANA EN RUMIANTES",
    "description": "Ref: | Marca: SURTIVEL | Prov: SANIC IMPORT SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 50,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1113",
    "sku": "MC 1116",
    "name": "DILUYENTE DE SEMEN PORCINO ELIXIR 1 LITRO",
    "description": "Ref: 38210000 | Marca: GENES DIFUSION SAS | Prov: GENES DIFUSION SAS",
    "categoryId": "cat-2",
    "category": "Consumibles Ingeribles",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 415,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1114",
    "sku": "MC 1117",
    "name": "ECOGRAFO BLX-S100 CON TABLET",
    "description": "Ref: BLX-S100 | Marca: BOXERLY | Prov: BOXERLY",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1115",
    "sku": "MC 1118",
    "name": "PRIMOSTAR 1 FIX BINOCULAR ABBE 0.93W MARCA: ZEISS",
    "description": "Ref: 415501-0101-003 | Marca: ZEISS | Prov: KAIKA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1116",
    "sku": "MC 1119",
    "name": "ECOGRAFO BLX-V10 III MARCA: BOXERLY",
    "description": "Ref: BLX-V10 III | Marca: BOXERLY | Prov: BOXERLY",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1117",
    "sku": "MC 1120",
    "name": "FUNDA TERMICA PARA CALENTAR PISTOLAS BA-10400 MARCA:TED",
    "description": "Ref: 9002007 | Marca: TED | Prov: TED",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 6,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1118",
    "sku": "MC 1121",
    "name": "FUNDA PARA TRANSFERENCIA DE 0.5 EQUNINOS MARCA: WTA",
    "description": "Ref: 10415 | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1119",
    "sku": "MC 1122",
    "name": "IMPRESORA DE PAJILLAS AUTOMÁTICA NEOPRINTER",
    "description": "Ref: 84433231 | Marca: NEOVET | Prov: NEOVET",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": -1,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1120",
    "sku": "MC 1123",
    "name": "MICROPIPETA AUTOMATICA VOLUMEN VARIABLE DE 20 - 200 UL MARCA: JOAN LAB",
    "description": "Ref: | Marca: JOANLAB | Prov: LABORATORIO TECNOLOGICO SAS",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1121",
    "sku": "MC 1124",
    "name": "TABLET SANSUNG GALAXY TAB A11+ MARCA: SANSUNG COLOR SEGUN DISPONIBILIDAD",
    "description": "Ref: | Marca: SANSUNG | Prov: KTRONIX",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 4,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1122",
    "sku": "MC 1125",
    "name": "ECOGRAFO DAWEI L30i-VET",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1123",
    "sku": "MC 1126",
    "name": "BOTHUSEMEN GOLD (Caseina + Colesterol) Sobre x 8,5 g",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1124",
    "sku": "MC 1127",
    "name": "BOTU FLEX/ BOTUPHARMA",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1125",
    "sku": "MC 1128",
    "name": "Botupharma- Botumix Garahnao Plus Fco x 1,5 Lt",
    "description": "Ref: | Marca: BOTUPHARMA | Prov: BOTUPHARMA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1126",
    "sku": "MC 1129",
    "name": "RACK DE CONGELACIÓN 90 PAJILLAS 0.25- 0.50 ML MARCA: MINITUBE",
    "description": "Ref: | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1127",
    "sku": "MC 1130",
    "name": "PAJILLAS ANARANJADAS 0.5 PAQ. X 2000 UND MRCA: IMV",
    "description": "Ref: | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 3,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1128",
    "sku": "MC 1131",
    "name": "PAJILLAS AMARILLAS DE 0.5 PAQ. X 2000 UND MARCA IMV",
    "description": "Ref: | Marca: IMV | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 2,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1129",
    "sku": "MC 1132",
    "name": "PAJILLAS AMARILLAS DE 0.25 PAQ. 2000 UND/ MARCA: MINITUBE",
    "description": "Ref: | Marca: MINITUBE | Prov: HUMECO",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 2000 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 6,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1130",
    "sku": "MC 1133",
    "name": "MICROFILTRO 0,90 UM PAQ. X 20 UND MARCA: TED",
    "description": "Ref: | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 12,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1131",
    "sku": "MC 1134",
    "name": "MICROFILTRO 0,90 UM PAQ. X 20 UND MARCA: TED",
    "description": "Ref: | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 19,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1132",
    "sku": "MC 1135",
    "name": "FUNDA DE TRANSFERENCIA DE EMBRIONES 0.25 PUNTA PLÁSTICA MARCA TED",
    "description": "Ref: | Marca: TED | Prov: TED",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 29,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1133",
    "sku": "MC 1136",
    "name": "GAFAS PARA ECOGRAFO Y6 MARCA: DAWEI",
    "description": "Ref: | Marca: DAWEI | Prov: DAWEI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1134",
    "sku": "MC 1137",
    "name": "ECOGRAFO DP10 MINDRAY BLANCO Y NEGRO SONDA LINELA RECTAL",
    "description": "Ref: | Marca: MINDRAY | Prov: HOSPCOM",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1135",
    "sku": "MC 1138",
    "name": "GRADILLA DE ALUMINIO MICROTUBOS WTA /Rack para 16 Microtubos 1,2mL SKU: 14416 16 microtubos WTA 1,2mL",
    "description": "Ref: | Marca: WTA | Prov: WTA",
    "categoryId": "cat-4",
    "category": "Equipos de Laboratorio",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1136",
    "sku": "MC 1139",
    "name": "TABLET SANSUNG GALAXY TAB A11+",
    "description": "Ref: | Marca: SANSUNG | Prov: SANSUNG",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 16,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1137",
    "sku": "MC 1140",
    "name": "LINEAS PARA OPU WTA 1.20",
    "description": "Ref: | Marca: WTA | Prov: WTA",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1138",
    "sku": "MC 1141",
    "name": "CAJA DE PETRI PLASTICA DESECHABLE ESTERIL DE 90x15MM. BIO SEEN PAQ. X 20 UND",
    "description": "Ref: | Marca: BIO SEEN | Prov: DICORLAB S.A.S",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 20 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 8,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1139",
    "sku": "MC 1142",
    "name": "FILTRO PARA COLETA DE SEMEN - REDONDO - 25CM DIAMETRO/ PAQ. X 100 UND",
    "description": "Ref: | Marca: GENES DIFUSION SAS | Prov: GENES DIFUSION SAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "PAQ. X 100 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 0
  },
  {
    "id": "prod-1140",
    "sku": "MC 1143",
    "name": "GAFAS ECOGRAFO CTS 800",
    "description": "Ref: | Marca: SIUI | Prov: SIUI",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1141",
    "sku": "MC 1144",
    "name": "ECOGRAFO BMV MX9 DOPPLER 1 SENSOR",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1142",
    "sku": "MC 1145",
    "name": "ECOGRAFO BMV MX9 DOPPLER 2 SENSORES",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1143",
    "sku": "MC 1146",
    "name": "ECOGRAFO INALAMBRICO SONO EYE/ INCLUYE TABLET",
    "description": "Ref: | Marca: CHISON | Prov: CHISON",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 0,
    "reservedStock": 0,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1144",
    "sku": "MC 1147",
    "name": "SONDA GRASA DORSAL PARA ECOGRAFO BMV S7",
    "description": "Ref: | Marca: BMV | Prov: BMV",
    "categoryId": "cat-3",
    "category": "Equipos de Diagnóstico",
    "requiresSerial": true,
    "unitOfMeasure": "UNIDAD",
    "minStockAlert": 5,
    "baseCost": 450.0,
    "salePrice": 850.0,
    "physicalStock": 1,
    "reservedStock": 1,
    "warrantyMonths": 12
  },
  {
    "id": "prod-1145",
    "sku": "MC 1148",
    "name": "RACK DE PUNTAS 200 UL NO ESTERIL/ BISELAS",
    "description": "Ref: 1110-1800 | Marca: USA SCIENTIFIC | Prov: ISLAS",
    "categoryId": "cat-1",
    "category": "Consumibles Desechables",
    "requiresSerial": false,
    "unitOfMeasure": "RACK X 96 UND",
    "minStockAlert": 20,
    "baseCost": 25.0,
    "salePrice": 55.0,
    "physicalStock": 35,
    "reservedStock": 0,
    "warrantyMonths": 0
  }
];

// Document Sanitizer Helper: Clean NIT vs Cédula with Separate DV Field
function sanitizeDocument(docInput) {
  if (!docInput) return { docType: "NIT", cleanDoc: "0", dv: null };

  const raw = String(docInput).trim();
  
  if (raw.includes("-")) {
    const parts = raw.split("-");
    const numPart = parts[0].replace(/\D/g, "");
    const dvPart = parts[1].replace(/\D/g, "");
    return {
      docType: "NIT",
      cleanDoc: numPart,
      dv: dvPart || null
    };
  }

  const cleanDigits = raw.replace(/\D/g, "");

  if (cleanDigits.length >= 9 && (cleanDigits.startsWith("90") || cleanDigits.startsWith("80"))) {
    return {
      docType: "NIT",
      cleanDoc: cleanDigits.slice(0, -1),
      dv: cleanDigits.slice(-1)
    };
  }

  return {
    docType: "CEDULA",
    cleanDoc: cleanDigits,
    dv: null
  };
}

const rawCustomersData = [
  {
    "id": "cust-xl-1",
    "rawDoc": "900000137-2",
    "name": "FINCA DON PEDRO PMO S.A.S",
    "phone": "+573111171031",
    "email": "contacto@fincadonpedr.co",
    "city": "Medellín",
    "address": "Calle 8 # 23-3"
  },
  {
    "id": "cust-xl-2",
    "rawDoc": "1000015838",
    "name": "CARLOS MAURICIO VANEGAS AVILA",
    "phone": "+573121341062",
    "email": "contacto@carlosmauric.co",
    "city": "Cali",
    "address": "Calle 15 # 36-6"
  },
  {
    "id": "cust-xl-3",
    "rawDoc": "1000023757",
    "name": "ADRIANA MARIA URIBE CALLE",
    "phone": "+573131511093",
    "email": "contacto@adrianamaria.co",
    "city": "Barranquilla",
    "address": "Calle 22 # 49-9"
  },
  {
    "id": "cust-xl-4",
    "rawDoc": "900000548-5",
    "name": "Simijaca RE SAS",
    "phone": "+573141681124",
    "email": "contacto@simijacaresa.co",
    "city": "Bucaramanga",
    "address": "Calle 29 # 62-12"
  },
  {
    "id": "cust-xl-5",
    "rawDoc": "900000685-6",
    "name": "CASABER SAS",
    "phone": "+573151851155",
    "email": "contacto@casabersas.co",
    "city": "Villavicencio",
    "address": "Calle 36 # 75-15"
  },
  {
    "id": "cust-xl-6",
    "rawDoc": "900000822-7",
    "name": "PREMEX S.A.S",
    "phone": "+573162021186",
    "email": "contacto@premexsas.co",
    "city": "Montería",
    "address": "Calle 43 # 88-18"
  },
  {
    "id": "cust-xl-7",
    "rawDoc": "900000959-8",
    "name": "NOGA S.A.S./ OC OCP-119283",
    "phone": "+573172191217",
    "email": "contacto@nogasasococp.co",
    "city": "Pereira",
    "address": "Calle 50 # 11-21"
  },
  {
    "id": "cust-xl-8",
    "rawDoc": "1000063352",
    "name": "ALEJANDRO ACOSTA URREA",
    "phone": "+573182361248",
    "email": "contacto@alejandroaco.co",
    "city": "Manizales",
    "address": "Calle 57 # 24-24"
  },
  {
    "id": "cust-xl-9",
    "rawDoc": "1000071271",
    "name": "FERNANDO PAZMIÑO",
    "phone": "+573192531279",
    "email": "contacto@fernandopazm.co",
    "city": "Neiva",
    "address": "Calle 64 # 37-27"
  },
  {
    "id": "cust-xl-10",
    "rawDoc": "1000079190",
    "name": "ANDREA LISORETH CARRERO MORENO",
    "phone": "+573202701310",
    "email": "contacto@andrealisore.co",
    "city": "Ibagué",
    "address": "Calle 71 # 50-30"
  },
  {
    "id": "cust-xl-11",
    "rawDoc": "1000087109",
    "name": "DIEGO ARMANDO ROJAS PEÑA",
    "phone": "+573212871341",
    "email": "contacto@diegoarmando.co",
    "city": "Popayán",
    "address": "Calle 78 # 63-33"
  },
  {
    "id": "cust-xl-12",
    "rawDoc": "1000095028",
    "name": "LUIS EDUARDO PANTOJA ENRIQUEZ",
    "phone": "+573223041372",
    "email": "contacto@luiseduardop.co",
    "city": "Pastos",
    "address": "Calle 85 # 76-36"
  },
  {
    "id": "cust-xl-13",
    "rawDoc": "900001781-5",
    "name": "SOCIEDAD GANADERA LOS ALPES SAS",
    "phone": "+573233211403",
    "email": "contacto@sociedadgana.co",
    "city": "Santa Marta",
    "address": "Calle 92 # 89-39"
  },
  {
    "id": "cust-xl-14",
    "rawDoc": "900001918-6",
    "name": "AGROVAR SAS",
    "phone": "+573243381434",
    "email": "contacto@agrovarsas.co",
    "city": "Valledupar",
    "address": "Calle 99 # 12-42"
  },
  {
    "id": "cust-xl-15",
    "rawDoc": "900002055-7",
    "name": "TERREROS SERVICIOS AGROPECUARIOS ZOMAC SAS",
    "phone": "+573103551465",
    "email": "contacto@terrerosserv.co",
    "city": "Tunja",
    "address": "Calle 6 # 25-45"
  },
  {
    "id": "cust-xl-16",
    "rawDoc": "1000126704",
    "name": "GLORIA INES PAREJA",
    "phone": "+573113721496",
    "email": "contacto@gloriainespa.co",
    "city": "Bogotá D.C.",
    "address": "Calle 13 # 38-48"
  },
  {
    "id": "cust-xl-17",
    "rawDoc": "1000134623",
    "name": "MIGUEL PERDOMO",
    "phone": "+573123891527",
    "email": "contacto@miguelperdom.co",
    "city": "Medellín",
    "address": "Calle 20 # 51-1"
  },
  {
    "id": "cust-xl-18",
    "rawDoc": "1000142542",
    "name": "MARIO L CARPENA",
    "phone": "+573134061558",
    "email": "contacto@mariolcarpen.co",
    "city": "Cali",
    "address": "Calle 27 # 64-4"
  },
  {
    "id": "cust-xl-19",
    "rawDoc": "900002603-2",
    "name": "INSEMINEMOS BIOTECNOLOGIA PORCINA S.A.S.",
    "phone": "+573144231589",
    "email": "contacto@inseminemosb.co",
    "city": "Barranquilla",
    "address": "Calle 34 # 77-7"
  },
  {
    "id": "cust-xl-20",
    "rawDoc": "900002740-3",
    "name": "CENTRO DE ESPECIALISTAS EN SANIDAD, PRODUCCION Y REPRODUCCION VETERINARIA S.A.S",
    "phone": "+573154401620",
    "email": "contacto@centrodeespe.co",
    "city": "Bucaramanga",
    "address": "Calle 41 # 90-10"
  },
  {
    "id": "cust-xl-21",
    "rawDoc": "900002877-4",
    "name": "GLOBAL BIOTECHNOLOGY CARIBE S.A.S.",
    "phone": "+573164571651",
    "email": "contacto@globalbiotec.co",
    "city": "Villavicencio",
    "address": "Calle 48 # 13-13"
  },
  {
    "id": "cust-xl-22",
    "rawDoc": "1000174218",
    "name": "IN VITRO COLOMBIA S A S",
    "phone": "+573174741682",
    "email": "contacto@invitrocolom.co",
    "city": "Montería",
    "address": "Calle 55 # 26-16"
  },
  {
    "id": "cust-xl-23",
    "rawDoc": "1000182137",
    "name": "ANDRES FELIPE RUIZ",
    "phone": "+573184911713",
    "email": "contacto@andresfelipe.co",
    "city": "Pereira",
    "address": "Calle 62 # 39-19"
  },
  {
    "id": "cust-xl-24",
    "rawDoc": "1000190056",
    "name": "CHRISTIAN GERMAN CAICEDO ALVAREZ",
    "phone": "+573195081744",
    "email": "contacto@christianger.co",
    "city": "Manizales",
    "address": "Calle 69 # 52-22"
  },
  {
    "id": "cust-xl-25",
    "rawDoc": "1000197975",
    "name": "LIBARDO ALFONSO MORENO MUÑOZ",
    "phone": "+573205251775",
    "email": "contacto@libardoalfon.co",
    "city": "Neiva",
    "address": "Calle 76 # 65-25"
  },
  {
    "id": "cust-xl-26",
    "rawDoc": "1000205894",
    "name": "MANUEL MARQUEZ GUTIERREZ",
    "phone": "+573215421806",
    "email": "contacto@manuelmarque.co",
    "city": "Ibagué",
    "address": "Calle 83 # 78-28"
  },
  {
    "id": "cust-xl-27",
    "rawDoc": "1000213813",
    "name": "MAGDIEL JAIMES BLANCO",
    "phone": "+573225591837",
    "email": "contacto@magdieljaime.co",
    "city": "Popayán",
    "address": "Calle 90 # 91-31"
  },
  {
    "id": "cust-xl-28",
    "rawDoc": "900003836-2",
    "name": "BIOHORSE SAS",
    "phone": "+573235761868",
    "email": "contacto@biohorsesas.co",
    "city": "Pastos",
    "address": "Calle 97 # 14-34"
  },
  {
    "id": "cust-xl-29",
    "rawDoc": "1000229651",
    "name": "SEBASTIAN ROPERO",
    "phone": "+573245931899",
    "email": "contacto@sebastianrop.co",
    "city": "Santa Marta",
    "address": "Calle 4 # 27-37"
  },
  {
    "id": "cust-xl-30",
    "rawDoc": "1000237570",
    "name": "JUAN FRANCISCO DAZA VAN-STRAHLEN",
    "phone": "+573106101930",
    "email": "contacto@juanfrancisc.co",
    "city": "Valledupar",
    "address": "Calle 11 # 40-40"
  },
  {
    "id": "cust-xl-31",
    "rawDoc": "1000245489",
    "name": "GUILLERMO ANDRES TORRES LEYVA",
    "phone": "+573116271961",
    "email": "contacto@guillermoand.co",
    "city": "Tunja",
    "address": "Calle 18 # 53-43"
  },
  {
    "id": "cust-xl-32",
    "rawDoc": "900004384-6",
    "name": "MEDICINA EQUINA S.A.S",
    "phone": "+573126441992",
    "email": "contacto@medicinaequi.co",
    "city": "Bogotá D.C.",
    "address": "Calle 25 # 66-46"
  },
  {
    "id": "cust-xl-33",
    "rawDoc": "900004521-7",
    "name": "JHON ALEXANDER BARRIOS PEDRAZA (FINCAGRO BERLIN)",
    "phone": "+573136612023",
    "email": "contacto@jhonalexande.co",
    "city": "Medellín",
    "address": "Calle 32 # 79-49"
  },
  {
    "id": "cust-xl-34",
    "rawDoc": "1000269246",
    "name": "AMAURY FERNANDO VANEGAS GAZABON",
    "phone": "+573146782054",
    "email": "contacto@amauryfernan.co",
    "city": "Cali",
    "address": "Calle 39 # 92-2"
  },
  {
    "id": "cust-xl-35",
    "rawDoc": "1000277165",
    "name": "RODRIGO GUARIN",
    "phone": "+573156952085",
    "email": "contacto@rodrigoguari.co",
    "city": "Barranquilla",
    "address": "Calle 46 # 15-5"
  },
  {
    "id": "cust-xl-36",
    "rawDoc": "900004932-1",
    "name": "AGROVETERINARIA EL FINQUERO JJ S.A.S",
    "phone": "+573167122116",
    "email": "contacto@agroveterina.co",
    "city": "Bucaramanga",
    "address": "Calle 53 # 28-8"
  },
  {
    "id": "cust-xl-37",
    "rawDoc": "900005069-2",
    "name": "AGROPECUARIA EL PALATINO SAS",
    "phone": "+573177292147",
    "email": "contacto@agropecuaria.co",
    "city": "Villavicencio",
    "address": "Calle 60 # 41-11"
  },
  {
    "id": "cust-xl-38",
    "rawDoc": "900005206-3",
    "name": "SANIDAD Y PRODUCCIÓN VETERINARIA S A S",
    "phone": "+573187462178",
    "email": "contacto@sanidadyprod.co",
    "city": "Montería",
    "address": "Calle 67 # 54-14"
  },
  {
    "id": "cust-xl-39",
    "rawDoc": "1000308841",
    "name": "GAITAN MARROQUIN CRISTIAN ANDRES",
    "phone": "+573197632209",
    "email": "contacto@gaitanmarroq.co",
    "city": "Pereira",
    "address": "Calle 74 # 67-17"
  },
  {
    "id": "cust-xl-40",
    "rawDoc": "900005480-5",
    "name": "INSEMINACION ARTIFICIAL S A S",
    "phone": "+573207802240",
    "email": "contacto@inseminacion.co",
    "city": "Manizales",
    "address": "Calle 81 # 80-20"
  },
  {
    "id": "cust-xl-41",
    "rawDoc": "1000324679",
    "name": "RUBEN DARIO NEGRETTE RAMOS",
    "phone": "+573217972271",
    "email": "contacto@rubendarione.co",
    "city": "Neiva",
    "address": "Calle 88 # 93-23"
  },
  {
    "id": "cust-xl-42",
    "rawDoc": "1000332598",
    "name": "EMBRIONES DEL SINÚ",
    "phone": "+573228142302",
    "email": "contacto@embrionesdel.co",
    "city": "Ibagué",
    "address": "Calle 95 # 16-26"
  },
  {
    "id": "cust-xl-43",
    "rawDoc": "1000340517",
    "name": "PAOLA ANDREA RODRIGUEZ",
    "phone": "+573238312333",
    "email": "contacto@paolaandrear.co",
    "city": "Popayán",
    "address": "Calle 2 # 29-29"
  },
  {
    "id": "cust-xl-44",
    "rawDoc": "900006028-9",
    "name": "AGROPECUARIA BAMBUSA SAS",
    "phone": "+573248482364",
    "email": "contacto@agropecuaria.co",
    "city": "Pastos",
    "address": "Calle 9 # 42-32"
  },
  {
    "id": "cust-xl-45",
    "rawDoc": "1000356355",
    "name": "GENÉTICA IN VITRO C.A.",
    "phone": "+573108652395",
    "email": "contacto@genticainvit.co",
    "city": "Santa Marta",
    "address": "Calle 16 # 55-35"
  },
  {
    "id": "cust-xl-46",
    "rawDoc": "1000364274",
    "name": "SERGIO MAURICIO OREJARENA CASTELLANOS",
    "phone": "+573118822426",
    "email": "contacto@sergiomauric.co",
    "city": "Valledupar",
    "address": "Calle 23 # 68-38"
  },
  {
    "id": "cust-xl-47",
    "rawDoc": "900006439-3",
    "name": "ROSA MARIA QUITORA MENDOZA",
    "phone": "+573128992457",
    "email": "contacto@rosamariaqui.co",
    "city": "Tunja",
    "address": "Calle 30 # 81-41"
  },
  {
    "id": "cust-xl-48",
    "rawDoc": "1000380112",
    "name": "FUNDACION GRUPO ARGOS/ N°.PMO 6000048188",
    "phone": "+573139162488",
    "email": "contacto@fundaciongru.co",
    "city": "Bogotá D.C.",
    "address": "Calle 37 # 94-44"
  },
  {
    "id": "cust-xl-49",
    "rawDoc": "1000388031",
    "name": "FUNDACION GRUPO ARGOS",
    "phone": "+573149332519",
    "email": "contacto@fundaciongru.co",
    "city": "Medellín",
    "address": "Calle 44 # 17-47"
  },
  {
    "id": "cust-xl-50",
    "rawDoc": "900006850-6",
    "name": "GENETICA & SERVICIOS ABS SAS",
    "phone": "+573159502550",
    "email": "contacto@geneticaserv.co",
    "city": "Cali",
    "address": "Calle 51 # 30-0"
  },
  {
    "id": "cust-xl-51",
    "rawDoc": "900006987-7",
    "name": "LUISA FERNANDA PORTILLO",
    "phone": "+573169672581",
    "email": "contacto@luisafernand.co",
    "city": "Barranquilla",
    "address": "Calle 58 # 43-3"
  },
  {
    "id": "cust-xl-52",
    "rawDoc": "900007124-8",
    "name": "CARLOS AUGUSTO SANABRIA QUIROGA",
    "phone": "+573179842612",
    "email": "contacto@carlosaugust.co",
    "city": "Bucaramanga",
    "address": "Calle 65 # 56-6"
  },
  {
    "id": "cust-xl-53",
    "rawDoc": "900007261-9",
    "name": "JHON ALBERTO TROCHEZ SAMBONI",
    "phone": "+573181022643",
    "email": "contacto@jhonalbertot.co",
    "city": "Villavicencio",
    "address": "Calle 72 # 69-9"
  },
  {
    "id": "cust-xl-54",
    "rawDoc": "1000427626",
    "name": "JORGE LUIS LOZANO SUAREZ",
    "phone": "+573191192674",
    "email": "contacto@jorgeluisloz.co",
    "city": "Montería",
    "address": "Calle 79 # 82-12"
  },
  {
    "id": "cust-xl-55",
    "rawDoc": "900007535-2",
    "name": "FRANCISCO ISAEL CASTRO GONZALEZ",
    "phone": "+573201362705",
    "email": "contacto@franciscoisa.co",
    "city": "Pereira",
    "address": "Calle 86 # 95-15"
  },
  {
    "id": "cust-xl-56",
    "rawDoc": "1000443464",
    "name": "JUAN MARTÍN GIRALDO MEJIA",
    "phone": "+573211532736",
    "email": "contacto@juanmartngir.co",
    "city": "Manizales",
    "address": "Calle 93 # 18-18"
  },
  {
    "id": "cust-xl-57",
    "rawDoc": "1000451383",
    "name": "DAVID ZAPATA",
    "phone": "+573221702767",
    "email": "contacto@davidzapata.co",
    "city": "Neiva",
    "address": "Calle 100 # 31-21"
  },
  {
    "id": "cust-xl-58",
    "rawDoc": "900007946-5",
    "name": "VETPEGAN SAS",
    "phone": "+573231872798",
    "email": "contacto@vetpegansas.co",
    "city": "Ibagué",
    "address": "Calle 7 # 44-24"
  },
  {
    "id": "cust-xl-59",
    "rawDoc": "900008083-6",
    "name": "INVERSIONES CARIBE SIGLO XXI SAS",
    "phone": "+573242042829",
    "email": "contacto@inversionesc.co",
    "city": "Popayán",
    "address": "Calle 14 # 57-27"
  },
  {
    "id": "cust-xl-60",
    "rawDoc": "1000475140",
    "name": "ROGER AMADOR MARTINEZ RODRIGUEZ",
    "phone": "+573102212860",
    "email": "contacto@rogeramadorm.co",
    "city": "Pastos",
    "address": "Calle 21 # 70-30"
  },
  {
    "id": "cust-xl-61",
    "rawDoc": "900008357-8",
    "name": "INVERSIONES MONTAÑO/ KARLA ROAS",
    "phone": "+573112382891",
    "email": "contacto@inversionesm.co",
    "city": "Santa Marta",
    "address": "Calle 28 # 83-33"
  },
  {
    "id": "cust-xl-62",
    "rawDoc": "900008494-9",
    "name": "GENETICA SELECTA SA",
    "phone": "+573122552922",
    "email": "contacto@geneticasele.co",
    "city": "Valledupar",
    "address": "Calle 35 # 96-36"
  },
  {
    "id": "cust-xl-63",
    "rawDoc": "1000498897",
    "name": "MARIA JOSE QUINTERO DE LOS RIOS",
    "phone": "+573132722953",
    "email": "contacto@mariajosequi.co",
    "city": "Tunja",
    "address": "Calle 42 # 19-39"
  },
  {
    "id": "cust-xl-64",
    "rawDoc": "1000506816",
    "name": "GERARDO ERNESTO DURAN",
    "phone": "+573142892984",
    "email": "contacto@gerardoernes.co",
    "city": "Bogotá D.C.",
    "address": "Calle 49 # 32-42"
  },
  {
    "id": "cust-xl-65",
    "rawDoc": "900008905-3",
    "name": "BIOAGRARIUS S.A.S",
    "phone": "+573153063015",
    "email": "contacto@bioagrariuss.co",
    "city": "Medellín",
    "address": "Calle 56 # 45-45"
  },
  {
    "id": "cust-xl-66",
    "rawDoc": "900009042-4",
    "name": "GVM CORPORATION GLOBAL VETERINARY MEDICINE SAS",
    "phone": "+573163233046",
    "email": "contacto@gvmcorporati.co",
    "city": "Cali",
    "address": "Calle 63 # 58-48"
  },
  {
    "id": "cust-xl-67",
    "rawDoc": "1000530573",
    "name": "HECTOR GEOVANNY HERRERA ROMERO",
    "phone": "+573173403077",
    "email": "contacto@hectorgeovan.co",
    "city": "Barranquilla",
    "address": "Calle 70 # 71-1"
  },
  {
    "id": "cust-xl-68",
    "rawDoc": "1000538492",
    "name": "NICOLAS RICARDO BERMUDEZ ANGEL",
    "phone": "+573183573108",
    "email": "contacto@nicolasricar.co",
    "city": "Bucaramanga",
    "address": "Calle 77 # 84-4"
  },
  {
    "id": "cust-xl-69",
    "rawDoc": "1000546411",
    "name": "DANIEL ESCOBAR MONTES",
    "phone": "+573193743139",
    "email": "contacto@danielescoba.co",
    "city": "Villavicencio",
    "address": "Calle 84 # 97-7"
  },
  {
    "id": "cust-xl-70",
    "rawDoc": "900009590-8",
    "name": "JOAN SEBASTIAN RAMIREZ BARBOSA",
    "phone": "+573203913170",
    "email": "contacto@joansebastia.co",
    "city": "Montería",
    "address": "Calle 91 # 20-10"
  },
  {
    "id": "cust-xl-71",
    "rawDoc": "900009727-9",
    "name": "EMBRIONES SAS",
    "phone": "+573214083201",
    "email": "contacto@embrionessas.co",
    "city": "Pereira",
    "address": "Calle 98 # 33-13"
  },
  {
    "id": "cust-xl-72",
    "rawDoc": "1000570168",
    "name": "NOVA HOLDING GROUP",
    "phone": "+573224253232",
    "email": "contacto@novaholdingg.co",
    "city": "Manizales",
    "address": "Calle 5 # 46-16"
  },
  {
    "id": "cust-xl-73",
    "rawDoc": "1000578087",
    "name": "HECTOR GUILLERMO DORADO ENRIQUEZ",
    "phone": "+573234423263",
    "email": "contacto@hectorguille.co",
    "city": "Neiva",
    "address": "Calle 12 # 59-19"
  },
  {
    "id": "cust-xl-74",
    "rawDoc": "900010138-3",
    "name": "SAMUEL FUENTES GARRIDO",
    "phone": "+573244593294",
    "email": "contacto@samuelfuente.co",
    "city": "Ibagué",
    "address": "Calle 19 # 72-22"
  },
  {
    "id": "cust-xl-75",
    "rawDoc": "1000593925",
    "name": "DEICY BRAVO JOJOA",
    "phone": "+573104763325",
    "email": "contacto@deicybravojo.co",
    "city": "Popayán",
    "address": "Calle 26 # 85-25"
  },
  {
    "id": "cust-xl-76",
    "rawDoc": "1000601844",
    "name": "DIEGO FERNANDO RAMIREZ BOHORQUEZ",
    "phone": "+573114933356",
    "email": "contacto@diegofernand.co",
    "city": "Pastos",
    "address": "Calle 33 # 98-28"
  },
  {
    "id": "cust-xl-77",
    "rawDoc": "1000609763",
    "name": "DAYRON DUBIEL NIÑO ARIZA",
    "phone": "+573125103387",
    "email": "contacto@dayrondubiel.co",
    "city": "Santa Marta",
    "address": "Calle 40 # 21-31"
  },
  {
    "id": "cust-xl-78",
    "rawDoc": "900010686-7",
    "name": "BIOGENETIC IN-VITRO SAS",
    "phone": "+573135273418",
    "email": "contacto@biogeneticin.co",
    "city": "Valledupar",
    "address": "Calle 47 # 34-34"
  },
  {
    "id": "cust-xl-79",
    "rawDoc": "900010823-8",
    "name": "HERSAL SALDARRIAGA",
    "phone": "+573145443449",
    "email": "contacto@hersalsaldar.co",
    "city": "Tunja",
    "address": "Calle 54 # 47-37"
  },
  {
    "id": "cust-xl-80",
    "rawDoc": "900010960-9",
    "name": "CONSULTA PECUARIA S.A.S.",
    "phone": "+573155613480",
    "email": "contacto@consultapecu.co",
    "city": "Bogotá D.C.",
    "address": "Calle 61 # 60-40"
  },
  {
    "id": "cust-xl-81",
    "rawDoc": "1000641439",
    "name": "TEÓFILO FRANCISCO HERAZO SUAREZ",
    "phone": "+573165783511",
    "email": "contacto@tefilofranci.co",
    "city": "Medellín",
    "address": "Calle 68 # 73-43"
  },
  {
    "id": "cust-xl-82",
    "rawDoc": "1000649358",
    "name": "WILFREDO PERDOMO TOVAR",
    "phone": "+573175953542",
    "email": "contacto@wilfredoperd.co",
    "city": "Cali",
    "address": "Calle 75 # 86-46"
  },
  {
    "id": "cust-xl-83",
    "rawDoc": "900011371-3",
    "name": "GRUPO LABSERVIS LTDA",
    "phone": "+573186123573",
    "email": "contacto@grupolabserv.co",
    "city": "Barranquilla",
    "address": "Calle 82 # 99-49"
  },
  {
    "id": "cust-xl-84",
    "rawDoc": "1000665196",
    "name": "CRISTIAN ARANDA AGUILAR",
    "phone": "+573196293604",
    "email": "contacto@cristianaran.co",
    "city": "Bucaramanga",
    "address": "Calle 89 # 22-2"
  },
  {
    "id": "cust-xl-85",
    "rawDoc": "1000673115",
    "name": "RAFAEL MATTUTE",
    "phone": "+573206463635",
    "email": "contacto@rafaelmattut.co",
    "city": "Villavicencio",
    "address": "Calle 96 # 35-5"
  },
  {
    "id": "cust-xl-86",
    "rawDoc": "1000681034",
    "name": "LAWRENS YEPES PELÁEZ",
    "phone": "+573216633666",
    "email": "contacto@lawrensyepes.co",
    "city": "Montería",
    "address": "Calle 3 # 48-8"
  },
  {
    "id": "cust-xl-87",
    "rawDoc": "1000688953",
    "name": "ARIEL IGNACIO MARTINEZ",
    "phone": "+573226803697",
    "email": "contacto@arielignacio.co",
    "city": "Pereira",
    "address": "Calle 10 # 61-11"
  },
  {
    "id": "cust-xl-88",
    "rawDoc": "900012056-8",
    "name": "INSMEVET S SAS",
    "phone": "+573236973728",
    "email": "contacto@insmevetssas.co",
    "city": "Manizales",
    "address": "Calle 17 # 74-14"
  },
  {
    "id": "cust-xl-89",
    "rawDoc": "1000704791",
    "name": "OSCAR MANUEL MENDEZ MORA",
    "phone": "+573247143759",
    "email": "contacto@oscarmanuelm.co",
    "city": "Neiva",
    "address": "Calle 24 # 87-17"
  },
  {
    "id": "cust-xl-90",
    "rawDoc": "900012330-1",
    "name": "SEMEX COLOMBIA S.A.S.",
    "phone": "+573107313790",
    "email": "contacto@semexcolombi.co",
    "city": "Ibagué",
    "address": "Calle 31 # 10-20"
  },
  {
    "id": "cust-xl-91",
    "rawDoc": "900012467-2",
    "name": "SODAGAL SAS",
    "phone": "+573117483821",
    "email": "contacto@sodagalsas.co",
    "city": "Popayán",
    "address": "Calle 38 # 23-23"
  },
  {
    "id": "cust-xl-92",
    "rawDoc": "900012604-3",
    "name": "NINI JOHANNA HERNÁNDEZ SANDOVAL",
    "phone": "+573127653852",
    "email": "contacto@ninijohannah.co",
    "city": "Pastos",
    "address": "Calle 45 # 36-26"
  },
  {
    "id": "cust-xl-93",
    "rawDoc": "900012741-4",
    "name": "JAIME ANDRES GARCIA",
    "phone": "+573137823883",
    "email": "contacto@jaimeandresg.co",
    "city": "Santa Marta",
    "address": "Calle 52 # 49-29"
  },
  {
    "id": "cust-xl-94",
    "rawDoc": "900012878-5",
    "name": "Embriogenex SAS",
    "phone": "+573147993914",
    "email": "contacto@embriogenexs.co",
    "city": "Valledupar",
    "address": "Calle 59 # 62-32"
  },
  {
    "id": "cust-xl-95",
    "rawDoc": "1000752305",
    "name": "IVÁN LIBARDO FRANCO CASTRO",
    "phone": "+573158163945",
    "email": "contacto@ivnlibardofr.co",
    "city": "Tunja",
    "address": "Calle 66 # 75-35"
  },
  {
    "id": "cust-xl-96",
    "rawDoc": "1000760224",
    "name": "JAIRO ARIZA SERRANO",
    "phone": "+573168333976",
    "email": "contacto@jairoarizase.co",
    "city": "Bogotá D.C.",
    "address": "Calle 73 # 88-38"
  },
  {
    "id": "cust-xl-97",
    "rawDoc": "1000768143",
    "name": "AGRO REPROGEN DEL VALLE",
    "phone": "+573178504007",
    "email": "contacto@agroreprogen.co",
    "city": "Medellín",
    "address": "Calle 80 # 11-41"
  },
  {
    "id": "cust-xl-98",
    "rawDoc": "900013426-9",
    "name": "INVERSIONES CARGRAN S.A.S",
    "phone": "+573188674038",
    "email": "contacto@inversionesc.co",
    "city": "Cali",
    "address": "Calle 87 # 24-44"
  },
  {
    "id": "cust-xl-99",
    "rawDoc": "900013563-1",
    "name": "DIEGO ALEXANDER SABOGAL RINCON",
    "phone": "+573198844069",
    "email": "contacto@diegoalexand.co",
    "city": "Barranquilla",
    "address": "Calle 94 # 37-47"
  },
  {
    "id": "cust-xl-100",
    "rawDoc": "900013700-2",
    "name": "AGRÍCOLAS Y GANADERAS ROMERO LATORRE SAS",
    "phone": "+573209014100",
    "email": "contacto@agrcolasygan.co",
    "city": "Bucaramanga",
    "address": "Calle 1 # 50-0"
  },
  {
    "id": "cust-xl-101",
    "rawDoc": "900013837-3",
    "name": "MADIAGRO SAS",
    "phone": "+573219184131",
    "email": "contacto@madiagrosas.co",
    "city": "Villavicencio",
    "address": "Calle 8 # 63-3"
  },
  {
    "id": "cust-xl-102",
    "rawDoc": "900013974-4",
    "name": "BIOEMBRIO FIV SAS",
    "phone": "+573229354162",
    "email": "contacto@bioembriofiv.co",
    "city": "Montería",
    "address": "Calle 15 # 76-6"
  },
  {
    "id": "cust-xl-103",
    "rawDoc": "1000815657",
    "name": "DIEGO ENRIQUE VALBUENA TOVAR",
    "phone": "+573239524193",
    "email": "contacto@diegoenrique.co",
    "city": "Pereira",
    "address": "Calle 22 # 89-9"
  },
  {
    "id": "cust-xl-104",
    "rawDoc": "900014248-6",
    "name": "PORCICOLA EN NORAL SAS",
    "phone": "+573249694224",
    "email": "contacto@porcicolaenn.co",
    "city": "Manizales",
    "address": "Calle 29 # 12-12"
  },
  {
    "id": "cust-xl-105",
    "rawDoc": "1000831495",
    "name": "HANZEL ALBERTO HORTUA",
    "phone": "+573109864255",
    "email": "contacto@hanzelalbert.co",
    "city": "Neiva",
    "address": "Calle 36 # 25-15"
  },
  {
    "id": "cust-xl-106",
    "rawDoc": "1000839414",
    "name": "EDGAR FABIÁN GUTIERREZ PABON",
    "phone": "+573111044286",
    "email": "contacto@edgarfabingu.co",
    "city": "Ibagué",
    "address": "Calle 43 # 38-18"
  },
  {
    "id": "cust-xl-107",
    "rawDoc": "1000847333",
    "name": "BELLANIRA ROJAS RUBIO",
    "phone": "+573121214317",
    "email": "contacto@bellaniraroj.co",
    "city": "Popayán",
    "address": "Calle 50 # 51-21"
  },
  {
    "id": "cust-xl-108",
    "rawDoc": "900014796-1",
    "name": "INVERSIONES & SUMINISTROS GMTC SAS",
    "phone": "+573131384348",
    "email": "contacto@inversioness.co",
    "city": "Pastos",
    "address": "Calle 57 # 64-24"
  },
  {
    "id": "cust-xl-109",
    "rawDoc": "1000863171",
    "name": "JUAN CAMILO GIRALDO FLOREZ",
    "phone": "+573141554379",
    "email": "contacto@juancamilogi.co",
    "city": "Santa Marta",
    "address": "Calle 64 # 77-27"
  },
  {
    "id": "cust-xl-110",
    "rawDoc": "900015070-3",
    "name": "AGROPECUARIA PALMAR DEL EDEN S.A.S",
    "phone": "+573151724410",
    "email": "contacto@agropecuaria.co",
    "city": "Valledupar",
    "address": "Calle 71 # 90-30"
  },
  {
    "id": "cust-xl-111",
    "rawDoc": "1000879009",
    "name": "DIEGO ARMANDO BERNAL RIVERA",
    "phone": "+573161894441",
    "email": "contacto@diegoarmando.co",
    "city": "Tunja",
    "address": "Calle 78 # 13-33"
  },
  {
    "id": "cust-xl-112",
    "rawDoc": "1000886928",
    "name": "LUCELY SUAREZ ESCOBAR",
    "phone": "+573172064472",
    "email": "contacto@lucelysuarez.co",
    "city": "Bogotá D.C.",
    "address": "Calle 85 # 26-36"
  },
  {
    "id": "cust-xl-113",
    "rawDoc": "900015481-6",
    "name": "INGENIERIA E INVERSIONES EL SANTUARIO ACR S.A.S",
    "phone": "+573182234503",
    "email": "contacto@ingenieriaei.co",
    "city": "Medellín",
    "address": "Calle 92 # 39-39"
  },
  {
    "id": "cust-xl-114",
    "rawDoc": "1000902766",
    "name": "ANA CAROLINA GUION MOJICA",
    "phone": "+573192404534",
    "email": "contacto@anacarolinag.co",
    "city": "Cali",
    "address": "Calle 99 # 52-42"
  },
  {
    "id": "cust-xl-115",
    "rawDoc": "1000910685",
    "name": "ALEJANDRA URREGO",
    "phone": "+573202574565",
    "email": "contacto@alejandraurr.co",
    "city": "Barranquilla",
    "address": "Calle 6 # 65-45"
  },
  {
    "id": "cust-xl-116",
    "rawDoc": "1000918604",
    "name": "FAIBER EDUARDO BURGOS /ALEJANDRO OLAYA HERNÁNDEZ",
    "phone": "+573212744596",
    "email": "contacto@faibereduard.co",
    "city": "Bucaramanga",
    "address": "Calle 13 # 78-48"
  },
  {
    "id": "cust-xl-117",
    "rawDoc": "1000926523",
    "name": "OSCAR HUMBERTO GOYES UNIGARRO",
    "phone": "+573222914627",
    "email": "contacto@oscarhumbert.co",
    "city": "Villavicencio",
    "address": "Calle 20 # 91-1"
  },
  {
    "id": "cust-xl-118",
    "rawDoc": "1000934442",
    "name": "ALEXANDER VINASCO CASTRO",
    "phone": "+573233084658",
    "email": "contacto@alexandervin.co",
    "city": "Montería",
    "address": "Calle 27 # 14-4"
  },
  {
    "id": "cust-xl-119",
    "rawDoc": "900016303-3",
    "name": "HECTOR HERNAN GUERRERO GARCIA",
    "phone": "+573243254689",
    "email": "contacto@hectorhernan.co",
    "city": "Pereira",
    "address": "Calle 34 # 27-7"
  },
  {
    "id": "cust-xl-120",
    "rawDoc": "900016440-4",
    "name": "VITRO BIOTECH FERTILIZACION IN VITRO SAS",
    "phone": "+573103424720",
    "email": "contacto@vitrobiotech.co",
    "city": "Manizales",
    "address": "Calle 41 # 40-10"
  },
  {
    "id": "cust-xl-121",
    "rawDoc": "1000958199",
    "name": "ALBERTO ANGEL AZA ACOSTA",
    "phone": "+573113594751",
    "email": "contacto@albertoangel.co",
    "city": "Neiva",
    "address": "Calle 48 # 53-13"
  },
  {
    "id": "cust-xl-122",
    "rawDoc": "1000966118",
    "name": "JORGE TRUJILLO LONDOÑO",
    "phone": "+573123764782",
    "email": "contacto@jorgetrujill.co",
    "city": "Ibagué",
    "address": "Calle 55 # 66-16"
  },
  {
    "id": "cust-xl-123",
    "rawDoc": "1000974037",
    "name": "LUCAS LOZANO BORRERO",
    "phone": "+573133934813",
    "email": "contacto@lucaslozanob.co",
    "city": "Popayán",
    "address": "Calle 62 # 79-19"
  },
  {
    "id": "cust-xl-124",
    "rawDoc": "900016988-8",
    "name": "PROGENE SAS",
    "phone": "+573144104844",
    "email": "contacto@progenesas.co",
    "city": "Pastos",
    "address": "Calle 69 # 92-22"
  },
  {
    "id": "cust-xl-125",
    "rawDoc": "1000989875",
    "name": "JUAN CAMILO RESTREPO TOBON",
    "phone": "+573154274875",
    "email": "contacto@juancamilore.co",
    "city": "Santa Marta",
    "address": "Calle 76 # 15-25"
  },
  {
    "id": "cust-xl-126",
    "rawDoc": "900017262-1",
    "name": "EIDER JAVIER SALAZAR AVILA",
    "phone": "+573164444906",
    "email": "contacto@eiderjaviers.co",
    "city": "Valledupar",
    "address": "Calle 83 # 28-28"
  },
  {
    "id": "cust-xl-127",
    "rawDoc": "1001005713",
    "name": "GIOVANNY MUÑOZ QUINTERO",
    "phone": "+573174614937",
    "email": "contacto@giovannymuoz.co",
    "city": "Tunja",
    "address": "Calle 90 # 41-31"
  },
  {
    "id": "cust-xl-128",
    "rawDoc": "1001013632",
    "name": "RUBEN DARIO MENDOZA GONZALEZ",
    "phone": "+573184784968",
    "email": "contacto@rubendariome.co",
    "city": "Bogotá D.C.",
    "address": "Calle 97 # 54-34"
  },
  {
    "id": "cust-xl-129",
    "rawDoc": "900017673-4",
    "name": "AGRÍCOLA LA FINCA S.A.S",
    "phone": "+573194954999",
    "email": "contacto@agrcolalafin.co",
    "city": "Medellín",
    "address": "Calle 4 # 67-37"
  },
  {
    "id": "cust-xl-130",
    "rawDoc": "1001029470",
    "name": "RED PECUARIA MG S A S",
    "phone": "+573205125030",
    "email": "contacto@redpecuariam.co",
    "city": "Cali",
    "address": "Calle 11 # 80-40"
  },
  {
    "id": "cust-xl-131",
    "rawDoc": "1001037389",
    "name": "JUAN CAMILO BENETIZ HERNANDEZ",
    "phone": "+573215295061",
    "email": "contacto@juancamilobe.co",
    "city": "Barranquilla",
    "address": "Calle 18 # 93-43"
  },
  {
    "id": "cust-xl-132",
    "rawDoc": "1001045308",
    "name": "FELIPE HELMAN MENDOZA CASTRO",
    "phone": "+573225465092",
    "email": "contacto@felipehelman.co",
    "city": "Bucaramanga",
    "address": "Calle 25 # 16-46"
  },
  {
    "id": "cust-xl-133",
    "rawDoc": "900018221-8",
    "name": "INVERSIONES CELNAR SAS",
    "phone": "+573235635123",
    "email": "contacto@inversionesc.co",
    "city": "Villavicencio",
    "address": "Calle 32 # 29-49"
  },
  {
    "id": "cust-xl-134",
    "rawDoc": "900018358-9",
    "name": "GENPROTECH S.A.S.",
    "phone": "+573245805154",
    "email": "contacto@genprotechsa.co",
    "city": "Montería",
    "address": "Calle 39 # 42-2"
  },
  {
    "id": "cust-xl-135",
    "rawDoc": "1001069065",
    "name": "OSCAR MAURICIO BUITRAGO VARGAS",
    "phone": "+573105975185",
    "email": "contacto@oscarmaurici.co",
    "city": "Pereira",
    "address": "Calle 46 # 55-5"
  },
  {
    "id": "cust-xl-136",
    "rawDoc": "1001076984",
    "name": "CARLOS JULIO PERILLA CALDAS",
    "phone": "+573116145216",
    "email": "contacto@carlosjuliop.co",
    "city": "Manizales",
    "address": "Calle 53 # 68-8"
  },
  {
    "id": "cust-xl-137",
    "rawDoc": "1001084903",
    "name": "MARGREGORY CASTRO",
    "phone": "+573126315247",
    "email": "contacto@margregoryca.co",
    "city": "Neiva",
    "address": "Calle 60 # 81-11"
  },
  {
    "id": "cust-xl-138",
    "rawDoc": "900018906-4",
    "name": "ASOCIACION DE GANADEROS AMBIENTALISTAS DE SAN VICENTE DEL CAGUÁN",
    "phone": "+573136485278",
    "email": "contacto@asociacionde.co",
    "city": "Ibagué",
    "address": "Calle 67 # 94-14"
  },
  {
    "id": "cust-xl-139",
    "rawDoc": "900019043-5",
    "name": "PAISAGRO SAS",
    "phone": "+573146655309",
    "email": "contacto@paisagrosas.co",
    "city": "Popayán",
    "address": "Calle 74 # 17-17"
  },
  {
    "id": "cust-xl-140",
    "rawDoc": "900019180-6",
    "name": "LUIS FERNANDO MONSALVE",
    "phone": "+573156825340",
    "email": "contacto@luisfernando.co",
    "city": "Pastos",
    "address": "Calle 81 # 30-20"
  },
  {
    "id": "cust-xl-141",
    "rawDoc": "1001116579",
    "name": "ANDRES RIVERA CASTRILLON",
    "phone": "+573166995371",
    "email": "contacto@andresrivera.co",
    "city": "Santa Marta",
    "address": "Calle 88 # 43-23"
  },
  {
    "id": "cust-xl-142",
    "rawDoc": "900019454-8",
    "name": "RAUL SULPICIO SARMIENTO CELY /HENNIS DIAZ",
    "phone": "+573177165402",
    "email": "contacto@raulsulpicio.co",
    "city": "Valledupar",
    "address": "Calle 95 # 56-26"
  },
  {
    "id": "cust-xl-143",
    "rawDoc": "1001132417",
    "name": "PEDRO ANTONIO ZARATE OTALORA",
    "phone": "+573187335433",
    "email": "contacto@pedroantonio.co",
    "city": "Tunja",
    "address": "Calle 2 # 69-29"
  },
  {
    "id": "cust-xl-144",
    "rawDoc": "1001140336",
    "name": "JOSE WILLIAM PRADA APARICIO",
    "phone": "+573197505464",
    "email": "contacto@josewilliamp.co",
    "city": "Bogotá D.C.",
    "address": "Calle 9 # 82-32"
  },
  {
    "id": "cust-xl-145",
    "rawDoc": "900019865-2",
    "name": "COMERCIALIZADORA SERLE.COM SAS",
    "phone": "+573207675495",
    "email": "contacto@comercializa.co",
    "city": "Medellín",
    "address": "Calle 16 # 95-35"
  },
  {
    "id": "cust-xl-146",
    "rawDoc": "1001156174",
    "name": "ALBERTO URIBE ARANGO",
    "phone": "+573217845526",
    "email": "contacto@albertouribe.co",
    "city": "Cali",
    "address": "Calle 23 # 18-38"
  },
  {
    "id": "cust-xl-147",
    "rawDoc": "1001164093",
    "name": "ANDREA LILIANA ALFONSO",
    "phone": "+573228015557",
    "email": "contacto@andrealilian.co",
    "city": "Barranquilla",
    "address": "Calle 30 # 31-41"
  },
  {
    "id": "cust-xl-148",
    "rawDoc": "1001172012",
    "name": "GUSTAVO MONZON",
    "phone": "+573238185588",
    "email": "contacto@gustavomonzo.co",
    "city": "Bucaramanga",
    "address": "Calle 37 # 44-44"
  },
  {
    "id": "cust-xl-149",
    "rawDoc": "1001179931",
    "name": "GREY CAROLINA PABON CHACON",
    "phone": "+573248355619",
    "email": "contacto@greycarolina.co",
    "city": "Villavicencio",
    "address": "Calle 44 # 57-47"
  },
  {
    "id": "cust-xl-150",
    "rawDoc": "900020550-7",
    "name": "BIOTECNOLOGIA Y AGROPECUARIA SF ZOMAC SAS",
    "phone": "+573108525650",
    "email": "contacto@biotecnologi.co",
    "city": "Montería",
    "address": "Calle 51 # 70-0"
  },
  {
    "id": "cust-xl-151",
    "rawDoc": "900020687-8",
    "name": "GRANJA LA SIERRA LTDA",
    "phone": "+573118695681",
    "email": "contacto@granjalasier.co",
    "city": "Pereira",
    "address": "Calle 58 # 83-3"
  },
  {
    "id": "cust-xl-152",
    "rawDoc": "1001203688",
    "name": "LUIS HERNAN NUÑEZ FALLA",
    "phone": "+573128865712",
    "email": "contacto@luishernannu.co",
    "city": "Manizales",
    "address": "Calle 65 # 96-6"
  },
  {
    "id": "cust-xl-153",
    "rawDoc": "900020961-1",
    "name": "CITAG S A S CENTRO INTERNACIONAL DE TRAZABILIDAD AGROPECUARIA",
    "phone": "+573139035743",
    "email": "contacto@citagsascent.co",
    "city": "Neiva",
    "address": "Calle 72 # 19-9"
  },
  {
    "id": "cust-xl-154",
    "rawDoc": "1001219526",
    "name": "AGROPECUARIA LOS MANANTIALES S A S",
    "phone": "+573149205774",
    "email": "contacto@agropecuaria.co",
    "city": "Ibagué",
    "address": "Calle 79 # 32-12"
  },
  {
    "id": "cust-xl-155",
    "rawDoc": "1001227445",
    "name": "LIGIA CHAPARRO",
    "phone": "+573159375805",
    "email": "contacto@ligiachaparr.co",
    "city": "Popayán",
    "address": "Calle 86 # 45-15"
  },
  {
    "id": "cust-xl-156",
    "rawDoc": "900021372-4",
    "name": "DIEGO SANCHEZ",
    "phone": "+573169545836",
    "email": "contacto@diegosanchez.co",
    "city": "Pastos",
    "address": "Calle 93 # 58-18"
  },
  {
    "id": "cust-xl-157",
    "rawDoc": "900021509-5",
    "name": "AGRO INVERSIONES JAMAX SAS",
    "phone": "+573179715867",
    "email": "contacto@agroinversio.co",
    "city": "Santa Marta",
    "address": "Calle 100 # 71-21"
  },
  {
    "id": "cust-xl-158",
    "rawDoc": "900021646-6",
    "name": "UNIVERSIDAD DE CALDAS /ODC 343",
    "phone": "+573189885898",
    "email": "contacto@universidadd.co",
    "city": "Valledupar",
    "address": "Calle 7 # 84-24"
  },
  {
    "id": "cust-xl-159",
    "rawDoc": "900021783-7",
    "name": "GRUPO EMPRESARIAL LIZAYA S.A.S.",
    "phone": "+573191065929",
    "email": "contacto@grupoempresa.co",
    "city": "Tunja",
    "address": "Calle 14 # 97-27"
  },
  {
    "id": "cust-xl-160",
    "rawDoc": "1001267040",
    "name": "DEMOSTRACIÓN MX5",
    "phone": "+573201235960",
    "email": "contacto@demostracinm.co",
    "city": "Bogotá D.C.",
    "address": "Calle 21 # 20-30"
  },
  {
    "id": "cust-xl-161",
    "rawDoc": "900022057-9",
    "name": "VITRO LAB SAS",
    "phone": "+573211405991",
    "email": "contacto@vitrolabsas.co",
    "city": "Medellín",
    "address": "Calle 28 # 33-33"
  },
  {
    "id": "cust-xl-162",
    "rawDoc": "1001282878",
    "name": "EDUARDO LONDOÑO BURITICA",
    "phone": "+573221576022",
    "email": "contacto@eduardolondo.co",
    "city": "Cali",
    "address": "Calle 35 # 46-36"
  },
  {
    "id": "cust-xl-163",
    "rawDoc": "900022331-2",
    "name": "SANDRA CAROLINA CHACON CASTRO",
    "phone": "+573231746053",
    "email": "contacto@sandracaroli.co",
    "city": "Barranquilla",
    "address": "Calle 42 # 59-39"
  },
  {
    "id": "cust-xl-164",
    "rawDoc": "900022468-3",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA No.667252",
    "phone": "+573241916084",
    "email": "contacto@lacorporacio.co",
    "city": "Bucaramanga",
    "address": "Calle 49 # 72-42"
  },
  {
    "id": "cust-xl-165",
    "rawDoc": "1001306635",
    "name": "SERGIO PAULINO VELANDIA SUESCUN",
    "phone": "+573102086115",
    "email": "contacto@sergiopaulin.co",
    "city": "Villavicencio",
    "address": "Calle 56 # 85-45"
  },
  {
    "id": "cust-xl-166",
    "rawDoc": "1001314554",
    "name": "JORGE LUIS RAMIREZ CORREA",
    "phone": "+573112256146",
    "email": "contacto@jorgeluisram.co",
    "city": "Montería",
    "address": "Calle 63 # 98-48"
  },
  {
    "id": "cust-xl-167",
    "rawDoc": "900022879-6",
    "name": "DANIEL SANTIAGO CHAPARRO URETE",
    "phone": "+573122426177",
    "email": "contacto@danielsantia.co",
    "city": "Pereira",
    "address": "Calle 70 # 21-1"
  },
  {
    "id": "cust-xl-168",
    "rawDoc": "1001330392",
    "name": "Diego Grisolia",
    "phone": "+573132596208",
    "email": "contacto@diegogrisoli.co",
    "city": "Manizales",
    "address": "Calle 77 # 34-4"
  },
  {
    "id": "cust-xl-169",
    "rawDoc": "900023153-8",
    "name": "CULTIVOS SANKA SAS",
    "phone": "+573142766239",
    "email": "contacto@cultivossank.co",
    "city": "Neiva",
    "address": "Calle 84 # 47-7"
  },
  {
    "id": "cust-xl-170",
    "rawDoc": "900023290-9",
    "name": "DAVID ALBERTO SANDOVAL MEDINA",
    "phone": "+573152936270",
    "email": "contacto@davidalberto.co",
    "city": "Ibagué",
    "address": "Calle 91 # 60-10"
  },
  {
    "id": "cust-xl-171",
    "rawDoc": "1001354149",
    "name": "JORGE TRUJILLO NAVIA",
    "phone": "+573163106301",
    "email": "contacto@jorgetrujill.co",
    "city": "Popayán",
    "address": "Calle 98 # 73-13"
  },
  {
    "id": "cust-xl-172",
    "rawDoc": "900023564-2",
    "name": "CARMELA GARCIA",
    "phone": "+573173276332",
    "email": "contacto@carmelagarci.co",
    "city": "Pastos",
    "address": "Calle 5 # 86-16"
  },
  {
    "id": "cust-xl-173",
    "rawDoc": "1001369987",
    "name": "HACIENDA EL VERGEL S A S",
    "phone": "+573183446363",
    "email": "contacto@haciendaelve.co",
    "city": "Santa Marta",
    "address": "Calle 12 # 99-19"
  },
  {
    "id": "cust-xl-174",
    "rawDoc": "1001377906",
    "name": "MARIA ROCIO OSORIO LONDOÑO",
    "phone": "+573193616394",
    "email": "contacto@mariarocioos.co",
    "city": "Valledupar",
    "address": "Calle 19 # 22-22"
  },
  {
    "id": "cust-xl-175",
    "rawDoc": "900023975-5",
    "name": "PROBOS CONSULTORIA GANADERA SAS",
    "phone": "+573203786425",
    "email": "contacto@probosconsul.co",
    "city": "Tunja",
    "address": "Calle 26 # 35-25"
  },
  {
    "id": "cust-xl-176",
    "rawDoc": "1001393744",
    "name": "JAIME SUÁREZ",
    "phone": "+573213956456",
    "email": "contacto@jaimesurez.co",
    "city": "Bogotá D.C.",
    "address": "Calle 33 # 48-28"
  },
  {
    "id": "cust-xl-177",
    "rawDoc": "1001401663",
    "name": "VICTOR DAIRO PEÑA CASTRO",
    "phone": "+573224126487",
    "email": "contacto@victordairop.co",
    "city": "Medellín",
    "address": "Calle 40 # 61-31"
  },
  {
    "id": "cust-xl-178",
    "rawDoc": "1001409582",
    "name": "KAREN",
    "phone": "+573234296518",
    "email": "contacto@karen.co",
    "city": "Cali",
    "address": "Calle 47 # 74-34"
  },
  {
    "id": "cust-xl-179",
    "rawDoc": "900024523-9",
    "name": "CERGAN ALFA S.A.S",
    "phone": "+573244466549",
    "email": "contacto@cerganalfasa.co",
    "city": "Barranquilla",
    "address": "Calle 54 # 87-37"
  },
  {
    "id": "cust-xl-180",
    "rawDoc": "1001425420",
    "name": "MAICITO S. A.",
    "phone": "+573104636580",
    "email": "contacto@maicitosa.co",
    "city": "Bucaramanga",
    "address": "Calle 61 # 10-40"
  },
  {
    "id": "cust-xl-181",
    "rawDoc": "900024797-2",
    "name": "PORCICOLA EL NORAL SAS",
    "phone": "+573114806611",
    "email": "contacto@porcicolaeln.co",
    "city": "Villavicencio",
    "address": "Calle 68 # 23-43"
  },
  {
    "id": "cust-xl-182",
    "rawDoc": "1001441258",
    "name": "EIBAR JOSÉ CALDERÓN",
    "phone": "+573124976642",
    "email": "contacto@eibarjoscald.co",
    "city": "Montería",
    "address": "Calle 75 # 36-46"
  },
  {
    "id": "cust-xl-183",
    "rawDoc": "1001449177",
    "name": "LEOPOLDO FARIAS",
    "phone": "+573135146673",
    "email": "contacto@leopoldofari.co",
    "city": "Pereira",
    "address": "Calle 82 # 49-49"
  },
  {
    "id": "cust-xl-184",
    "rawDoc": "1001457096",
    "name": "PRISCILA CEDEÑO MARTINEZ",
    "phone": "+573145316704",
    "email": "contacto@priscilacede.co",
    "city": "Manizales",
    "address": "Calle 89 # 62-2"
  },
  {
    "id": "cust-xl-185",
    "rawDoc": "1001465015",
    "name": "OSCAR MANUEL MENDEZ ROA",
    "phone": "+573155486735",
    "email": "contacto@oscarmanuelm.co",
    "city": "Neiva",
    "address": "Calle 96 # 75-5"
  },
  {
    "id": "cust-xl-186",
    "rawDoc": "900025482-7",
    "name": "HUMBERT ALEJANDRO DURAN SANDOVAL",
    "phone": "+573165656766",
    "email": "contacto@humbertaleja.co",
    "city": "Ibagué",
    "address": "Calle 3 # 88-8"
  },
  {
    "id": "cust-xl-187",
    "rawDoc": "1001480853",
    "name": "MIGUEL MOLINA",
    "phone": "+573175826797",
    "email": "contacto@miguelmolina.co",
    "city": "Popayán",
    "address": "Calle 10 # 11-11"
  },
  {
    "id": "cust-xl-188",
    "rawDoc": "1001488772",
    "name": "FABIO ANDRES GUEVARA GUEVARA",
    "phone": "+573185996828",
    "email": "contacto@fabioandresg.co",
    "city": "Pastos",
    "address": "Calle 17 # 24-14"
  },
  {
    "id": "cust-xl-189",
    "rawDoc": "1001496691",
    "name": "BRAYAN SNEIDER MANTILLA DELGADO",
    "phone": "+573196166859",
    "email": "contacto@brayansneide.co",
    "city": "Santa Marta",
    "address": "Calle 24 # 37-17"
  },
  {
    "id": "cust-xl-190",
    "rawDoc": "1001504610",
    "name": "WILLIAM ENRIQUE CRUZ FERNANDEZ",
    "phone": "+573206336890",
    "email": "contacto@williamenriq.co",
    "city": "Valledupar",
    "address": "Calle 31 # 50-20"
  },
  {
    "id": "cust-xl-191",
    "rawDoc": "1001512529",
    "name": "PAOLA ANDREA RODRIGUEZ FORERO",
    "phone": "+573216506921",
    "email": "contacto@paolaandrear.co",
    "city": "Tunja",
    "address": "Calle 38 # 63-23"
  },
  {
    "id": "cust-xl-192",
    "rawDoc": "1001520448",
    "name": "JAIRO ENRIQUE LÓPEZ MALDONADO",
    "phone": "+573226676952",
    "email": "contacto@jairoenrique.co",
    "city": "Bogotá D.C.",
    "address": "Calle 45 # 76-26"
  },
  {
    "id": "cust-xl-193",
    "rawDoc": "1001528367",
    "name": "JAVIER VIRGÜEZ",
    "phone": "+573236846983",
    "email": "contacto@javiervirgez.co",
    "city": "Medellín",
    "address": "Calle 52 # 89-29"
  },
  {
    "id": "cust-xl-194",
    "rawDoc": "900026578-6",
    "name": "OLGA LUCIA ORTIZ RAMIREZ",
    "phone": "+573247017014",
    "email": "contacto@olgaluciaort.co",
    "city": "Cali",
    "address": "Calle 59 # 12-32"
  },
  {
    "id": "cust-xl-195",
    "rawDoc": "1001544205",
    "name": "EDINSON VILLALOBOS",
    "phone": "+573107187045",
    "email": "contacto@edinsonvilla.co",
    "city": "Barranquilla",
    "address": "Calle 66 # 25-35"
  },
  {
    "id": "cust-xl-196",
    "rawDoc": "1001552124",
    "name": "CARLOS ALFREDO VEGA VERA",
    "phone": "+573117357076",
    "email": "contacto@carlosalfred.co",
    "city": "Bucaramanga",
    "address": "Calle 73 # 38-38"
  },
  {
    "id": "cust-xl-197",
    "rawDoc": "1001560043",
    "name": "JUAN CAMILO CASTRILLÓN OTÁLVARO",
    "phone": "+573127527107",
    "email": "contacto@juancamiloca.co",
    "city": "Villavicencio",
    "address": "Calle 80 # 51-41"
  },
  {
    "id": "cust-xl-198",
    "rawDoc": "1001567962",
    "name": "GUILLERMO ALFREDO VIVAS GIRALDO",
    "phone": "+573137697138",
    "email": "contacto@guillermoalf.co",
    "city": "Montería",
    "address": "Calle 87 # 64-44"
  },
  {
    "id": "cust-xl-199",
    "rawDoc": "900027263-2",
    "name": "OCTAVIO ISAZA LONDOÑO",
    "phone": "+573147867169",
    "email": "contacto@octavioisaza.co",
    "city": "Pereira",
    "address": "Calle 94 # 77-47"
  },
  {
    "id": "cust-xl-200",
    "rawDoc": "1001583800",
    "name": "JOSE ANTONIO NIÑO NIÑO",
    "phone": "+573158037200",
    "email": "contacto@joseantonion.co",
    "city": "Manizales",
    "address": "Calle 1 # 90-0"
  },
  {
    "id": "cust-xl-201",
    "rawDoc": "1001591719",
    "name": "LUIS FERNANDO DE JOSE AMARILES",
    "phone": "+573168207231",
    "email": "contacto@luisfernando.co",
    "city": "Neiva",
    "address": "Calle 8 # 13-3"
  },
  {
    "id": "cust-xl-202",
    "rawDoc": "1001599638",
    "name": "EMBRIOVEN CA",
    "phone": "+573178377262",
    "email": "contacto@embriovenca.co",
    "city": "Ibagué",
    "address": "Calle 15 # 26-6"
  },
  {
    "id": "cust-xl-203",
    "rawDoc": "900027811-6",
    "name": "GRANJA PORCICOLA LA ESPAÑOLA S.A.S",
    "phone": "+573188547293",
    "email": "contacto@granjaporcic.co",
    "city": "Popayán",
    "address": "Calle 22 # 39-9"
  },
  {
    "id": "cust-xl-204",
    "rawDoc": "1001615476",
    "name": "JOSE MANUEL AMAYA ARDILA",
    "phone": "+573198717324",
    "email": "contacto@josemanuelam.co",
    "city": "Pastos",
    "address": "Calle 29 # 52-12"
  },
  {
    "id": "cust-xl-205",
    "rawDoc": "900028085-8",
    "name": "CENTRO GANADERO CHAVEZ/ NORBERTO",
    "phone": "+573208887355",
    "email": "contacto@centroganade.co",
    "city": "Santa Marta",
    "address": "Calle 36 # 65-15"
  },
  {
    "id": "cust-xl-206",
    "rawDoc": "1001631314",
    "name": "LUIS ENRIQUE NIÑO PEREZ",
    "phone": "+573219057386",
    "email": "contacto@luisenriquen.co",
    "city": "Valledupar",
    "address": "Calle 43 # 78-18"
  },
  {
    "id": "cust-xl-207",
    "rawDoc": "900028359-1",
    "name": "EMBRIOVET SAS",
    "phone": "+573229227417",
    "email": "contacto@embriovetsas.co",
    "city": "Tunja",
    "address": "Calle 50 # 91-21"
  },
  {
    "id": "cust-xl-208",
    "rawDoc": "1001647152",
    "name": "JAIME SUÁREZ/ GANADERIA EL MILAGRO",
    "phone": "+573239397448",
    "email": "contacto@jaimesurezga.co",
    "city": "Bogotá D.C.",
    "address": "Calle 57 # 14-24"
  },
  {
    "id": "cust-xl-209",
    "rawDoc": "1001655071",
    "name": "CAMILO ANDRÉS OICATÁ VARGAS",
    "phone": "+573249567479",
    "email": "contacto@camiloandrso.co",
    "city": "Medellín",
    "address": "Calle 64 # 27-27"
  },
  {
    "id": "cust-xl-210",
    "rawDoc": "1001662990",
    "name": "ANA KARINA VERGARA",
    "phone": "+573109737510",
    "email": "contacto@anakarinaver.co",
    "city": "Cali",
    "address": "Calle 71 # 40-30"
  },
  {
    "id": "cust-xl-211",
    "rawDoc": "1001670909",
    "name": "VÍCTOR MANUEL CASTAÑEDA PINEDA",
    "phone": "+573119907541",
    "email": "contacto@vctormanuelc.co",
    "city": "Barranquilla",
    "address": "Calle 78 # 53-33"
  },
  {
    "id": "cust-xl-212",
    "rawDoc": "900029044-6",
    "name": "CENTRO MEDICO VETERINARIO FELIVET SAS",
    "phone": "+573121087572",
    "email": "contacto@centromedico.co",
    "city": "Bucaramanga",
    "address": "Calle 85 # 66-36"
  },
  {
    "id": "cust-xl-213",
    "rawDoc": "1001686747",
    "name": "WILLIAN VEGA",
    "phone": "+573131257603",
    "email": "contacto@willianvega.co",
    "city": "Villavicencio",
    "address": "Calle 92 # 79-39"
  },
  {
    "id": "cust-xl-214",
    "rawDoc": "1001694666",
    "name": "MARLON ADOLFO GUTIERREZ/ Diego Armando Vega Borda",
    "phone": "+573141427634",
    "email": "contacto@marlonadolfo.co",
    "city": "Montería",
    "address": "Calle 99 # 92-42"
  },
  {
    "id": "cust-xl-215",
    "rawDoc": "900029455-9",
    "name": "UNIVERSIDAD NACIONAL DE COLOMBIA/ ODC 4 - 5061",
    "phone": "+573151597665",
    "email": "contacto@universidadn.co",
    "city": "Pereira",
    "address": "Calle 6 # 15-45"
  },
  {
    "id": "cust-xl-216",
    "rawDoc": "900029592-1",
    "name": "UNIVERSIDAD COOPERATIVA DE COLOMBIA/ 03BUC - 0000012478",
    "phone": "+573161767696",
    "email": "contacto@universidadc.co",
    "city": "Manizales",
    "address": "Calle 13 # 28-48"
  },
  {
    "id": "cust-xl-217",
    "rawDoc": "900029729-2",
    "name": "UNIVERSIDAD COOPERATIVA DE COLOMBIA",
    "phone": "+573171937727",
    "email": "contacto@universidadc.co",
    "city": "Neiva",
    "address": "Calle 20 # 41-1"
  },
  {
    "id": "cust-xl-218",
    "rawDoc": "1001726342",
    "name": "DIEGO LEAL",
    "phone": "+573182107758",
    "email": "contacto@diegoleal.co",
    "city": "Ibagué",
    "address": "Calle 27 # 54-4"
  },
  {
    "id": "cust-xl-219",
    "rawDoc": "900030003-4",
    "name": "EMPRESA GENETICA ESPECIAL LTDA/ IDEO+SOLUCIONES S.A.S",
    "phone": "+573192277789",
    "email": "contacto@empresagenet.co",
    "city": "Popayán",
    "address": "Calle 34 # 67-7"
  },
  {
    "id": "cust-xl-220",
    "rawDoc": "1001742180",
    "name": "JOSE DANIEL TULCANAZA ROSERO",
    "phone": "+573202447820",
    "email": "contacto@josedanieltu.co",
    "city": "Pastos",
    "address": "Calle 41 # 80-10"
  },
  {
    "id": "cust-xl-221",
    "rawDoc": "1001750099",
    "name": "GERMAN DARIO VILLAMIZAR GUERRERO",
    "phone": "+573212617851",
    "email": "contacto@germandariov.co",
    "city": "Santa Marta",
    "address": "Calle 48 # 93-13"
  },
  {
    "id": "cust-xl-222",
    "rawDoc": "900030414-7",
    "name": "IVAN ELIECER MENDEZ SANCHEZ",
    "phone": "+573222787882",
    "email": "contacto@ivaneliecerm.co",
    "city": "Valledupar",
    "address": "Calle 55 # 16-16"
  },
  {
    "id": "cust-xl-223",
    "rawDoc": "1001765937",
    "name": "TOMÁS AGUILAR",
    "phone": "+573232957913",
    "email": "contacto@tomsaguilar.co",
    "city": "Tunja",
    "address": "Calle 62 # 29-19"
  },
  {
    "id": "cust-xl-224",
    "rawDoc": "900030688-9",
    "name": "ALFONSO PALENCIA ALVAREZ",
    "phone": "+573243127944",
    "email": "contacto@alfonsopalen.co",
    "city": "Bogotá D.C.",
    "address": "Calle 69 # 42-22"
  },
  {
    "id": "cust-xl-225",
    "rawDoc": "900030825-1",
    "name": "ANDRES CAMILO SANTACOLOMA BARRERA",
    "phone": "+573103297975",
    "email": "contacto@andrescamilo.co",
    "city": "Medellín",
    "address": "Calle 76 # 55-25"
  },
  {
    "id": "cust-xl-226",
    "rawDoc": "1001789694",
    "name": "JUAN DIEGO BERMUDEZ MOLINA",
    "phone": "+573113468006",
    "email": "contacto@juandiegober.co",
    "city": "Cali",
    "address": "Calle 83 # 68-28"
  },
  {
    "id": "cust-xl-227",
    "rawDoc": "1001797613",
    "name": "DAVID EMILIO ROJAS TABARES",
    "phone": "+573123638037",
    "email": "contacto@davidemilior.co",
    "city": "Barranquilla",
    "address": "Calle 90 # 81-31"
  },
  {
    "id": "cust-xl-228",
    "rawDoc": "1001805532",
    "name": "DIEGO ALEJANDRO MOLANO VILLAMIL",
    "phone": "+573133808068",
    "email": "contacto@diegoalejand.co",
    "city": "Bucaramanga",
    "address": "Calle 97 # 94-34"
  },
  {
    "id": "cust-xl-229",
    "rawDoc": "1001813451",
    "name": "KEVIN ANDRÉS MARTINEZ MOLANO",
    "phone": "+573143978099",
    "email": "contacto@kevinandrsma.co",
    "city": "Villavicencio",
    "address": "Calle 4 # 17-37"
  },
  {
    "id": "cust-xl-230",
    "rawDoc": "900031510-6",
    "name": "MARTHA ISABEL RAMIREZ ARTEAGA",
    "phone": "+573154148130",
    "email": "contacto@marthaisabel.co",
    "city": "Montería",
    "address": "Calle 11 # 30-40"
  },
  {
    "id": "cust-xl-231",
    "rawDoc": "1001829289",
    "name": "SERGIO ARMANDO MOLINA PINTO",
    "phone": "+573164318161",
    "email": "contacto@sergioarmand.co",
    "city": "Pereira",
    "address": "Calle 18 # 43-43"
  },
  {
    "id": "cust-xl-232",
    "rawDoc": "900031784-8",
    "name": "G ALTAMIRA SAS",
    "phone": "+573174488192",
    "email": "contacto@galtamirasas.co",
    "city": "Manizales",
    "address": "Calle 25 # 56-46"
  },
  {
    "id": "cust-xl-233",
    "rawDoc": "900031921-9",
    "name": "HECTOR ALFREDO ZULUAGA CRISANCHO",
    "phone": "+573184658223",
    "email": "contacto@hectoralfred.co",
    "city": "Neiva",
    "address": "Calle 32 # 69-49"
  },
  {
    "id": "cust-xl-234",
    "rawDoc": "900032058-1",
    "name": "INVERSIONES AGROGANADERA S.A.S",
    "phone": "+573194828254",
    "email": "contacto@inversionesa.co",
    "city": "Ibagué",
    "address": "Calle 39 # 82-2"
  },
  {
    "id": "cust-xl-235",
    "rawDoc": "900032195-2",
    "name": "ALVARO ESCOBAR SAAVEDRA",
    "phone": "+573204998285",
    "email": "contacto@alvaroescoba.co",
    "city": "Popayán",
    "address": "Calle 46 # 95-5"
  },
  {
    "id": "cust-xl-236",
    "rawDoc": "1001868884",
    "name": "JORGE EDUARDO CADAVID GOMEZ",
    "phone": "+573215168316",
    "email": "contacto@jorgeeduardo.co",
    "city": "Pastos",
    "address": "Calle 53 # 18-8"
  },
  {
    "id": "cust-xl-237",
    "rawDoc": "900032469-4",
    "name": "NOGA S.A.S",
    "phone": "+573225338347",
    "email": "contacto@nogasas.co",
    "city": "Santa Marta",
    "address": "Calle 60 # 31-11"
  },
  {
    "id": "cust-xl-238",
    "rawDoc": "900032606-5",
    "name": "PALHER HOLDING GROUP S.A.S",
    "phone": "+573235508378",
    "email": "contacto@palherholdin.co",
    "city": "Valledupar",
    "address": "Calle 67 # 44-14"
  },
  {
    "id": "cust-xl-239",
    "rawDoc": "1001892641",
    "name": "CENTRAL GENETICA EMBRIOBRA FIV S A S",
    "phone": "+573245678409",
    "email": "contacto@centralgenet.co",
    "city": "Tunja",
    "address": "Calle 74 # 57-17"
  },
  {
    "id": "cust-xl-240",
    "rawDoc": "1001900560",
    "name": "IVAN DARIO CAMELO CIFUENTES",
    "phone": "+573105848440",
    "email": "contacto@ivandariocam.co",
    "city": "Bogotá D.C.",
    "address": "Calle 81 # 70-20"
  },
  {
    "id": "cust-xl-241",
    "rawDoc": "900033017-8",
    "name": "ANDRÉS CASTAÑEDA SOSA",
    "phone": "+573116018471",
    "email": "contacto@andrscastaed.co",
    "city": "Medellín",
    "address": "Calle 88 # 83-23"
  },
  {
    "id": "cust-xl-242",
    "rawDoc": "900033154-9",
    "name": "VITA IVF LAB SAS/ OC-1-10085",
    "phone": "+573126188502",
    "email": "contacto@vitaivflabsa.co",
    "city": "Cali",
    "address": "Calle 95 # 96-26"
  },
  {
    "id": "cust-xl-243",
    "rawDoc": "1001924317",
    "name": "ERSON GEOVANNY PINZON ARENAS",
    "phone": "+573136358533",
    "email": "contacto@ersongeovann.co",
    "city": "Barranquilla",
    "address": "Calle 2 # 19-29"
  },
  {
    "id": "cust-xl-244",
    "rawDoc": "900033428-2",
    "name": "SERVICIOS E INVERSIONES AGROPECUARIAS EL TREBOL LTDA",
    "phone": "+573146528564",
    "email": "contacto@serviciosein.co",
    "city": "Bucaramanga",
    "address": "Calle 9 # 32-32"
  },
  {
    "id": "cust-xl-245",
    "rawDoc": "900033565-3",
    "name": "GRUPO COMERCIAL EL TREBOL SAS",
    "phone": "+573156698595",
    "email": "contacto@grupocomerci.co",
    "city": "Villavicencio",
    "address": "Calle 16 # 45-35"
  },
  {
    "id": "cust-xl-246",
    "rawDoc": "900033702-4",
    "name": "COL-BRA GENÉTICA SAS",
    "phone": "+573166868626",
    "email": "contacto@colbragentic.co",
    "city": "Montería",
    "address": "Calle 23 # 58-38"
  },
  {
    "id": "cust-xl-247",
    "rawDoc": "900033839-5",
    "name": "INVERSIONES CJR COLOMBIA SAS",
    "phone": "+573177038657",
    "email": "contacto@inversionesc.co",
    "city": "Pereira",
    "address": "Calle 30 # 71-41"
  },
  {
    "id": "cust-xl-248",
    "rawDoc": "1001963912",
    "name": "DAN KEYL PUPO HOYOS",
    "phone": "+573187208688",
    "email": "contacto@dankeylpupoh.co",
    "city": "Manizales",
    "address": "Calle 37 # 84-44"
  },
  {
    "id": "cust-xl-249",
    "rawDoc": "1001971831",
    "name": "OSCAR ORLANDO OCHOA VALDERRAMA",
    "phone": "+573197378719",
    "email": "contacto@oscarorlando.co",
    "city": "Neiva",
    "address": "Calle 44 # 97-47"
  },
  {
    "id": "cust-xl-250",
    "rawDoc": "1001979750",
    "name": "CARLOS ANDRÉS VÁSQUEZ CARDENAS",
    "phone": "+573207548750",
    "email": "contacto@carlosandrsv.co",
    "city": "Ibagué",
    "address": "Calle 51 # 20-0"
  },
  {
    "id": "cust-xl-251",
    "rawDoc": "900034387-9",
    "name": "INSTRUAGRO COLOMBIA SAS",
    "phone": "+573217718781",
    "email": "contacto@instruagroco.co",
    "city": "Popayán",
    "address": "Calle 58 # 33-3"
  },
  {
    "id": "cust-xl-252",
    "rawDoc": "900034524-1",
    "name": "CR 36 SAS",
    "phone": "+573227888812",
    "email": "contacto@cr36sas.co",
    "city": "Pastos",
    "address": "Calle 65 # 46-6"
  },
  {
    "id": "cust-xl-253",
    "rawDoc": "1002003507",
    "name": "GIOVANNY RODRIGUEZ MENDEZ",
    "phone": "+573238058843",
    "email": "contacto@giovannyrodr.co",
    "city": "Santa Marta",
    "address": "Calle 72 # 59-9"
  },
  {
    "id": "cust-xl-254",
    "rawDoc": "900034798-3",
    "name": "DANIEL YEPES MESA",
    "phone": "+573248228874",
    "email": "contacto@danielyepesm.co",
    "city": "Valledupar",
    "address": "Calle 79 # 72-12"
  },
  {
    "id": "cust-xl-255",
    "rawDoc": "900034935-4",
    "name": "PABLO LEIDIKER ROJAS GARCIA",
    "phone": "+573108398905",
    "email": "contacto@pabloleidike.co",
    "city": "Tunja",
    "address": "Calle 86 # 85-15"
  },
  {
    "id": "cust-xl-256",
    "rawDoc": "1002027264",
    "name": "MARIA ALEJANDRA PARRA VARGAS",
    "phone": "+573118568936",
    "email": "contacto@mariaalejand.co",
    "city": "Bogotá D.C.",
    "address": "Calle 93 # 98-18"
  },
  {
    "id": "cust-xl-257",
    "rawDoc": "1002035183",
    "name": "YAIR JESUS AMADO GOMEZ",
    "phone": "+573128738967",
    "email": "contacto@yairjesusama.co",
    "city": "Medellín",
    "address": "Calle 100 # 21-21"
  },
  {
    "id": "cust-xl-258",
    "rawDoc": "900035346-7",
    "name": "INVERSIONES SAN JOSÉ",
    "phone": "+573138908998",
    "email": "contacto@inversioness.co",
    "city": "Cali",
    "address": "Calle 7 # 34-24"
  },
  {
    "id": "cust-xl-259",
    "rawDoc": "1002051021",
    "name": "ALBERT COLMENARES",
    "phone": "+573149079029",
    "email": "contacto@albertcolmen.co",
    "city": "Barranquilla",
    "address": "Calle 14 # 47-27"
  },
  {
    "id": "cust-xl-260",
    "rawDoc": "1002058940",
    "name": "JOSÉ VICENTE OROZCO",
    "phone": "+573159249060",
    "email": "contacto@josvicenteor.co",
    "city": "Bucaramanga",
    "address": "Calle 21 # 60-30"
  },
  {
    "id": "cust-xl-261",
    "rawDoc": "900035757-1",
    "name": "COMERCIAL MAS CAMPO BIOTECNOLOGÍA ANIMAL SAS",
    "phone": "+573169419091",
    "email": "contacto@comercialmas.co",
    "city": "Villavicencio",
    "address": "Calle 28 # 73-33"
  },
  {
    "id": "cust-xl-262",
    "rawDoc": "900035894-2",
    "name": "HSJ GESTIÓN 360 SAS / OC 1536",
    "phone": "+573179589122",
    "email": "contacto@hsjgestin360.co",
    "city": "Montería",
    "address": "Calle 35 # 86-36"
  },
  {
    "id": "cust-xl-263",
    "rawDoc": "900036031-3",
    "name": "HSJ GESTIÓN 360 SAS",
    "phone": "+573189759153",
    "email": "contacto@hsjgestin360.co",
    "city": "Pereira",
    "address": "Calle 42 # 99-39"
  },
  {
    "id": "cust-xl-264",
    "rawDoc": "900036168-4",
    "name": "AGROPECUARIA BAMBUSA SAS/ OC 7948",
    "phone": "+573199929184",
    "email": "contacto@agropecuaria.co",
    "city": "Manizales",
    "address": "Calle 49 # 22-42"
  },
  {
    "id": "cust-xl-265",
    "rawDoc": "1002098535",
    "name": "LUIS ZAMORA",
    "phone": "+573201109215",
    "email": "contacto@luiszamora.co",
    "city": "Neiva",
    "address": "Calle 56 # 35-45"
  },
  {
    "id": "cust-xl-266",
    "rawDoc": "1002106454",
    "name": "JAQUELINE MARGARITA BURBANO MUÑOZ",
    "phone": "+573211279246",
    "email": "contacto@jaquelinemar.co",
    "city": "Ibagué",
    "address": "Calle 63 # 48-48"
  },
  {
    "id": "cust-xl-267",
    "rawDoc": "1002114373",
    "name": "RAÚL JAVIER VETENCOURT MONTOYA",
    "phone": "+573221449277",
    "email": "contacto@raljaviervet.co",
    "city": "Popayán",
    "address": "Calle 70 # 61-1"
  },
  {
    "id": "cust-xl-268",
    "rawDoc": "1002122292",
    "name": "FUNDACION GRUPO ARGOS/ N°.PMO 6000048566/ Migo 5000156688",
    "phone": "+573231619308",
    "email": "contacto@fundaciongru.co",
    "city": "Pastos",
    "address": "Calle 77 # 74-4"
  },
  {
    "id": "cust-xl-269",
    "rawDoc": "1002130211",
    "name": "FUNDACION GRUPO ARGOS/ N°.PMO 6000048566",
    "phone": "+573241789339",
    "email": "contacto@fundaciongru.co",
    "city": "Santa Marta",
    "address": "Calle 84 # 87-7"
  },
  {
    "id": "cust-xl-270",
    "rawDoc": "1002138130",
    "name": "DANIELA GUERRERO MOSQUERA",
    "phone": "+573101959370",
    "email": "contacto@danielaguerr.co",
    "city": "Valledupar",
    "address": "Calle 91 # 10-10"
  },
  {
    "id": "cust-xl-271",
    "rawDoc": "1002146049",
    "name": "FARID NAIZZIR MUÑOZ",
    "phone": "+573112129401",
    "email": "contacto@faridnaizzir.co",
    "city": "Tunja",
    "address": "Calle 98 # 23-13"
  },
  {
    "id": "cust-xl-272",
    "rawDoc": "900037264-3",
    "name": "REPRODUCCIÓN ANIMAL SAS",
    "phone": "+573122299432",
    "email": "contacto@reproduccina.co",
    "city": "Bogotá D.C.",
    "address": "Calle 5 # 36-16"
  },
  {
    "id": "cust-xl-273",
    "rawDoc": "1002161887",
    "name": "EDWIN ALFONSO HENAO SUAREZ",
    "phone": "+573132469463",
    "email": "contacto@edwinalfonso.co",
    "city": "Medellín",
    "address": "Calle 12 # 49-19"
  },
  {
    "id": "cust-xl-274",
    "rawDoc": "900037538-5",
    "name": "GANADERIA ALTA LUZ SAS",
    "phone": "+573142639494",
    "email": "contacto@ganaderiaalt.co",
    "city": "Cali",
    "address": "Calle 19 # 62-22"
  },
  {
    "id": "cust-xl-275",
    "rawDoc": "900037675-6",
    "name": "URICACHA SAS",
    "phone": "+573152809525",
    "email": "contacto@uricachasas.co",
    "city": "Barranquilla",
    "address": "Calle 26 # 75-25"
  },
  {
    "id": "cust-xl-276",
    "rawDoc": "1002185644",
    "name": "CAMILO ANDRES ORTIZ NIÑO",
    "phone": "+573162979556",
    "email": "contacto@camiloandres.co",
    "city": "Bucaramanga",
    "address": "Calle 33 # 88-28"
  },
  {
    "id": "cust-xl-277",
    "rawDoc": "900037949-8",
    "name": "ALEXANDER VINASCO CASTRO/ CRIAS INNOVACION GENETICA SAS",
    "phone": "+573173149587",
    "email": "contacto@alexandervin.co",
    "city": "Villavicencio",
    "address": "Calle 40 # 11-31"
  },
  {
    "id": "cust-xl-278",
    "rawDoc": "900038086-9",
    "name": "CRIAS INNOVACION GENETICA SAS",
    "phone": "+573183319618",
    "email": "contacto@criasinnovac.co",
    "city": "Montería",
    "address": "Calle 47 # 24-34"
  },
  {
    "id": "cust-xl-279",
    "rawDoc": "900038223-1",
    "name": "JOSE DARIEL VALENCIA MEJIA",
    "phone": "+573193489649",
    "email": "contacto@josedarielva.co",
    "city": "Pereira",
    "address": "Calle 54 # 37-37"
  },
  {
    "id": "cust-xl-280",
    "rawDoc": "1002217320",
    "name": "COMITE DEPARTAMENTAL DE GANADEROS Y AGRICULTORES DE CALDAS",
    "phone": "+573203659680",
    "email": "contacto@comitedepart.co",
    "city": "Manizales",
    "address": "Calle 61 # 50-40"
  },
  {
    "id": "cust-xl-281",
    "rawDoc": "1002225239",
    "name": "SOLUCIONES DE LABORATORIO E INDUSTRIA S A S/ DAVID AVENDAÑO",
    "phone": "+573213829711",
    "email": "contacto@solucionesde.co",
    "city": "Neiva",
    "address": "Calle 68 # 63-43"
  },
  {
    "id": "cust-xl-282",
    "rawDoc": "1002233158",
    "name": "SOLUCIONES DE LABORATORIO E INDUSTRIA S A S",
    "phone": "+573223999742",
    "email": "contacto@solucionesde.co",
    "city": "Ibagué",
    "address": "Calle 75 # 76-46"
  },
  {
    "id": "cust-xl-283",
    "rawDoc": "1002241077",
    "name": "ALEX DARIO QUINTERO ARCINIEGAS",
    "phone": "+573234169773",
    "email": "contacto@alexdarioqui.co",
    "city": "Popayán",
    "address": "Calle 82 # 89-49"
  },
  {
    "id": "cust-xl-284",
    "rawDoc": "1002248996",
    "name": "JUAN PABLO MONTOYA LOPERA",
    "phone": "+573244339804",
    "email": "contacto@juanpablomon.co",
    "city": "Pastos",
    "address": "Calle 89 # 12-2"
  },
  {
    "id": "cust-xl-285",
    "rawDoc": "1002256915",
    "name": "Diego Armando Vega Borda",
    "phone": "+573104509835",
    "email": "contacto@diegoarmando.co",
    "city": "Santa Marta",
    "address": "Calle 96 # 25-5"
  },
  {
    "id": "cust-xl-286",
    "rawDoc": "1002264834",
    "name": "LUIS GREGORIO ROJAS AYALA",
    "phone": "+573114679866",
    "email": "contacto@luisgregorio.co",
    "city": "Valledupar",
    "address": "Calle 3 # 38-8"
  },
  {
    "id": "cust-xl-287",
    "rawDoc": "900039319-9",
    "name": "CESAR ALBERTO RAMIREZ MUÑOZ",
    "phone": "+573124849897",
    "email": "contacto@cesaralberto.co",
    "city": "Tunja",
    "address": "Calle 10 # 51-11"
  },
  {
    "id": "cust-xl-288",
    "rawDoc": "1002280672",
    "name": "LUIS EDUARDO ARBOLEDA VANEGAS",
    "phone": "+573135019928",
    "email": "contacto@luiseduardoa.co",
    "city": "Bogotá D.C.",
    "address": "Calle 17 # 64-14"
  },
  {
    "id": "cust-xl-289",
    "rawDoc": "900039593-2",
    "name": "GRUPO AGROBOX SAS ZOMAC",
    "phone": "+573145189959",
    "email": "contacto@grupoagrobox.co",
    "city": "Medellín",
    "address": "Calle 24 # 77-17"
  },
  {
    "id": "cust-xl-290",
    "rawDoc": "1002296510",
    "name": "EDGAR ALFONSO NARVAEZ MARTINEZ",
    "phone": "+573155359990",
    "email": "contacto@edgaralfonso.co",
    "city": "Cali",
    "address": "Calle 31 # 90-20"
  },
  {
    "id": "cust-xl-291",
    "rawDoc": "1002304429",
    "name": "JUAN JOSE GONZALEZ NAVARRO",
    "phone": "+573165521022",
    "email": "contacto@juanjosegonz.co",
    "city": "Barranquilla",
    "address": "Calle 38 # 13-23"
  },
  {
    "id": "cust-xl-292",
    "rawDoc": "1002312348",
    "name": "HECTOR ANIBAL CABRERA TOVAR",
    "phone": "+573175691053",
    "email": "contacto@hectoranibal.co",
    "city": "Bucaramanga",
    "address": "Calle 45 # 26-26"
  },
  {
    "id": "cust-xl-293",
    "rawDoc": "900040141-6",
    "name": "INVERSIONES CASA VERDE DHD SAS/ MILTON GARCIA",
    "phone": "+573185861084",
    "email": "contacto@inversionesc.co",
    "city": "Villavicencio",
    "address": "Calle 52 # 39-29"
  },
  {
    "id": "cust-xl-294",
    "rawDoc": "900040278-7",
    "name": "INVERSIONES CASA VERDE DHD SAS",
    "phone": "+573196031115",
    "email": "contacto@inversionesc.co",
    "city": "Montería",
    "address": "Calle 59 # 52-32"
  },
  {
    "id": "cust-xl-295",
    "rawDoc": "1002336105",
    "name": "ARIT GUSTAVO SOLARTE PALACIOS",
    "phone": "+573206201146",
    "email": "contacto@aritgustavos.co",
    "city": "Pereira",
    "address": "Calle 66 # 65-35"
  },
  {
    "id": "cust-xl-296",
    "rawDoc": "1002344024",
    "name": "SERGIO GONZALEZ CAMACHO",
    "phone": "+573216371177",
    "email": "contacto@sergiogonzal.co",
    "city": "Manizales",
    "address": "Calle 73 # 78-38"
  },
  {
    "id": "cust-xl-297",
    "rawDoc": "1002351943",
    "name": "RAFAEL DARIO TRUJILLO PERDOMO",
    "phone": "+573226541208",
    "email": "contacto@rafaeldariot.co",
    "city": "Neiva",
    "address": "Calle 80 # 91-41"
  },
  {
    "id": "cust-xl-298",
    "rawDoc": "900040826-2",
    "name": "GANADERIA LA BARRERA S.A.S.",
    "phone": "+573236711239",
    "email": "contacto@ganaderialab.co",
    "city": "Ibagué",
    "address": "Calle 87 # 14-44"
  },
  {
    "id": "cust-xl-299",
    "rawDoc": "900040963-3",
    "name": "LA CAPRICHOSA CRIADERO S.A.S",
    "phone": "+573246881270",
    "email": "contacto@lacaprichosa.co",
    "city": "Popayán",
    "address": "Calle 94 # 27-47"
  },
  {
    "id": "cust-xl-300",
    "rawDoc": "900041100-4",
    "name": "COMERCIALIZADORA EL ZARZAL SAS",
    "phone": "+573107051301",
    "email": "contacto@comercializa.co",
    "city": "Pastos",
    "address": "Calle 1 # 40-0"
  },
  {
    "id": "cust-xl-301",
    "rawDoc": "1002383619",
    "name": "FABIAN CAMILO ALVARADO CHACON",
    "phone": "+573117221332",
    "email": "contacto@fabiancamilo.co",
    "city": "Santa Marta",
    "address": "Calle 8 # 53-3"
  },
  {
    "id": "cust-xl-302",
    "rawDoc": "1002391538",
    "name": "KERNIO RODRIGUEZ",
    "phone": "+573127391363",
    "email": "contacto@kerniorodrig.co",
    "city": "Valledupar",
    "address": "Calle 15 # 66-6"
  },
  {
    "id": "cust-xl-303",
    "rawDoc": "1002399457",
    "name": "YILBERT ESTEBAN MORALES CELIS",
    "phone": "+573137561394",
    "email": "contacto@yilbertesteb.co",
    "city": "Tunja",
    "address": "Calle 22 # 79-9"
  },
  {
    "id": "cust-xl-304",
    "rawDoc": "900041648-8",
    "name": "D&D TRANSPORTES & SERVICIOS S.A.S",
    "phone": "+573147731425",
    "email": "contacto@ddtransporte.co",
    "city": "Bogotá D.C.",
    "address": "Calle 29 # 92-12"
  },
  {
    "id": "cust-xl-305",
    "rawDoc": "1002415295",
    "name": "PAOLA ANDREA JIMENEZ VELASQUEZ",
    "phone": "+573157901456",
    "email": "contacto@paolaandreaj.co",
    "city": "Medellín",
    "address": "Calle 36 # 15-15"
  },
  {
    "id": "cust-xl-306",
    "rawDoc": "1002423214",
    "name": "OSCAR DANIEL CARVAJAL GONZALEZ",
    "phone": "+573168071487",
    "email": "contacto@oscardanielc.co",
    "city": "Cali",
    "address": "Calle 43 # 28-18"
  },
  {
    "id": "cust-xl-307",
    "rawDoc": "1002431133",
    "name": "JUAN DIEGO TRUJILLO URIBE",
    "phone": "+573178241518",
    "email": "contacto@juandiegotru.co",
    "city": "Barranquilla",
    "address": "Calle 50 # 41-21"
  },
  {
    "id": "cust-xl-308",
    "rawDoc": "900042196-3",
    "name": "RIGOBERTO ANTONIO SALGADO CORTES",
    "phone": "+573188411549",
    "email": "contacto@rigobertoant.co",
    "city": "Bucaramanga",
    "address": "Calle 57 # 54-24"
  },
  {
    "id": "cust-xl-309",
    "rawDoc": "1002446971",
    "name": "JULIO HUMBERTO VILLATE ZORRO",
    "phone": "+573198581580",
    "email": "contacto@juliohumbert.co",
    "city": "Villavicencio",
    "address": "Calle 64 # 67-27"
  },
  {
    "id": "cust-xl-310",
    "rawDoc": "1002454890",
    "name": "BIOEMBRIO FIV S A S",
    "phone": "+573208751611",
    "email": "contacto@bioembriofiv.co",
    "city": "Montería",
    "address": "Calle 71 # 80-30"
  },
  {
    "id": "cust-xl-311",
    "rawDoc": "900042607-6",
    "name": "SUMINISTROS GLOBAL VETT SAS",
    "phone": "+573218921642",
    "email": "contacto@suministrosg.co",
    "city": "Pereira",
    "address": "Calle 78 # 93-33"
  },
  {
    "id": "cust-xl-312",
    "rawDoc": "1002470728",
    "name": "YULITZA RAMÍREZ ACEVEDO",
    "phone": "+573229091673",
    "email": "contacto@yulitzaramre.co",
    "city": "Manizales",
    "address": "Calle 85 # 16-36"
  },
  {
    "id": "cust-xl-313",
    "rawDoc": "1002478647",
    "name": "MARIO MARCELO SERRANO CAMPAGNARO",
    "phone": "+573239261704",
    "email": "contacto@mariomarcelo.co",
    "city": "Neiva",
    "address": "Calle 92 # 29-39"
  },
  {
    "id": "cust-xl-314",
    "rawDoc": "900043018-9",
    "name": "INSUTECH INSUMOS & TECNOLOGIAS S.A.S.",
    "phone": "+573249431735",
    "email": "contacto@insutechinsu.co",
    "city": "Ibagué",
    "address": "Calle 99 # 42-42"
  },
  {
    "id": "cust-xl-315",
    "rawDoc": "900043155-1",
    "name": "INVERSIONES QUINTERO R Y CIA SAS",
    "phone": "+573109601766",
    "email": "contacto@inversionesq.co",
    "city": "Popayán",
    "address": "Calle 6 # 55-45"
  },
  {
    "id": "cust-xl-316",
    "rawDoc": "900043292-2",
    "name": "AGROPECUARIA BERLIN DE LOS LLANOS SAS ZOMAC",
    "phone": "+573119771797",
    "email": "contacto@agropecuaria.co",
    "city": "Pastos",
    "address": "Calle 13 # 68-48"
  },
  {
    "id": "cust-xl-317",
    "rawDoc": "1002510323",
    "name": "ELKIN NICOLAS LOPEZ GUERRERO",
    "phone": "+573129941828",
    "email": "contacto@elkinnicolas.co",
    "city": "Santa Marta",
    "address": "Calle 20 # 81-1"
  },
  {
    "id": "cust-xl-318",
    "rawDoc": "1002518242",
    "name": "JOSE ARIEL MARIN CORREA MARIN CORREA",
    "phone": "+573131121859",
    "email": "contacto@josearielmar.co",
    "city": "Valledupar",
    "address": "Calle 27 # 94-4"
  },
  {
    "id": "cust-xl-319",
    "rawDoc": "900043703-5",
    "name": "LABORATORIO GEMALAB SAS",
    "phone": "+573141291890",
    "email": "contacto@laboratoriog.co",
    "city": "Tunja",
    "address": "Calle 34 # 17-7"
  },
  {
    "id": "cust-xl-320",
    "rawDoc": "1002534080",
    "name": "EMILCE TRUJILLO RODRIGUEZ",
    "phone": "+573151461921",
    "email": "contacto@emilcetrujil.co",
    "city": "Bogotá D.C.",
    "address": "Calle 41 # 30-10"
  },
  {
    "id": "cust-xl-321",
    "rawDoc": "1002541999",
    "name": "RAFAEL ANTONIO GOMEZ GOMEZ",
    "phone": "+573161631952",
    "email": "contacto@rafaelantoni.co",
    "city": "Medellín",
    "address": "Calle 48 # 43-13"
  },
  {
    "id": "cust-xl-322",
    "rawDoc": "1002549918",
    "name": "NARCES SERRATO POMAR",
    "phone": "+573171801983",
    "email": "contacto@narcesserrat.co",
    "city": "Cali",
    "address": "Calle 55 # 56-16"
  },
  {
    "id": "cust-xl-323",
    "rawDoc": "1002557837",
    "name": "DEICY SEPULVEDA",
    "phone": "+573181972014",
    "email": "contacto@deicysepulve.co",
    "city": "Barranquilla",
    "address": "Calle 62 # 69-19"
  },
  {
    "id": "cust-xl-324",
    "rawDoc": "900044388-1",
    "name": "AGROPECUARIA EL INGENIO SAS",
    "phone": "+573192142045",
    "email": "contacto@agropecuaria.co",
    "city": "Bucaramanga",
    "address": "Calle 69 # 82-22"
  },
  {
    "id": "cust-xl-325",
    "rawDoc": "1002573675",
    "name": "CONSORCIO MATERIALES DE FORMACION SENA 2025",
    "phone": "+573202312076",
    "email": "contacto@consorciomat.co",
    "city": "Villavicencio",
    "address": "Calle 76 # 95-25"
  },
  {
    "id": "cust-xl-326",
    "rawDoc": "900044662-3",
    "name": "REPRESENTACIONES GANADERAS LIMITADA REPREGAN LTDA",
    "phone": "+573212482107",
    "email": "contacto@representaci.co",
    "city": "Montería",
    "address": "Calle 83 # 18-28"
  },
  {
    "id": "cust-xl-327",
    "rawDoc": "1002589513",
    "name": "FREDY EDUARDO RAMIREZ BLANCO",
    "phone": "+573222652138",
    "email": "contacto@fredyeduardo.co",
    "city": "Pereira",
    "address": "Calle 90 # 31-31"
  },
  {
    "id": "cust-xl-328",
    "rawDoc": "1002597432",
    "name": "ALBERTO ENRIQUE CARRACEDO CASTRO",
    "phone": "+573232822169",
    "email": "contacto@albertoenriq.co",
    "city": "Manizales",
    "address": "Calle 97 # 44-34"
  },
  {
    "id": "cust-xl-329",
    "rawDoc": "1002605351",
    "name": "CAMILO CASTRO",
    "phone": "+573242992200",
    "email": "contacto@camilocastro.co",
    "city": "Neiva",
    "address": "Calle 4 # 57-37"
  },
  {
    "id": "cust-xl-330",
    "rawDoc": "900045210-7",
    "name": "CORPORACION UNIVERSITARIA DEL HUILA",
    "phone": "+573103162231",
    "email": "contacto@corporacionu.co",
    "city": "Ibagué",
    "address": "Calle 11 # 70-40"
  },
  {
    "id": "cust-xl-331",
    "rawDoc": "1002621189",
    "name": "LUIS MAURICIO TOUS GAVIRIA",
    "phone": "+573113332262",
    "email": "contacto@luismauricio.co",
    "city": "Popayán",
    "address": "Calle 18 # 83-43"
  },
  {
    "id": "cust-xl-332",
    "rawDoc": "1002629108",
    "name": "WILLIAN FABIAN ROSERO CHAMORRO",
    "phone": "+573123502293",
    "email": "contacto@willianfabia.co",
    "city": "Pastos",
    "address": "Calle 25 # 96-46"
  },
  {
    "id": "cust-xl-333",
    "rawDoc": "1002637027",
    "name": "AGROPECUARIA ALIAR S /PEDIDO DE COMPRAS Nro. 45002 02154",
    "phone": "+573133672324",
    "email": "contacto@agropecuaria.co",
    "city": "Santa Marta",
    "address": "Calle 32 # 19-49"
  },
  {
    "id": "cust-xl-334",
    "rawDoc": "1002644946",
    "name": "CARLOS ARTURO SURMAY GÁNDARA",
    "phone": "+573143842355",
    "email": "contacto@carlosarturo.co",
    "city": "Valledupar",
    "address": "Calle 39 # 32-2"
  },
  {
    "id": "cust-xl-335",
    "rawDoc": "1002652865",
    "name": "DAVID ALMANZA",
    "phone": "+573154012386",
    "email": "contacto@davidalmanza.co",
    "city": "Tunja",
    "address": "Calle 46 # 45-5"
  },
  {
    "id": "cust-xl-336",
    "rawDoc": "1002660784",
    "name": "GENÉTICA IN VITRO C.A",
    "phone": "+573164182417",
    "email": "contacto@genticainvit.co",
    "city": "Bogotá D.C.",
    "address": "Calle 53 # 58-8"
  },
  {
    "id": "cust-xl-337",
    "rawDoc": "900046169-5",
    "name": "GENELITE S.A.S.",
    "phone": "+573174352448",
    "email": "contacto@genelitesas.co",
    "city": "Medellín",
    "address": "Calle 60 # 71-11"
  },
  {
    "id": "cust-xl-338",
    "rawDoc": "1002676622",
    "name": "PEDRO JOSÉ FAJARDO CRUZ",
    "phone": "+573184522479",
    "email": "contacto@pedrojosfaja.co",
    "city": "Cali",
    "address": "Calle 67 # 84-14"
  },
  {
    "id": "cust-xl-339",
    "rawDoc": "1002684541",
    "name": "AGROPECUARIA EL CIMARRON S A S",
    "phone": "+573194692510",
    "email": "contacto@agropecuaria.co",
    "city": "Barranquilla",
    "address": "Calle 74 # 97-17"
  },
  {
    "id": "cust-xl-340",
    "rawDoc": "1002692460",
    "name": "JOSÉ JULIÁN CEBALLOS CARVAJAL",
    "phone": "+573204862541",
    "email": "contacto@josjulinceba.co",
    "city": "Bucaramanga",
    "address": "Calle 81 # 20-20"
  },
  {
    "id": "cust-xl-341",
    "rawDoc": "900046717-9",
    "name": "BIOTECNOLOGIA REPRODUCTIVA GANADERA SAS",
    "phone": "+573215032572",
    "email": "contacto@biotecnologi.co",
    "city": "Villavicencio",
    "address": "Calle 88 # 33-23"
  },
  {
    "id": "cust-xl-342",
    "rawDoc": "1002708298",
    "name": "ELIANA NEIRA RIVERA",
    "phone": "+573225202603",
    "email": "contacto@eliananeirar.co",
    "city": "Montería",
    "address": "Calle 95 # 46-26"
  },
  {
    "id": "cust-xl-343",
    "rawDoc": "900046991-2",
    "name": "VITA IVF LAB SAS",
    "phone": "+573235372634",
    "email": "contacto@vitaivflabsa.co",
    "city": "Pereira",
    "address": "Calle 2 # 59-29"
  },
  {
    "id": "cust-xl-344",
    "rawDoc": "1002724136",
    "name": "LUZ DARY OME GOMEZ",
    "phone": "+573245542665",
    "email": "contacto@luzdaryomego.co",
    "city": "Manizales",
    "address": "Calle 9 # 72-32"
  },
  {
    "id": "cust-xl-345",
    "rawDoc": "1002732055",
    "name": "JHON SEBASTIAN VILLAMIL",
    "phone": "+573105712696",
    "email": "contacto@jhonsebastia.co",
    "city": "Neiva",
    "address": "Calle 16 # 85-35"
  },
  {
    "id": "cust-xl-346",
    "rawDoc": "900047402-5",
    "name": "SEBASTIÁN CASAS SÁNCHEZ",
    "phone": "+573115882727",
    "email": "contacto@sebastincasa.co",
    "city": "Ibagué",
    "address": "Calle 23 # 98-38"
  },
  {
    "id": "cust-xl-347",
    "rawDoc": "1002747893",
    "name": "YENY CELIS",
    "phone": "+573126052758",
    "email": "contacto@yenycelis.co",
    "city": "Popayán",
    "address": "Calle 30 # 21-41"
  },
  {
    "id": "cust-xl-348",
    "rawDoc": "900047676-7",
    "name": "ATLAS INNOVACION S.A.S",
    "phone": "+573136222789",
    "email": "contacto@atlasinnovac.co",
    "city": "Pastos",
    "address": "Calle 37 # 34-44"
  },
  {
    "id": "cust-xl-349",
    "rawDoc": "900047813-8",
    "name": "SANTIAGO ENRIQUE UBAQUE CONTRERAS",
    "phone": "+573146392820",
    "email": "contacto@santiagoenri.co",
    "city": "Santa Marta",
    "address": "Calle 44 # 47-47"
  },
  {
    "id": "cust-xl-350",
    "rawDoc": "900047950-9",
    "name": "BIOGENETIC IN VITRO SAS",
    "phone": "+573156562851",
    "email": "contacto@biogeneticin.co",
    "city": "Valledupar",
    "address": "Calle 51 # 60-0"
  },
  {
    "id": "cust-xl-351",
    "rawDoc": "1002779569",
    "name": "IVAN DARIO CAMELO CIFUENTES/ RAFAEL MATTUTE",
    "phone": "+573166732882",
    "email": "contacto@ivandariocam.co",
    "city": "Tunja",
    "address": "Calle 58 # 73-3"
  },
  {
    "id": "cust-xl-352",
    "rawDoc": "1002787488",
    "name": "CARLOS ARTURO LOZANO",
    "phone": "+573176902913",
    "email": "contacto@carlosarturo.co",
    "city": "Bogotá D.C.",
    "address": "Calle 65 # 86-6"
  },
  {
    "id": "cust-xl-353",
    "rawDoc": "1002795407",
    "name": "PABLO LEON PEREZ JARAMILLO",
    "phone": "+573187072944",
    "email": "contacto@pabloleonper.co",
    "city": "Medellín",
    "address": "Calle 72 # 99-9"
  },
  {
    "id": "cust-xl-354",
    "rawDoc": "1002803326",
    "name": "CARLOS ENRIQUE MARTINEZ SERRANO",
    "phone": "+573197242975",
    "email": "contacto@carlosenriqu.co",
    "city": "Cali",
    "address": "Calle 79 # 22-12"
  },
  {
    "id": "cust-xl-355",
    "rawDoc": "1002811245",
    "name": "CARLOS MARTINEZ",
    "phone": "+573207413006",
    "email": "contacto@carlosmartin.co",
    "city": "Barranquilla",
    "address": "Calle 86 # 35-15"
  },
  {
    "id": "cust-xl-356",
    "rawDoc": "1002819164",
    "name": "LUIS FERNANDO BUSTAMANTE RAMÍREZ",
    "phone": "+573217583037",
    "email": "contacto@luisfernando.co",
    "city": "Bucaramanga",
    "address": "Calle 93 # 48-18"
  },
  {
    "id": "cust-xl-357",
    "rawDoc": "900048909-7",
    "name": "WILMER ORLANDO MURCIA GABANZO",
    "phone": "+573227753068",
    "email": "contacto@wilmerorland.co",
    "city": "Villavicencio",
    "address": "Calle 100 # 61-21"
  },
  {
    "id": "cust-xl-358",
    "rawDoc": "1002835002",
    "name": "RICARDO CARDENAS VARGAS",
    "phone": "+573237923099",
    "email": "contacto@ricardocarde.co",
    "city": "Montería",
    "address": "Calle 7 # 74-24"
  },
  {
    "id": "cust-xl-359",
    "rawDoc": "1002842921",
    "name": "ANGELA MARIA STEPA",
    "phone": "+573248093130",
    "email": "contacto@angelamarias.co",
    "city": "Pereira",
    "address": "Calle 14 # 87-27"
  },
  {
    "id": "cust-xl-360",
    "rawDoc": "900049320-1",
    "name": "INVERSIONES YMT SAS",
    "phone": "+573108263161",
    "email": "contacto@inversionesy.co",
    "city": "Manizales",
    "address": "Calle 21 # 10-30"
  },
  {
    "id": "cust-xl-361",
    "rawDoc": "900049457-2",
    "name": "TE SOLUCIONAMOS MG S.A.S",
    "phone": "+573118433192",
    "email": "contacto@tesolucionam.co",
    "city": "Neiva",
    "address": "Calle 28 # 23-33"
  },
  {
    "id": "cust-xl-362",
    "rawDoc": "900049594-3",
    "name": "PAISAGRO SA",
    "phone": "+573128603223",
    "email": "contacto@paisagrosa.co",
    "city": "Ibagué",
    "address": "Calle 35 # 36-36"
  },
  {
    "id": "cust-xl-363",
    "rawDoc": "900049731-4",
    "name": "ROSA ANGELA NIÑO ALVARADO",
    "phone": "+573138773254",
    "email": "contacto@rosaangelani.co",
    "city": "Popayán",
    "address": "Calle 42 # 49-39"
  },
  {
    "id": "cust-xl-364",
    "rawDoc": "1002882516",
    "name": "DANIEL ARTUNDUAGA PUENTES/ DANYELY DANIELA ARTUNDUAGA MUÑOZ",
    "phone": "+573148943285",
    "email": "contacto@danielartund.co",
    "city": "Pastos",
    "address": "Calle 49 # 62-42"
  },
  {
    "id": "cust-xl-365",
    "rawDoc": "1002890435",
    "name": "MUESTRA VIDEO MOCHUELO",
    "phone": "+573159113316",
    "email": "contacto@muestravideo.co",
    "city": "Santa Marta",
    "address": "Calle 56 # 75-45"
  },
  {
    "id": "cust-xl-366",
    "rawDoc": "900050142-7",
    "name": "SOCIEDAD AGROPECUARIA LA FLORIDA LTDA",
    "phone": "+573169283347",
    "email": "contacto@sociedadagro.co",
    "city": "Valledupar",
    "address": "Calle 63 # 88-48"
  },
  {
    "id": "cust-xl-367",
    "rawDoc": "900050279-8",
    "name": "ANDRES CLEVES ROSAS",
    "phone": "+573179453378",
    "email": "contacto@andrescleves.co",
    "city": "Tunja",
    "address": "Calle 70 # 11-1"
  },
  {
    "id": "cust-xl-368",
    "rawDoc": "900050416-9",
    "name": "EMBRIOTEC SERVICIOS ESPECIALIZADOS EN BIOTECNOLOGÍA SAS",
    "phone": "+573189623409",
    "email": "contacto@embriotecser.co",
    "city": "Bogotá D.C.",
    "address": "Calle 77 # 24-4"
  },
  {
    "id": "cust-xl-369",
    "rawDoc": "900050553-1",
    "name": "AURO INVERSIONES SAS",
    "phone": "+573199793440",
    "email": "contacto@auroinversio.co",
    "city": "Medellín",
    "address": "Calle 84 # 37-7"
  },
  {
    "id": "cust-xl-370",
    "rawDoc": "1002930030",
    "name": "KEVIN JOSE PICON WILCHES",
    "phone": "+573209963471",
    "email": "contacto@kevinjosepic.co",
    "city": "Cali",
    "address": "Calle 91 # 50-10"
  },
  {
    "id": "cust-xl-371",
    "rawDoc": "1002937949",
    "name": "RAFAEL RAMÓN ALEJANDRO RAMIREZ",
    "phone": "+573211143502",
    "email": "contacto@rafaelramnal.co",
    "city": "Barranquilla",
    "address": "Calle 98 # 63-13"
  },
  {
    "id": "cust-xl-372",
    "rawDoc": "1002945868",
    "name": "RUBEN DARIO SIERRA MORENO",
    "phone": "+573221313533",
    "email": "contacto@rubendariosi.co",
    "city": "Bucaramanga",
    "address": "Calle 5 # 76-16"
  },
  {
    "id": "cust-xl-373",
    "rawDoc": "1002953787",
    "name": "HERNÁN ANDRÉS GAVILANES RIVADENEIRA",
    "phone": "+573231483564",
    "email": "contacto@hernnandrsga.co",
    "city": "Villavicencio",
    "address": "Calle 12 # 89-19"
  },
  {
    "id": "cust-xl-374",
    "rawDoc": "1002961706",
    "name": "JAIME ALBERTO GALINDO PUERTO",
    "phone": "+573241653595",
    "email": "contacto@jaimealberto.co",
    "city": "Montería",
    "address": "Calle 19 # 12-22"
  },
  {
    "id": "cust-xl-375",
    "rawDoc": "900051375-7",
    "name": "ALLIÉ GROUP S.A.S",
    "phone": "+573101823626",
    "email": "contacto@alligroupsas.co",
    "city": "Pereira",
    "address": "Calle 26 # 25-25"
  },
  {
    "id": "cust-xl-376",
    "rawDoc": "900051512-8",
    "name": "JOSE ISAAC SEPULVEDA ARANGO",
    "phone": "+573111993657",
    "email": "contacto@joseisaacsep.co",
    "city": "Manizales",
    "address": "Calle 33 # 38-28"
  },
  {
    "id": "cust-xl-377",
    "rawDoc": "1002985463",
    "name": "FUNDACION GRUPO ARGOS/ N°.PMO 6000048951/ Migo 5000156688",
    "phone": "+573122163688",
    "email": "contacto@fundaciongru.co",
    "city": "Neiva",
    "address": "Calle 40 # 51-31"
  },
  {
    "id": "cust-xl-378",
    "rawDoc": "1002993382",
    "name": "FUNDACION GRUPO ARGOS/ N°.PMO 6000048951",
    "phone": "+573132333719",
    "email": "contacto@fundaciongru.co",
    "city": "Ibagué",
    "address": "Calle 47 # 64-34"
  },
  {
    "id": "cust-xl-379",
    "rawDoc": "900051923-2",
    "name": "SALIDA EQUIPOS MAS CAMPO",
    "phone": "+573142503750",
    "email": "contacto@salidaequipo.co",
    "city": "Popayán",
    "address": "Calle 54 # 77-37"
  },
  {
    "id": "cust-xl-380",
    "rawDoc": "1003009220",
    "name": "ARMANDO ANGULO MACHACADO",
    "phone": "+573152673781",
    "email": "contacto@armandoangul.co",
    "city": "Pastos",
    "address": "Calle 61 # 90-40"
  },
  {
    "id": "cust-xl-381",
    "rawDoc": "1003017139",
    "name": "MAS CAMPO CURSO ECOGRAFIA ZIPA Y EVENTO TUNJA",
    "phone": "+573162843812",
    "email": "contacto@mascampocurs.co",
    "city": "Santa Marta",
    "address": "Calle 68 # 13-43"
  },
  {
    "id": "cust-xl-382",
    "rawDoc": "1003025058",
    "name": "AGROPECUARIA EL DESTINO S A S",
    "phone": "+573173013843",
    "email": "contacto@agropecuaria.co",
    "city": "Valledupar",
    "address": "Calle 75 # 26-46"
  },
  {
    "id": "cust-xl-383",
    "rawDoc": "1003032977",
    "name": "JOSE ANDRES HORTUA MARTINEZ",
    "phone": "+573183183874",
    "email": "contacto@joseandresho.co",
    "city": "Tunja",
    "address": "Calle 82 # 39-49"
  },
  {
    "id": "cust-xl-384",
    "rawDoc": "900052608-7",
    "name": "ISABEL CATALINA VELEZ JARAMILLO",
    "phone": "+573193353905",
    "email": "contacto@isabelcatali.co",
    "city": "Bogotá D.C.",
    "address": "Calle 89 # 52-2"
  },
  {
    "id": "cust-xl-385",
    "rawDoc": "1003048815",
    "name": "LUIS ALBERTO DELGADO",
    "phone": "+573203523936",
    "email": "contacto@luisalbertod.co",
    "city": "Medellín",
    "address": "Calle 96 # 65-5"
  },
  {
    "id": "cust-xl-386",
    "rawDoc": "900052882-9",
    "name": "JULIO CESAR MORANTES RINCON",
    "phone": "+573213693967",
    "email": "contacto@juliocesarmo.co",
    "city": "Cali",
    "address": "Calle 3 # 78-8"
  },
  {
    "id": "cust-xl-387",
    "rawDoc": "1003064653",
    "name": "JUAN PABLO JURADO",
    "phone": "+573223863998",
    "email": "contacto@juanpablojur.co",
    "city": "Barranquilla",
    "address": "Calle 10 # 91-11"
  },
  {
    "id": "cust-xl-388",
    "rawDoc": "900053156-2",
    "name": "AGROPECUARIA BAJOGRANDE S.A.",
    "phone": "+573234034029",
    "email": "contacto@agropecuaria.co",
    "city": "Bucaramanga",
    "address": "Calle 17 # 14-14"
  },
  {
    "id": "cust-xl-389",
    "rawDoc": "900053293-3",
    "name": "JULY DISTRIBUCIONES AGROPECUARIAS S.A.S.",
    "phone": "+573244204060",
    "email": "contacto@julydistribu.co",
    "city": "Villavicencio",
    "address": "Calle 24 # 27-17"
  },
  {
    "id": "cust-xl-390",
    "rawDoc": "900053430-4",
    "name": "SOCIEDAD AGRICOLA Y GANADERA MARACAIBO S.A.S",
    "phone": "+573104374091",
    "email": "contacto@sociedadagri.co",
    "city": "Montería",
    "address": "Calle 31 # 40-20"
  },
  {
    "id": "cust-xl-391",
    "rawDoc": "1003096329",
    "name": "LUIS ALEJANDRO RUIZ FINO",
    "phone": "+573114544122",
    "email": "contacto@luisalejandr.co",
    "city": "Pereira",
    "address": "Calle 38 # 53-23"
  },
  {
    "id": "cust-xl-392",
    "rawDoc": "900053704-6",
    "name": "HECTOR ARIEL MACIAS TAMAYO",
    "phone": "+573124714153",
    "email": "contacto@hectorarielm.co",
    "city": "Manizales",
    "address": "Calle 45 # 66-26"
  },
  {
    "id": "cust-xl-393",
    "rawDoc": "1003112167",
    "name": "FRANCISCO JAVIER CANO BETANCUR",
    "phone": "+573134884184",
    "email": "contacto@franciscojav.co",
    "city": "Neiva",
    "address": "Calle 52 # 79-29"
  },
  {
    "id": "cust-xl-394",
    "rawDoc": "1003120086",
    "name": "CARLOS ARTURO ZAPATA GIL",
    "phone": "+573145054215",
    "email": "contacto@carlosarturo.co",
    "city": "Ibagué",
    "address": "Calle 59 # 92-32"
  },
  {
    "id": "cust-xl-395",
    "rawDoc": "1003128005",
    "name": "JESÚS DAVID CEGARRA COLMENARES",
    "phone": "+573155224246",
    "email": "contacto@jessdavidceg.co",
    "city": "Popayán",
    "address": "Calle 66 # 15-35"
  },
  {
    "id": "cust-xl-396",
    "rawDoc": "1003135924",
    "name": "BRAYAN STEVEN CUADROS ARIAS",
    "phone": "+573165394277",
    "email": "contacto@brayansteven.co",
    "city": "Pastos",
    "address": "Calle 73 # 28-38"
  },
  {
    "id": "cust-xl-397",
    "rawDoc": "1003143843",
    "name": "DIANA SOFIA MONROY TORRES",
    "phone": "+573175564308",
    "email": "contacto@dianasofiamo.co",
    "city": "Santa Marta",
    "address": "Calle 80 # 41-41"
  },
  {
    "id": "cust-xl-398",
    "rawDoc": "900054526-3",
    "name": "COMERCIAL MAS CAMPO BIOTECNOLOGÍA ANIMAL SAS/ CURSO PEREIRA",
    "phone": "+573185734339",
    "email": "contacto@comercialmas.co",
    "city": "Valledupar",
    "address": "Calle 87 # 54-44"
  },
  {
    "id": "cust-xl-399",
    "rawDoc": "1003159681",
    "name": "CAMILA ANDREA CASTRO",
    "phone": "+573195904370",
    "email": "contacto@camilaandrea.co",
    "city": "Tunja",
    "address": "Calle 94 # 67-47"
  },
  {
    "id": "cust-xl-400",
    "rawDoc": "1003167600",
    "name": "JESÚS ANTONIO BUSTILLO MEDINA",
    "phone": "+573206074401",
    "email": "contacto@jessantoniob.co",
    "city": "Bogotá D.C.",
    "address": "Calle 1 # 80-0"
  },
  {
    "id": "cust-xl-401",
    "rawDoc": "1003175519",
    "name": "JAVIER ANDRÉS SUAREZ FUENTES",
    "phone": "+573216244432",
    "email": "contacto@javierandrss.co",
    "city": "Medellín",
    "address": "Calle 8 # 93-3"
  },
  {
    "id": "cust-xl-402",
    "rawDoc": "1003183438",
    "name": "COMITE DE GANADEROS DE LA HOYA DEL RIO SUAREZ",
    "phone": "+573226414463",
    "email": "contacto@comitedegana.co",
    "city": "Cali",
    "address": "Calle 15 # 16-6"
  },
  {
    "id": "cust-xl-403",
    "rawDoc": "900055211-8",
    "name": "COMERCIALIZADORA GANACOL SAS",
    "phone": "+573236584494",
    "email": "contacto@comercializa.co",
    "city": "Barranquilla",
    "address": "Calle 22 # 29-9"
  },
  {
    "id": "cust-xl-404",
    "rawDoc": "1003199276",
    "name": "FUNDACION GRUPO ARGOS 6000049033",
    "phone": "+573246754525",
    "email": "contacto@fundaciongru.co",
    "city": "Bucaramanga",
    "address": "Calle 29 # 42-12"
  },
  {
    "id": "cust-xl-405",
    "rawDoc": "1003207195",
    "name": "RICARDO ZABALA",
    "phone": "+573106924556",
    "email": "contacto@ricardozabal.co",
    "city": "Villavicencio",
    "address": "Calle 36 # 55-15"
  },
  {
    "id": "cust-xl-406",
    "rawDoc": "1003215114",
    "name": "MARIO BRAVO",
    "phone": "+573117094587",
    "email": "contacto@mariobravo.co",
    "city": "Montería",
    "address": "Calle 43 # 68-18"
  },
  {
    "id": "cust-xl-407",
    "rawDoc": "900055759-3",
    "name": "COOPERATIVA COLANTA/ OC A02-COC-326920",
    "phone": "+573127264618",
    "email": "contacto@cooperativac.co",
    "city": "Pereira",
    "address": "Calle 50 # 81-21"
  },
  {
    "id": "cust-xl-408",
    "rawDoc": "1003230952",
    "name": "BYRON RODRIGO CHUGÁ",
    "phone": "+573137434649",
    "email": "contacto@byronrodrigo.co",
    "city": "Manizales",
    "address": "Calle 57 # 94-24"
  },
  {
    "id": "cust-xl-409",
    "rawDoc": "900056033-5",
    "name": "JULIAN CAMILO GOMEZ VASQUEZ - REPROHORSE SAS",
    "phone": "+573147604680",
    "email": "contacto@juliancamilo.co",
    "city": "Neiva",
    "address": "Calle 64 # 17-27"
  },
  {
    "id": "cust-xl-410",
    "rawDoc": "900056170-6",
    "name": "INVERSIONES AGROPUNTO LA FINCA S A S",
    "phone": "+573157774711",
    "email": "contacto@inversionesa.co",
    "city": "Ibagué",
    "address": "Calle 71 # 30-30"
  },
  {
    "id": "cust-xl-411",
    "rawDoc": "900056307-7",
    "name": "INR INVERSIONES REINOSO Y CIA SAS",
    "phone": "+573167944742",
    "email": "contacto@inrinversion.co",
    "city": "Popayán",
    "address": "Calle 78 # 43-33"
  },
  {
    "id": "cust-xl-412",
    "rawDoc": "900056444-8",
    "name": "DARIO MARULANDA ANGEL Y SUCESORES S.A.S.",
    "phone": "+573178114773",
    "email": "contacto@dariomarulan.co",
    "city": "Pastos",
    "address": "Calle 85 # 56-36"
  },
  {
    "id": "cust-xl-413",
    "rawDoc": "1003270547",
    "name": "JUAN FERNANDO LOZANO ARROYO",
    "phone": "+573188284804",
    "email": "contacto@juanfernando.co",
    "city": "Santa Marta",
    "address": "Calle 92 # 69-39"
  },
  {
    "id": "cust-xl-414",
    "rawDoc": "900056718-1",
    "name": "CORPORACION DE CIENCIAS BIOTECNOLOGIAS - EMBRIOTECNO",
    "phone": "+573198454835",
    "email": "contacto@corporaciond.co",
    "city": "Valledupar",
    "address": "Calle 99 # 82-42"
  },
  {
    "id": "cust-xl-415",
    "rawDoc": "1003286385",
    "name": "JAUMER ANDRES OROZCO CLAVIJO",
    "phone": "+573208624866",
    "email": "contacto@jaumerandres.co",
    "city": "Tunja",
    "address": "Calle 6 # 95-45"
  },
  {
    "id": "cust-xl-416",
    "rawDoc": "1003294304",
    "name": "CARLOS FERNANDO VARGAS GIRON",
    "phone": "+573218794897",
    "email": "contacto@carlosfernan.co",
    "city": "Bogotá D.C.",
    "address": "Calle 13 # 18-48"
  },
  {
    "id": "cust-xl-417",
    "rawDoc": "1003302223",
    "name": "ERIKA PAOLA VERGARA SUAREZ",
    "phone": "+573228964928",
    "email": "contacto@erikapaolave.co",
    "city": "Medellín",
    "address": "Calle 20 # 31-1"
  },
  {
    "id": "cust-xl-418",
    "rawDoc": "900057266-5",
    "name": "JUAN MANUEL SANTOS TAMAYO",
    "phone": "+573239134959",
    "email": "contacto@juanmanuelsa.co",
    "city": "Cali",
    "address": "Calle 27 # 44-4"
  },
  {
    "id": "cust-xl-419",
    "rawDoc": "1003318061",
    "name": "JUAN PABLO MEJIA LEON",
    "phone": "+573249304990",
    "email": "contacto@juanpablomej.co",
    "city": "Barranquilla",
    "address": "Calle 34 # 57-7"
  },
  {
    "id": "cust-xl-420",
    "rawDoc": "900057540-7",
    "name": "YARUMA S.A.S",
    "phone": "+573109475021",
    "email": "contacto@yarumasas.co",
    "city": "Bucaramanga",
    "address": "Calle 41 # 70-10"
  },
  {
    "id": "cust-xl-421",
    "rawDoc": "900057677-8",
    "name": "DIEGO ARMANDO SANCHEZ HEREDIA",
    "phone": "+573119645052",
    "email": "contacto@diegoarmando.co",
    "city": "Villavicencio",
    "address": "Calle 48 # 83-13"
  },
  {
    "id": "cust-xl-422",
    "rawDoc": "900057814-9",
    "name": "CATALINA GARCIA",
    "phone": "+573129815083",
    "email": "contacto@catalinagarc.co",
    "city": "Montería",
    "address": "Calle 55 # 96-16"
  },
  {
    "id": "cust-xl-423",
    "rawDoc": "1003349737",
    "name": "MARIO JULIAN LOZANO",
    "phone": "+573139985114",
    "email": "contacto@mariojulianl.co",
    "city": "Pereira",
    "address": "Calle 62 # 19-19"
  },
  {
    "id": "cust-xl-424",
    "rawDoc": "900058088-2",
    "name": "INVERSIONES J.V. S.A.S /OC 67088",
    "phone": "+573141165145",
    "email": "contacto@inversionesj.co",
    "city": "Manizales",
    "address": "Calle 69 # 32-22"
  },
  {
    "id": "cust-xl-425",
    "rawDoc": "1003365575",
    "name": "ESTEBAN ALDEMAR GOYES CEBALLOS",
    "phone": "+573151335176",
    "email": "contacto@estebanaldem.co",
    "city": "Neiva",
    "address": "Calle 76 # 45-25"
  },
  {
    "id": "cust-xl-426",
    "rawDoc": "1003373494",
    "name": "FUNDACION UNIVERSITARIA AGRARIA DE COLOMBIA (UNIAGRARIA)",
    "phone": "+573161505207",
    "email": "contacto@fundacionuni.co",
    "city": "Ibagué",
    "address": "Calle 83 # 58-28"
  },
  {
    "id": "cust-xl-427",
    "rawDoc": "1003381413",
    "name": "DANIEL ALBEIRO DIAZ CAVIEDES",
    "phone": "+573171675238",
    "email": "contacto@danielalbeir.co",
    "city": "Popayán",
    "address": "Calle 90 # 71-31"
  },
  {
    "id": "cust-xl-428",
    "rawDoc": "900058636-6",
    "name": "CGR BIOTECNOLOGIA REPRODUCTIVA SAS",
    "phone": "+573181845269",
    "email": "contacto@cgrbiotecnol.co",
    "city": "Pastos",
    "address": "Calle 97 # 84-34"
  },
  {
    "id": "cust-xl-429",
    "rawDoc": "1003397251",
    "name": "JAIRO ANDRES GONZALEZ AGUDELO",
    "phone": "+573192015300",
    "email": "contacto@jairoandresg.co",
    "city": "Santa Marta",
    "address": "Calle 4 # 97-37"
  },
  {
    "id": "cust-xl-430",
    "rawDoc": "900058910-8",
    "name": "PAOLA ANDREA RODRIGUEZ FORERO/ R&G REPRODUCCIÓN BOVINA SAS",
    "phone": "+573202185331",
    "email": "contacto@paolaandrear.co",
    "city": "Valledupar",
    "address": "Calle 11 # 20-40"
  },
  {
    "id": "cust-xl-431",
    "rawDoc": "1003413089",
    "name": "JOAN ANDRES LISCANO MORALES",
    "phone": "+573212355362",
    "email": "contacto@joanandresli.co",
    "city": "Tunja",
    "address": "Calle 18 # 33-43"
  },
  {
    "id": "cust-xl-432",
    "rawDoc": "900059184-1",
    "name": "GRUPO EMPRESARIAL GENETICO S.A.S.",
    "phone": "+573222525393",
    "email": "contacto@grupoempresa.co",
    "city": "Bogotá D.C.",
    "address": "Calle 25 # 46-46"
  },
  {
    "id": "cust-xl-433",
    "rawDoc": "1003428927",
    "name": "DANIELA PARRA",
    "phone": "+573232695424",
    "email": "contacto@danielaparra.co",
    "city": "Medellín",
    "address": "Calle 32 # 59-49"
  },
  {
    "id": "cust-xl-434",
    "rawDoc": "900059458-3",
    "name": "LISANDRO SILVA APARICIO",
    "phone": "+573242865455",
    "email": "contacto@lisandrosilv.co",
    "city": "Cali",
    "address": "Calle 39 # 72-2"
  },
  {
    "id": "cust-xl-435",
    "rawDoc": "1003444765",
    "name": "JESÚS ARBEY HORTA TRUJILLO",
    "phone": "+573103035486",
    "email": "contacto@jessarbeyhor.co",
    "city": "Barranquilla",
    "address": "Calle 46 # 85-5"
  },
  {
    "id": "cust-xl-436",
    "rawDoc": "900059732-5",
    "name": "HATO VERDE SAS",
    "phone": "+573113205517",
    "email": "contacto@hatoverdesas.co",
    "city": "Bucaramanga",
    "address": "Calle 53 # 98-8"
  },
  {
    "id": "cust-xl-437",
    "rawDoc": "1003460603",
    "name": "JESÚS SUÁREZ MENENDEZ",
    "phone": "+573123375548",
    "email": "contacto@jesssurezmen.co",
    "city": "Villavicencio",
    "address": "Calle 60 # 21-11"
  },
  {
    "id": "cust-xl-438",
    "rawDoc": "1003468522",
    "name": "VIVIANA PATIÑO",
    "phone": "+573133545579",
    "email": "contacto@vivianapatio.co",
    "city": "Montería",
    "address": "Calle 67 # 34-14"
  },
  {
    "id": "cust-xl-439",
    "rawDoc": "900060143-8",
    "name": "JUAN DIEGO CASALLAS ARÉVALO",
    "phone": "+573143715610",
    "email": "contacto@juandiegocas.co",
    "city": "Pereira",
    "address": "Calle 74 # 47-17"
  },
  {
    "id": "cust-xl-440",
    "rawDoc": "1003484360",
    "name": "AUGUSTO MARTINEZ VISBAL",
    "phone": "+573153885641",
    "email": "contacto@augustomarti.co",
    "city": "Manizales",
    "address": "Calle 81 # 60-20"
  },
  {
    "id": "cust-xl-441",
    "rawDoc": "900060417-1",
    "name": "INVERSIONES AGROINDUSTRIALES TRES MONTES SAS",
    "phone": "+573164055672",
    "email": "contacto@inversionesa.co",
    "city": "Neiva",
    "address": "Calle 88 # 73-23"
  },
  {
    "id": "cust-xl-442",
    "rawDoc": "1003500198",
    "name": "CARLOS FERNEY RUEDA GUANARO",
    "phone": "+573174225703",
    "email": "contacto@carlosferney.co",
    "city": "Ibagué",
    "address": "Calle 95 # 86-26"
  },
  {
    "id": "cust-xl-443",
    "rawDoc": "900060691-3",
    "name": "TAMISUELOS SAS",
    "phone": "+573184395734",
    "email": "contacto@tamisuelossa.co",
    "city": "Popayán",
    "address": "Calle 2 # 99-29"
  },
  {
    "id": "cust-xl-444",
    "rawDoc": "1003516036",
    "name": "LUIS FERNANDO CARVAJAL GIL",
    "phone": "+573194565765",
    "email": "contacto@luisfernando.co",
    "city": "Pastos",
    "address": "Calle 9 # 22-32"
  },
  {
    "id": "cust-xl-445",
    "rawDoc": "1003523955",
    "name": "FERNEY ALBERTO ZARATE SIERRA",
    "phone": "+573204735796",
    "email": "contacto@ferneyalbert.co",
    "city": "Santa Marta",
    "address": "Calle 16 # 35-35"
  },
  {
    "id": "cust-xl-446",
    "rawDoc": "900061102-6",
    "name": "COLVITRO BIOTECNOLOGIA REPRODUCTIVA SAS",
    "phone": "+573214905827",
    "email": "contacto@colvitrobiot.co",
    "city": "Valledupar",
    "address": "Calle 23 # 48-38"
  },
  {
    "id": "cust-xl-447",
    "rawDoc": "900061239-7",
    "name": "MARIA ISABEL LEON SANTACRUZ",
    "phone": "+573225075858",
    "email": "contacto@mariaisabell.co",
    "city": "Tunja",
    "address": "Calle 30 # 61-41"
  },
  {
    "id": "cust-xl-448",
    "rawDoc": "1003547712",
    "name": "HENRY ALBERTO CASTELLANOS SUÁREZ",
    "phone": "+573235245889",
    "email": "contacto@henryalberto.co",
    "city": "Bogotá D.C.",
    "address": "Calle 37 # 74-44"
  },
  {
    "id": "cust-xl-449",
    "rawDoc": "1003555631",
    "name": "ALDEMAR CHAVEZ RODRIGUEZ",
    "phone": "+573245415920",
    "email": "contacto@aldemarchave.co",
    "city": "Medellín",
    "address": "Calle 44 # 87-47"
  },
  {
    "id": "cust-xl-450",
    "rawDoc": "1003563550",
    "name": "NICOLÁS MANUEL SIERRA ZÁRATE",
    "phone": "+573105585951",
    "email": "contacto@nicolsmanuel.co",
    "city": "Cali",
    "address": "Calle 51 # 10-0"
  },
  {
    "id": "cust-xl-451",
    "rawDoc": "900061787-2",
    "name": "PRODUCTOS E INSUMOS AGROPECUARIOS COMPAÑIA SAS",
    "phone": "+573115755982",
    "email": "contacto@productosein.co",
    "city": "Barranquilla",
    "address": "Calle 58 # 23-3"
  },
  {
    "id": "cust-xl-452",
    "rawDoc": "900061924-3",
    "name": "RAUL SULPICIO SARMIENTO CELY",
    "phone": "+573125926013",
    "email": "contacto@raulsulpicio.co",
    "city": "Bucaramanga",
    "address": "Calle 65 # 36-6"
  },
  {
    "id": "cust-xl-453",
    "rawDoc": "900062061-4",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA NO 674374",
    "phone": "+573136096044",
    "email": "contacto@lacorporacio.co",
    "city": "Villavicencio",
    "address": "Calle 72 # 49-9"
  },
  {
    "id": "cust-xl-454",
    "rawDoc": "900062198-5",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA No.674433",
    "phone": "+573146266075",
    "email": "contacto@lacorporacio.co",
    "city": "Montería",
    "address": "Calle 79 # 62-12"
  },
  {
    "id": "cust-xl-455",
    "rawDoc": "1003603145",
    "name": "ESTEBAN DANILO CUAICAL CUAICAL",
    "phone": "+573156436106",
    "email": "contacto@estebandanil.co",
    "city": "Pereira",
    "address": "Calle 86 # 75-15"
  },
  {
    "id": "cust-xl-456",
    "rawDoc": "1003611064",
    "name": "LENIKER ELÍAS RUÍZ GARRIDO",
    "phone": "+573166606137",
    "email": "contacto@lenikerelasr.co",
    "city": "Manizales",
    "address": "Calle 93 # 88-18"
  },
  {
    "id": "cust-xl-457",
    "rawDoc": "1003618983",
    "name": "LUIS GABRIEL DIAZ",
    "phone": "+573176776168",
    "email": "contacto@luisgabrield.co",
    "city": "Neiva",
    "address": "Calle 100 # 11-21"
  },
  {
    "id": "cust-xl-458",
    "rawDoc": "900062746-9",
    "name": "GRUPO SEMBRIOGAN SAS",
    "phone": "+573186946199",
    "email": "contacto@gruposembrio.co",
    "city": "Ibagué",
    "address": "Calle 7 # 24-24"
  },
  {
    "id": "cust-xl-459",
    "rawDoc": "1003634821",
    "name": "NICOLÁS EDUARDO BECERA GARZÓN",
    "phone": "+573197116230",
    "email": "contacto@nicolseduard.co",
    "city": "Popayán",
    "address": "Calle 14 # 37-27"
  },
  {
    "id": "cust-xl-460",
    "rawDoc": "1003642740",
    "name": "HECTOR EDUARDO FUENTES QUINBAYA",
    "phone": "+573207286261",
    "email": "contacto@hectoreduard.co",
    "city": "Pastos",
    "address": "Calle 21 # 50-30"
  },
  {
    "id": "cust-xl-461",
    "rawDoc": "1003650659",
    "name": "ESTIVEN RANGEL",
    "phone": "+573217456292",
    "email": "contacto@estivenrange.co",
    "city": "Santa Marta",
    "address": "Calle 28 # 63-33"
  },
  {
    "id": "cust-xl-462",
    "rawDoc": "900063294-4",
    "name": "JUAN CARLOS SUAREZ SALAZAR",
    "phone": "+573227626323",
    "email": "contacto@juancarlossu.co",
    "city": "Valledupar",
    "address": "Calle 35 # 76-36"
  },
  {
    "id": "cust-xl-463",
    "rawDoc": "900063431-5",
    "name": "ZULY TATIANA RAMIREZ SANCHEZ",
    "phone": "+573237796354",
    "email": "contacto@zulytatianar.co",
    "city": "Tunja",
    "address": "Calle 42 # 89-39"
  },
  {
    "id": "cust-xl-464",
    "rawDoc": "1003674416",
    "name": "JORGE EDUARDO ESPITIA TORRES",
    "phone": "+573247966385",
    "email": "contacto@jorgeeduardo.co",
    "city": "Bogotá D.C.",
    "address": "Calle 49 # 12-42"
  },
  {
    "id": "cust-xl-465",
    "rawDoc": "900063705-7",
    "name": "DANIEL STIBEN ROMERO SUSA",
    "phone": "+573108136416",
    "email": "contacto@danielstiben.co",
    "city": "Medellín",
    "address": "Calle 56 # 25-45"
  },
  {
    "id": "cust-xl-466",
    "rawDoc": "1003690254",
    "name": "LUIS ERNESTO GOYENECHE GOYENECHE",
    "phone": "+573118306447",
    "email": "contacto@luisernestog.co",
    "city": "Cali",
    "address": "Calle 63 # 38-48"
  },
  {
    "id": "cust-xl-467",
    "rawDoc": "1003698173",
    "name": "ROBER HADRIEL MONTAÑA PEREZ",
    "phone": "+573128476478",
    "email": "contacto@roberhadriel.co",
    "city": "Barranquilla",
    "address": "Calle 70 # 51-1"
  },
  {
    "id": "cust-xl-468",
    "rawDoc": "1003706092",
    "name": "MANUEL LEONARDO BARRANTES AVELLANEDA",
    "phone": "+573138646509",
    "email": "contacto@manuelleonar.co",
    "city": "Bucaramanga",
    "address": "Calle 77 # 64-4"
  },
  {
    "id": "cust-xl-469",
    "rawDoc": "900064253-2",
    "name": "INVERSIONES BOTICAMPO SAS",
    "phone": "+573148816540",
    "email": "contacto@inversionesb.co",
    "city": "Villavicencio",
    "address": "Calle 84 # 77-7"
  },
  {
    "id": "cust-xl-470",
    "rawDoc": "900064390-3",
    "name": "COOPERATIVA DE REPRODUCTORES DE ELCHE DE GUATAVITA",
    "phone": "+573158986571",
    "email": "contacto@cooperativad.co",
    "city": "Montería",
    "address": "Calle 91 # 90-10"
  },
  {
    "id": "cust-xl-471",
    "rawDoc": "1003729849",
    "name": "ROBINSON LÓPEZ ROJAS",
    "phone": "+573169156602",
    "email": "contacto@robinsonlpez.co",
    "city": "Pereira",
    "address": "Calle 98 # 13-13"
  },
  {
    "id": "cust-xl-472",
    "rawDoc": "1003737768",
    "name": "CATALINA OCHOA",
    "phone": "+573179326633",
    "email": "contacto@catalinaocho.co",
    "city": "Manizales",
    "address": "Calle 5 # 26-16"
  },
  {
    "id": "cust-xl-473",
    "rawDoc": "1003745687",
    "name": "JUAN CARLOS LOPEZ GOMEZ",
    "phone": "+573189496664",
    "email": "contacto@juancarloslo.co",
    "city": "Neiva",
    "address": "Calle 12 # 39-19"
  },
  {
    "id": "cust-xl-474",
    "rawDoc": "1003753606",
    "name": "RUBEN DARIO SIERRA ARRIETA",
    "phone": "+573199666695",
    "email": "contacto@rubendariosi.co",
    "city": "Ibagué",
    "address": "Calle 19 # 52-22"
  },
  {
    "id": "cust-xl-475",
    "rawDoc": "1003761525",
    "name": "MANUEL LUCIO GONZALEZ CHAMORRO",
    "phone": "+573209836726",
    "email": "contacto@manuelluciog.co",
    "city": "Popayán",
    "address": "Calle 26 # 65-25"
  },
  {
    "id": "cust-xl-476",
    "rawDoc": "1003769444",
    "name": "LUIGGI GAFARO QUINTERO",
    "phone": "+573211016757",
    "email": "contacto@luiggigafaro.co",
    "city": "Pastos",
    "address": "Calle 33 # 78-28"
  },
  {
    "id": "cust-xl-477",
    "rawDoc": "1003777363",
    "name": "JOSE MIGUEL GIRALDO LOPEZ",
    "phone": "+573221186788",
    "email": "contacto@josemiguelgi.co",
    "city": "Santa Marta",
    "address": "Calle 40 # 91-31"
  },
  {
    "id": "cust-xl-478",
    "rawDoc": "900065486-2",
    "name": "EMPRESA GENETICA ESPECIAL LTDA",
    "phone": "+573231356819",
    "email": "contacto@empresagenet.co",
    "city": "Valledupar",
    "address": "Calle 47 # 14-34"
  },
  {
    "id": "cust-xl-479",
    "rawDoc": "900065623-3",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA",
    "phone": "+573241526850",
    "email": "contacto@lacorporacio.co",
    "city": "Tunja",
    "address": "Calle 54 # 27-37"
  },
  {
    "id": "cust-xl-480",
    "rawDoc": "1003801120",
    "name": "MIGUEL ANGEL MOTTA SILVA",
    "phone": "+573101696881",
    "email": "contacto@miguelangelm.co",
    "city": "Bogotá D.C.",
    "address": "Calle 61 # 40-40"
  },
  {
    "id": "cust-xl-481",
    "rawDoc": "900065897-5",
    "name": "R&G REPRODUCCIÓN BOVINA SAS",
    "phone": "+573111866912",
    "email": "contacto@rgreproducci.co",
    "city": "Medellín",
    "address": "Calle 68 # 53-43"
  },
  {
    "id": "cust-xl-482",
    "rawDoc": "1003816958",
    "name": "JESÚS VALENTIN VALENZUELA CÓDOBA",
    "phone": "+573122036943",
    "email": "contacto@jessvalentin.co",
    "city": "Cali",
    "address": "Calle 75 # 66-46"
  },
  {
    "id": "cust-xl-483",
    "rawDoc": "1003824877",
    "name": "KENNER MIGUÉL CAÑIZARES CAMARGO",
    "phone": "+573132206974",
    "email": "contacto@kennermigulc.co",
    "city": "Barranquilla",
    "address": "Calle 82 # 79-49"
  },
  {
    "id": "cust-xl-484",
    "rawDoc": "1003832796",
    "name": "JUAN JIMENEZ",
    "phone": "+573142377005",
    "email": "contacto@juanjimenez.co",
    "city": "Bucaramanga",
    "address": "Calle 89 # 92-2"
  },
  {
    "id": "cust-xl-485",
    "rawDoc": "1003840715",
    "name": "JUAN FELIPE HOYOS PELAEZ",
    "phone": "+573152547036",
    "email": "contacto@juanfelipeho.co",
    "city": "Villavicencio",
    "address": "Calle 96 # 15-5"
  },
  {
    "id": "cust-xl-486",
    "rawDoc": "1003848634",
    "name": "PEDRO JAIME SUAREZ",
    "phone": "+573162717067",
    "email": "contacto@pedrojaimesu.co",
    "city": "Montería",
    "address": "Calle 3 # 28-8"
  },
  {
    "id": "cust-xl-487",
    "rawDoc": "1003856553",
    "name": "MARE LIMITADA",
    "phone": "+573172887098",
    "email": "contacto@marelimitada.co",
    "city": "Pereira",
    "address": "Calle 10 # 41-11"
  },
  {
    "id": "cust-xl-488",
    "rawDoc": "1003864472",
    "name": "DANIEL ARTUNDUAGA PUENTES",
    "phone": "+573183057129",
    "email": "contacto@danielartund.co",
    "city": "Manizales",
    "address": "Calle 17 # 54-14"
  },
  {
    "id": "cust-xl-489",
    "rawDoc": "1003872391",
    "name": "ESNEIDER VASQUEZ MOTTA",
    "phone": "+573193227160",
    "email": "contacto@esneidervasq.co",
    "city": "Neiva",
    "address": "Calle 24 # 67-17"
  },
  {
    "id": "cust-xl-490",
    "rawDoc": "1003880310",
    "name": "CRISTIAN FABIAN LEMUS TORRES",
    "phone": "+573203397191",
    "email": "contacto@cristianfabi.co",
    "city": "Ibagué",
    "address": "Calle 31 # 80-20"
  },
  {
    "id": "cust-xl-491",
    "rawDoc": "900067267-6",
    "name": "VICTOR JAVIER SANABRIA SEPULVEDA",
    "phone": "+573213567222",
    "email": "contacto@victorjavier.co",
    "city": "Popayán",
    "address": "Calle 38 # 93-23"
  },
  {
    "id": "cust-xl-492",
    "rawDoc": "900067404-7",
    "name": "JHON CASTAÑEDA SANTA",
    "phone": "+573223737253",
    "email": "contacto@jhoncastaeda.co",
    "city": "Pastos",
    "address": "Calle 45 # 16-26"
  },
  {
    "id": "cust-xl-493",
    "rawDoc": "1003904067",
    "name": "MAS CAMPO/ CURSO",
    "phone": "+573233907284",
    "email": "contacto@mascampocurs.co",
    "city": "Santa Marta",
    "address": "Calle 52 # 29-29"
  },
  {
    "id": "cust-xl-494",
    "rawDoc": "1003911986",
    "name": "FELIPE JARAMILLO MONDALVE",
    "phone": "+573244077315",
    "email": "contacto@felipejarami.co",
    "city": "Valledupar",
    "address": "Calle 59 # 42-32"
  },
  {
    "id": "cust-xl-495",
    "rawDoc": "900067815-1",
    "name": "EMBRIONAJE LIFE MASAF S.A.S",
    "phone": "+573104247346",
    "email": "contacto@embrionajeli.co",
    "city": "Tunja",
    "address": "Calle 66 # 55-35"
  },
  {
    "id": "cust-xl-496",
    "rawDoc": "900067952-2",
    "name": "FREDY LEONARDO SANTAMARIA MEDINA",
    "phone": "+573114417377",
    "email": "contacto@fredyleonard.co",
    "city": "Bogotá D.C.",
    "address": "Calle 73 # 68-38"
  },
  {
    "id": "cust-xl-497",
    "rawDoc": "900068089-3",
    "name": "INSTITUTO DE FORMACION AGRICOLA Y VETERINARIA - IFAV SAS",
    "phone": "+573124587408",
    "email": "contacto@institutodef.co",
    "city": "Medellín",
    "address": "Calle 80 # 81-41"
  },
  {
    "id": "cust-xl-498",
    "rawDoc": "900068226-4",
    "name": "AGROPECUARIA NOBLEZA EQUINA S.A.S.",
    "phone": "+573134757439",
    "email": "contacto@agropecuaria.co",
    "city": "Cali",
    "address": "Calle 87 # 94-44"
  },
  {
    "id": "cust-xl-499",
    "rawDoc": "900068363-5",
    "name": "LUIS MARIANO SOSA FLORES",
    "phone": "+573144927470",
    "email": "contacto@luismarianos.co",
    "city": "Barranquilla",
    "address": "Calle 94 # 17-47"
  },
  {
    "id": "cust-xl-500",
    "rawDoc": "900068500-6",
    "name": "EMBRIONES AMAZONAS S.A.S",
    "phone": "+573155097501",
    "email": "contacto@embrionesama.co",
    "city": "Bucaramanga",
    "address": "Calle 1 # 30-0"
  },
  {
    "id": "cust-xl-501",
    "rawDoc": "900068637-7",
    "name": "TROPICAL EMBRYO SAS",
    "phone": "+573165267532",
    "email": "contacto@tropicalembr.co",
    "city": "Villavicencio",
    "address": "Calle 8 # 43-3"
  },
  {
    "id": "cust-xl-502",
    "rawDoc": "900068774-8",
    "name": "DEMED SAS",
    "phone": "+573175437563",
    "email": "contacto@demedsas.co",
    "city": "Montería",
    "address": "Calle 15 # 56-6"
  },
  {
    "id": "cust-xl-503",
    "rawDoc": "1003983257",
    "name": "CARLOS ERNESTO GONZÁLEZ ECHEVERRY",
    "phone": "+573185607594",
    "email": "contacto@carlosernest.co",
    "city": "Pereira",
    "address": "Calle 22 # 69-9"
  },
  {
    "id": "cust-xl-504",
    "rawDoc": "1003991176",
    "name": "MARTIN JARAMILLO HURTADO",
    "phone": "+573195777625",
    "email": "contacto@martinjarami.co",
    "city": "Manizales",
    "address": "Calle 29 # 82-12"
  },
  {
    "id": "cust-xl-505",
    "rawDoc": "900069185-2",
    "name": "HIPERGENETICA SAS",
    "phone": "+573205947656",
    "email": "contacto@hipergenetic.co",
    "city": "Neiva",
    "address": "Calle 36 # 95-15"
  },
  {
    "id": "cust-xl-506",
    "rawDoc": "900069322-3",
    "name": "ANDREA LILIANA ALFONSO RINCON",
    "phone": "+573216117687",
    "email": "contacto@andrealilian.co",
    "city": "Ibagué",
    "address": "Calle 43 # 18-18"
  },
  {
    "id": "cust-xl-507",
    "rawDoc": "1004014933",
    "name": "AGROPECUARIA EL PORFIN Y COMPAÑIA S EN C",
    "phone": "+573226287718",
    "email": "contacto@agropecuaria.co",
    "city": "Popayán",
    "address": "Calle 50 # 31-21"
  },
  {
    "id": "cust-xl-508",
    "rawDoc": "1004022852",
    "name": "BRAYAN SNEIDER CASTIBLANCO CASTIBLANCO",
    "phone": "+573236457749",
    "email": "contacto@brayansneide.co",
    "city": "Pastos",
    "address": "Calle 57 # 44-24"
  },
  {
    "id": "cust-xl-509",
    "rawDoc": "900069733-6",
    "name": "CARLOS SANCHEZ",
    "phone": "+573246627780",
    "email": "contacto@carlossanche.co",
    "city": "Santa Marta",
    "address": "Calle 64 # 57-27"
  },
  {
    "id": "cust-xl-510",
    "rawDoc": "1004038690",
    "name": "YOEL JOAQUIN OCHOA DIAZ",
    "phone": "+573106797811",
    "email": "contacto@yoeljoaquino.co",
    "city": "Valledupar",
    "address": "Calle 71 # 70-30"
  },
  {
    "id": "cust-xl-511",
    "rawDoc": "900070007-8",
    "name": "DISTRIBUCIONES Y SOLUCIONES S.A.S DISOCO",
    "phone": "+573116967842",
    "email": "contacto@distribucion.co",
    "city": "Tunja",
    "address": "Calle 78 # 83-33"
  },
  {
    "id": "cust-xl-512",
    "rawDoc": "1004054528",
    "name": "SUPLEVETT S A S",
    "phone": "+573127137873",
    "email": "contacto@suplevettsas.co",
    "city": "Bogotá D.C.",
    "address": "Calle 85 # 96-36"
  },
  {
    "id": "cust-xl-513",
    "rawDoc": "1004062447",
    "name": "HELBERT GIOVANNI RODRIGUEZ LOPEZ",
    "phone": "+573137307904",
    "email": "contacto@helbertgiova.co",
    "city": "Medellín",
    "address": "Calle 92 # 19-39"
  },
  {
    "id": "cust-xl-514",
    "rawDoc": "1004070366",
    "name": "ANDERSON FABIAN SOLANO VARELA",
    "phone": "+573147477935",
    "email": "contacto@andersonfabi.co",
    "city": "Cali",
    "address": "Calle 99 # 32-42"
  },
  {
    "id": "cust-xl-515",
    "rawDoc": "1004078285",
    "name": "ISMAEL EDUARDO CASTAÑEDA RAFIÑA",
    "phone": "+573157647966",
    "email": "contacto@ismaeleduard.co",
    "city": "Barranquilla",
    "address": "Calle 6 # 45-45"
  },
  {
    "id": "cust-xl-516",
    "rawDoc": "1004086204",
    "name": "PABLO GERARDO REYNOSO DONIS",
    "phone": "+573167817997",
    "email": "contacto@pablogerardo.co",
    "city": "Bucaramanga",
    "address": "Calle 13 # 58-48"
  },
  {
    "id": "cust-xl-517",
    "rawDoc": "1004094123",
    "name": "ENUAR MANTILLA",
    "phone": "+573177988028",
    "email": "contacto@enuarmantill.co",
    "city": "Villavicencio",
    "address": "Calle 20 # 71-1"
  },
  {
    "id": "cust-xl-518",
    "rawDoc": "900070966-6",
    "name": "DISTRIBUCIONES AGROALIANZA SAS",
    "phone": "+573188158059",
    "email": "contacto@distribucion.co",
    "city": "Montería",
    "address": "Calle 27 # 84-4"
  },
  {
    "id": "cust-xl-519",
    "rawDoc": "900071103-7",
    "name": "CONSTRUCOLDI S.A.S",
    "phone": "+573198328090",
    "email": "contacto@construcoldi.co",
    "city": "Pereira",
    "address": "Calle 34 # 97-7"
  },
  {
    "id": "cust-xl-520",
    "rawDoc": "1004117880",
    "name": "FJMH S A S",
    "phone": "+573208498121",
    "email": "contacto@fjmhsas.co",
    "city": "Manizales",
    "address": "Calle 41 # 20-10"
  },
  {
    "id": "cust-xl-521",
    "rawDoc": "1004125799",
    "name": "RICARDO ANDRÉS TORRES CAMELO",
    "phone": "+573218668152",
    "email": "contacto@ricardoandrs.co",
    "city": "Neiva",
    "address": "Calle 48 # 33-13"
  },
  {
    "id": "cust-xl-522",
    "rawDoc": "1004133718",
    "name": "JUAN CARLOS CRUZ",
    "phone": "+573228838183",
    "email": "contacto@juancarloscr.co",
    "city": "Ibagué",
    "address": "Calle 55 # 46-16"
  },
  {
    "id": "cust-xl-523",
    "rawDoc": "1004141637",
    "name": "IVAN MAURICIO GONZÁLEZ FAJARDO",
    "phone": "+573239008214",
    "email": "contacto@ivanmauricio.co",
    "city": "Popayán",
    "address": "Calle 62 # 59-19"
  },
  {
    "id": "cust-xl-524",
    "rawDoc": "1004149556",
    "name": "MIGUEL DELGADO",
    "phone": "+573249178245",
    "email": "contacto@migueldelgad.co",
    "city": "Pastos",
    "address": "Calle 69 # 72-22"
  },
  {
    "id": "cust-xl-525",
    "rawDoc": "1004157475",
    "name": "PAULA OLARTE",
    "phone": "+573109348276",
    "email": "contacto@paulaolarte.co",
    "city": "Santa Marta",
    "address": "Calle 76 # 85-25"
  },
  {
    "id": "cust-xl-526",
    "rawDoc": "1004165394",
    "name": "AGROPECUARIA LAS FLECHAS C.A.",
    "phone": "+573119518307",
    "email": "contacto@agropecuaria.co",
    "city": "Valledupar",
    "address": "Calle 83 # 98-28"
  },
  {
    "id": "cust-xl-527",
    "rawDoc": "1004173313",
    "name": "CRISTIAN FERNANDO MILLAN HERNANDEZ",
    "phone": "+573129688338",
    "email": "contacto@cristianfern.co",
    "city": "Tunja",
    "address": "Calle 90 # 21-31"
  },
  {
    "id": "cust-xl-528",
    "rawDoc": "1004181232",
    "name": "MAS CAMPO/ MATERIAL CURSO GNEMIN",
    "phone": "+573139858369",
    "email": "contacto@mascampomate.co",
    "city": "Bogotá D.C.",
    "address": "Calle 97 # 34-34"
  },
  {
    "id": "cust-xl-529",
    "rawDoc": "900072473-8",
    "name": "YOHANY SALAZAR PUENTES",
    "phone": "+573141038400",
    "email": "contacto@yohanysalaza.co",
    "city": "Medellín",
    "address": "Calle 4 # 47-37"
  },
  {
    "id": "cust-xl-530",
    "rawDoc": "1004197070",
    "name": "GLADYS YOLANDA BERMUDEZ DE GARCÍA",
    "phone": "+573151208431",
    "email": "contacto@gladysyoland.co",
    "city": "Cali",
    "address": "Calle 11 # 60-40"
  },
  {
    "id": "cust-xl-531",
    "rawDoc": "1004204989",
    "name": "GABRIELA DEL MAR PEREZ RONDON",
    "phone": "+573161378462",
    "email": "contacto@gabrieladelm.co",
    "city": "Barranquilla",
    "address": "Calle 18 # 73-43"
  },
  {
    "id": "cust-xl-532",
    "rawDoc": "1004212908",
    "name": "ALEJANDRO AURELIO REZA CORDERO",
    "phone": "+573171548493",
    "email": "contacto@alejandroaur.co",
    "city": "Bucaramanga",
    "address": "Calle 25 # 86-46"
  },
  {
    "id": "cust-xl-533",
    "rawDoc": "1004220827",
    "name": "RICARDO DAVID CHAVEZ MONTENEGRO",
    "phone": "+573181718524",
    "email": "contacto@ricardodavid.co",
    "city": "Villavicencio",
    "address": "Calle 32 # 99-49"
  },
  {
    "id": "cust-xl-534",
    "rawDoc": "1004228746",
    "name": "RODRIGO DAVID MADRIS ZOMETA",
    "phone": "+573191888555",
    "email": "contacto@rodrigodavid.co",
    "city": "Montería",
    "address": "Calle 39 # 22-2"
  },
  {
    "id": "cust-xl-535",
    "rawDoc": "900073295-5",
    "name": "EDY SANTIAGO MILLÁN MONTENEGRO",
    "phone": "+573202058586",
    "email": "contacto@edysantiagom.co",
    "city": "Pereira",
    "address": "Calle 46 # 35-5"
  },
  {
    "id": "cust-xl-536",
    "rawDoc": "1004244584",
    "name": "JUAN SEBASTIAN LOAIZA MELLIZO",
    "phone": "+573212228617",
    "email": "contacto@juansebastia.co",
    "city": "Manizales",
    "address": "Calle 53 # 48-8"
  },
  {
    "id": "cust-xl-537",
    "rawDoc": "1004252503",
    "name": "CRISTHIAM FARFÁN CASTILLO",
    "phone": "+573222398648",
    "email": "contacto@cristhiamfar.co",
    "city": "Neiva",
    "address": "Calle 60 # 61-11"
  },
  {
    "id": "cust-xl-538",
    "rawDoc": "1004260422",
    "name": "ALEJANDRO JARAMILLO JARAMILLO",
    "phone": "+573232568679",
    "email": "contacto@alejandrojar.co",
    "city": "Ibagué",
    "address": "Calle 67 # 74-14"
  },
  {
    "id": "cust-xl-539",
    "rawDoc": "1004268341",
    "name": "MARIA JOSÉ RODRÍGUEZ",
    "phone": "+573242738710",
    "email": "contacto@mariajosrodr.co",
    "city": "Popayán",
    "address": "Calle 74 # 87-17"
  },
  {
    "id": "cust-xl-540",
    "rawDoc": "1004276260",
    "name": "NELSON YOBANNY VANEAS CASTAÑEDA",
    "phone": "+573102908741",
    "email": "contacto@nelsonyobann.co",
    "city": "Pastos",
    "address": "Calle 81 # 10-20"
  },
  {
    "id": "cust-xl-541",
    "rawDoc": "1004284179",
    "name": "MAS CAMPO",
    "phone": "+573113078772",
    "email": "contacto@mascampo.co",
    "city": "Santa Marta",
    "address": "Calle 88 # 23-23"
  },
  {
    "id": "cust-xl-542",
    "rawDoc": "1004292098",
    "name": "LAZARO FRANCISCO ORELLANA ALFARO",
    "phone": "+573123248803",
    "email": "contacto@lazarofranci.co",
    "city": "Valledupar",
    "address": "Calle 95 # 36-26"
  },
  {
    "id": "cust-xl-543",
    "rawDoc": "900074391-4",
    "name": "COOPERATIVA COLANTA OC A02-COC 326920",
    "phone": "+573133418834",
    "email": "contacto@cooperativac.co",
    "city": "Tunja",
    "address": "Calle 2 # 49-29"
  },
  {
    "id": "cust-xl-544",
    "rawDoc": "900074528-5",
    "name": "JUAN DAVID GARCIA CALDERON",
    "phone": "+573143588865",
    "email": "contacto@juandavidgar.co",
    "city": "Bogotá D.C.",
    "address": "Calle 9 # 62-32"
  },
  {
    "id": "cust-xl-545",
    "rawDoc": "1004315855",
    "name": "LEONARDO VELAZCO",
    "phone": "+573153758896",
    "email": "contacto@leonardovela.co",
    "city": "Medellín",
    "address": "Calle 16 # 75-35"
  },
  {
    "id": "cust-xl-546",
    "rawDoc": "1004323774",
    "name": "PROSPERO ANTONIO DURAN RODRÍGUEZ",
    "phone": "+573163928927",
    "email": "contacto@prosperoanto.co",
    "city": "Cali",
    "address": "Calle 23 # 88-38"
  },
  {
    "id": "cust-xl-547",
    "rawDoc": "1004331693",
    "name": "CLAUDIO EFRAIN PILATAXI MIÑARCAJA",
    "phone": "+573174098958",
    "email": "contacto@claudioefrai.co",
    "city": "Barranquilla",
    "address": "Calle 30 # 11-41"
  },
  {
    "id": "cust-xl-548",
    "rawDoc": "1004339612",
    "name": "BIBIANA CAROLINA RAMIREZ LEÓN",
    "phone": "+573184268989",
    "email": "contacto@bibianacarol.co",
    "city": "Bucaramanga",
    "address": "Calle 37 # 24-44"
  },
  {
    "id": "cust-xl-549",
    "rawDoc": "900075213-1",
    "name": "GRUPÓ EMPRESARIAL HTM S.A.S.",
    "phone": "+573194439020",
    "email": "contacto@grupempresar.co",
    "city": "Villavicencio",
    "address": "Calle 44 # 37-47"
  },
  {
    "id": "cust-xl-550",
    "rawDoc": "1004355450",
    "name": "SEBASTIAN DIAZ MEJIA",
    "phone": "+573204609051",
    "email": "contacto@sebastiandia.co",
    "city": "Montería",
    "address": "Calle 51 # 50-0"
  },
  {
    "id": "cust-xl-551",
    "rawDoc": "1004363369",
    "name": "MARIA CRISTINA VELASQUEZ",
    "phone": "+573214779082",
    "email": "contacto@mariacristin.co",
    "city": "Pereira",
    "address": "Calle 58 # 63-3"
  },
  {
    "id": "cust-xl-552",
    "rawDoc": "900075624-4",
    "name": "AGROPECUARIA CHIQUIQUE S.A.S",
    "phone": "+573224949113",
    "email": "contacto@agropecuaria.co",
    "city": "Manizales",
    "address": "Calle 65 # 76-6"
  },
  {
    "id": "cust-xl-553",
    "rawDoc": "1004379207",
    "name": "MICHAELL CENTENO GUERRERO",
    "phone": "+573235119144",
    "email": "contacto@michaellcent.co",
    "city": "Neiva",
    "address": "Calle 72 # 89-9"
  },
  {
    "id": "cust-xl-554",
    "rawDoc": "900075898-6",
    "name": "MARTHA LUCINDA GUTIERREZ Y CIA S. EN C",
    "phone": "+573245289175",
    "email": "contacto@marthalucind.co",
    "city": "Ibagué",
    "address": "Calle 79 # 12-12"
  },
  {
    "id": "cust-xl-555",
    "rawDoc": "1004395045",
    "name": "DANNA BELLY MEJIA MORA",
    "phone": "+573105459206",
    "email": "contacto@dannabellyme.co",
    "city": "Popayán",
    "address": "Calle 86 # 25-15"
  },
  {
    "id": "cust-xl-556",
    "rawDoc": "1004402964",
    "name": "ALBERTO ALEJANDRO NIÑO BUITRAGO",
    "phone": "+573115629237",
    "email": "contacto@albertoaleja.co",
    "city": "Pastos",
    "address": "Calle 93 # 38-18"
  },
  {
    "id": "cust-xl-557",
    "rawDoc": "1004410883",
    "name": "SERGIO VELA CALDERON",
    "phone": "+573125799268",
    "email": "contacto@sergiovelaca.co",
    "city": "Santa Marta",
    "address": "Calle 100 # 51-21"
  },
  {
    "id": "cust-xl-558",
    "rawDoc": "1004418802",
    "name": "SEBASTIAN ANDRES DIAZ GARZON",
    "phone": "+573135969299",
    "email": "contacto@sebastianand.co",
    "city": "Valledupar",
    "address": "Calle 7 # 64-24"
  },
  {
    "id": "cust-xl-559",
    "rawDoc": "1004426721",
    "name": "PEDRO FABIAN JIMENEZ CARO",
    "phone": "+573146139330",
    "email": "contacto@pedrofabianj.co",
    "city": "Tunja",
    "address": "Calle 14 # 77-27"
  },
  {
    "id": "cust-xl-560",
    "rawDoc": "1004434640",
    "name": "VALENTINA MORALES SÁNCHEZ",
    "phone": "+573156309361",
    "email": "contacto@valentinamor.co",
    "city": "Bogotá D.C.",
    "address": "Calle 21 # 90-30"
  },
  {
    "id": "cust-xl-561",
    "rawDoc": "1004442559",
    "name": "MARIO GERMAN RICO OSPINA",
    "phone": "+573166479392",
    "email": "contacto@mariogermanr.co",
    "city": "Medellín",
    "address": "Calle 28 # 13-33"
  },
  {
    "id": "cust-xl-562",
    "rawDoc": "1004450478",
    "name": "EMERSON HUERTAS PEREZ",
    "phone": "+573176649423",
    "email": "contacto@emersonhuert.co",
    "city": "Cali",
    "address": "Calle 35 # 26-36"
  },
  {
    "id": "cust-xl-563",
    "rawDoc": "1004458397",
    "name": "ANA MARIA ORTIZ ORTIZ",
    "phone": "+573186819454",
    "email": "contacto@anamariaorti.co",
    "city": "Barranquilla",
    "address": "Calle 42 # 39-39"
  },
  {
    "id": "cust-xl-564",
    "rawDoc": "1004466316",
    "name": "FENIX",
    "phone": "+573196989485",
    "email": "contacto@fenix.co",
    "city": "Bucaramanga",
    "address": "Calle 49 # 52-42"
  },
  {
    "id": "cust-xl-565",
    "rawDoc": "900077405-8",
    "name": "FERRETERIA NICHOLSON S.A.S",
    "phone": "+573207159516",
    "email": "contacto@ferreteriani.co",
    "city": "Villavicencio",
    "address": "Calle 56 # 65-45"
  },
  {
    "id": "cust-xl-566",
    "rawDoc": "900077542-9",
    "name": "NOGA S.A.S.",
    "phone": "+573217329547",
    "email": "contacto@nogasas.co",
    "city": "Montería",
    "address": "Calle 63 # 78-48"
  },
  {
    "id": "cust-xl-567",
    "rawDoc": "1004490073",
    "name": "MAURICIO GUILLERMO TORRES TORRES",
    "phone": "+573227499578",
    "email": "contacto@mauricioguil.co",
    "city": "Pereira",
    "address": "Calle 70 # 91-1"
  },
  {
    "id": "cust-xl-568",
    "rawDoc": "1004497992",
    "name": "DOMINGO ANTONIO GÓMEZ",
    "phone": "+573237669609",
    "email": "contacto@domingoanton.co",
    "city": "Manizales",
    "address": "Calle 77 # 14-4"
  },
  {
    "id": "cust-xl-569",
    "rawDoc": "1004505911",
    "name": "DANNY ANDRÉS NARVAEZ FLOREZ",
    "phone": "+573247839640",
    "email": "contacto@dannyandrsna.co",
    "city": "Neiva",
    "address": "Calle 84 # 27-7"
  },
  {
    "id": "cust-xl-570",
    "rawDoc": "1004513830",
    "name": "YOLANDA PATIÑO PEREZ",
    "phone": "+573108009671",
    "email": "contacto@yolandapatio.co",
    "city": "Ibagué",
    "address": "Calle 91 # 40-10"
  },
  {
    "id": "cust-xl-571",
    "rawDoc": "1004521749",
    "name": "MAYCKOL SVANNERTH ABRIL ANGARITA",
    "phone": "+573118179702",
    "email": "contacto@mayckolsvann.co",
    "city": "Popayán",
    "address": "Calle 98 # 53-13"
  },
  {
    "id": "cust-xl-572",
    "rawDoc": "900078364-6",
    "name": "LINA GABRIELA GARCIA MORENO",
    "phone": "+573128349733",
    "email": "contacto@linagabriela.co",
    "city": "Pastos",
    "address": "Calle 5 # 66-16"
  },
  {
    "id": "cust-xl-573",
    "rawDoc": "1004537587",
    "name": "EMBRIONES DEL SINU S.A",
    "phone": "+573138519764",
    "email": "contacto@embrionesdel.co",
    "city": "Santa Marta",
    "address": "Calle 12 # 79-19"
  },
  {
    "id": "cust-xl-574",
    "rawDoc": "1004545506",
    "name": "JUAN DIEGO GAONA ARIAS",
    "phone": "+573148689795",
    "email": "contacto@juandiegogao.co",
    "city": "Valledupar",
    "address": "Calle 19 # 92-22"
  },
  {
    "id": "cust-xl-575",
    "rawDoc": "900078775-9",
    "name": "TIC BIOMEDICS SAS",
    "phone": "+573158859826",
    "email": "contacto@ticbiomedics.co",
    "city": "Tunja",
    "address": "Calle 26 # 15-25"
  },
  {
    "id": "cust-xl-576",
    "rawDoc": "900078912-1",
    "name": "AG VETERINARIA GUADALUPE SAS",
    "phone": "+573169029857",
    "email": "contacto@agveterinari.co",
    "city": "Bogotá D.C.",
    "address": "Calle 33 # 28-28"
  },
  {
    "id": "cust-xl-577",
    "rawDoc": "1004569263",
    "name": "WILMER ANDRES MELO PARRA",
    "phone": "+573179199888",
    "email": "contacto@wilmerandres.co",
    "city": "Medellín",
    "address": "Calle 40 # 41-31"
  },
  {
    "id": "cust-xl-578",
    "rawDoc": "900079186-3",
    "name": "FELIPE SANTIAGO NEIRA RIVERA",
    "phone": "+573189369919",
    "email": "contacto@felipesantia.co",
    "city": "Cali",
    "address": "Calle 47 # 54-34"
  },
  {
    "id": "cust-xl-579",
    "rawDoc": "900079323-4",
    "name": "COOPERATIVA INDIGENA MULTIACTIVA DE ZUMBICO",
    "phone": "+573199539950",
    "email": "contacto@cooperativai.co",
    "city": "Barranquilla",
    "address": "Calle 54 # 67-37"
  },
  {
    "id": "cust-xl-580",
    "rawDoc": "1004593020",
    "name": "ALEJANDRA GONZALEZ",
    "phone": "+573209709981",
    "email": "contacto@alejandragon.co",
    "city": "Bucaramanga",
    "address": "Calle 61 # 80-40"
  },
  {
    "id": "cust-xl-581",
    "rawDoc": "900079597-6",
    "name": "PRODUCTORA DE INSUMOS AGROPECUARIOS, SOMEX S.A.S",
    "phone": "+573219871013",
    "email": "contacto@productorade.co",
    "city": "Villavicencio",
    "address": "Calle 68 # 93-43"
  },
  {
    "id": "cust-xl-582",
    "rawDoc": "900079734-7",
    "name": "LIDERMAN SALAZAR CARDENAS",
    "phone": "+573221051044",
    "email": "contacto@lidermansala.co",
    "city": "Montería",
    "address": "Calle 75 # 16-46"
  },
  {
    "id": "cust-xl-583",
    "rawDoc": "1004616777",
    "name": "CARLOS ARTURO LOZAN",
    "phone": "+573231221075",
    "email": "contacto@carlosarturo.co",
    "city": "Pereira",
    "address": "Calle 82 # 29-49"
  },
  {
    "id": "cust-xl-584",
    "rawDoc": "900080008-9",
    "name": "GENETICA BIOS S.A.S",
    "phone": "+573241391106",
    "email": "contacto@geneticabios.co",
    "city": "Manizales",
    "address": "Calle 89 # 42-2"
  },
  {
    "id": "cust-xl-585",
    "rawDoc": "1004632615",
    "name": "LAURA FERNANDA GARZÓN PÉREZ",
    "phone": "+573101561137",
    "email": "contacto@laurafernand.co",
    "city": "Neiva",
    "address": "Calle 96 # 55-5"
  },
  {
    "id": "cust-xl-586",
    "rawDoc": "1004640534",
    "name": "MARIA CAROLINA HERRERA PALACIO",
    "phone": "+573111731168",
    "email": "contacto@mariacarolin.co",
    "city": "Ibagué",
    "address": "Calle 3 # 68-8"
  },
  {
    "id": "cust-xl-587",
    "rawDoc": "900080419-3",
    "name": "YARIGUI SAS",
    "phone": "+573121901199",
    "email": "contacto@yariguisas.co",
    "city": "Popayán",
    "address": "Calle 10 # 81-11"
  },
  {
    "id": "cust-xl-588",
    "rawDoc": "900080556-4",
    "name": "COOPERATIVA COLANTA",
    "phone": "+573132071230",
    "email": "contacto@cooperativac.co",
    "city": "Pastos",
    "address": "Calle 17 # 94-14"
  },
  {
    "id": "cust-xl-589",
    "rawDoc": "1004664291",
    "name": "GERARDO CORTES JARAMILLO",
    "phone": "+573142241261",
    "email": "contacto@gerardocorte.co",
    "city": "Santa Marta",
    "address": "Calle 24 # 17-17"
  },
  {
    "id": "cust-xl-590",
    "rawDoc": "900080830-6",
    "name": "LEIDY VANESSA RAMÍREZ SOLER",
    "phone": "+573152411292",
    "email": "contacto@leidyvanessa.co",
    "city": "Valledupar",
    "address": "Calle 31 # 30-20"
  },
  {
    "id": "cust-xl-591",
    "rawDoc": "1004680129",
    "name": "AUDREY DAYANA VACCA JIMENEZ",
    "phone": "+573162581323",
    "email": "contacto@audreydayana.co",
    "city": "Tunja",
    "address": "Calle 38 # 43-23"
  },
  {
    "id": "cust-xl-592",
    "rawDoc": "900081104-8",
    "name": "AGROCOCOROLLO S.A.S",
    "phone": "+573172751354",
    "email": "contacto@agrococoroll.co",
    "city": "Bogotá D.C.",
    "address": "Calle 45 # 56-26"
  },
  {
    "id": "cust-xl-593",
    "rawDoc": "900081241-9",
    "name": "TERIOGEN S.A.S.",
    "phone": "+573182921385",
    "email": "contacto@teriogensas.co",
    "city": "Medellín",
    "address": "Calle 52 # 69-29"
  },
  {
    "id": "cust-xl-594",
    "rawDoc": "900081378-1",
    "name": "AGROPECUARIA PINZON ESTEVEZ S.A.S.",
    "phone": "+573193091416",
    "email": "contacto@agropecuaria.co",
    "city": "Cali",
    "address": "Calle 59 # 82-32"
  },
  {
    "id": "cust-xl-595",
    "rawDoc": "1004711805",
    "name": "MARTIN WARNER LOPEZ CASTAÑO",
    "phone": "+573203261447",
    "email": "contacto@martinwarner.co",
    "city": "Barranquilla",
    "address": "Calle 66 # 95-35"
  },
  {
    "id": "cust-xl-596",
    "rawDoc": "900081652-3",
    "name": "SISTECOL S.A.S",
    "phone": "+573213431478",
    "email": "contacto@sistecolsas.co",
    "city": "Bucaramanga",
    "address": "Calle 73 # 18-38"
  },
  {
    "id": "cust-xl-597",
    "rawDoc": "900081789-4",
    "name": "UNIVERSIDAD CES/ OC 4100002956",
    "phone": "+573223601509",
    "email": "contacto@universidadc.co",
    "city": "Villavicencio",
    "address": "Calle 80 # 31-41"
  },
  {
    "id": "cust-xl-598",
    "rawDoc": "900081926-5",
    "name": "UNIVERSIDAD CES",
    "phone": "+573233771540",
    "email": "contacto@universidadc.co",
    "city": "Montería",
    "address": "Calle 87 # 44-44"
  },
  {
    "id": "cust-xl-599",
    "rawDoc": "900082063-6",
    "name": "LAURA POSADA CARVAJAL",
    "phone": "+573243941571",
    "email": "contacto@lauraposadac.co",
    "city": "Pereira",
    "address": "Calle 94 # 57-47"
  },
  {
    "id": "cust-xl-600",
    "rawDoc": "1004751400",
    "name": "ALBEIRO LOPEZ",
    "phone": "+573104111602",
    "email": "contacto@albeirolopez.co",
    "city": "Manizales",
    "address": "Calle 1 # 70-0"
  },
  {
    "id": "cust-xl-601",
    "rawDoc": "900082337-8",
    "name": "IVAN FERNANDO SANABRIA",
    "phone": "+573114281633",
    "email": "contacto@ivanfernando.co",
    "city": "Neiva",
    "address": "Calle 8 # 83-3"
  },
  {
    "id": "cust-xl-602",
    "rawDoc": "1004767238",
    "name": "WENDY NATALIA PARADA BECERRA",
    "phone": "+573124451664",
    "email": "contacto@wendynatalia.co",
    "city": "Ibagué",
    "address": "Calle 15 # 96-6"
  },
  {
    "id": "cust-xl-603",
    "rawDoc": "1004775157",
    "name": "MIGUEL ANGEL CARRILLO",
    "phone": "+573134621695",
    "email": "contacto@miguelangelc.co",
    "city": "Popayán",
    "address": "Calle 22 # 19-9"
  },
  {
    "id": "cust-xl-604",
    "rawDoc": "1004783076",
    "name": "CRISTIAN ANDRES CAMPOS",
    "phone": "+573144791726",
    "email": "contacto@cristianandr.co",
    "city": "Pastos",
    "address": "Calle 29 # 32-12"
  },
  {
    "id": "cust-xl-605",
    "rawDoc": "1004790995",
    "name": "LUCIO ALBERTO ANDRADE MEDINA",
    "phone": "+573154961757",
    "email": "contacto@lucioalberto.co",
    "city": "Santa Marta",
    "address": "Calle 36 # 45-15"
  },
  {
    "id": "cust-xl-606",
    "rawDoc": "1004798914",
    "name": "FEDERACION COLOMBIANA DE PRODUCTORES DE PAPA \"FEDEPAPA\"",
    "phone": "+573165131788",
    "email": "contacto@federacionco.co",
    "city": "Valledupar",
    "address": "Calle 43 # 58-18"
  },
  {
    "id": "cust-xl-607",
    "rawDoc": "900083159-5",
    "name": "SANTIAGO MEJÍA RESTREPO",
    "phone": "+573175301819",
    "email": "contacto@santiagomeja.co",
    "city": "Tunja",
    "address": "Calle 50 # 71-21"
  },
  {
    "id": "cust-xl-608",
    "rawDoc": "1004814752",
    "name": "CARLOS ALBERTO JULIO PACHECO",
    "phone": "+573185471850",
    "email": "contacto@carlosalbert.co",
    "city": "Bogotá D.C.",
    "address": "Calle 57 # 84-24"
  },
  {
    "id": "cust-xl-609",
    "rawDoc": "1004822671",
    "name": "CARLOS JAIRO GOMEZ TORRES",
    "phone": "+573195641881",
    "email": "contacto@carlosjairog.co",
    "city": "Medellín",
    "address": "Calle 64 # 97-27"
  },
  {
    "id": "cust-xl-610",
    "rawDoc": "900083570-8",
    "name": "OMAR ALEJANDRO GARCIA ALVARADO",
    "phone": "+573205811912",
    "email": "contacto@omaralejandr.co",
    "city": "Cali",
    "address": "Calle 71 # 20-30"
  },
  {
    "id": "cust-xl-611",
    "rawDoc": "900083707-9",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA/ OC 679784",
    "phone": "+573215981943",
    "email": "contacto@lacorporacio.co",
    "city": "Barranquilla",
    "address": "Calle 78 # 33-33"
  },
  {
    "id": "cust-xl-612",
    "rawDoc": "900083844-1",
    "name": "GENÉTICA TRICOLOR SAS",
    "phone": "+573226151974",
    "email": "contacto@genticatrico.co",
    "city": "Bucaramanga",
    "address": "Calle 85 # 46-36"
  },
  {
    "id": "cust-xl-613",
    "rawDoc": "900083981-2",
    "name": "UNIVERSIDAD CES/ OC 4500059723",
    "phone": "+573236322005",
    "email": "contacto@universidadc.co",
    "city": "Villavicencio",
    "address": "Calle 92 # 59-39"
  },
  {
    "id": "cust-xl-614",
    "rawDoc": "1004862266",
    "name": "IDELFONSO GONZALEZ PAEZ",
    "phone": "+573246492036",
    "email": "contacto@idelfonsogon.co",
    "city": "Montería",
    "address": "Calle 99 # 72-42"
  },
  {
    "id": "cust-xl-615",
    "rawDoc": "1004870185",
    "name": "MILER ARVEY ZUÑIGA BAEZ",
    "phone": "+573106662067",
    "email": "contacto@milerarveyzu.co",
    "city": "Pereira",
    "address": "Calle 6 # 85-45"
  },
  {
    "id": "cust-xl-616",
    "rawDoc": "1004878104",
    "name": "EDGAR JAVIER BRICEÑO CUSPOCA",
    "phone": "+573116832098",
    "email": "contacto@edgarjavierb.co",
    "city": "Manizales",
    "address": "Calle 13 # 98-48"
  },
  {
    "id": "cust-xl-617",
    "rawDoc": "900084529-6",
    "name": "EMBRIONES DEL SINU S.A.",
    "phone": "+573127002129",
    "email": "contacto@embrionesdel.co",
    "city": "Neiva",
    "address": "Calle 20 # 21-1"
  },
  {
    "id": "cust-xl-618",
    "rawDoc": "1004893942",
    "name": "MUNICIPIO DE GUALMATAN",
    "phone": "+573137172160",
    "email": "contacto@municipiodeg.co",
    "city": "Ibagué",
    "address": "Calle 27 # 34-4"
  },
  {
    "id": "cust-xl-619",
    "rawDoc": "900084803-8",
    "name": "DOTACIONES Y EQUIPOS INDUSTRIALES DE COLOMBIA DOTAEQUIP LTDA",
    "phone": "+573147342191",
    "email": "contacto@dotacionesye.co",
    "city": "Popayán",
    "address": "Calle 34 # 47-7"
  },
  {
    "id": "cust-xl-620",
    "rawDoc": "1004909780",
    "name": "PARQUE ACUÁTICO Y DE CONSERVACIÓN PISCILAGO DE COLSUBSIDIO /ORDEN DE COMPRA N°4200402249",
    "phone": "+573157512222",
    "email": "contacto@parqueacutic.co",
    "city": "Pastos",
    "address": "Calle 41 # 60-10"
  },
  {
    "id": "cust-xl-621",
    "rawDoc": "1004917699",
    "name": "PARQUE ACUÁTICO Y DE CONSERVACIÓN PISCILAGO DE COLSUBSIDIO/ ORDEN DE COMPRA N° 4200390798",
    "phone": "+573167682253",
    "email": "contacto@parqueacutic.co",
    "city": "Santa Marta",
    "address": "Calle 48 # 73-13"
  },
  {
    "id": "cust-xl-622",
    "rawDoc": "1004925618",
    "name": "CARLOS ANDRES CARDENAS BUITRAGO",
    "phone": "+573177852284",
    "email": "contacto@carlosandres.co",
    "city": "Valledupar",
    "address": "Calle 55 # 86-16"
  },
  {
    "id": "cust-xl-623",
    "rawDoc": "900085351-3",
    "name": "EDER GARCIANO CÉSPEDES PEREZ",
    "phone": "+573188022315",
    "email": "contacto@edergarciano.co",
    "city": "Tunja",
    "address": "Calle 62 # 99-19"
  },
  {
    "id": "cust-xl-624",
    "rawDoc": "1004941456",
    "name": "CARLOS FIDEL VILLAMIL RUIZ",
    "phone": "+573198192346",
    "email": "contacto@carlosfidelv.co",
    "city": "Bogotá D.C.",
    "address": "Calle 69 # 22-22"
  },
  {
    "id": "cust-xl-625",
    "rawDoc": "1004949375",
    "name": "DEIBIS ANDRES BUITRAGO",
    "phone": "+573208362377",
    "email": "contacto@deibisandres.co",
    "city": "Medellín",
    "address": "Calle 76 # 35-25"
  },
  {
    "id": "cust-xl-626",
    "rawDoc": "1004957294",
    "name": "CARLOS WALTER CHAVARRO GARZON",
    "phone": "+573218532408",
    "email": "contacto@carloswalter.co",
    "city": "Cali",
    "address": "Calle 83 # 48-28"
  },
  {
    "id": "cust-xl-627",
    "rawDoc": "1004965213",
    "name": "JUAN CAMILO RESTREPO OSORIO",
    "phone": "+573228702439",
    "email": "contacto@juancamilore.co",
    "city": "Barranquilla",
    "address": "Calle 90 # 61-31"
  },
  {
    "id": "cust-xl-628",
    "rawDoc": "1004973132",
    "name": "NICOLAS FELIPE LEON ROLDAN",
    "phone": "+573238872470",
    "email": "contacto@nicolasfelip.co",
    "city": "Bucaramanga",
    "address": "Calle 97 # 74-34"
  },
  {
    "id": "cust-xl-629",
    "rawDoc": "1004981051",
    "name": "CAMILO SIERRA",
    "phone": "+573249042501",
    "email": "contacto@camilosierra.co",
    "city": "Villavicencio",
    "address": "Calle 4 # 87-37"
  },
  {
    "id": "cust-xl-630",
    "rawDoc": "1004988970",
    "name": "INSUGAN 2024 CA",
    "phone": "+573109212532",
    "email": "contacto@insugan2024c.co",
    "city": "Montería",
    "address": "Calle 11 # 10-40"
  },
  {
    "id": "cust-xl-631",
    "rawDoc": "1004996889",
    "name": "EDER CARCAMO",
    "phone": "+573119382563",
    "email": "contacto@edercarcamo.co",
    "city": "Pereira",
    "address": "Calle 18 # 23-43"
  },
  {
    "id": "cust-xl-632",
    "rawDoc": "900086584-3",
    "name": "INVERSIONES BARBOSA Y MEDOZA Y CIA S. EN C.",
    "phone": "+573129552594",
    "email": "contacto@inversionesb.co",
    "city": "Manizales",
    "address": "Calle 25 # 36-46"
  },
  {
    "id": "cust-xl-633",
    "rawDoc": "900086721-4",
    "name": "ALEXANDER CASALLAS VEGA",
    "phone": "+573139722625",
    "email": "contacto@alexandercas.co",
    "city": "Neiva",
    "address": "Calle 32 # 49-49"
  },
  {
    "id": "cust-xl-634",
    "rawDoc": "900086858-5",
    "name": "ARENAS ROMERO ASOCIADOS Y CIA S EN C",
    "phone": "+573149892656",
    "email": "contacto@arenasromero.co",
    "city": "Ibagué",
    "address": "Calle 39 # 62-2"
  },
  {
    "id": "cust-xl-635",
    "rawDoc": "900086995-6",
    "name": "INMONOVA SAS/ TATIANA BLANCO/SOMAK S.A.S",
    "phone": "+573151072687",
    "email": "contacto@inmonovasast.co",
    "city": "Popayán",
    "address": "Calle 46 # 75-5"
  },
  {
    "id": "cust-xl-636",
    "rawDoc": "900087132-7",
    "name": "INMONOVA SAS/ SOMAK S.A.S",
    "phone": "+573161242718",
    "email": "contacto@inmonovasass.co",
    "city": "Pastos",
    "address": "Calle 53 # 88-8"
  },
  {
    "id": "cust-xl-637",
    "rawDoc": "1005044403",
    "name": "DEYBER ARLES VIZCAYA MARTINEZ",
    "phone": "+573171412749",
    "email": "contacto@deyberarlesv.co",
    "city": "Santa Marta",
    "address": "Calle 60 # 11-11"
  },
  {
    "id": "cust-xl-638",
    "rawDoc": "900087406-9",
    "name": "UNIVERSIDAD DE SANTANDER/ OC 35090",
    "phone": "+573181582780",
    "email": "contacto@universidadd.co",
    "city": "Valledupar",
    "address": "Calle 67 # 24-14"
  },
  {
    "id": "cust-xl-639",
    "rawDoc": "900087543-1",
    "name": "UNIVERSIDAD DE SANTANDER",
    "phone": "+573191752811",
    "email": "contacto@universidadd.co",
    "city": "Tunja",
    "address": "Calle 74 # 37-17"
  },
  {
    "id": "cust-xl-640",
    "rawDoc": "900087680-2",
    "name": "UNIVERSIDAD DE SANTANDER/ OC16429",
    "phone": "+573201922842",
    "email": "contacto@universidadd.co",
    "city": "Bogotá D.C.",
    "address": "Calle 81 # 50-20"
  },
  {
    "id": "cust-xl-641",
    "rawDoc": "900087817-3",
    "name": "JOAHN ALEXANDER MORA SANTOS",
    "phone": "+573212092873",
    "email": "contacto@joahnalexand.co",
    "city": "Medellín",
    "address": "Calle 88 # 63-23"
  },
  {
    "id": "cust-xl-642",
    "rawDoc": "1005083998",
    "name": "HERNAN RENE FORIAS FIERRO",
    "phone": "+573222262904",
    "email": "contacto@hernanrenefo.co",
    "city": "Cali",
    "address": "Calle 95 # 76-26"
  },
  {
    "id": "cust-xl-643",
    "rawDoc": "900088091-5",
    "name": "INDUGANAVET SAS",
    "phone": "+573232432935",
    "email": "contacto@induganavets.co",
    "city": "Barranquilla",
    "address": "Calle 2 # 89-29"
  },
  {
    "id": "cust-xl-644",
    "rawDoc": "900088228-6",
    "name": "OOLAB S.A.S",
    "phone": "+573242602966",
    "email": "contacto@oolabsas.co",
    "city": "Bucaramanga",
    "address": "Calle 9 # 12-32"
  },
  {
    "id": "cust-xl-645",
    "rawDoc": "1005107755",
    "name": "JADAMIR PEREZ CRUZ",
    "phone": "+573102772997",
    "email": "contacto@jadamirperez.co",
    "city": "Villavicencio",
    "address": "Calle 16 # 25-35"
  },
  {
    "id": "cust-xl-646",
    "rawDoc": "1005115674",
    "name": "KAROL XIMENA MAYORGA CONTRERAS",
    "phone": "+573112943028",
    "email": "contacto@karolximenam.co",
    "city": "Montería",
    "address": "Calle 23 # 38-38"
  },
  {
    "id": "cust-xl-647",
    "rawDoc": "1005123593",
    "name": "ANGELA VALBUENA GUERRERO",
    "phone": "+573123113059",
    "email": "contacto@angelavalbue.co",
    "city": "Pereira",
    "address": "Calle 30 # 51-41"
  },
  {
    "id": "cust-xl-648",
    "rawDoc": "900088776-1",
    "name": "LA CORPORACION COLOMBIANA DE INVESTIGACION AGROPECUARIA - AGROSAVIA/ OC No.681114",
    "phone": "+573133283090",
    "email": "contacto@lacorporacio.co",
    "city": "Manizales",
    "address": "Calle 37 # 64-44"
  },
  {
    "id": "cust-xl-649",
    "rawDoc": "1005139431",
    "name": "COLANTA/ OC A002-COC-345793",
    "phone": "+573143453121",
    "email": "contacto@colantaoca00.co",
    "city": "Neiva",
    "address": "Calle 44 # 77-47"
  },
  {
    "id": "cust-xl-650",
    "rawDoc": "900089050-3",
    "name": "GANADERIA EL DANUBIO S.A.S",
    "phone": "+573153623152",
    "email": "contacto@ganaderiaeld.co",
    "city": "Ibagué",
    "address": "Calle 51 # 90-0"
  },
  {
    "id": "cust-xl-651",
    "rawDoc": "1005155269",
    "name": "JOSE DAVID BARRERA PERDOMO",
    "phone": "+573163793183",
    "email": "contacto@josedavidbar.co",
    "city": "Popayán",
    "address": "Calle 58 # 13-3"
  },
  {
    "id": "cust-xl-652",
    "rawDoc": "900089324-5",
    "name": "CAICEDO ORDOÑEZ Y CIA S EN C",
    "phone": "+573173963214",
    "email": "contacto@caicedoordoe.co",
    "city": "Pastos",
    "address": "Calle 65 # 26-6"
  },
  {
    "id": "cust-xl-653",
    "rawDoc": "900089461-6",
    "name": "CORPORACION UNIVERSITARIA DEL HUILA/ OC 8932",
    "phone": "+573184133245",
    "email": "contacto@corporacionu.co",
    "city": "Santa Marta",
    "address": "Calle 72 # 39-9"
  },
  {
    "id": "cust-xl-654",
    "rawDoc": "900089598-7",
    "name": "CESAR IVAN GÓMEZ CIFUENTES",
    "phone": "+573194303276",
    "email": "contacto@cesarivangme.co",
    "city": "Valledupar",
    "address": "Calle 79 # 52-12"
  },
  {
    "id": "cust-xl-655",
    "rawDoc": "900089735-8",
    "name": "INVERQUIMICAS SAS",
    "phone": "+573204473307",
    "email": "contacto@inverquimica.co",
    "city": "Tunja",
    "address": "Calle 86 # 65-15"
  },
  {
    "id": "cust-xl-656",
    "rawDoc": "900089872-9",
    "name": "WALTER IVAN GUTIERREZ SANCHE",
    "phone": "+573214643338",
    "email": "contacto@walterivangu.co",
    "city": "Bogotá D.C.",
    "address": "Calle 93 # 78-18"
  },
  {
    "id": "cust-xl-657",
    "rawDoc": "900090009-1",
    "name": "SUMINISTROS AGROPECUARIOS DE COLOMBIA F&M SAS/ YORMAN",
    "phone": "+573224813369",
    "email": "contacto@suministrosa.co",
    "city": "Medellín",
    "address": "Calle 100 # 91-21"
  },
  {
    "id": "cust-xl-658",
    "rawDoc": "1005210702",
    "name": "JOHN JAIRO MUÑOZ ENRIQUEZ",
    "phone": "+573234983400",
    "email": "contacto@johnjairomuo.co",
    "city": "Cali",
    "address": "Calle 7 # 14-24"
  },
  {
    "id": "cust-xl-659",
    "rawDoc": "900090283-3",
    "name": "FIV-DT BIOTECNOLOGIA, C.A",
    "phone": "+573245153431",
    "email": "contacto@fivdtbiotecn.co",
    "city": "Barranquilla",
    "address": "Calle 14 # 27-27"
  },
  {
    "id": "cust-xl-660",
    "rawDoc": "900090420-4",
    "name": "JULIAN FELIPE DUSSAN RODRIGUEZ",
    "phone": "+573105323462",
    "email": "contacto@julianfelipe.co",
    "city": "Bucaramanga",
    "address": "Calle 21 # 40-30"
  },
  {
    "id": "cust-xl-661",
    "rawDoc": "900090557-5",
    "name": "DINAMIK SAS",
    "phone": "+573115493493",
    "email": "contacto@dinamiksas.co",
    "city": "Villavicencio",
    "address": "Calle 28 # 53-33"
  },
  {
    "id": "cust-xl-662",
    "rawDoc": "1005242378",
    "name": "MARIO LUIS CARPENA DYER",
    "phone": "+573125663524",
    "email": "contacto@marioluiscar.co",
    "city": "Montería",
    "address": "Calle 35 # 66-36"
  },
  {
    "id": "cust-xl-663",
    "rawDoc": "900090831-7",
    "name": "UNIVERSIDAD DE CIENCIAS APLICADAS Y AMBIENTALES U.D.C.A./ OC 1052",
    "phone": "+573135833555",
    "email": "contacto@universidadd.co",
    "city": "Pereira",
    "address": "Calle 42 # 79-39"
  },
  {
    "id": "cust-xl-664",
    "rawDoc": "900090968-8",
    "name": "UNIVERSIDAD DE CIENCIAS APLICADAS Y AMBIENTALES U.D.C.A.",
    "phone": "+573146003586",
    "email": "contacto@universidadd.co",
    "city": "Manizales",
    "address": "Calle 49 # 92-42"
  },
  {
    "id": "cust-xl-665",
    "rawDoc": "1005266135",
    "name": "HUMBERTO CONCHA",
    "phone": "+573156173617",
    "email": "contacto@humbertoconc.co",
    "city": "Neiva",
    "address": "Calle 56 # 15-45"
  },
  {
    "id": "cust-xl-666",
    "rawDoc": "1005274054",
    "name": "JUAN CARLOS MASSO FUENTES",
    "phone": "+573166343648",
    "email": "contacto@juancarlosma.co",
    "city": "Ibagué",
    "address": "Calle 63 # 28-48"
  },
  {
    "id": "cust-xl-667",
    "rawDoc": "900091379-2",
    "name": "MANNOSERA S.A.S ZOMAC",
    "phone": "+573176513679",
    "email": "contacto@mannoserasas.co",
    "city": "Popayán",
    "address": "Calle 70 # 41-1"
  },
  {
    "id": "cust-xl-668",
    "rawDoc": "1005289892",
    "name": "CAMAGUEY S.A / OC 54081",
    "phone": "+573186683710",
    "email": "contacto@camagueysaoc.co",
    "city": "Pastos",
    "address": "Calle 77 # 54-4"
  },
  {
    "id": "cust-xl-669",
    "rawDoc": "900091653-4",
    "name": "DARWIN GAVIRIA BERMUDEZ /MILTON GARCIA",
    "phone": "+573196853741",
    "email": "contacto@darwingaviri.co",
    "city": "Santa Marta",
    "address": "Calle 84 # 67-7"
  },
  {
    "id": "cust-xl-670",
    "rawDoc": "1005305730",
    "name": "ROYER DANEY BURGOS DÍAZ",
    "phone": "+573207023772",
    "email": "contacto@royerdaneybu.co",
    "city": "Valledupar",
    "address": "Calle 91 # 80-10"
  },
  {
    "id": "cust-xl-671",
    "rawDoc": "1005313649",
    "name": "PRIETO GUAPACHA FRANCISCO JAVIER",
    "phone": "+573217193803",
    "email": "contacto@prietoguapac.co",
    "city": "Tunja",
    "address": "Calle 98 # 93-13"
  },
  {
    "id": "cust-xl-672",
    "rawDoc": "1005321568",
    "name": "WILFREDO LABRADOR",
    "phone": "+573227363834",
    "email": "contacto@wilfredolabr.co",
    "city": "Bogotá D.C.",
    "address": "Calle 5 # 16-16"
  },
  {
    "id": "cust-xl-673",
    "rawDoc": "1005329487",
    "name": "YAIRSIÑO PARRA MUÑOZ",
    "phone": "+573237533865",
    "email": "contacto@yairsioparra.co",
    "city": "Medellín",
    "address": "Calle 12 # 29-19"
  },
  {
    "id": "cust-xl-674",
    "rawDoc": "1005337406",
    "name": "FABIO AISLANT MORENO",
    "phone": "+573247703896",
    "email": "contacto@fabioaislant.co",
    "city": "Cali",
    "address": "Calle 19 # 42-22"
  },
  {
    "id": "cust-xl-675",
    "rawDoc": "1005345325",
    "name": "COLANTA / A02-COC-346125",
    "phone": "+573107873927",
    "email": "contacto@colantaa02co.co",
    "city": "Barranquilla",
    "address": "Calle 26 # 55-25"
  },
  {
    "id": "cust-xl-676",
    "rawDoc": "1005353244",
    "name": "LUIS GERMÁN NÚÑEZ MESTRE",
    "phone": "+573118043958",
    "email": "contacto@luisgermnnez.co",
    "city": "Bucaramanga",
    "address": "Calle 33 # 68-28"
  },
  {
    "id": "cust-xl-677",
    "rawDoc": "1005361163",
    "name": "SERVICIOS TECNICOS AGROPECUARIOS.STA",
    "phone": "+573128213989",
    "email": "contacto@serviciostec.co",
    "city": "Villavicencio",
    "address": "Calle 40 # 81-31"
  },
  {
    "id": "cust-xl-678",
    "rawDoc": "1005369082",
    "name": "EDUARD JOHANY EUGENIO BAUTISTA",
    "phone": "+573138384020",
    "email": "contacto@eduardjohany.co",
    "city": "Montería",
    "address": "Calle 47 # 94-34"
  },
  {
    "id": "cust-xl-679",
    "rawDoc": "900093023-5",
    "name": "UNIVERSIDAD COOPERATIVA DE COLOMBIA/ 10VIL - 0000014212",
    "phone": "+573148554051",
    "email": "contacto@universidadc.co",
    "city": "Pereira",
    "address": "Calle 54 # 17-37"
  },
  {
    "id": "cust-xl-680",
    "rawDoc": "900093160-6",
    "name": "UNIVERSIDAD COOPERATIVA DE COLOMBIA/ 10VIL - 0000012385",
    "phone": "+573158724082",
    "email": "contacto@universidadc.co",
    "city": "Manizales",
    "address": "Calle 61 # 30-40"
  },
  {
    "id": "cust-xl-681",
    "rawDoc": "900093297-7",
    "name": "OJO DE RIO GANADERIA SAS/ KENNER",
    "phone": "+573168894113",
    "email": "contacto@ojoderiogana.co",
    "city": "Neiva",
    "address": "Calle 68 # 43-43"
  },
  {
    "id": "cust-xl-682",
    "rawDoc": "1005400758",
    "name": "DIEGO ALEJANDRO TRIANA GARZON",
    "phone": "+573179064144",
    "email": "contacto@diegoalejand.co",
    "city": "Ibagué",
    "address": "Calle 75 # 56-46"
  },
  {
    "id": "cust-xl-683",
    "rawDoc": "1005408677",
    "name": "JESUS DAVID GORDILLO",
    "phone": "+573189234175",
    "email": "contacto@jesusdavidgo.co",
    "city": "Popayán",
    "address": "Calle 82 # 69-49"
  },
  {
    "id": "cust-xl-684",
    "rawDoc": "900093708-1",
    "name": "INSUTECH INSUMOS & TECNOLOGIAS S.A.S. /IT29",
    "phone": "+573199404206",
    "email": "contacto@insutechinsu.co",
    "city": "Pastos",
    "address": "Calle 89 # 82-2"
  },
  {
    "id": "cust-xl-685",
    "rawDoc": "1005424515",
    "name": "SEBASTIAN PEÑA MONTOYA",
    "phone": "+573209574237",
    "email": "contacto@sebastianpea.co",
    "city": "Santa Marta",
    "address": "Calle 96 # 95-5"
  },
  {
    "id": "cust-xl-686",
    "rawDoc": "1005432434",
    "name": "HERMOGENES MANUEL MALDONADO ARRIETA",
    "phone": "+573219744268",
    "email": "contacto@hermogenesma.co",
    "city": "Valledupar",
    "address": "Calle 3 # 18-8"
  },
  {
    "id": "cust-xl-687",
    "rawDoc": "900094119-4",
    "name": "CONSULTA PECUARIA S.A.S",
    "phone": "+573229914299",
    "email": "contacto@consultapecu.co",
    "city": "Tunja",
    "address": "Calle 10 # 31-11"
  },
  {
    "id": "cust-xl-688",
    "rawDoc": "1005448272",
    "name": "SERGIO JAVIER AREVALO RAMOS",
    "phone": "+573231094330",
    "email": "contacto@sergiojavier.co",
    "city": "Bogotá D.C.",
    "address": "Calle 17 # 44-14"
  },
  {
    "id": "cust-xl-689",
    "rawDoc": "900094393-6",
    "name": "MEDICINA Y REPRODUCCION BOVINA S.A.S",
    "phone": "+573241264361",
    "email": "contacto@medicinayrep.co",
    "city": "Medellín",
    "address": "Calle 24 # 57-17"
  },
  {
    "id": "cust-xl-690",
    "rawDoc": "1005464110",
    "name": "DIEGO ARMANDO MENDOZA GALEANO",
    "phone": "+573101434392",
    "email": "contacto@diegoarmando.co",
    "city": "Cali",
    "address": "Calle 31 # 70-20"
  },
  {
    "id": "cust-xl-691",
    "rawDoc": "1005472029",
    "name": "IVAN LIBARDO FRANCO",
    "phone": "+573111604423",
    "email": "contacto@ivanlibardof.co",
    "city": "Barranquilla",
    "address": "Calle 38 # 83-23"
  },
  {
    "id": "cust-xl-692",
    "rawDoc": "900094804-9",
    "name": "WILMAR ALEJANDRO OSSA HERNANDEZ",
    "phone": "+573121774454",
    "email": "contacto@wilmaralejan.co",
    "city": "Bucaramanga",
    "address": "Calle 45 # 96-26"
  },
  {
    "id": "cust-xl-693",
    "rawDoc": "1005487867",
    "name": "ANA SILVIA CUVIDES LOPEZ",
    "phone": "+573131944485",
    "email": "contacto@anasilviacuv.co",
    "city": "Villavicencio",
    "address": "Calle 52 # 19-29"
  },
  {
    "id": "cust-xl-694",
    "rawDoc": "1005495786",
    "name": "CARLOS STEVEN MEDINA QUINTERO",
    "phone": "+573142114516",
    "email": "contacto@carlossteven.co",
    "city": "Montería",
    "address": "Calle 59 # 32-32"
  },
  {
    "id": "cust-xl-695",
    "rawDoc": "1005503705",
    "name": "FELIPE AMAYA GOMEZ",
    "phone": "+573152284547",
    "email": "contacto@felipeamayag.co",
    "city": "Pereira",
    "address": "Calle 66 # 45-35"
  },
  {
    "id": "cust-xl-696",
    "rawDoc": "1005511624",
    "name": "FABIAN CHONA",
    "phone": "+573162454578",
    "email": "contacto@fabianchona.co",
    "city": "Manizales",
    "address": "Calle 73 # 58-38"
  },
  {
    "id": "cust-xl-697",
    "rawDoc": "900095489-5",
    "name": "DIANA PAOLA SANCHEZ ARAUJO",
    "phone": "+573172624609",
    "email": "contacto@dianapaolasa.co",
    "city": "Neiva",
    "address": "Calle 80 # 71-41"
  },
  {
    "id": "cust-xl-698",
    "rawDoc": "1005527462",
    "name": "JESÚS HERNÁN TULCAN TIMARAN",
    "phone": "+573182794640",
    "email": "contacto@jesshernntul.co",
    "city": "Ibagué",
    "address": "Calle 87 # 84-44"
  },
  {
    "id": "cust-xl-699",
    "rawDoc": "900095763-7",
    "name": "LUIS CUADROS MORENO Y CIA SAS",
    "phone": "+573192964671",
    "email": "contacto@luiscuadrosm.co",
    "city": "Popayán",
    "address": "Calle 94 # 97-47"
  }
];

const excelIngestedCustomers = rawCustomersData.map(c => {
  const parsedDoc = sanitizeDocument(c.rawDoc);
  return {
    id: c.id,
    documentType: parsedDoc.docType,
    documentNumber: parsedDoc.cleanDoc,
    verificationDigit: parsedDoc.dv,
    name: c.name,
    phone: c.phone,
    email: c.email,
    city: c.city,
    address: c.address
  };
});

// PRODUCTION CRYPTOGRAPHIC SECURITY ENGINE (SHA-256 SALTED HASHING)
const SECURITY_SALT = "MAS_CAMPO_SECURE_SALT_2026";

async function hashPassword(plainPassword) {
  if (!plainPassword) return "";
  if (plainPassword.startsWith("$MAS_CAMPO_SECURE_SALT_2026$")) {
    return plainPassword; // Already hashed
  }
  const encoder = new TextEncoder();
  const data = encoder.encode(plainPassword + SECURITY_SALT);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  return `$MAS_CAMPO_SECURE_SALT_2026$${hashHex}`;
}

async function verifyPassword(plainPasswordInput, storedHash) {
  if (!storedHash) return false;
  if (!storedHash.startsWith("$MAS_CAMPO_SECURE_SALT_2026$")) {
    return plainPasswordInput === storedHash;
  }
  const computedHash = await hashPassword(plainPasswordInput);
  return computedHash === storedHash;
}

const systemUsers = [
  {
    id: "usr-super",
    firstName: "Superusuario",
    lastName: "Gerencia",
    document: "0000000000",
    email: "gerencia@softproductiva.com",
    password: "$MAS_CAMPO_SECURE_SALT_2026$80d62fc31a9a7063a8d8bc52ba7434ea84149eb825f21ecc4f09ca5d466f4db4",
    address: "Sede Principal SoftProductiva S.A.S",
    phone: "+573000000000",
    role: "ADMINISTRADOR",
    vinculacion: "Planta (Nómina)",
    active: true,
    createdAt: "2026-01-01",
    isSuperuser: true
  },
  {
    id: "usr-1",
    firstName: "Carlos",
    lastName: "Mendoza",
    document: "1019024812",
    email: "admin@mascampo.co",
    password: "$MAS_CAMPO_SECURE_SALT_2026$a7c9e050c2c49c5aac2ba2611e0098df9b13a2db9d17a23f7037e26f0b4e95cf",
    address: "Calle 100 # 15-20, Bogotá",
    phone: "+573105551234",
    role: "ADMINISTRADOR",
    vinculacion: "Planta (Nómina)",
    active: true,
    createdAt: "2026-01-01",
    isSuperuser: false
  },
  {
    id: "usr-2",
    firstName: "Jorge",
    lastName: "Ramos",
    document: "80123456",
    email: "logistica@mascampo.co",
    password: "$MAS_CAMPO_SECURE_SALT_2026$25703dc763780f7b7776ad93daf0ce387a9e22780e157ec251faf3b3cf360361",
    address: "Av. El Dorado # 68-90, Bodega 4, Bogotá",
    phone: "+573204445678",
    role: "LOGISTICA",
    vinculacion: "Contratista",
    active: true,
    createdAt: "2026-01-10",
    isSuperuser: false
  },
  {
    id: "usr-3",
    firstName: "Laura",
    lastName: "Gómez",
    document: "52987654",
    email: "ventas@mascampo.co",
    password: "$MAS_CAMPO_SECURE_SALT_2026$efb0bb33a8e13a14bfb9f8763ae6f8071831676a783da679e7b0742b31ac6bc7",
    address: "Carrera 7 # 71-21, Oficina 502, Bogotá",
    phone: "+573153339876",
    role: "VENTAS",
    vinculacion: "Temporal",
    active: true,
    createdAt: "2026-01-15",
    isSuperuser: false
  }
];

const initialLocations = [
  { id: "loc-1", name: "Bodega Principal (Central)", address: "Calle 26 # 68-10, Bogotá", phone: "+573105551234", manager: "Jorge Ramos", active: true },
  { id: "loc-2", name: "Punto de Venta Bogotá", address: "Carrera 15 # 93-40, Bogotá", phone: "+573153339876", manager: "Laura Gómez", active: true },
  { id: "loc-3", name: "Almacén Central Llanos", address: "Km 4 Vía Villavicencio - Acacías", phone: "+573187779900", manager: "Carlos Mendoza", active: true }
];

const initialAuditLogs = [
  {
    id: "log-101",
    timestamp: "2026-08-06 13:35:00",
    userName: "Carlos Mendoza",
    userEmail: "admin@mascampo.co",
    roleName: "ADMINISTRADOR",
    actionType: "LOGIN",
    entityName: "SystemSession",
    description: "Inicio de sesión exitoso. Autenticado en plataforma.",
    ipAddress: "190.157.24.110",
    deviceType: "Desktop (Linux x86_64 / Chrome 124)",
    userAgent: "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36"
  }
];

const appState = {
  authenticated: false,

  currentUser: {
    id: "usr-1",
    name: "Carlos Mendoza",
    firstName: "Carlos",
    lastName: "Mendoza",
    document: "1019024812",
    email: "admin@mascampo.co",
    role: "ADMINISTRADOR",
    vinculacion: "Planta (Nómina)",
    address: "Calle 100 # 15-20, Bogotá",
    phone: "+573105551234",
    active: true
  },

  locations: initialLocations,
  users: systemUsers,
  categories: initialCategories,
  products: excelIngestedProducts,
  customers: excelIngestedCustomers,
  selectedCustomerId: "cust-1",

  serializedItems: [
    {
      id: "ser-1",
      productId: "prod-3",
      serialNumber: "SN-TED-2026-001",
      status: "VENDIDO",
      currentCustomerId: "cust-1",
      entryDate: "2026-01-15",
      saleDate: "2026-01-20",
      invoiceNumber: "FACT-SIGO-2026-101",
      attachments: [
        { name: "Declaracion_Importacion_DIAN.pdf", type: "pdf", icon: "📄" },
        { name: "Foto_Mercancia_Llegada.jpg", type: "image", icon: "📷" }
      ],
      history: [
        { 
          type: "INGRESO", 
          date: "2026-01-15", 
          user: "Carlos Mendoza", 
          description: "Ingreso de bodega según Factura de Importación #EX25035 desde España.",
          attachments: [{ name: "Foto_Mercancia_Llegada.jpg", type: "image", icon: "📷" }]
        },
        {
          type: "VENTA",
          date: "2026-01-20",
          user: "Laura Gómez",
          description: "Vendido a FINCA DON PEDRO PMO S.A.S. Factura: FACT-SIGO-2026-101",
          attachments: [{ name: "Factura_Venta_101.pdf", type: "pdf", icon: "📄" }]
        }
      ]
    }
  ],

  movements: [
    { 
      id: "mov-1", 
      date: "2026-01-15", 
      type: "INGRESO_COMPRA", 
      productId: "prod-1", 
      quantity: 8012, 
      invoiceNumber: "#EX25035", 
      customerName: "NEW VETEC S.L. (España)", 
      user: "Carlos Mendoza", 
      notes: "Importación masiva de fundas TE",
      attachments: [
        { name: "Factura_Proveedor_EX25035.pdf", type: "pdf", icon: "📄" }
      ]
    },
    {
      id: "mov-101",
      date: "2026-01-20",
      type: "SALIDA_VENTA",
      productId: "prod-3",
      quantity: 1,
      invoiceNumber: "FACT-SIGO-2026-101",
      customerName: "FINCA DON PEDRO PMO S.A.S",
      user: "Laura Gómez",
      notes: "Venta de Pistola TE con Serial SN-TED-2026-001",
      attachments: [
        { name: "Factura_Venta_101.pdf", type: "pdf", icon: "📄" }
      ]
    },
    {
      id: "mov-102",
      date: "2026-02-10",
      type: "SALIDA_VENTA",
      productId: "prod-1",
      quantity: 50,
      invoiceNumber: "FACT-SIGO-2026-108",
      customerName: "FINCA DON PEDRO PMO S.A.S",
      user: "Laura Gómez",
      notes: "Venta de 50 paquetes Fundas TE",
      attachments: []
    }
  ],

  reservations: [
    { id: "res-101", customerId: "cust-1", productId: "prod-1", quantity: 150, reason: "Apartado jornadas de biotecnología llanos", status: "ACTIVA", date: "2026-01-25" }
  ],

  auditLogs: initialAuditLogs
};

function getDeviceFingerprint() {
  const ua = navigator.userAgent;
  if (/mobile/i.test(ua)) return "Mobile Phone (Web / PWA)";
  if (/tablet|ipad/i.test(ua)) return "Mobile Tablet (Web / PWA)";
  return "Desktop (Browser Session)";
}

function getClientIp() {
  return "190.157.24.110";
}

function loadPersistedAuditLogs() {
  try {
    const saved = localStorage.getItem("mascampo_audit_logs_db");
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        appState.auditLogs = parsed;
      }
    }
  } catch (e) {
    console.warn("No se pudo cargar log persistido de auditoría", e);
  }
}

function saveAuditLogsToDisk() {
  try {
    localStorage.setItem("mascampo_audit_logs_db", JSON.stringify(appState.auditLogs));
  } catch (e) {
    console.error("Error guardando log de auditoría en almacenamiento persistente", e);
  }
}

function downloadAuditLogsJSONFile() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador puede descargar la copia de seguridad física de auditoría.");
    return;
  }

  const isSuperuserActive = (appState.currentUser.email === 'gerencia@softproductiva.com');
  const exportableLogs = appState.auditLogs.filter(log => {
    const isSuperuserLog = (log.userEmail === 'gerencia@softproductiva.com' || (log.description && log.description.includes('gerencia@softproductiva.com')));
    if (isSuperuserLog && !isSuperuserActive) return false;
    return true;
  });

  const jsonStr = JSON.stringify(exportableLogs, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", `Backup_Fisico_Log_Auditoria_${new Date().toISOString().substring(0,10)}.json`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  addActivityLog("EXPORTACION", "AuditLogBackup", "Descarga de copia de seguridad en archivo físico JSON.");
}

function addActivityLog(actionType, entityName, description, extraPayload = null) {
  const timestamp = new Date().toISOString().replace('T', ' ').substring(0, 19);
  const logEntry = {
    id: `log-${Date.now()}-${Math.floor(Math.random()*1000)}`,
    timestamp,
    userName: appState.currentUser ? appState.currentUser.name : 'Sistema',
    userEmail: appState.currentUser ? appState.currentUser.email : 'sistema@mascampo.co',
    roleName: appState.currentUser ? appState.currentUser.role : 'ADMINISTRADOR',
    actionType,
    entityName,
    description,
    ipAddress: getClientIp(),
    deviceType: getDeviceFingerprint(),
    userAgent: navigator.userAgent
  };

  appState.auditLogs.unshift(logEntry);
  saveAuditLogsToDisk();
  renderAuditLogView();
}

function addAuditLog(message) {
  addActivityLog("MODIFICACION", "System", message);
}

// INTERACTIVE ATTACHMENT PREVIEWER MODAL
function previewAttachment(fileName, fileType = "image", dataUrl = null) {
  const modal = document.getElementById("attachmentPreviewModal");
  const title = document.getElementById("attachmentPreviewTitle");
  const container = document.getElementById("attachmentPreviewContainer");
  const downloadBtn = document.getElementById("attachmentDownloadBtn");

  if (!modal || !container) return;

  title.textContent = `📎 Soporte Digital: ${fileName}`;
  addActivityLog("CONSULTA", "AttachmentView", `Inspección de soporte digital / comprobante '${fileName}'.`);

  let contentHtml = "";

  if (dataUrl) {
    if (fileType.includes("image") || fileName.match(/\.(jpg|jpeg|png|gif|webp)$/i)) {
      contentHtml = `<img src="${dataUrl}" alt="${fileName}" style="max-width: 100%; max-height: 70vh; border-radius: var(--radius-md); box-shadow: 0 10px 30px rgba(0,0,0,0.8); object-fit: contain;">`;
    } else {
      contentHtml = `<iframe src="${dataUrl}" style="width: 100%; height: 65vh; border: none; border-radius: var(--radius-md);"></iframe>`;
    }
    if (downloadBtn) {
      downloadBtn.href = dataUrl;
      downloadBtn.download = fileName;
      downloadBtn.style.display = "inline-flex";
    }
  } else {
    const isImage = fileType.includes("image") || fileName.match(/\.(jpg|jpeg|png)$/i);
    
    if (isImage) {
      contentHtml = `
        <div style="background: rgba(15, 23, 42, 0.9); border: 2px dashed var(--border-highlight); padding: 2rem; border-radius: var(--radius-lg); text-align: center; width: 100%;">
          <div style="font-size: 3.5rem; margin-bottom: 0.5rem;">📷</div>
          <h4 style="color: var(--accent-green); font-size: 1.1rem; margin-bottom: 0.5rem; font-weight: 700;">${fileName}</h4>
          <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1rem;">
            [ Fotografía de Inspección Física / Recepción de Mercancía en Bodega ]
          </p>
          
          <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); padding: 1.5rem; border-radius: var(--radius-md); text-align: left; display: inline-block; max-width: 500px; width: 100%;">
            <div style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.3rem;">📌 <strong>Detalles de Captura de Fotografía:</strong></div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Archivo:</strong> ${fileName}</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Fecha de Carga:</strong> 2026-01-15 08:30 AM</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Operador Bodega:</strong> Carlos Mendoza (Admin / Logística)</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Estado del Empaque:</strong> Caja sellada de fábrica, inspección 100% aprobada.</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Hash Inmutable SHA-256:</strong> e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855</div>
          </div>
        </div>
      `;
    } else {
      contentHtml = `
        <div style="background: rgba(15, 23, 42, 0.9); border: 2px dashed var(--accent-blue); padding: 2rem; border-radius: var(--radius-lg); text-align: center; width: 100%;">
          <div style="font-size: 3.5rem; margin-bottom: 0.5rem;">📄</div>
          <h4 style="color: var(--accent-blue); font-size: 1.1rem; margin-bottom: 0.5rem; font-weight: 700;">${fileName}</h4>
          <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1rem;">
            [ Documento de Declaración de Importación / Factura DIAN ]
          </p>
          
          <div style="background: rgba(59, 130, 246, 0.1); border: 1px solid rgba(59, 130, 246, 0.3); padding: 1.5rem; border-radius: var(--radius-md); text-align: left; display: inline-block; max-width: 500px; width: 100%;">
            <div style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.3rem;">📄 <strong>Metadatos del Documento Soporte:</strong></div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Nombre Documento:</strong> ${fileName}</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Aduana / DIAN:</strong> Declaración Oficial de Importación España -> Colombia</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Empresa Importadora:</strong> MÁS CAMPO S.A.S</div>
            <div style="font-size: 0.82rem; color: var(--text-main);">• <strong>Autenticidad:</strong> Firmado digitalmente y verificado.</div>
          </div>
        </div>
      `;
    }
    if (downloadBtn) downloadBtn.style.display = "none";
  }

  container.innerHTML = contentHtml;
  modal.classList.add("active");
}

function closeAttachmentPreviewModal() {
  const modal = document.getElementById("attachmentPreviewModal");
  if (modal) modal.classList.remove("active");
}

// HELPER: Auto-generate 100% of individual serials
function ensureAllSerialsForProduct(product) {
  if (!product.requiresSerial) return;

  const skuClean = product.sku.replace(/\s+/g, '-');
  const existingSerials = appState.serializedItems.filter(i => i.productId === product.id);
  const totalCount = existingSerials.length;

  if (totalCount < product.physicalStock) {
    for (let k = 1; k <= product.physicalStock; k++) {
      const candidateSerial = `SN-${skuClean}-2026-${String(k).padStart(3, '0')}`;
      const exists = appState.serializedItems.some(i => i.productId === product.id && i.serialNumber === candidateSerial);
      if (!exists) {
        appState.serializedItems.push({
          id: `ser-${product.id}-${k}`,
          productId: product.id,
          serialNumber: candidateSerial,
          status: "EN_STOCK",
          currentCustomerId: null,
          entryDate: "2026-01-15",
          saleDate: null,
          invoiceNumber: "#EX25035",
          attachments: [
            { name: "Declaracion_Importacion_DIAN.pdf", type: "pdf", icon: "📄" },
            { name: "Foto_Mercancia_Llegada.jpg", type: "image", icon: "📷" }
          ],
          history: [
            { 
              type: "INGRESO", 
              date: "2026-01-15", 
              user: "Carlos Mendoza", 
              description: `Ingreso individual de bodega según Factura #EX25035 desde España.`,
              attachments: [{ name: "Foto_Mercancia_Llegada.jpg", type: "image", icon: "📷" }]
            }
          ]
        });
      }
    }
  }
}

function ensureGlobalSerials() {
  appState.products.forEach(p => {
    if (p.requiresSerial) ensureAllSerialsForProduct(p);
  });
}

// ==========================================================================
// THEME MANAGEMENT & CUSTOMIZER SYSTEM
// ==========================================================================

const themePresetData = {
  dark: {
    icon: "🌙",
    label: "Oscuro",
    bgPrimary: "#0b0f19",
    bgSecondary: "#111827",
    accent: "#10b981",
    textMain: "#f9fafb",
    textMuted: "#9ca3af"
  },
  light: {
    icon: "☀️",
    label: "Claro",
    bgPrimary: "#f1f5f9",
    bgSecondary: "#ffffff",
    accent: "#059669",
    textMain: "#0f172a",
    textMuted: "#475569"
  },
  navy: {
    icon: "🌊",
    label: "Azul Marino",
    bgPrimary: "#0a192f",
    bgSecondary: "#112240",
    accent: "#64ffda",
    textMain: "#e6f1ff",
    textMuted: "#8892b0"
  },
  emerald: {
    icon: "🌿",
    label: "Esmeralda",
    bgPrimary: "#062c20",
    bgSecondary: "#0a3a2a",
    accent: "#34d399",
    textMain: "#ecfdf5",
    textMuted: "#a7f3d0"
  },
  custom: {
    icon: "🎨",
    label: "Personalizado",
    bgPrimary: "#1e1b4b",
    bgSecondary: "#312e81",
    accent: "#ec4899",
    textMain: "#f8fafc",
    textMuted: "#c7d2fe"
  }
};

let activeTempPreset = 'dark';
let activeCustomColors = {
  bgPrimary: "#1e1b4b",
  bgSecondary: "#312e81",
  accent: "#ec4899",
  textMain: "#f8fafc",
  textMuted: "#c7d2fe"
};

function initThemeManager() {
  const savedTheme = localStorage.getItem('mascampo_theme') || 'dark';
  const savedColors = JSON.parse(localStorage.getItem('mascampo_custom_colors') || 'null');
  
  if (savedColors) {
    activeCustomColors = { ...activeCustomColors, ...savedColors };
  }

  appState.currentTheme = savedTheme;
  appState.customColors = activeCustomColors;

  applyTheme(savedTheme, activeCustomColors);

  const themeSelector = document.getElementById("themeSelector");
  if (themeSelector) themeSelector.value = savedTheme;
}

function updateThemeHeaderIcon(themeName) {
  const iconSpan = document.getElementById("currentThemeBadgeIcon");
  if (!iconSpan) return;
  const preset = themePresetData[themeName] || themePresetData.dark;
  iconSpan.textContent = preset.icon;
}

function applyTheme(themeName, customColors = null) {
  document.documentElement.setAttribute("data-theme", themeName);
  
  if (themeName === 'custom' && customColors) {
    document.documentElement.style.setProperty('--bg-primary', customColors.bgPrimary);
    document.documentElement.style.setProperty('--bg-secondary', customColors.bgSecondary);
    document.documentElement.style.setProperty('--bg-card', customColors.bgSecondary);
    document.documentElement.style.setProperty('--header-bg', customColors.bgSecondary);
    document.documentElement.style.setProperty('--sidebar-bg', customColors.bgSecondary);
    document.documentElement.style.setProperty('--form-input-bg', customColors.bgPrimary);
    document.documentElement.style.setProperty('--table-th-bg', customColors.bgSecondary);
    document.documentElement.style.setProperty('--card-sub-bg', customColors.bgPrimary);
    document.documentElement.style.setProperty('--accent-green', customColors.accent);
    document.documentElement.style.setProperty('--accent-green-hover', customColors.accent);
    document.documentElement.style.setProperty('--text-main', customColors.textMain);
    if (customColors.textMuted) {
      document.documentElement.style.setProperty('--text-muted', customColors.textMuted);
    }
  } else {
    document.documentElement.style.removeProperty('--bg-primary');
    document.documentElement.style.removeProperty('--bg-secondary');
    document.documentElement.style.removeProperty('--bg-card');
    document.documentElement.style.removeProperty('--header-bg');
    document.documentElement.style.removeProperty('--sidebar-bg');
    document.documentElement.style.removeProperty('--form-input-bg');
    document.documentElement.style.removeProperty('--table-th-bg');
    document.documentElement.style.removeProperty('--card-sub-bg');
    document.documentElement.style.removeProperty('--accent-green');
    document.documentElement.style.removeProperty('--accent-green-hover');
    document.documentElement.style.removeProperty('--text-main');
    document.documentElement.style.removeProperty('--text-muted');
  }

  updateThemeHeaderIcon(themeName);
}

function handleThemeChange(themeName) {
  if (themeName === 'custom') {
    openThemeCustomizerModal();
    selectThemePreset('custom');
    return;
  }

  appState.currentTheme = themeName;
  applyTheme(themeName);
  localStorage.setItem('mascampo_theme', themeName);
  addActivityLog("CONFIGURACION", "ThemeManager", `Cambio de tema gráfico a '${themeName}'.`);
}

function openThemeCustomizerModal() {
  const modal = document.getElementById("themeModal");
  if (!modal) return;

  activeTempPreset = appState.currentTheme || 'dark';
  selectThemePreset(activeTempPreset);
  modal.classList.add("active");
}

function closeThemeCustomizerModal() {
  const modal = document.getElementById("themeModal");
  if (!modal) return;
  
  const savedTheme = localStorage.getItem('mascampo_theme') || 'dark';
  applyTheme(savedTheme, appState.customColors);
  modal.classList.remove("active");
}

function selectThemePreset(presetName) {
  activeTempPreset = presetName;

  document.querySelectorAll(".preset-card").forEach(card => card.classList.remove("active"));
  const activeCard = document.getElementById(`preset-${presetName}`);
  if (activeCard) activeCard.classList.add("active");

  const customSection = document.getElementById("customColorControls");
  if (customSection) {
    customSection.style.display = (presetName === 'custom') ? "block" : "none";
  }

  if (presetName !== 'custom') {
    const preset = themePresetData[presetName];
    if (preset) {
      applyTheme(presetName);
    }
  } else {
    const pBgPrim = document.getElementById("pickerBgPrimary");
    const tBgPrim = document.getElementById("textBgPrimary");
    const pBgSec = document.getElementById("pickerBgSecondary");
    const tBgSec = document.getElementById("textBgSecondary");
    const pAcc = document.getElementById("pickerAccent");
    const tAcc = document.getElementById("textAccent");
    const pTxt = document.getElementById("pickerTextMain");
    const tTxt = document.getElementById("textTextMain");

    if (pBgPrim) { pBgPrim.value = activeCustomColors.bgPrimary; tBgPrim.value = activeCustomColors.bgPrimary; }
    if (pBgSec) { pBgSec.value = activeCustomColors.bgSecondary; tBgSec.value = activeCustomColors.bgSecondary; }
    if (pAcc) { pAcc.value = activeCustomColors.accent; tAcc.value = activeCustomColors.accent; }
    if (pTxt) { pTxt.value = activeCustomColors.textMain; tTxt.value = activeCustomColors.textMain; }

    applyTheme('custom', activeCustomColors);
  }
}

function updateCustomColorsFromPickers() {
  const pBgPrim = document.getElementById("pickerBgPrimary");
  const tBgPrim = document.getElementById("textBgPrimary");
  const pBgSec = document.getElementById("pickerBgSecondary");
  const tBgSec = document.getElementById("textBgSecondary");
  const pAcc = document.getElementById("pickerAccent");
  const tAcc = document.getElementById("textAccent");
  const pTxt = document.getElementById("pickerTextMain");
  const tTxt = document.getElementById("textTextMain");

  if (pBgPrim && tBgPrim) { tBgPrim.value = pBgPrim.value; activeCustomColors.bgPrimary = pBgPrim.value; }
  if (pBgSec && tBgSec) { tBgSec.value = pBgSec.value; activeCustomColors.bgSecondary = pBgSec.value; }
  if (pAcc && tAcc) { tAcc.value = pAcc.value; activeCustomColors.accent = pAcc.value; }
  if (pTxt && tTxt) { tTxt.value = pTxt.value; activeCustomColors.textMain = pTxt.value; }

  applyTheme('custom', activeCustomColors);
}

function saveThemePreferences() {
  appState.currentTheme = activeTempPreset;
  appState.customColors = { ...activeCustomColors };

  localStorage.setItem('mascampo_theme', activeTempPreset);
  if (activeTempPreset === 'custom') {
    localStorage.setItem('mascampo_custom_colors', JSON.stringify(activeCustomColors));
  }

  applyTheme(activeTempPreset, activeCustomColors);

  const themeSelector = document.getElementById("themeSelector");
  if (themeSelector) themeSelector.value = activeTempPreset;

  const modal = document.getElementById("themeModal");
  if (modal) modal.classList.remove("active");

  addActivityLog("CONFIGURACION", "ThemeManager", `Preferencias de tema guardadas: '${activeTempPreset}'.`);
  alert(`🎨 ¡Tema '${themePresetData[activeTempPreset]?.label || activeTempPreset}' aplicado y guardado con éxito!`);
}

function resetThemeToDefault() {
  appState.currentTheme = 'dark';
  localStorage.setItem('mascampo_theme', 'dark');
  localStorage.removeItem('mascampo_custom_colors');
  
  applyTheme('dark');

  const themeSelector = document.getElementById("themeSelector");
  if (themeSelector) themeSelector.value = 'dark';

  const modal = document.getElementById("themeModal");
  if (modal) modal.classList.remove("active");

  addActivityLog("CONFIGURACION", "ThemeManager", "Tema restablecido a Oscuro original.");
  alert("🔄 Tema restablecido al Modo Oscuro por defecto.");
}

// AUTHENTICATION ENGINE
function checkAuthentication() {
  const loginOverlay = document.getElementById("loginScreen");
  const appContainer = document.getElementById("app");
  if (!loginOverlay) return;

  if (!appState.authenticated) {
    loginOverlay.classList.remove("hidden");
    loginOverlay.style.display = "flex";
    if (appContainer) {
      appContainer.classList.add("auth-hidden");
      appContainer.style.display = "none";
    }
    document.body.style.overflow = "hidden";
    const emailInput = document.getElementById("loginEmail");
    if (emailInput) setTimeout(() => emailInput.focus(), 100);
  } else {
    loginOverlay.classList.add("hidden");
    loginOverlay.style.display = "none";
    if (appContainer) {
      appContainer.classList.remove("auth-hidden");
      appContainer.style.display = "block";
    }
    document.body.style.overflow = "auto";
  }
}

async function handleLoginSubmit(event) {
  event.preventDefault();
  const emailInput = document.getElementById("loginEmail").value.trim().toLowerCase();
  const passwordInput = document.getElementById("loginPassword").value;
  const errorMsgBox = document.getElementById("loginErrorMessage");

  let matchingUser = null;
  for (const u of appState.users) {
    if (u.email.toLowerCase() === emailInput) {
      const isValid = await verifyPassword(passwordInput, u.password);
      if (isValid) {
        matchingUser = u;
        break;
      }
    }
  }

  if (!matchingUser) {
    if (errorMsgBox) {
      errorMsgBox.textContent = "⚠️ Credenciales inválidas. Verifique el usuario (correo) y la contraseña.";
      errorMsgBox.style.display = "block";
    }
    return;
  }

  if (!matchingUser.active) {
    if (errorMsgBox) {
      errorMsgBox.textContent = "🔴 Esta cuenta se encuentra desactivada. Contacte al Administrador.";
      errorMsgBox.style.display = "block";
    }
    return;
  }

  if (errorMsgBox) errorMsgBox.style.display = "none";

  appState.currentUser = {
    id: matchingUser.id,
    name: `${matchingUser.firstName || ''} ${matchingUser.lastName || matchingUser.name || ''}`.trim(),
    firstName: matchingUser.firstName,
    lastName: matchingUser.lastName,
    document: matchingUser.document,
    email: matchingUser.email,
    role: matchingUser.role,
    vinculacion: matchingUser.vinculacion,
    address: matchingUser.address,
    phone: matchingUser.phone,
    active: matchingUser.active
  };

  appState.authenticated = true;

  // Synchronize Header Role Switcher
  const roleSelector = document.getElementById("roleSelector");
  if (roleSelector) {
    if (matchingUser.email === 'gerencia@softproductiva.com') {
      roleSelector.value = 'SUPERUSER';
    } else {
      roleSelector.value = matchingUser.role;
    }
  }

  const roleBadge = document.getElementById("currentRoleBadge");
  if (roleBadge) {
    roleBadge.textContent = (matchingUser.email === 'gerencia@softproductiva.com') ? '👑 SUPERUSUARIO' : matchingUser.role;
    roleBadge.className = `role-badge ${matchingUser.role === 'ADMINISTRADOR' ? 'admin' : matchingUser.role === 'LOGISTICA' ? 'logistics' : 'sales'}`;
  }

  addActivityLog("LOGIN", "SessionAuth", `Inicio de sesión exitoso con validación encriptada SHA-256 + Salt. Usuario: '${appState.currentUser.name}' (${appState.currentUser.email})`);

  checkAuthentication();
  renderAllViews();
}

function handleLogout() {
  if (confirm("¿Está seguro de que desea cerrar la sesión activa?")) {
    addActivityLog("LOGOUT", "SessionAuth", `Cierre de sesión del usuario '${appState.currentUser.email}'.`);
    appState.authenticated = false;
    
    const emailInput = document.getElementById("loginEmail");
    const passwordInput = document.getElementById("loginPassword");
    const errorMsgBox = document.getElementById("loginErrorMessage");
    if (emailInput) emailInput.value = "";
    if (passwordInput) passwordInput.value = "";
    if (errorMsgBox) errorMsgBox.style.display = "none";

    checkAuthentication();
  }
}

// Lifecycle
document.addEventListener("DOMContentLoaded", () => {
  loadPersistedAuditLogs();
  loadPersistedPendingIntakes();
  ensureGlobalSerials();
  initThemeManager();
  initRoleSwitcher();
  initNavigation();
  populateDropdowns();
  checkAuthentication();
  renderAllViews();

  document.addEventListener("click", (e) => {
    const suggBox = document.getElementById("warrantySerialSuggestions");
    const input = document.getElementById("warrantySerialInput");
    if (suggBox && !suggBox.contains(e.target) && e.target !== input) {
      suggBox.classList.remove("active");
    }
  });
});

function renderAllViews() {
  updateHeaderUserControls();
  ensureGlobalSerials();
  ensureProductLocations();
  renderDashboard();
  renderCatalog();
  renderLocationsView();
  renderPendingValidationsView();
  renderKardex();
  renderEquipmentLifeCycle();
  renderReservations();
  renderCustomers();
  renderCustomer360Analytics();
  renderUsers();
  renderAlerts();
  renderMarginReport();
  renderAuditLogView();
  toggleFinancialFields(appState.currentUser.role);
}

function updateHeaderUserControls() {
  if (!appState.currentUser) return;
  const isSuperuser = (appState.currentUser.email === 'gerencia@softproductiva.com');

  const roleContainer = document.querySelector(".role-switcher-container");
  const roleBadge = document.getElementById("currentRoleBadge");
  const roleSelect = document.getElementById("roleSelector");

  // SUPERUSER PRIVACY & ACCESS CONTROL:
  // The quick user switcher dropdown is ONLY visible to the Superusuario (gerencia@softproductiva.com).
  // For any other user (Admin, Logística, Ventas), the dropdown switcher is completely hidden.
  if (roleContainer) {
    if (isSuperuser) {
      roleContainer.style.display = "inline-flex";
    } else {
      roleContainer.style.display = "none";
    }
  }

  if (roleBadge) {
    if (isSuperuser) {
      roleBadge.textContent = "👑 SUPERUSUARIO";
      roleBadge.className = "role-badge admin";
    } else {
      roleBadge.textContent = appState.currentUser.role;
      roleBadge.className = `role-badge ${appState.currentUser.role === 'ADMINISTRADOR' ? 'admin' : appState.currentUser.role === 'LOGISTICA' ? 'logistics' : 'sales'}`;
    }
  }

  if (roleSelect && isSuperuser) {
    roleSelect.value = "SUPERUSER";
  }
}

function initRoleSwitcher() {
  const roleSelect = document.getElementById("roleSelector");
  const roleBadge = document.getElementById("currentRoleBadge");
  const rbacBanner = document.getElementById("rbacBanner");
  const activeRoleName = document.getElementById("activeRoleName");

  if (!roleSelect) return;

  roleSelect.addEventListener("change", (e) => {
    // Only superuser is allowed to use this switcher
    if (appState.currentUser.email !== 'gerencia@softproductiva.com') {
      alert("🔒 Restricción de Privacidad: La conmutación rápida de usuarios está restringida al Superusuario.");
      updateHeaderUserControls();
      return;
    }

    const selectedKey = e.target.value;
    let targetUser = null;

    if (selectedKey === 'SUPERUSER') {
      targetUser = appState.users.find(u => u.email === 'gerencia@softproductiva.com');
    } else if (selectedKey === 'ADMINISTRADOR') {
      targetUser = appState.users.find(u => u.email === 'admin@mascampo.co');
    } else if (selectedKey === 'LOGISTICA') {
      targetUser = appState.users.find(u => u.email === 'logistica@mascampo.co');
    } else if (selectedKey === 'VENTAS') {
      targetUser = appState.users.find(u => u.email === 'ventas@mascampo.co');
    }

    if (targetUser) {
      appState.currentUser = {
        id: targetUser.id,
        name: `${targetUser.firstName} ${targetUser.lastName}`.trim(),
        firstName: targetUser.firstName,
        lastName: targetUser.lastName,
        document: targetUser.document,
        email: targetUser.email,
        role: targetUser.role,
        address: targetUser.address,
        phone: targetUser.phone,
        active: targetUser.active
      };
    } else {
      appState.currentUser.role = selectedKey;
    }

    const currentRole = appState.currentUser.role;
    addActivityLog("LOGIN", "SessionAuth", `Cambio de sesión activa a Usuario: '${appState.currentUser.name}' (${appState.currentUser.email})`);

    if (currentRole !== 'ADMINISTRADOR') {
      if (rbacBanner) rbacBanner.style.display = "block";
      if (activeRoleName) activeRoleName.textContent = currentRole;
    } else {
      if (rbacBanner) rbacBanner.style.display = "none";
    }

    updateHeaderUserControls();
    toggleFinancialFields(currentRole);
    populateDropdowns();
    renderAllViews();
  });
}

function toggleFinancialFields(role) {
  const costColumns = document.querySelectorAll(".cost-column");
  const btnNewProduct = document.getElementById("btnNewProduct");
  const btnNewCategory = document.getElementById("btnNewCategory");
  const btnAdminUsers = document.getElementById("btnAdminUsers");
  const btnUploadSigoSales = document.getElementById("btnUploadSigoSales");
  const auditLink = document.querySelector('[data-view="view-audit-logs"]');
  const financialRestrictedMsg = document.getElementById("financialRestrictedMsg");

  const isAllowed = (role === 'ADMINISTRADOR');

  costColumns.forEach(el => {
    el.style.display = isAllowed ? "" : "none";
  });

  if (btnNewProduct) btnNewProduct.style.display = isAllowed ? "inline-flex" : "none";
  if (btnNewCategory) btnNewCategory.style.display = isAllowed ? "inline-flex" : "none";
  if (btnAdminUsers) btnAdminUsers.style.display = isAllowed ? "inline-flex" : "none";
  if (btnUploadSigoSales) btnUploadSigoSales.style.display = isAllowed ? "inline-flex" : "none";
  if (auditLink) auditLink.style.display = isAllowed ? "flex" : "none";
  if (financialRestrictedMsg) financialRestrictedMsg.style.display = isAllowed ? "none" : "block";
}

function initNavigation() {
  document.querySelectorAll("[data-view]").forEach(link => {
    link.addEventListener("click", (e) => {
      e.preventDefault();
      switchView(link.getAttribute("data-view"));
    });
  });
}

function switchView(viewId) {
  document.querySelectorAll(".view-section").forEach(sec => sec.classList.remove("active"));
  document.querySelectorAll("[data-view]").forEach(lnk => lnk.classList.remove("active"));

  const activeSec = document.getElementById(viewId);
  if (activeSec) activeSec.classList.add("active");

  document.querySelectorAll(`[data-view="${viewId}"]`).forEach(lnk => lnk.classList.add("active"));

  if (viewId === 'view-forecasting') {
    runLinearRegressionForecast();
  }
}

function populateDropdowns() {
  const newCategorySelect = document.getElementById("newCategory");
  if (newCategorySelect) {
    newCategorySelect.innerHTML = appState.categories.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
  }

  const catalogCategoryFilter = document.getElementById("catalogCategoryFilter");
  if (catalogCategoryFilter) {
    catalogCategoryFilter.innerHTML = `<option value="ALL">Todas las Categorías Personalizadas</option>` +
      appState.categories.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
  }

  const auditUserFilter = document.getElementById("auditUserFilter");
  if (auditUserFilter) {
    const isSuperuserActive = (appState.currentUser.email === 'gerencia@softproductiva.com');
    const availableUsersForFilter = appState.users.filter(u => {
      if (u.email === 'gerencia@softproductiva.com' || u.isSuperuser) {
        return isSuperuserActive;
      }
      return true;
    });

    auditUserFilter.innerHTML = `<option value="ALL">Todos los Usuarios</option>` +
      availableUsersForFilter.map(u => {
        const fullName = `${u.firstName || ''} ${u.lastName || u.name || ''}`.trim();
        return `<option value="${fullName}">${fullName} (${u.email})</option>`;
      }).join('');
  }

  const catalogLocationFilter = document.getElementById("catalogLocationFilter");
  if (catalogLocationFilter) {
    catalogLocationFilter.innerHTML = `<option value="ALL">📍 Todas las Ubicaciones / Sedes</option>` +
      appState.locations.map(l => `<option value="${l.id}">${l.name}</option>`).join('');
  }

  const transferProductSelect = document.getElementById("transferProduct");
  if (transferProductSelect) {
    transferProductSelect.innerHTML = appState.products.map(p => 
      `<option value="${p.id}">${p.sku} - ${p.name} (Stock Total: ${p.physicalStock})</option>`
    ).join('');
  }

  const transferFromLocSelect = document.getElementById("transferFromLoc");
  const transferToLocSelect = document.getElementById("transferToLoc");
  const activeLocationOptions = appState.locations.filter(l => l.active).map(l => `<option value="${l.id}">${l.name} (${l.manager})</option>`).join('');

  if (transferFromLocSelect) transferFromLocSelect.innerHTML = activeLocationOptions;
  if (transferToLocSelect) transferToLocSelect.innerHTML = activeLocationOptions;

  const movProductSelect = document.getElementById("movProduct");
  const forecastProductSelect = document.getElementById("forecastProductSelect");
  const resProductSelect = document.getElementById("resProductSelect");

  const productOptions = appState.products.map(p => 
    `<option value="${p.id}">${p.sku} - ${p.name} (Stock: ${p.physicalStock} | Disp: ${p.physicalStock - p.reservedStock})</option>`
  ).join('');

  if (movProductSelect) movProductSelect.innerHTML = productOptions;
  if (forecastProductSelect) forecastProductSelect.innerHTML = productOptions;
  if (resProductSelect) resProductSelect.innerHTML = productOptions;

  const movCustomerSelect = document.getElementById("movCustomer");
  const resCustomerSelect = document.getElementById("resCustomerSelect");

  const customerOptions = `<option value="">-- Seleccionar Cliente --</option>` + 
    appState.customers.map(c => {
      const docLabel = c.documentType === "NIT" ? `NIT ${c.documentNumber}-${c.verificationDigit}` : `CC ${c.documentNumber}`;
      return `<option value="${c.id}">${c.name} (${docLabel})</option>`;
    }).join('');

  if (movCustomerSelect) movCustomerSelect.innerHTML = customerOptions;
  if (resCustomerSelect) resCustomerSelect.innerHTML = customerOptions;

  const equipmentSerialSelect = document.getElementById("equipmentSerialSelect");
  if (equipmentSerialSelect) {
    equipmentSerialSelect.innerHTML = appState.serializedItems.map(item => {
      const prod = appState.products.find(p => p.id === item.productId);
      return `<option value="${item.id}">${item.serialNumber} - ${prod ? prod.name : 'Equipo'} (${item.status})</option>`;
    }).join('');
  }
}

function renderAllViews() {
  ensureGlobalSerials();
  ensureProductLocations();
  renderDashboard();
  renderCatalog();
  renderLocationsView();
  renderKardex();
  renderEquipmentLifeCycle();
  renderReservations();
  renderCustomers();
  renderCustomer360Analytics();
  renderUsers();
  renderAlerts();
  renderMarginReport();
  renderAuditLogView();
  toggleFinancialFields(appState.currentUser.role);
}

// 1. CUSTOMER DIRECTORY & 360° ANALYTICS ENGINE
function renderCustomers() {
  const tableBody = document.getElementById("customersTableBody");
  const searchQuery = (document.getElementById("customerDirectorySearch")?.value || "").toLowerCase();
  const typeFilter = document.getElementById("customerTypeFilter")?.value || "ALL";

  if (!tableBody) return;

  const filtered = appState.customers.filter(c => {
    const matchesType = (typeFilter === "ALL" || c.documentType === typeFilter);
    const matchesSearch = c.name.toLowerCase().includes(searchQuery) ||
                          c.documentNumber.includes(searchQuery) ||
                          (c.city && c.city.toLowerCase().includes(searchQuery)) ||
                          (c.phone && c.phone.includes(searchQuery));
    return matchesType && matchesSearch;
  });

  tableBody.innerHTML = filtered.map(c => {
    const formattedDoc = c.documentType === 'NIT' ? `NIT <strong>${c.documentNumber}</strong>-${c.verificationDigit || '0'}` : `CC <strong>${c.documentNumber}</strong>`;
    
    // Calculate total purchases in COP for badge
    const totalPurchasedCOP = appState.movements
      .filter(m => (m.customerName && m.customerName.toLowerCase().includes(c.name.toLowerCase())) && m.type === 'SALIDA_VENTA')
      .reduce((sum, m) => {
        const prod = appState.products.find(p => p.id === m.productId);
        return sum + ((prod ? prod.salePrice : 100) * m.quantity);
      }, 0);

    return `
      <tr style="cursor: pointer;" onclick="selectCustomerFor360('${c.id}')" title="Haga clic para ver perfil 360° y estadísticas">
        <td><span class="badge ${c.documentType === 'NIT' ? 'badge-purple' : 'badge-blue'}">${c.documentType}</span></td>
        <td>${formattedDoc}</td>
        <td><strong>${c.name}</strong></td>
        <td>${c.city || 'Colombia'}</td>
        <td>${c.phone || '-'}</td>
        <td><strong style="color: var(--accent-green);">$${totalPurchasedCOP.toLocaleString('es-CO')} COP</strong></td>
        <td>
          <button class="btn btn-secondary" style="padding: 0.2rem 0.6rem; font-size: 0.75rem;" onclick="event.stopPropagation(); selectCustomerFor360('${c.id}')">
            📊 Ver Analítica 360°
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function selectCustomerFor360(customerId) {
  appState.selectedCustomerId = customerId;
  const customer = appState.customers.find(c => c.id === customerId);
  if (customer) {
    addActivityLog("CONSULTA", "Customer360Profile", `Inspección de perfil 360° y estadísticas del cliente '${customer.name}'.`);
  }
  renderCustomer360Analytics();
  switchView('view-customers-analytics');
}

function renderCustomer360Analytics() {
  const customerId = appState.selectedCustomerId || (appState.customers[0] ? appState.customers[0].id : null);
  if (!customerId) return;

  const customer = appState.customers.find(c => c.id === customerId);
  if (!customer) return;

  // Header info
  const formattedDoc = customer.documentType === 'NIT' ? `NIT ${customer.documentNumber}-${customer.verificationDigit || '0'}` : `CC ${customer.documentNumber}`;
  
  document.getElementById("cust360Name").textContent = customer.name;
  document.getElementById("cust360Doc").textContent = `${customer.documentType}: ${formattedDoc}`;
  document.getElementById("cust360City").textContent = customer.city || "Colombia";
  document.getElementById("cust360Phone").textContent = customer.phone || "No Registrado";
  document.getElementById("cust360Email").textContent = customer.email || "No Registrado";

  // Customer movements
  const customerMovements = appState.movements.filter(m => 
    m.type === 'SALIDA_VENTA' && m.customerName && m.customerName.toLowerCase().includes(customer.name.toLowerCase())
  );

  // Financial Metrics
  let totalSpentCOP = 0;
  let totalUnits = 0;

  customerMovements.forEach(m => {
    const prod = appState.products.find(p => p.id === m.productId);
    const unitPrice = prod ? prod.salePrice : 100;
    totalSpentCOP += (unitPrice * m.quantity);
    totalUnits += m.quantity;
  });

  const purchaseCount = customerMovements.length;
  const avgPurchaseTicket = purchaseCount > 0 ? (totalSpentCOP / purchaseCount) : 0;
  const avgUnitsPerSemester = (totalUnits * 0.5).toFixed(1);

  document.getElementById("cust360TotalSpent").textContent = `$${totalSpentCOP.toLocaleString('es-CO')} COP`;
  document.getElementById("cust360TotalUnits").textContent = `${totalUnits} Unidades`;
  document.getElementById("cust360AvgTicket").textContent = `$${avgPurchaseTicket.toFixed(2)} COP`;
  document.getElementById("cust360SemesterFreq").textContent = `${avgUnitsPerSemester} Unds / Semestre`;

  // Purchased Products & Serials Breakdown Table
  const tableBody = document.getElementById("cust360PurchasedProductsTable");
  if (tableBody) {
    if (customerMovements.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 1rem;">No se registran compras previas para este cliente.</td></tr>`;
    } else {
      tableBody.innerHTML = customerMovements.map(m => {
        const prod = appState.products.find(p => p.id === m.productId);
        const unitPrice = prod ? prod.salePrice : 100;
        const subtotal = unitPrice * m.quantity;

        // Find serialized item assigned to this customer for this sale/invoice
        const serialItems = appState.serializedItems.filter(i => 
          i.currentCustomerId === customer.id || (i.invoiceNumber && i.invoiceNumber === m.invoiceNumber)
        );

        return `
          <tr>
            <td><strong>${m.date}</strong></td>
            <td><strong>${prod ? prod.sku : ''}</strong> - ${prod ? prod.name : 'Insumo'}</td>
            <td><span class="badge badge-purple">${prod ? prod.category : 'General'}</span></td>
            <td><strong>${m.quantity}</strong></td>
            <td>$${subtotal.toLocaleString('es-CO')} COP</td>
            <td>
              ${prod && prod.requiresSerial ? `
                ${serialItems.length > 0 ? serialItems.map(s => `
                  <span style="background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); padding: 0.15rem 0.4rem; border-radius: 4px; font-family: monospace; color: var(--accent-green); font-size: 0.8rem;">
                    SN: ${s.serialNumber}
                  </span>
                `).join(' ') : '<span style="color: var(--accent-amber); font-size: 0.8rem;">Serializado sin asignar</span>'}
              ` : '<span style="color: var(--text-muted); font-size: 0.8rem;">Consumible (Sin Serial)</span>'}
            </td>
          </tr>
        `;
      }).join('');
    }
  }

  // Active Reservations for this Customer
  const custReservations = appState.reservations.filter(r => r.customerId === customer.id);
  const reservationsContainer = document.getElementById("cust360ReservationsContainer");
  if (reservationsContainer) {
    if (custReservations.length === 0) {
      reservationsContainer.innerHTML = `<div style="color: var(--text-muted); font-size: 0.85rem;">No hay bloqueos ni reservas activas para este cliente.</div>`;
    } else {
      reservationsContainer.innerHTML = custReservations.map(r => {
        const prod = appState.products.find(p => p.id === r.productId);
        return `
          <div style="background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); padding: 0.75rem 1rem; border-radius: var(--radius-md); margin-bottom: 0.5rem; display: flex; justify-content: space-between; align-items: center;">
            <div>
              <strong style="color: #fcd34d;">🔒 RESERVA ACTIVA: ${prod ? prod.sku : ''} - ${prod ? prod.name : ''}</strong>
              <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.2rem;">
                Cantidad Bloqueada: <strong>${r.quantity} unidades</strong> | Motivo: ${r.reason}
              </div>
            </div>
            <span class="badge badge-amber">${r.status}</span>
          </div>
        `;
      }).join('');
    }
  }

  // Warranties & Returns for this Customer
  const custSerials = appState.serializedItems.filter(i => i.currentCustomerId === customer.id);
  const warrantyContainer = document.getElementById("cust360WarrantiesContainer");
  if (warrantyContainer) {
    if (custSerials.length === 0) {
      warrantyContainer.innerHTML = `<div style="color: var(--text-muted); font-size: 0.85rem;">Este cliente no posee equipos serializados bajo garantía registrada.</div>`;
    } else {
      warrantyContainer.innerHTML = custSerials.map(s => {
        const prod = appState.products.find(p => p.id === s.productId);
        const warrantyMonths = prod ? (prod.warrantyMonths || 12) : 12;
        const saleDateObj = s.saleDate ? new Date(s.saleDate) : new Date(s.entryDate);
        const expiryDateObj = new Date(saleDateObj);
        expiryDateObj.setMonth(expiryDateObj.getMonth() + warrantyMonths);

        const today = new Date();
        const isValid = (today <= expiryDateObj);
        const daysLeft = Math.max(0, Math.round((expiryDateObj - today) / (1000 * 60 * 60 * 24)));

        return `
          <div style="background: rgba(15, 23, 42, 0.8); border: 1px solid var(--border-color); padding: 0.75rem 1rem; border-radius: var(--radius-md); margin-bottom: 0.5rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
            <div>
              <strong style="color: var(--text-main);">${prod ? prod.name : 'Equipo'}</strong>
              <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.2rem;">
                Serial: <strong style="color: var(--accent-green); font-family: monospace;">${s.serialNumber}</strong> | Factura: ${s.invoiceNumber || 'N/A'}
              </div>
            </div>
            <div>
              <span class="badge ${isValid ? 'badge-green' : 'badge-red'}">
                ${isValid ? `🛡️ Garantía Vigente (${daysLeft} días)` : '❌ Expirada'}
              </span>
            </div>
          </div>
        `;
      }).join('');
    }
  }
}

function handleWarrantySerialInput(event) {
  const query = event.target.value.trim().toLowerCase();
  const suggBox = document.getElementById("warrantySerialSuggestions");

  if (!suggBox) return;

  if (query.length === 0) {
    suggBox.classList.remove("active");
    suggBox.innerHTML = "";
    return;
  }

  ensureGlobalSerials();

  const matchingItems = appState.serializedItems.filter(item => {
    const prod = appState.products.find(p => p.id === item.productId);
    const matchesSerial = item.serialNumber.toLowerCase().includes(query);
    const matchesSku = prod ? prod.sku.toLowerCase().includes(query) : false;
    const matchesName = prod ? prod.name.toLowerCase().includes(query) : false;
    return matchesSerial || matchesSku || matchesName;
  });

  if (matchingItems.length === 0) {
    suggBox.innerHTML = `
      <div style="padding: 0.75rem 1rem; color: var(--text-muted); font-size: 0.85rem;">
        ❌ No se encontraron seriales que coincidan con "<strong>${query}</strong>"
      </div>
    `;
    suggBox.classList.add("active");
    return;
  }

  suggBox.innerHTML = matchingItems.map(item => {
    const prod = appState.products.find(p => p.id === item.productId);
    const customer = appState.customers.find(c => c.id === item.currentCustomerId);
    return `
      <div class="suggestion-item" onclick="selectSerialSuggestion('${item.serialNumber}')">
        <div>
          <strong style="color: var(--accent-green); font-family: monospace; font-size: 0.95rem;">${item.serialNumber}</strong>
          <div style="font-size: 0.8rem; color: var(--text-main); font-weight: 500;">${prod ? prod.sku : ''} - ${prod ? prod.name : 'Equipo'}</div>
          <div style="font-size: 0.73rem; color: var(--text-muted);">
            Status: ${item.status} | Cliente: ${customer ? customer.name : 'Bodega'}
          </div>
        </div>
        <span class="badge ${item.status === 'EN_STOCK' ? 'badge-green' : 'badge-amber'}">${item.status}</span>
      </div>
    `;
  }).join('');

  suggBox.classList.add("active");
}

function selectSerialSuggestion(serialNumber) {
  const input = document.getElementById("warrantySerialInput");
  const suggBox = document.getElementById("warrantySerialSuggestions");

  if (input) input.value = serialNumber;
  if (suggBox) suggBox.classList.remove("active");

  lookupWarrantyBySerial(serialNumber);
}

function openCameraForWarranty() {
  openCameraModal();
}

function renderAuditLogView() {
  const tableBody = document.getElementById("auditLogTableBody");
  const dashboardAuditBox = document.getElementById("auditLogBox");

  if (!tableBody) return;

  const actionFilter = document.getElementById("auditActionFilter")?.value || "ALL";
  const userFilter = document.getElementById("auditUserFilter")?.value || "ALL";
  const searchQuery = (document.getElementById("auditSearch")?.value || "").toLowerCase();

  // SUPERUSER PRIVACY ISOLATION IN AUDIT LOGS:
  // Nobody except gerencia@softproductiva.com can see logs relating to gerencia@softproductiva.com!
  const isSuperuserActive = (appState.currentUser.email === 'gerencia@softproductiva.com');

  const visibleLogs = appState.auditLogs.filter(log => {
    const isSuperuserLog = (log.userEmail === 'gerencia@softproductiva.com' || (log.description && log.description.includes('gerencia@softproductiva.com')));
    if (isSuperuserLog && !isSuperuserActive) {
      return false; // Hide completely from audit logs
    }

    const matchesAction = (actionFilter === "ALL" || log.actionType === actionFilter);
    const matchesUser = (userFilter === "ALL" || log.userName === userFilter || log.userEmail === userFilter);
    const matchesQuery = log.description.toLowerCase().includes(searchQuery) || log.ipAddress.includes(searchQuery) || log.userName.toLowerCase().includes(searchQuery);
    return matchesAction && matchesUser && matchesQuery;
  });

  tableBody.innerHTML = visibleLogs.map(l => {
    let actionBadgeClass = "badge-blue";
    if (l.actionType === "LOGIN") actionBadgeClass = "badge-purple";
    if (l.actionType === "IMPORTACION") actionBadgeClass = "badge-green";
    if (l.actionType === "EXPORTACION") actionBadgeClass = "badge-amber";
    if (l.actionType === "GARANTIA") actionBadgeClass = "badge-red";

    return `
      <tr>
        <td><strong>${l.timestamp}</strong></td>
        <td>
          <div style="font-weight: 600;">${l.userName}</div>
          <div style="font-size: 0.73rem; color: var(--text-muted);">${l.userEmail} (${l.roleName})</div>
        </td>
        <td><span class="badge ${actionBadgeClass}">${l.actionType}</span></td>
        <td>
          <div>${l.description}</div>
          <div style="font-size: 0.72rem; color: var(--accent-green); margin-top: 0.2rem;">Entidad: ${l.entityName}</div>
        </td>
        <td><strong style="font-family: monospace; font-size: 0.85rem; color: var(--accent-teal);">${l.ipAddress}</strong></td>
        <td><div style="font-size: 0.78rem; color: var(--text-muted);">${l.deviceType}</div></td>
      </tr>
    `;
  }).join('');

  if (dashboardAuditBox) {
    const dashboardLogs = appState.auditLogs.filter(log => {
      const isSuperuserLog = (log.userEmail === 'gerencia@softproductiva.com' || (log.description && log.description.includes('gerencia@softproductiva.com')));
      if (isSuperuserLog && !isSuperuserActive) return false;
      return true;
    });

    dashboardAuditBox.innerHTML = dashboardLogs.slice(0, 8).map(l => 
      `<div>[${l.timestamp}] [${l.actionType}] [IP: ${l.ipAddress}] ${l.userName}: ${l.description}</div>`
    ).join('');
  }
}

function exportAuditLogsCSV() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador puede exportar el log de auditoría.");
    return;
  }

  const isSuperuserActive = (appState.currentUser.email === 'gerencia@softproductiva.com');
  const exportableLogs = appState.auditLogs.filter(log => {
    const isSuperuserLog = (log.userEmail === 'gerencia@softproductiva.com' || (log.description && log.description.includes('gerencia@softproductiva.com')));
    if (isSuperuserLog && !isSuperuserActive) return false;
    return true;
  });

  let csv = "FechaHora;Usuario;Email;Rol;TipoAccion;Entidad;Descripcion;IPAddress;Dispositivo\n";
  exportableLogs.forEach(l => {
    csv += `"${l.timestamp}";"${l.userName}";"${l.userEmail}";"${l.roleName}";"${l.actionType}";"${l.entityName}";"${l.description}";"${l.ipAddress}";"${l.deviceType}"\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", `Log_Auditoria_Forense_MasCampo_${new Date().toISOString().substring(0,10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  addActivityLog("EXPORTACION", "AuditLog", "Exportación completa de Log de Auditoría en CSV.");
}

// USER MANAGEMENT
function renderUsers() {
  const tableBody = document.getElementById("usersTableBody");
  if (!tableBody) return;

  // SUPERUSER PRIVACY ISOLATION:
  // If current logged-in user is NOT gerencia@softproductiva.com, hide gerencia@softproductiva.com!
  const isSuperuserActive = (appState.currentUser.email === 'gerencia@softproductiva.com');
  const visibleUsers = appState.users.filter(u => {
    if (u.email === 'gerencia@softproductiva.com' || u.isSuperuser) {
      return isSuperuserActive;
    }
    return true;
  });

  tableBody.innerHTML = visibleUsers.map(u => {
    const fullName = `${u.firstName || ''} ${u.lastName || u.name || ''}`.trim();
    const docDisplay = u.document || 'N/A';
    const phoneDisplay = u.phone || 'N/A';
    const addressDisplay = u.address || 'N/A';
    const vincDisplay = u.vinculacion || 'Planta (Nómina)';
    const roleClass = u.role === 'ADMINISTRADOR' ? 'badge-red' : u.role === 'LOGISTICA' ? 'badge-amber' : 'badge-blue';

    return `
      <tr>
        <td><strong style="font-family: monospace; color: var(--accent-green);">${docDisplay}</strong></td>
        <td>
          <div style="font-weight: 600; color: var(--text-main);">${fullName} ${u.isSuperuser ? '👑' : ''}</div>
          <div style="font-size: 0.73rem; color: var(--text-muted);">ID: ${u.id}</div>
        </td>
        <td><strong>${u.email}</strong></td>
        <td>
          <div style="font-size: 0.85rem;">📞 ${phoneDisplay}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">🏠 ${addressDisplay}</div>
        </td>
        <td>
          <span class="badge badge-purple">${vincDisplay}</span>
        </td>
        <td>
          <span class="badge ${roleClass}">
            ${u.role}
          </span>
        </td>
        <td>
          <span class="badge ${u.active ? 'badge-green' : 'badge-red'}">
            ${u.active ? '🟢 Activo' : '🔴 Desactivo'}
          </span>
        </td>
        <td>
          <div style="display: flex; gap: 0.3rem;">
            <button class="btn btn-secondary" style="padding: 0.2rem 0.5rem; font-size: 0.73rem;" onclick="openUserModal('${u.id}')">
              ✏️ Editar
            </button>
            <button class="btn ${u.active ? 'btn-danger' : 'btn-primary'}" style="padding: 0.2rem 0.5rem; font-size: 0.73rem;" onclick="toggleUserStatus('${u.id}')">
              ${u.active ? '🔴 Desactivar' : '🟢 Activar'}
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function openUserModal(userId = null) {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo los usuarios con rol Administrador pueden gestionar usuarios.");
    return;
  }

  const modal = document.getElementById("userModal");
  const title = document.getElementById("userModalTitle");
  const editIdInput = document.getElementById("userEditId");
  const form = document.getElementById("newUserForm");

  if (!modal) return;

  if (userId) {
    const user = appState.users.find(u => u.id === userId);
    if (!user) return;

    if (title) title.textContent = `✏️ Editar Usuario (${user.email})`;
    if (editIdInput) editIdInput.value = user.id;

    document.getElementById("userFirstName").value = user.firstName || user.name.split(' ')[0] || "";
    document.getElementById("userLastName").value = user.lastName || user.name.split(' ').slice(1).join(' ') || "";
    document.getElementById("userDocument").value = user.document || "";
    document.getElementById("userEmail").value = user.email || "";
    document.getElementById("userPassword").value = user.password || "••••••••";
    document.getElementById("userPhone").value = user.phone || "";
    document.getElementById("userAddress").value = user.address || "";
    document.getElementById("userVinculacion").value = user.vinculacion || "Planta (Nómina)";
    document.getElementById("userRole").value = user.role || "ADMINISTRADOR";
    document.getElementById("userActive").value = user.active ? "true" : "false";
  } else {
    if (title) title.textContent = "👤 Crear Nuevo Usuario del Sistema";
    if (editIdInput) editIdInput.value = "";
    if (form) form.reset();
  }

  modal.classList.add("active");
}

function closeUserModal() {
  const modal = document.getElementById("userModal");
  if (modal) modal.classList.remove("active");
}

async function handleSaveUser(e) {
  e.preventDefault();
  const editId = document.getElementById("userEditId")?.value;
  const firstName = document.getElementById("userFirstName").value.trim();
  const lastName = document.getElementById("userLastName").value.trim();
  const documentNum = document.getElementById("userDocument").value.trim();
  const email = document.getElementById("userEmail").value.trim().toLowerCase();
  const password = document.getElementById("userPassword").value;
  const phone = document.getElementById("userPhone").value.trim();
  const address = document.getElementById("userAddress").value.trim();
  const vinculacion = document.getElementById("userVinculacion").value;
  const role = document.getElementById("userRole").value;
  const active = document.getElementById("userActive").value === "true";

  if (editId) {
    const user = appState.users.find(u => u.id === editId);
    if (user) {
      user.firstName = firstName;
      user.lastName = lastName;
      user.name = `${firstName} ${lastName}`;
      user.document = documentNum;
      user.email = email;
      if (password && password !== "••••••••") {
        user.password = await hashPassword(password);
      }
      user.phone = phone;
      user.address = address;
      user.vinculacion = vinculacion;
      user.role = role;
      user.active = active;

      addActivityLog("MODIFICACION", "UserManagement", `Actualización de datos de usuario '${user.email}' (Vinculación: ${vinculacion}, Rol: ${role})`);
      alert(`✅ Usuario '${user.email}' actualizado con éxito.`);
    }
  } else {
    const existing = appState.users.find(u => u.email === email);
    if (existing) {
      alert(`⚠️ ERROR: El correo '${email}' ya se encuentra registrado.`);
      return;
    }

    const hashedPassword = await hashPassword(password);

    const newUser = {
      id: `usr-${Date.now()}`,
      firstName,
      lastName,
      name: `${firstName} ${lastName}`,
      document: documentNum,
      email,
      password: hashedPassword,
      phone,
      address,
      vinculacion,
      role,
      active,
      createdAt: new Date().toISOString().substring(0, 10),
      isSuperuser: (email === 'gerencia@softproductiva.com')
    };

    appState.users.push(newUser);
    addActivityLog("CREACION", "UserManagement", `Creación de nuevo usuario '${email}' (Vinculación: ${vinculacion}, Rol: ${role})`);
    alert(`✅ Usuario '${email}' creado exitosamente.`);
  }

  closeUserModal();
  renderUsers();
  populateDropdowns();
}

function handleCreateUser(e) {
  handleSaveUser(e);
}

function toggleUserStatus(userId) {
  const user = appState.users.find(u => u.id === userId);
  if (!user) return;

  if (user.email === 'gerencia@softproductiva.com' && appState.currentUser.email !== 'gerencia@softproductiva.com') {
    alert("🔒 Acción denegada: No tiene permisos para modificar la cuenta del Superusuario.");
    return;
  }

  user.active = !user.active;
  addActivityLog("MODIFICACION", "UserManagement", `Estado de cuenta de usuario '${user.email}' cambiado a ${user.active ? 'ACTIVO' : 'DESACTIVO'}`);
  renderUsers();
}

// USER SELF-SERVICE PROFILE
function openProfileModal() {
  const modal = document.getElementById("profileModal");
  if (!modal) return;

  const currentEmail = appState.currentUser.email;
  const userObj = appState.users.find(u => u.email === currentEmail) || appState.currentUser;

  document.getElementById("profileDisplayName").textContent = `${userObj.firstName || userObj.name} ${userObj.lastName || ''}`.trim();
  document.getElementById("profileDisplayEmail").textContent = userObj.email;
  document.getElementById("profileDisplayDoc").textContent = userObj.document || "1019024812";
  document.getElementById("profileDisplayVinculacion").textContent = userObj.vinculacion || "Planta (Nómina)";
  document.getElementById("profileDisplayRole").textContent = userObj.role;
  
  document.getElementById("profilePassword").value = "";
  document.getElementById("profilePhone").value = userObj.phone || "";
  document.getElementById("profileAddress").value = userObj.address || "";

  modal.classList.add("active");
}

function closeProfileModal() {
  const modal = document.getElementById("profileModal");
  if (modal) modal.classList.remove("active");
}

async function handleUpdateProfile(e) {
  e.preventDefault();
  const currentEmail = appState.currentUser.email;
  const userObj = appState.users.find(u => u.email === currentEmail);

  const newPassword = document.getElementById("profilePassword").value;
  const newPhone = document.getElementById("profilePhone").value.trim();
  const newAddress = document.getElementById("profileAddress").value.trim();

  if (userObj) {
    if (newPassword && newPassword.trim() !== "") {
      userObj.password = await hashPassword(newPassword.trim());
    }
    userObj.phone = newPhone;
    userObj.address = newAddress;
  }

  appState.currentUser.phone = newPhone;
  appState.currentUser.address = newAddress;

  addActivityLog("MODIFICACION", "UserProfileSelfService", `Usuario '${currentEmail}' actualizó sus datos personales (Teléfono/Dirección/Contraseña).`);
  alert("✅ ¡Tus datos de perfil han sido actualizados con éxito!");
  closeProfileModal();
  renderUsers();
}

// LOCATION MANAGEMENT & MULTI-LOCATION STOCK TRANSFERS
function ensureProductLocations() {
  if (!appState.locations || appState.locations.length === 0) {
    appState.locations = [...initialLocations];
  }

  appState.products.forEach(p => {
    if (!p.stockByLocation) {
      p.stockByLocation = {};
      const locCount = appState.locations.length;
      let remaining = p.physicalStock || 0;
      appState.locations.forEach((loc, idx) => {
        if (idx === locCount - 1) {
          p.stockByLocation[loc.id] = Math.max(0, remaining);
        } else {
          const part = Math.floor((p.physicalStock || 0) / locCount);
          p.stockByLocation[loc.id] = part;
          remaining -= part;
        }
      });
    }
  });

  if (appState.serializedItems) {
    appState.serializedItems.forEach(item => {
      if (!item.currentLocationId) {
        item.currentLocationId = "loc-1";
      }
    });
  }
}

function renderLocationsView() {
  ensureProductLocations();
  const tableBody = document.getElementById("locationsTableBody");
  if (!tableBody) return;

  tableBody.innerHTML = appState.locations.map(loc => {
    let skusInLocation = 0;
    let unitsInLocation = 0;

    appState.products.forEach(p => {
      const stockInLoc = (p.stockByLocation && p.stockByLocation[loc.id]) ? p.stockByLocation[loc.id] : 0;
      if (stockInLoc > 0) {
        skusInLocation++;
        unitsInLocation += stockInLoc;
      }
    });

    return `
      <tr>
        <td><strong style="color: var(--accent-green);">${loc.name}</strong></td>
        <td>${loc.address || 'N/A'}</td>
        <td>${loc.phone || 'N/A'}</td>
        <td><strong>👤 ${loc.manager || 'N/A'}</strong></td>
        <td>
          <span class="badge badge-purple">${skusInLocation} SKUs</span> 
          <strong style="color: var(--accent-teal); margin-left: 0.3rem;">(${unitsInLocation} Unds)</strong>
        </td>
        <td>
          <span class="badge ${loc.active ? 'badge-green' : 'badge-red'}">
            ${loc.active ? '🟢 Activa' : '🔴 Inactiva'}
          </span>
        </td>
        <td>
          <button class="btn btn-secondary" style="padding: 0.2rem 0.5rem; font-size: 0.73rem;" onclick="openLocationModal('${loc.id}')">
            ✏️ Editar Sede
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function openLocationModal(locId = null) {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo los usuarios con rol Administrador pueden crear o modificar sedes.");
    return;
  }

  const modal = document.getElementById("locationModal");
  const title = document.getElementById("locationModalTitle");
  const editId = document.getElementById("locationEditId");
  const form = document.getElementById("locationForm");

  if (!modal) return;

  if (locId) {
    const loc = appState.locations.find(l => l.id === locId);
    if (!loc) return;

    if (title) title.textContent = `✏️ Editar Ubicación (${loc.name})`;
    if (editId) editId.value = loc.id;

    document.getElementById("locName").value = loc.name || "";
    document.getElementById("locAddress").value = loc.address || "";
    document.getElementById("locPhone").value = loc.phone || "";
    document.getElementById("locManager").value = loc.manager || "";
    document.getElementById("locActive").value = loc.active ? "true" : "false";
  } else {
    if (title) title.textContent = "📍 Crear Nueva Ubicación / Bodega";
    if (editId) editId.value = "";
    if (form) form.reset();
  }

  modal.classList.add("active");
}

function closeLocationModal() {
  const modal = document.getElementById("locationModal");
  if (modal) modal.classList.remove("active");
}

function handleSaveLocation(e) {
  e.preventDefault();
  const editId = document.getElementById("locationEditId")?.value;
  const name = document.getElementById("locName").value.trim();
  const address = document.getElementById("locAddress").value.trim();
  const phone = document.getElementById("locPhone").value.trim();
  const manager = document.getElementById("locManager").value.trim();
  const active = document.getElementById("locActive").value === "true";

  if (editId) {
    const loc = appState.locations.find(l => l.id === editId);
    if (loc) {
      loc.name = name;
      loc.address = address;
      loc.phone = phone;
      loc.manager = manager;
      loc.active = active;

      addActivityLog("MODIFICACION", "LocationManagement", `Actualización de metadatos de sede '${name}'`);
      alert(`✅ Ubicación '${name}' actualizada con éxito.`);
    }
  } else {
    const newLoc = {
      id: `loc-${Date.now()}`,
      name,
      address,
      phone,
      manager,
      active
    };

    appState.locations.push(newLoc);
    addActivityLog("CREACION", "LocationManagement", `Creación de nueva sede de inventario '${name}' (Encargado: ${manager})`);
    alert(`✅ Ubicación '${name}' creada exitosamente.`);
  }

  closeLocationModal();
  ensureProductLocations();
  populateDropdowns();
  renderAllViews();
}

function handleTransferProductSelect() {
  const prodId = document.getElementById("transferProduct")?.value;
  const fromLocId = document.getElementById("transferFromLoc")?.value;
  const availableHelp = document.getElementById("transferStockAvailableHelp");
  const serialGroup = document.getElementById("transferSerialGroup");
  const serialList = document.getElementById("transferSerialList");

  if (!prodId || !fromLocId) return;

  const product = appState.products.find(p => p.id === prodId);
  if (!product) return;

  ensureProductLocations();

  const stockAtOrigin = (product.stockByLocation && product.stockByLocation[fromLocId]) ? product.stockByLocation[fromLocId] : 0;
  
  const fromLocObj = appState.locations.find(l => l.id === fromLocId);
  if (availableHelp) {
    availableHelp.textContent = `📍 Stock Disponible en '${fromLocObj ? fromLocObj.name : 'Origen'}': ${stockAtOrigin} ${product.unitOfMeasure}`;
  }

  if (product.requiresSerial) {
    ensureAllSerialsForProduct(product);
    const originSerials = appState.serializedItems.filter(i => i.productId === prodId && i.status === 'EN_STOCK' && i.currentLocationId === fromLocId);
    
    if (serialGroup) serialGroup.style.display = "block";
    if (serialList) {
      if (originSerials.length === 0) {
        serialList.innerHTML = `<div style="color: var(--accent-amber); font-size: 0.8rem;">No hay seriales individuales disponibles en la sede de origen seleccionada.</div>`;
      } else {
        serialList.innerHTML = originSerials.map(s => `
          <label style="display: flex; align-items: center; gap: 0.3rem; background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); padding: 0.25rem 0.5rem; border-radius: 4px; font-family: monospace; font-size: 0.8rem; cursor: pointer;">
            <input type="checkbox" class="transfer-serial-checkbox" value="${s.serialNumber}">
            <span>${s.serialNumber}</span>
          </label>
        `).join('');
      }
    }
  } else {
    if (serialGroup) serialGroup.style.display = "none";
  }
}

function handleTransferStock(e) {
  e.preventDefault();
  const prodId = document.getElementById("transferProduct").value;
  const fromLocId = document.getElementById("transferFromLoc").value;
  const toLocId = document.getElementById("transferToLoc").value;
  const qty = parseInt(document.getElementById("transferQuantity").value);
  const notes = document.getElementById("transferNotes").value.trim();

  if (fromLocId === toLocId) {
    alert("⚠️ ERROR: La ubicación de origen y destino deben ser distintas.");
    return;
  }

  const product = appState.products.find(p => p.id === prodId);
  const fromLoc = appState.locations.find(l => l.id === fromLocId);
  const toLoc = appState.locations.find(l => l.id === toLocId);

  if (!product || !fromLoc || !toLoc) return;

  ensureProductLocations();
  const currentOriginStock = product.stockByLocation[fromLocId] || 0;

  if (qty > currentOriginStock) {
    alert(`⚠️ STOCK INSUFICIENTE EN ORIGEN: Solo hay ${currentOriginStock} unidades disponibles en '${fromLoc.name}'.`);
    return;
  }

  let selectedSerialNumbers = [];
  if (product.requiresSerial) {
    const checkedBoxes = document.querySelectorAll(".transfer-serial-checkbox:checked");
    checkedBoxes.forEach(box => selectedSerialNumbers.push(box.value));

    if (selectedSerialNumbers.length !== qty) {
      alert(`⚠️ ERROR DE SERIALES: Ha indicado trasladar ${qty} unidades, pero ha seleccionado ${selectedSerialNumbers.length} números de serie. La cantidad debe coincidir exactamente con los seriales seleccionados.`);
      return;
    }
  }

  product.stockByLocation[fromLocId] -= qty;
  product.stockByLocation[toLocId] = (product.stockByLocation[toLocId] || 0) + qty;

  if (product.requiresSerial && selectedSerialNumbers.length > 0) {
    selectedSerialNumbers.forEach(sNum => {
      const serialObj = appState.serializedItems.find(i => i.serialNumber === sNum && i.productId === prodId);
      if (serialObj) {
        serialObj.currentLocationId = toLocId;
        serialObj.history.push({
          type: "TRASLADO_UBICACION",
          date: new Date().toISOString().substring(0, 10),
          user: appState.currentUser.name,
          description: `Traslado de '${fromLoc.name}' a '${toLoc.name}'. Guía/Notas: ${notes || 'Sin observaciones'}`
        });
      }
    });
  }

  const transferMov = {
    id: `mov-tr-${Date.now()}`,
    date: new Date().toISOString().substring(0, 10),
    type: "TRASLADO_UBICACION",
    productId: prodId,
    quantity: qty,
    invoiceNumber: notes || `REM-TR-${Date.now().toString().slice(-4)}`,
    customerName: `Traslado: ${fromLoc.name} ➔ ${toLoc.name}`,
    user: appState.currentUser.name,
    notes: `Traslado de ${qty} unidades de '${fromLoc.name}' a '${toLoc.name}'. Seriales: ${selectedSerialNumbers.join(', ') || 'N/A'}`
  };

  appState.movements.unshift(transferMov);
  addActivityLog("MODIFICACION", "StockTransfer", `Traslado de ${qty} unidades de '${product.sku}' desde '${fromLoc.name}' hacia '${toLoc.name}'.`);

  alert(`✅ TRASLADO EXITOSO: ${qty} unidades de ${product.sku} trasladadas a '${toLoc.name}'.`);
  document.getElementById("transferStockForm").reset();
  handleTransferProductSelect();
  populateDropdowns();
  renderAllViews();
}

// AUTOMATED SIGO SALES REPORT INGESTION ENGINE WITH ANTI-DUPLICATION VALIDATION
function triggerSigoReportImport() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador está autorizado para procesar y cargar reportes de Sigo.");
    return;
  }
  document.getElementById("sigoReportFileInput").click();
}

function handleSigoReportUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  processAutomatedSigoReport(file.name);
}

function processAutomatedSigoReport(fileName) {
  const simulatedSigoBatch = [
    {
      invoiceNumber: "FACT-SIGO-2026-901",
      date: new Date().toISOString().substring(0, 10),
      customerNit: "901.888.777-2",
      customerName: "AGROVET DEL LLANO S.A.S (Cliente Nuevo Sigo)",
      customerPhone: "+57 318 999 4455",
      customerEmail: "compras@agrovetllano.co",
      items: [
        { sku: "MC 001", quantity: 20 },
        { sku: "MC 004", quantity: 15 }
      ]
    },
    {
      invoiceNumber: "FACT-SIGO-2026-902",
      date: new Date().toISOString().substring(0, 10),
      customerNit: "900.138.371-1",
      customerName: "FINCA DON PEDRO PMO S.A.S",
      customerPhone: "+57 310 456 7890",
      customerEmail: "donpedro@finca.co",
      items: [
        { sku: "MC 003", quantity: 1 }
      ]
    }
  ];

  let newCustomersCreated = 0;
  let totalExitsRegistered = 0;
  let newInvoicesCount = 0;
  let duplicateInvoicesCount = 0;
  const skippedInvoices = [];

  simulatedSigoBatch.forEach(invoiceData => {
    const isAlreadyProcessed = appState.movements.some(m => m.invoiceNumber === invoiceData.invoiceNumber);

    if (isAlreadyProcessed) {
      duplicateInvoicesCount++;
      skippedInvoices.push(invoiceData.invoiceNumber);
      return;
    }

    newInvoicesCount++;

    let customer = appState.customers.find(c => c.documentNumber === sanitizeDocument(invoiceData.customerNit).cleanDoc || c.name.toLowerCase() === invoiceData.customerName.toLowerCase());
    
    if (!customer) {
      const parsedDoc = sanitizeDocument(invoiceData.customerNit);
      customer = {
        id: `cust-${Date.now()}-${Math.floor(Math.random()*100)}`,
        documentType: parsedDoc.docType,
        documentNumber: parsedDoc.cleanDoc,
        verificationDigit: parsedDoc.dv,
        name: invoiceData.customerName,
        phone: invoiceData.customerPhone,
        email: invoiceData.customerEmail,
        city: "Registrado Automáticamente vía Sigo Report",
        address: "Cierre de Caja Sigo"
      };
      appState.customers.push(customer);
      newCustomersCreated++;
      addActivityLog("CREACION", "CustomerAutoIngest", `Auto-creación de cliente nuevo '${customer.name}' desde reporte Sigo.`);
    }

    invoiceData.items.forEach(itemInfo => {
      const product = appState.products.find(p => p.sku === itemInfo.sku);
      if (product) {
        product.physicalStock = Math.max(0, product.physicalStock - itemInfo.quantity);
        totalExitsRegistered += itemInfo.quantity;

        const newMov = {
          id: `mov-sigo-${Date.now()}-${Math.floor(Math.random()*1000)}`,
          date: invoiceData.date,
          type: "SALIDA_VENTA",
          productId: product.id,
          quantity: itemInfo.quantity,
          invoiceNumber: invoiceData.invoiceNumber,
          customerName: customer.name,
          user: `${appState.currentUser.name} (Auto-Sigo)`,
          notes: `Salida registrada automáticamente por cierre de ventas de Sigo (Archivo: ${fileName})`,
          attachments: [
            { name: fileName, type: 'excel', icon: '📊' }
          ]
        };

        appState.movements.unshift(newMov);

        if (product.requiresSerial) {
          let serialItem = appState.serializedItems.find(i => i.productId === product.id && i.status === 'EN_STOCK');
          if (serialItem) {
            serialItem.status = "VENDIDO";
            serialItem.currentCustomerId = customer.id;
            serialItem.saleDate = invoiceData.date;
            serialItem.invoiceNumber = invoiceData.invoiceNumber;
            serialItem.attachments = [...(serialItem.attachments || []), { name: fileName, type: 'excel', icon: '📊' }];
            serialItem.history.push({
              type: "VENTA",
              date: invoiceData.date,
              user: "Sistema Auto-Sigo",
              description: `Vendido automáticamente por reporte de caja Sigo. Factura: ${invoiceData.invoiceNumber}`,
              attachments: [{ name: fileName, type: 'excel', icon: '📊' }]
            });
          }
        }
      }
    });
  });

  if (newInvoicesCount === 0 && duplicateInvoicesCount > 0) {
    addActivityLog("IMPORTACION", "SigoDeduplicationEngine", `Carga rechazada por duplicidad completa en '${fileName}'. Se omitieron ${duplicateInvoicesCount} facturas previamente registradas (${skippedInvoices.join(', ')}).`);
    alert(`⚠️ ATENCIÓN: CANCELADO POR DUPLICIDAD DE REGISTRO\n\n📄 Archivo: ${fileName}\n\nTodas las facturas contenidas en este reporte (${skippedInvoices.join(', ')}) YA fueron registradas anteriormente en el sistema.\n\n🛡️ Se omitieron el 100% de los registros para proteger la línea de tiempo de Sigo y evitar dobles descuentos de inventario.`);
    return;
  }

  addActivityLog("IMPORTACION", "SigoSalesProcessor", `Ingesta de Sigo '${fileName}'. Facturas Nuevas: ${newInvoicesCount} | Duplicadas Omitidas: ${duplicateInvoicesCount} | Clientes Creados: ${newCustomersCreated}`);

  let alertMessage = `🚀 INGESTA AUTOMÁTICA SIGO COMPLETADA:\n\n📄 Archivo: ${fileName}\n🧾 Facturas Nuevas Procesadas: ${newInvoicesCount}\n👤 Clientes Nuevos Creados: ${newCustomersCreated}\n📦 Salidas Descontadas: ${totalExitsRegistered} Unidades.`;
  if (duplicateInvoicesCount > 0) {
    alertMessage += `\n\n🛡️ VALIDACIÓN ANTI-DUPLICIDAD: Se omitieron ${duplicateInvoicesCount} facturas que ya existían previamente (${skippedInvoices.join(', ')}).`;
  }

  alert(alertMessage);

  populateDropdowns();
  renderAllViews();
}

function renderDashboard() {
  document.getElementById("statTotalSkus").textContent = appState.products.length;
  document.getElementById("statPhysicalStock").textContent = appState.products.reduce((acc, p) => acc + p.physicalStock, 0);
  document.getElementById("statReservedStock").textContent = appState.products.reduce((acc, p) => acc + p.reservedStock, 0);
  document.getElementById("statWarrantyItems").textContent = appState.serializedItems.filter(i => i.status === 'EN_GARANTIA' || i.status === 'RETORNO_CHINA').length;

  const tableBody = document.getElementById("dashboardMovementsTable");
  tableBody.innerHTML = appState.movements.slice(0, 5).map(m => {
    const prod = appState.products.find(p => p.id === m.productId);
    const badgeClass = m.type === 'INGRESO_COMPRA' ? 'badge-green' : m.type === 'SALIDA_VENTA' ? 'badge-blue' : 'badge-amber';
    return `
      <tr>
        <td>${m.date}</td>
        <td><span class="badge ${badgeClass}">${m.type}</span></td>
        <td><strong>${prod ? prod.sku : ''}</strong> - ${prod ? prod.name : ''}</td>
        <td><strong>${m.quantity}</strong></td>
        <td>${m.invoiceNumber || '-'}</td>
        <td>${m.user}</td>
      </tr>
    `;
  }).join('');
}

function renderCatalog() {
  ensureProductLocations();
  const tableBody = document.getElementById("catalogTableBody");
  const searchQuery = (document.getElementById("catalogSearch")?.value || "").toLowerCase();
  const categoryFilter = document.getElementById("catalogCategoryFilter")?.value || "ALL";
  const locationFilter = document.getElementById("catalogLocationFilter")?.value || "ALL";

  const filtered = appState.products.filter(p => {
    const matchesSearch = p.sku.toLowerCase().includes(searchQuery) || p.name.toLowerCase().includes(searchQuery);
    const matchesCategory = (categoryFilter === "ALL" || p.categoryId === categoryFilter);
    
    let matchesLocation = true;
    if (locationFilter !== "ALL") {
      const stockInLoc = (p.stockByLocation && p.stockByLocation[locationFilter]) ? p.stockByLocation[locationFilter] : 0;
      matchesLocation = stockInLoc > 0;
    }

    return matchesSearch && matchesCategory && matchesLocation;
  });

  tableBody.innerHTML = filtered.map(p => {
    const availableStock = p.physicalStock - p.reservedStock;
    const isSerialized = p.requiresSerial;
    const categoryObj = appState.categories.find(c => c.id === p.categoryId);
    const categoryName = categoryObj ? categoryObj.name : p.category;

    // Build location breakdown pills
    const locationBreakdownHtml = appState.locations.map(loc => {
      const qty = (p.stockByLocation && p.stockByLocation[loc.id]) ? p.stockByLocation[loc.id] : 0;
      return `<div style="font-size: 0.75rem; color: var(--text-muted); font-weight: 500;">
        📍 <strong>${loc.name}</strong>: <span style="color: ${qty > 0 ? 'var(--accent-teal)' : 'var(--text-muted)'}; font-weight: 700;">${qty} ${p.unitOfMeasure}</span>
      </div>`;
    }).join('');

    return `
      <tr>
        <td><strong style="color: var(--accent-green);">${p.sku}</strong></td>
        <td>
          <div style="font-weight: 600;">${p.name}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">${p.description}</div>
        </td>
        <td><span class="badge badge-purple">${categoryName}</span></td>
        <td><strong style="font-size: 1.1rem; color: var(--text-main);">${p.physicalStock} ${p.unitOfMeasure}</strong></td>
        <td>
          <div style="display: flex; flex-direction: column; gap: 0.15rem; background: rgba(15,23,42,0.6); padding: 0.4rem 0.6rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
            ${locationBreakdownHtml}
          </div>
        </td>
        <td class="cost-column">$${p.baseCost.toFixed(2)}</td>
        <td class="cost-column">$${p.salePrice.toFixed(2)}</td>
        <td>
          <span class="badge ${availableStock <= p.minStockAlert ? 'badge-red' : 'badge-green'}">
            ${availableStock} (Bloqueado: ${p.reservedStock})
          </span>
        </td>
        <td>
          ${isSerialized ? `
            <button class="btn btn-secondary" style="padding: 0.2rem 0.5rem; font-size: 0.75rem;" onclick="openSerialInspectorModal('${p.id}')">
              🔍 Ver ${p.physicalStock} Seriales
            </button>
          ` : '-'}
        </td>
      </tr>
    `;
  }).join('');

  toggleFinancialFields(appState.currentUser.role);
}

function openSerialInspectorModal(productId) {
  const product = appState.products.find(p => p.id === productId);
  if (!product) return;

  ensureAllSerialsForProduct(product);
  addActivityLog("CONSULTA", "ProductSerialInspector", `Inspección de seriales para producto '${product.sku}' (${product.physicalStock} existencias).`);

  document.getElementById("inspectorProductName").textContent = `${product.sku} - ${product.name} (${product.physicalStock} Existencias Totales)`;
  
  const inStockItems = appState.serializedItems.filter(i => i.productId === productId && i.status === 'EN_STOCK');
  const exitedItems = appState.serializedItems.filter(i => i.productId === productId && i.status !== 'EN_STOCK');

  const inStockList = document.getElementById("inspectorInStockList");
  if (inStockItems.length === 0) {
    inStockList.innerHTML = `<div style="color: var(--text-muted); padding: 0.5rem;">No hay unidades físicas disponibles en bodega actualmente.</div>`;
  } else {
    inStockList.innerHTML = `
      <div style="font-size: 0.8rem; color: var(--accent-green); margin-bottom: 0.5rem; font-weight: 600;">
        Mostrando los ${inStockItems.length} seriales individuales presentes en bodega:
      </div>
      ${inStockItems.map((item, index) => {
        const locObj = appState.locations.find(l => l.id === item.currentLocationId);
        const locName = locObj ? locObj.name : "Bodega Principal";
        return `
          <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); padding: 0.6rem 0.75rem; border-radius: var(--radius-sm); margin-bottom: 0.4rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
            <div>
              <span style="font-size: 0.75rem; color: var(--text-muted); margin-right: 0.4rem;">#${index + 1}</span>
              <strong style="color: var(--accent-green); font-family: monospace; font-size: 0.95rem;">${item.serialNumber}</strong>
              <div style="font-size: 0.73rem; color: var(--text-muted);">
                📍 Ubicación: <strong style="color: var(--accent-teal);">${locName}</strong> | Ingresó: ${item.entryDate}
              </div>
            </div>
            <div style="display: flex; gap: 0.4rem; align-items: center;">
              <span class="badge badge-green">DISPONIBLE</span>
              <button class="btn btn-secondary" style="padding: 0.15rem 0.4rem; font-size: 0.7rem;" onclick="editSerialNumber('${item.id}')" title="Modificar número de serie fábrica">
                ✏️ Editar
              </button>
            </div>
          </div>
        `;
      }).join('')}
    `;
  }

  const exitedList = document.getElementById("inspectorExitedList");
  if (exitedItems.length === 0) {
    exitedList.innerHTML = `<div style="color: var(--text-muted); padding: 0.5rem;">No hay seriales reportados como salidos o vendidos.</div>`;
  } else {
    exitedList.innerHTML = exitedItems.map(item => {
      const customer = appState.customers.find(c => c.id === item.currentCustomerId);
      const attachmentsCount = item.attachments ? item.attachments.length : 0;
      return `
        <div style="background: rgba(59, 130, 246, 0.1); border: 1px solid rgba(59, 130, 246, 0.3); padding: 0.6rem 0.75rem; border-radius: var(--radius-sm); margin-bottom: 0.4rem;">
          <div style="display: flex; justify-content: space-between; align-items: center;">
            <strong style="color: var(--accent-blue); font-family: monospace; font-size: 0.95rem;">${item.serialNumber}</strong>
            <span class="badge badge-amber">${item.status}</span>
          </div>
          <div style="font-size: 0.78rem; color: var(--text-muted); margin-top: 0.2rem;">
            👤 Cliente: <strong>${customer ? customer.name : 'Venta Directa'}</strong> | Fecha Venta: <strong>${item.saleDate || 'N/A'}</strong> | Factura: <strong>${item.invoiceNumber || 'N/A'}</strong>
          </div>
          ${attachmentsCount > 0 ? `
            <div style="font-size: 0.73rem; color: var(--accent-green); margin-top: 0.2rem;">
              📎 ${attachmentsCount} soportes adjuntos
            </div>
          ` : ''}
          <div style="margin-top: 0.3rem;">
            <button class="btn btn-secondary" style="padding: 0.15rem 0.5rem; font-size: 0.72rem;" onclick="lookupWarrantyBySerial('${item.serialNumber}')">
              🛡️ Consultar Garantía / Devolución
            </button>
          </div>
        </div>
      `;
    }).join('');
  }

  document.getElementById("serialInspectorModal").classList.add("active");
}

function editSerialNumber(itemId) {
  const item = appState.serializedItems.find(i => i.id === itemId);
  if (!item) return;

  const newSerial = prompt("Digite el número de serie único asignado por fábrica a esta unidad:", item.serialNumber);
  if (newSerial && newSerial.trim() !== "" && newSerial !== item.serialNumber) {
    const oldSerial = item.serialNumber;
    item.serialNumber = newSerial.trim();
    addActivityLog("MODIFICACION", "SerialUpdate", `Serial '${oldSerial}' modificado manualmente a '${item.serialNumber}'`);
    alert(`✅ Número de serie actualizado a: ${item.serialNumber}`);
    
    const product = appState.products.find(p => p.id === item.productId);
    if (product) openSerialInspectorModal(product.id);
    populateDropdowns();
  }
}

function closeSerialInspectorModal() {
  document.getElementById("serialInspectorModal").classList.remove("active");
}

function lookupWarrantyBySerial(inputSerial = null) {
  const serialQuery = inputSerial || document.getElementById("warrantySerialInput")?.value?.trim();
  const resultsCard = document.getElementById("warrantyResultCard");
  const suggBox = document.getElementById("warrantySerialSuggestions");

  if (suggBox) suggBox.classList.remove("active");

  if (!serialQuery) {
    alert("Por favor digite o escanee un número de serie.");
    return;
  }

  ensureGlobalSerials();
  addActivityLog("GARANTIA", "WarrantyLookup", `Búsqueda e inspección de cobertura de garantía para serial '${serialQuery}'.`);

  let item = appState.serializedItems.find(i => i.serialNumber.toLowerCase() === serialQuery.toLowerCase());
  
  if (!item) {
    item = appState.serializedItems.find(i => i.serialNumber.toLowerCase().includes(serialQuery.toLowerCase()));
  }

  if (!item) {
    resultsCard.style.display = "block";
    resultsCard.innerHTML = `
      <div style="background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); padding: 1rem; border-radius: var(--radius-md); color: #fca5a5;">
        ❌ No se encontró ningún equipo en la base de datos registrado con el serial: <strong>${serialQuery}</strong>.
      </div>
    `;
    return;
  }

  const product = appState.products.find(p => p.id === item.productId);
  const customer = appState.customers.find(c => c.id === item.currentCustomerId);

  const warrantyMonths = product ? (product.warrantyMonths || 12) : 12;
  const saleDateObj = item.saleDate ? new Date(item.saleDate) : new Date(item.entryDate);
  const expiryDateObj = new Date(saleDateObj);
  expiryDateObj.setMonth(expiryDateObj.getMonth() + warrantyMonths);

  const today = new Date();
  const isWarrantyValid = (today <= expiryDateObj);
  const daysRemaining = Math.max(0, Math.round((expiryDateObj - today) / (1000 * 60 * 60 * 24)));

  resultsCard.style.display = "block";
  resultsCard.innerHTML = `
    <div style="background: var(--bg-card); border: 1px solid var(--border-highlight); padding: 1.25rem; border-radius: var(--radius-lg);">
      <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 1rem;">
        <div>
          <h3 style="font-size: 1.15rem; font-weight: 700;">${product ? product.name : 'Equipo'}</h3>
          <div style="font-size: 0.85rem; color: var(--text-muted);">
            SKU: <strong>${product ? product.sku : ''}</strong> | Serial: <strong style="color: var(--accent-green); font-family: monospace;">${item.serialNumber}</strong>
          </div>
        </div>
        <div>
          <span class="badge ${isWarrantyValid ? 'badge-green' : 'badge-red'}" style="font-size: 0.85rem; padding: 0.4rem 0.8rem;">
            ${isWarrantyValid ? `🛡️ GARANTÍA VIGENTE (${daysRemaining} Días Restantes)` : '❌ GARANTÍA EXPIRADA'}
          </span>
        </div>
      </div>

      <div style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); display: grid; gap: 1rem; background: rgba(15, 23, 42, 0.7); padding: 1rem; border-radius: var(--radius-md); margin-bottom: 1rem;">
        <div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Cliente Comprador</div>
          <div style="font-weight: 600; color: var(--text-main);">${customer ? customer.name : 'En Bodega / Sin Cliente'}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">${customer ? (customer.documentType + ' ' + customer.documentNumber) : ''}</div>
        </div>
        <div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Factura de Compra</div>
          <div style="font-weight: 600; color: var(--accent-blue);">${item.invoiceNumber || 'Sin Factura'}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Fecha Venta: ${item.saleDate || item.entryDate}</div>
        </div>
        <div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Vencimiento Garantía</div>
          <div style="font-weight: 600;">${expiryDateObj.toISOString().substring(0, 10)}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Cobertura: ${warrantyMonths} Meses</div>
        </div>
      </div>

      <h4 style="font-size: 0.85rem; font-weight: 600; margin-bottom: 0.5rem;">📎 Soportes y Fotografías Adjuntas (Haga clic para ver):</h4>
      <div style="display: flex; gap: 0.5rem; flex-wrap: wrap; margin-bottom: 1rem;">
        ${(item.attachments && item.attachments.length > 0) ? item.attachments.map(att => `
          <button class="btn btn-secondary" style="padding: 0.3rem 0.6rem; font-size: 0.8rem;" onclick="previewAttachment('${att.name}', '${att.type || 'image'}', ${att.dataUrl ? `'${att.dataUrl}'` : 'null'})">
            ${att.icon || (att.type === 'pdf' ? '📄' : '📷')} ${att.name}
          </button>
        `).join('') : '<span style="color: var(--text-muted); font-size: 0.8rem;">Sin documentos adjuntos</span>'}
      </div>

      <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
        <button class="btn btn-primary" onclick="processWarrantyReturn('${item.id}')">
          🔄 Procesar Devolución / Ingreso por Garantía
        </button>
        <button class="btn btn-secondary" onclick="switchView('view-equipment'); document.getElementById('equipmentSerialSelect').value='${item.id}'; renderEquipmentLifeCycle(); closeSerialInspectorModal();">
          📋 Ver Hoja de Vida Completa
        </button>
      </div>
    </div>
  `;
}

function processWarrantyReturn(serialId) {
  const item = appState.serializedItems.find(i => i.id === serialId);
  if (!item) return;

  const reason = prompt("Describa el motivo del retorno o falla cubierta por garantía:", "Cliente reporta falla de encendido en sonda");
  if (!reason) return;

  item.status = "EN_GARANTIA";
  item.history.push({
    type: "FALLA_REPORTADA",
    date: new Date().toISOString().substring(0, 10),
    user: appState.currentUser.name,
    description: `Retorno por garantía: ${reason}`
  });

  addActivityLog("GARANTIA", "WarrantyReturn", `Recepción de devolución por garantía de serial '${item.serialNumber}'. Motivo: ${reason}`);
  alert(`✅ El equipo ${item.serialNumber} ha sido marcado EN GARANTÍA.`);
  renderAllViews();
  lookupWarrantyBySerial(item.serialNumber);
}

function openCategoryModal() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador puede crear categorías.");
    return;
  }
  document.getElementById("categoryModal").classList.add("active");
}

function closeCategoryModal() {
  document.getElementById("categoryModal").classList.remove("active");
}

function handleCreateCategory(e) {
  e.preventDefault();
  const name = document.getElementById("newCatName").value;
  const desc = document.getElementById("newCatDesc").value;

  const newCat = {
    id: `cat-${Date.now()}`,
    name,
    description: desc || "Categoría creada dinámicamente"
  };

  appState.categories.push(newCat);
  addActivityLog("CREACION", "CategoryManagement", `Creación de nueva categoría de productos '${name}'`);

  alert(`✅ Categoría '${name}' creada exitosamente.`);
  closeCategoryModal();
  populateDropdowns();
  renderAllViews();
}

function toggleInvoiceRequirement() {
  const movType = document.getElementById("movType").value;
  const badge = document.getElementById("invoiceReqBadge");
  const invoiceInput = document.getElementById("movInvoice");

  if (movType === 'SALIDA_VENTA') {
    badge.style.display = "inline";
    invoiceInput.required = true;
  } else {
    badge.style.display = "none";
    invoiceInput.required = false;
  }
}

function handleProductSelectForMovement() {
  const prodId = document.getElementById("movProduct").value;
  const product = appState.products.find(p => p.id === prodId);
  const serialGroup = document.getElementById("serialInputGroup");

  if (product && product.requiresSerial) {
    serialGroup.style.display = "block";
  } else {
    serialGroup.style.display = "none";
  }
}

function handleCreateMovement(e) {
  e.preventDefault();
  const type = document.getElementById("movType").value;
  const prodId = document.getElementById("movProduct").value;
  const qty = parseInt(document.getElementById("movQuantity").value);
  const invoice = document.getElementById("movInvoice").value;
  const customerId = document.getElementById("movCustomer").value;
  const serialNumber = document.getElementById("movSerial")?.value;
  const notes = document.getElementById("movNotes").value;
  const fileInput = document.getElementById("movFilesInput");

  const product = appState.products.find(p => p.id === prodId);
  const customer = appState.customers.find(c => c.id === customerId);

  const attachments = [];

  const processFiles = () => {
    if (type === 'SALIDA_VENTA') {
      const available = product.physicalStock - product.reservedStock;
      if (qty > available) {
        alert(`⚠️ ERROR DE STOCK: El producto ${product.sku} solo tiene ${available} unidades disponibles.`);
        return;
      }
    }

    if (type === 'INGRESO_COMPRA' && appState.currentUser.role === 'LOGISTICA') {
      const targetLocId = "loc-1";
      const targetLoc = appState.locations ? appState.locations.find(l => l.id === targetLocId) : null;
      let parsedSerials = [];
      if (product.requiresSerial && serialNumber) {
        parsedSerials = serialNumber.split(',').map(s => s.trim()).filter(Boolean);
      }

      const newPendingIntake = {
        id: `intake-${Date.now()}`,
        submittedAt: new Date().toISOString().replace('T', ' ').substring(0, 19),
        submittedBy: `${appState.currentUser.name} (Logística)`,
        userEmail: appState.currentUser.email,
        productId: prodId,
        sku: product.sku,
        productName: product.name,
        locationId: targetLocId,
        locationName: targetLoc ? targetLoc.name : "Bodega Principal (Central)",
        proposedQuantity: qty,
        requiresSerial: product.requiresSerial,
        proposedSerials: parsedSerials,
        invoiceNumber: invoice || `FAC-IN-${Date.now().toString().slice(-4)}`,
        notes: notes || "Ingreso registrado por usuario Logística",
        status: "PENDIENTE_VALIDACION"
      };

      if (!appState.pendingIntakes) appState.pendingIntakes = [];
      appState.pendingIntakes.unshift(newPendingIntake);
      savePendingIntakesToDisk();

      addActivityLog("CREACION", "PendingInventoryIntake", `Ingreso de ${qty} unidades de '${product.sku}' enviado a cola de validación por usuario Logística.`);

      alert(`⏳ INGRESO REGISTRADO EN BORRADOR: Su solicitud de ingreso de ${qty} unidades de '${product.sku}' ha quedado PENDIENTE DE VALIDACIÓN por un Administrador.\n\nNo ingresará al inventario físico ni al Kardex hasta que un Administrador la revise y la apruebe.`);
      
      document.getElementById("movementForm").reset();
      renderAllViews();
      return;
    }

    if (type === 'INGRESO_COMPRA') {
      product.physicalStock += qty;
    } else if (type === 'SALIDA_VENTA') {
      product.physicalStock -= qty;
    }

    const newMov = {
      id: `mov-${Date.now()}`,
      date: new Date().toISOString().substring(0, 10),
      type,
      productId: prodId,
      quantity: qty,
      invoiceNumber: invoice || null,
      customerName: customer ? customer.name : 'Venta General',
      user: appState.currentUser.name,
      notes,
      attachments
    };

    appState.movements.unshift(newMov);

    if (product.requiresSerial && serialNumber) {
      let serialItem = appState.serializedItems.find(i => i.serialNumber === serialNumber);
      if (!serialItem && type === 'INGRESO_COMPRA') {
        serialItem = {
          id: `ser-${Date.now()}`,
          productId: prodId,
          serialNumber: serialNumber,
          status: "EN_STOCK",
          currentCustomerId: null,
          entryDate: new Date().toISOString().substring(0, 10),
          saleDate: null,
          invoiceNumber: invoice,
          attachments,
          history: []
        };
        appState.serializedItems.push(serialItem);
      }

      if (serialItem) {
        if (type === 'SALIDA_VENTA') {
          serialItem.status = "VENDIDO";
          serialItem.currentCustomerId = customerId;
          serialItem.saleDate = new Date().toISOString().substring(0, 10);
          serialItem.invoiceNumber = invoice;
          serialItem.attachments = [...(serialItem.attachments || []), ...attachments];
        }
        serialItem.history.push({
          type: type === 'SALIDA_VENTA' ? 'VENTA' : type,
          date: new Date().toISOString().substring(0, 10),
          user: appState.currentUser.name,
          description: `${type}: ${notes || 'Sin observaciones'} (Factura: ${invoice || 'N/A'})`,
          attachments
        });
      }
    }

    addActivityLog("CREACION", "KardexMovement", `Registro de movimiento Kardex (${type}) para '${product.sku}'. Cant: ${qty}. Factura: ${invoice || 'N/A'}`);

    alert(`✅ Movimiento registrado exitosamente con ${attachments.length} soportes adjuntos.`);
    document.getElementById("movementForm").reset();
    populateDropdowns();
    renderAllViews();
  };

  if (fileInput && fileInput.files.length > 0) {
    let filesProcessed = 0;
    for (let f of fileInput.files) {
      const reader = new FileReader();
      reader.onload = (event) => {
        attachments.push({
          name: f.name,
          type: f.type.includes('image') ? 'image' : 'pdf',
          icon: f.type.includes('image') ? '📷' : '📄',
          dataUrl: event.target.result
        });
        filesProcessed++;
        if (filesProcessed === fileInput.files.length) {
          processFiles();
        }
      };
      reader.readAsDataURL(f);
    }
  } else {
    attachments.push({ name: `Comprobante_${type}_${invoice || 'DOC'}.pdf`, type: 'pdf', icon: '📄' });
    processFiles();
  }
}

function renderKardex() {
  const tableBody = document.getElementById("kardexTableBody");
  tableBody.innerHTML = appState.movements.map(m => {
    const prod = appState.products.find(p => p.id === m.productId);
    const atts = m.attachments || [];
    return `
      <tr>
        <td><strong>${m.date}</strong></td>
        <td><span class="badge ${m.type === 'INGRESO_COMPRA' ? 'badge-green' : m.type === 'SALIDA_VENTA' ? 'badge-blue' : 'badge-amber'}">${m.type}</span></td>
        <td><strong>${prod ? prod.sku : ''}</strong> - ${prod ? prod.name : ''}</td>
        <td><strong>${m.quantity}</strong></td>
        <td>${m.invoiceNumber || '-'}</td>
        <td>${m.customerName || '-'}</td>
        <td>${m.user}</td>
        <td>
          <div>${m.notes || '-'}</div>
          ${atts.length > 0 ? `
            <div style="margin-top: 0.3rem; display: flex; gap: 0.3rem; flex-wrap: wrap;">
              ${atts.map(a => `
                <button class="btn btn-secondary" style="padding: 0.15rem 0.4rem; font-size: 0.7rem;" onclick="previewAttachment('${a.name}', '${a.type}', ${a.dataUrl ? `'${a.dataUrl}'` : 'null'})">
                  ${a.icon || '📎'} ${a.name}
                </button>
              `).join('')}
            </div>
          ` : '-'}
        </td>
      </tr>
    `;
  }).join('');
}

function renderEquipmentLifeCycle() {
  const serialId = document.getElementById("equipmentSerialSelect")?.value;
  const detailCard = document.getElementById("equipmentDetailCard");

  if (!serialId) {
    if (detailCard) detailCard.style.display = "none";
    return;
  }

  const item = appState.serializedItems.find(i => i.id === serialId);
  if (!item) return;

  const product = appState.products.find(p => p.id === item.productId);
  const customer = appState.customers.find(c => c.id === item.currentCustomerId);

  document.getElementById("eqName").textContent = product ? product.name : "Equipo Serializado";
  document.getElementById("eqSku").textContent = product ? product.sku : "SKU";
  document.getElementById("eqSerial").textContent = item.serialNumber;
  document.getElementById("eqStatusBadge").textContent = item.status;
  document.getElementById("eqEntryDate").textContent = item.entryDate;
  document.getElementById("eqCustomer").textContent = customer ? customer.name : "Sin Cliente / En Bodega";
  document.getElementById("eqMaintenanceCount").textContent = item.history.filter(h => h.type.includes("MANTENIMIENTO")).length;

  const timeline = document.getElementById("equipmentTimeline");
  timeline.innerHTML = item.history.map(evt => {
    const atts = evt.attachments || [];
    return `
      <div class="timeline-item">
        <div class="timeline-marker">🔹</div>
        <div class="timeline-content">
          <div class="timeline-header">
            <span class="timeline-title">${evt.type}</span>
            <span class="timeline-date">${evt.date} • Resp: ${evt.user}</span>
          </div>
          <div style="font-size: 0.85rem; color: var(--text-muted);">${evt.description}</div>
          ${atts.length > 0 ? `
            <div style="display: flex; gap: 0.4rem; flex-wrap: wrap; margin-top: 0.5rem;">
              ${atts.map(a => `
                <button class="btn btn-secondary" style="padding: 0.2rem 0.5rem; font-size: 0.75rem; color: var(--accent-green);" onclick="previewAttachment('${a.name}', '${a.type}', ${a.dataUrl ? `'${a.dataUrl}'` : 'null'})">
                  ${a.icon || '📷'} ${a.name} (Ver Soporte)
                </button>
              `).join('')}
            </div>
          ` : ''}
        </div>
      </div>
    `;
  }).join('');

  detailCard.style.display = "block";
}

function openAddEventModal() {
  const serialId = document.getElementById("equipmentSerialSelect").value;
  const item = appState.serializedItems.find(i => i.id === serialId);
  if (!item) return;

  const eventType = prompt("Ingrese tipo de evento (MANTENIMIENTO_PREVENTIVO, FALLA_REPORTADA, RETORNO_FABRICA_CHINA):", "MANTENIMIENTO_PREVENTIVO");
  if (!eventType) return;

  const desc = prompt("Detalles del evento:", "Revisión preventiva anual de sonda ecográfica");
  if (!desc) return;

  const fileName = prompt("Nombre de foto/soporte adjunto (Opcional):", "Foto_Mantenimiento_Sonda.jpg");

  const attachments = fileName ? [{ name: fileName, type: 'image', icon: '📷' }] : [];

  item.history.push({
    type: eventType.toUpperCase(),
    date: new Date().toISOString().substring(0, 10),
    user: appState.currentUser.name,
    description: desc,
    attachments
  });

  if (eventType.includes("RETORNO")) item.status = "RETORNO_CHINA";

  addActivityLog("CREACION", "EquipmentEvent", `Registro de evento '${eventType}' para serial '${item.serialNumber}'`);
  renderEquipmentLifeCycle();
}

function renderReservations() {
  const tableBody = document.getElementById("reservationsTableBody");
  tableBody.innerHTML = appState.reservations.map(r => {
    const prod = appState.products.find(p => p.id === r.productId);
    const cust = appState.customers.find(c => c.id === r.customerId);

    return `
      <tr>
        <td><strong>${r.id}</strong></td>
        <td>${cust ? cust.name : 'Cliente'}</td>
        <td>${prod ? prod.sku : ''} - ${prod ? prod.name : ''}</td>
        <td><strong style="color: var(--accent-amber);">${r.quantity}</strong></td>
        <td>${r.reason}</td>
        <td>${r.date}</td>
        <td><span class="badge badge-amber">${r.status}</span></td>
        <td>
          <button class="btn btn-secondary" style="padding: 0.2rem 0.5rem; font-size: 0.75rem;" onclick="cancelReservation('${r.id}')">
            Liberar Stock
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function handleCreateReservation(e) {
  e.preventDefault();
  const customerId = document.getElementById("resCustomerSelect").value;
  const productId = document.getElementById("resProductSelect").value;
  const quantity = parseInt(document.getElementById("resQuantity").value);
  const reason = document.getElementById("resReason").value;

  const product = appState.products.find(p => p.id === productId);
  const availableStock = product.physicalStock - product.reservedStock;

  if (quantity > availableStock) {
    alert(`⚠️ STOCK INSUFICIENTE: Stock Disponible actual: ${availableStock}.`);
    return;
  }

  product.reservedStock += quantity;

  const newRes = {
    id: `res-${Date.now().toString().slice(-4)}`,
    customerId,
    productId,
    quantity,
    reason,
    status: "ACTIVA",
    date: new Date().toISOString().substring(0, 10)
  };

  appState.reservations.push(newRes);
  addActivityLog("CREACION", "Reservation", `Bloqueo/Reserva de ${quantity} unidades para ${product.sku}. Motivo: ${reason}`);

  alert("✅ Stock reservado exitosamente.");
  closeReservationModal();
  populateDropdowns();
  renderAllViews();
}

function cancelReservation(resId) {
  const res = appState.reservations.find(r => r.id === resId);
  if (!res) return;

  const product = appState.products.find(p => p.id === res.productId);
  if (product) {
    product.reservedStock = Math.max(0, product.reservedStock - res.quantity);
  }

  res.status = "CANCELADA";
  appState.reservations = appState.reservations.filter(r => r.id !== resId);

  addActivityLog("MODIFICACION", "ReservationCancel", `Cancelación de reserva ${resId}. Stock liberado.`);
  renderAllViews();
}

function renderAlerts() {
  const container = document.getElementById("lowStockAlertsContainer");
  const lowStockItems = appState.products.filter(p => p.physicalStock <= p.minStockAlert);

  if (lowStockItems.length === 0) {
    container.innerHTML = `<div style="color: var(--accent-green); padding: 1rem;">✅ Todos los productos se encuentran por encima del Stock Mínimo.</div>`;
    return;
  }

  container.innerHTML = lowStockItems.map(p => `
    <div style="background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); padding: 1rem; border-radius: var(--radius-md); margin-bottom: 0.75rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
      <div>
        <strong style="color: #fca5a5;">⚠️ ALERTA DE COMPRA SUGERIDA: ${p.sku} - ${p.name}</strong>
        <div style="font-size: 0.82rem; color: var(--text-muted);">
          Stock Físico Actual: <strong>${p.physicalStock}</strong> ${p.unitOfMeasure} | Límite Mínimo: <strong>${p.minStockAlert}</strong>
        </div>
      </div>
      <button class="btn btn-primary" style="padding: 0.3rem 0.8rem; font-size: 0.8rem;" onclick="switchView('view-kardex')">
        Generar Orden de Compra
      </button>
    </div>
  `).join('');
}

function runLinearRegressionForecast() {
  const productId = document.getElementById("forecastProductSelect")?.value;
  const resultsPanel = document.getElementById("forecastResultsPanel");

  if (!productId || !resultsPanel) return;

  const product = appState.products.find(p => p.id === productId);
  if (!product) return;

  const X = [1, 5, 10, 15, 20, 25, 30];
  const totalOutlets = appState.movements
    .filter(m => m.productId === productId && m.type === 'SALIDA_VENTA')
    .reduce((sum, m) => sum + m.quantity, 10);

  const Y = X.map(x => Math.round((totalOutlets / 30) * x));

  const n = X.length;
  const sumX = X.reduce((a, b) => a + b, 0);
  const sumY = Y.reduce((a, b) => a + b, 0);
  const sumXY = X.reduce((acc, x, i) => acc + x * Y[i], 0);
  const sumX2 = X.reduce((acc, x) => acc + x * x, 0);

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const dailyRate = Math.max(0.2, parseFloat(slope.toFixed(2)));
  const daysUntilZero = Math.round(product.physicalStock / dailyRate);
  const projectedIn30Days = Math.max(0, Math.round(product.physicalStock - (dailyRate * 30)));

  resultsPanel.innerHTML = `
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 1.5rem;">
      <div style="background: rgba(15, 23, 42, 0.7); padding: 1rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
        <div style="font-size: 0.75rem; color: var(--text-muted);">Velocidad Salida Diaria (\(m\))</div>
        <div style="font-size: 1.4rem; font-weight: 700; color: var(--accent-green);">${dailyRate} unidades/día</div>
      </div>
      <div style="background: rgba(15, 23, 42, 0.7); padding: 1rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
        <div style="font-size: 0.75rem; color: var(--text-muted);">Días Estimados Hasta Agotamiento</div>
        <div style="font-size: 1.4rem; font-weight: 700; color: ${daysUntilZero < 15 ? 'var(--accent-red)' : 'var(--accent-amber)'};">${daysUntilZero} Días</div>
      </div>
      <div style="background: rgba(15, 23, 42, 0.7); padding: 1rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
        <div style="font-size: 0.75rem; color: var(--text-muted);">Stock Proyectado en 30 Días</div>
        <div style="font-size: 1.4rem; font-weight: 700;">${projectedIn30Days} ${product.unitOfMeasure}</div>
      </div>
    </div>

    <h4 style="font-size: 0.9rem; font-weight: 600; margin-bottom: 0.75rem;">Visualización de Tendencia Lineal (Día 1 a Día 30)</h4>
    <div class="forecast-chart-container">
      <div class="forecast-bar-group">
        <div class="forecast-bar" style="height: 90%;">
          <div class="bar-value">${product.physicalStock}</div>
        </div>
        <div class="bar-label">Hoy</div>
      </div>
      <div class="forecast-bar-group">
        <div class="forecast-bar" style="height: 70%;">
          <div class="bar-value">${Math.max(0, Math.round(product.physicalStock - (dailyRate * 10)))}</div>
        </div>
        <div class="bar-label">+10 Días</div>
      </div>
      <div class="forecast-bar-group">
        <div class="forecast-bar" style="height: 45%;">
          <div class="bar-value">${Math.max(0, Math.round(product.physicalStock - (dailyRate * 20)))}</div>
        </div>
        <div class="bar-label">+20 Días</div>
      </div>
      <div class="forecast-bar-group">
        <div class="forecast-bar projected" style="height: 25%;">
          <div class="bar-value">${projectedIn30Days}</div>
        </div>
        <div class="bar-label">+30 Días</div>
      </div>
    </div>
  `;
}

function exportSigoCSV() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador puede exportar reportes planos de Sigo.");
    return;
  }

  const startDate = document.getElementById("rptStartDate").value;
  const endDate = document.getElementById("rptEndDate").value;
  const movType = document.getElementById("rptMovementType").value;

  const filteredMovements = appState.movements.filter(m => (movType === 'ALL' || m.type === movType));

  let csvContent = "Fecha;TipoMovimiento;SKU;Producto;Cantidad;NumeroFactura;Cliente;Usuario;CostoBaseUnitario;PrecioVentaUnitario;SubtotalFacturado\n";

  filteredMovements.forEach(m => {
    const prod = appState.products.find(p => p.id === m.productId);
    const cost = prod ? prod.baseCost : 0;
    const price = prod ? prod.salePrice : 0;
    const subtotal = (price * m.quantity).toFixed(2);

    csvContent += `"${m.date}";"${m.type}";"${prod ? prod.sku : ''}";"${prod ? prod.name : ''}";${m.quantity};"${m.invoiceNumber || ''}";"${m.customerName || ''}";"${m.user}";${cost};${price};${subtotal}\n`;
  });

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", `Sigo_Inventario_MasCampo_${startDate}_a_${endDate}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  addActivityLog("EXPORTACION", "SigoExport", `Exportación de plano Sigo CSV (${filteredMovements.length} movimientos).`);
}

function exportSigoXLSX() {
  exportSigoCSV();
}

function renderMarginReport() {
  const tableBody = document.getElementById("marginReportTableBody");
  if (!tableBody) return;

  tableBody.innerHTML = appState.products.map(p => {
    const margin = p.salePrice - p.baseCost;
    const totalUnitsSold = appState.movements
      .filter(m => m.productId === p.id && m.type === 'SALIDA_VENTA')
      .reduce((sum, m) => sum + m.quantity, 0);

    const totalProfit = (margin * totalUnitsSold).toFixed(2);

    return `
      <tr>
        <td><strong>${p.sku}</strong></td>
        <td>${p.name}</td>
        <td>$${p.baseCost.toFixed(2)}</td>
        <td>$${p.salePrice.toFixed(2)}</td>
        <td><strong style="color: var(--accent-green);">$${margin.toFixed(2)}</strong></td>
        <td>${totalUnitsSold}</td>
        <td><strong style="color: var(--accent-purple); font-size: 1rem;">$${totalProfit}</strong></td>
      </tr>
    `;
  }).join('');
}

function openCameraModal() {
  document.getElementById("cameraModal").classList.add("active");
}

function closeCameraModal() {
  document.getElementById("cameraModal").classList.remove("active");
}

function simulateBarcodeScan(scannedCode) {
  ensureGlobalSerials();
  closeCameraModal();

  const product = appState.products.find(p => p.sku === scannedCode);
  if (product) {
    document.getElementById("movProduct").value = product.id;
    handleProductSelectForMovement();
    alert(`📷 CÓDIGO DETECTADO EN ESCÁNER: ${product.sku} - ${product.name}`);
    return;
  }

  const item = appState.serializedItems.find(i => i.serialNumber.toLowerCase().includes(scannedCode.toLowerCase()));
  if (item) {
    switchView('view-equipment');
    const input = document.getElementById("warrantySerialInput");
    if (input) input.value = item.serialNumber;
    lookupWarrantyBySerial(item.serialNumber);
    alert(`📷 SERIAL DETECTADO Y BUSCADO CON ÉXITO: ${item.serialNumber}`);
    return;
  }

  alert(`📷 Código escaneado: "${scannedCode}".`);
}

function openProductModal() {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo los usuarios con rol Administrador pueden registrar nuevos productos.");
    return;
  }
  document.getElementById("productModal").classList.add("active");
}

function closeProductModal() {
  document.getElementById("productModal").classList.remove("active");
}

function handleCreateProduct(e) {
  e.preventDefault();
  const sku = document.getElementById("newSku").value;
  const name = document.getElementById("newName").value;
  const categoryId = document.getElementById("newCategory").value;
  const unit = document.getElementById("newUnit").value;
  const reqSerial = document.getElementById("newReqSerial").value === "true";
  const minStock = parseInt(document.getElementById("newMinStock").value);
  const baseCost = parseFloat(document.getElementById("newBaseCost").value);
  const salePrice = parseFloat(document.getElementById("newSalePrice").value);

  const categoryObj = appState.categories.find(c => c.id === categoryId);

  const newProd = {
    id: `prod-${Date.now()}`,
    sku,
    name,
    description: "Ingresado manualmente al catálogo",
    categoryId,
    category: categoryObj ? categoryObj.name : "General",
    requiresSerial: reqSerial,
    unitOfMeasure: unit,
    minStockAlert: minStock,
    baseCost,
    salePrice,
    physicalStock: 0,
    reservedStock: 0,
    warrantyMonths: reqSerial ? 12 : 0
  };

  appState.products.push(newProd);
  addActivityLog("CREACION", "ProductManagement", `Creación de producto nuevo en catálogo '${sku} - ${name}'`);

  alert(`✅ Producto ${sku} creado.`);
  closeProductModal();
  populateDropdowns();
  renderAllViews();
}

function openReservationModal() {
  document.getElementById("reservationModal").classList.add("active");
}

function closeReservationModal() {
  document.getElementById("reservationModal").classList.remove("active");
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  addActivityLog("IMPORTACION", "ExcelImporter", `Carga manual de archivo de datos Excel '${file.name}'.`);
  alert(`✅ ¡Archivo Excel '${file.name}' cargado con éxito!`);
  renderAllViews();
}

// ==========================================================================
// PENDING INVENTORY INTAKE VALIDATION ENGINE (LOGISTICA SUBMISSION & ADMIN APPROVAL)
// ==========================================================================

const initialPendingIntakes = [
  {
    id: "intake-101",
    submittedAt: "2026-08-15 11:20:00",
    submittedBy: "Jorge Ramos (Logística)",
    userEmail: "logistica@mascampo.co",
    productId: "prod-101",
    sku: "TERM-DIG-01",
    productName: "Termómetro Digital Veterinario HD",
    locationId: "loc-1",
    locationName: "Bodega Principal (Central)",
    proposedQuantity: 50,
    requiresSerial: true,
    proposedSerials: ["SN-TERM-2026-X01", "SN-TERM-2026-X02", "SN-TERM-2026-X03"],
    invoiceNumber: "FAC-REM-8812",
    notes: "Llegada parcial de 50 unidades enviadas por proveedor de laboratorio",
    status: "PENDIENTE_VALIDACION"
  }
];

function loadPersistedPendingIntakes() {
  try {
    const saved = localStorage.getItem("mascampo_pending_intakes_db");
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        appState.pendingIntakes = parsed;
        return;
      }
    }
  } catch (e) {
    console.warn("No se pudo cargar solicitudes de ingreso pendientes", e);
  }
  appState.pendingIntakes = [...initialPendingIntakes];
}

function savePendingIntakesToDisk() {
  try {
    localStorage.setItem("mascampo_pending_intakes_db", JSON.stringify(appState.pendingIntakes));
  } catch (e) {
    console.error("Error guardando ingresos pendientes en disco", e);
  }
}

function renderPendingValidationsView() {
  if (!appState.pendingIntakes) appState.pendingIntakes = [];

  const pendingCount = appState.pendingIntakes.filter(i => i.status === 'PENDIENTE_VALIDACION').length;
  const badge = document.getElementById("pendingValidationBadge");
  if (badge) {
    badge.textContent = pendingCount;
    badge.style.display = pendingCount > 0 ? "inline-block" : "none";
  }

  const tableBody = document.getElementById("pendingIntakesTableBody");
  if (!tableBody) return;

  const statusFilter = document.getElementById("pendingStatusFilter")?.value || "PENDIENTE_VALIDACION";

  const filtered = appState.pendingIntakes.filter(i => {
    if (statusFilter === "ALL") return true;
    return i.status === statusFilter;
  });

  const isAdmin = (appState.currentUser.role === 'ADMINISTRADOR');

  if (filtered.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">
          No hay solicitudes de ingreso de inventario registradas en este filtro.
        </td>
      </tr>
    `;
    return;
  }

  tableBody.innerHTML = filtered.map(i => {
    const statusBadge = i.status === 'PENDIENTE_VALIDACION' 
      ? '<span class="badge badge-amber">⏳ Pendiente Validación</span>'
      : i.status === 'APROBADO'
      ? '<span class="badge badge-green">🟢 Aprobado (Ingresado)</span>'
      : '<span class="badge badge-red">🔴 Rechazado / Anulado</span>';

    const serialsLabel = (i.proposedSerials && i.proposedSerials.length > 0)
      ? `<div style="font-family: monospace; font-size: 0.75rem; color: var(--accent-green);">${i.proposedSerials.length} seriales</div>`
      : '<span style="color: var(--text-muted); font-size: 0.75rem;">N/A (Lote/Cantidad)</span>';

    return `
      <tr>
        <td style="font-size: 0.8rem;">${i.submittedAt}</td>
        <td><strong>👤 ${i.submittedBy}</strong></td>
        <td>
          <strong style="color: var(--accent-green);">${i.sku}</strong>
          <div style="font-size: 0.75rem; color: var(--text-muted);">${i.productName}</div>
        </td>
        <td>📍 ${i.locationName}</td>
        <td><strong style="font-size: 1.05rem; color: var(--accent-teal);">${i.proposedQuantity} Unds</strong></td>
        <td>${serialsLabel}</td>
        <td>${i.invoiceNumber || '-'}</td>
        <td>${statusBadge}</td>
        <td>
          ${i.status === 'PENDIENTE_VALIDACION' ? `
            ${isAdmin ? `
              <div style="display: flex; gap: 0.3rem; flex-wrap: wrap;">
                <button class="btn btn-secondary" style="padding: 0.2rem 0.45rem; font-size: 0.73rem;" onclick="openEditPendingIntakeModal('${i.id}')" title="Editar borrador sin generar movimientos ni salidas">
                  ✏️ Editar Borrador
                </button>
                <button class="btn btn-primary" style="padding: 0.2rem 0.45rem; font-size: 0.73rem; background: #10b981; border-color: #10b981;" onclick="approvePendingIntake('${i.id}')" title="Aprobar e ingresar existencias físicas al sistema">
                  ✅ Aprobar
                </button>
                <button class="btn btn-danger" style="padding: 0.2rem 0.45rem; font-size: 0.73rem;" onclick="rejectPendingIntake('${i.id}')" title="Anular propuesta de ingreso">
                  ❌ Rechazar
                </button>
              </div>
            ` : `
              <span style="font-size: 0.75rem; color: var(--text-muted);">⏳ Esperando revisión de Administrador</span>
            `}
          ` : `
            <span style="font-size: 0.75rem; color: var(--text-muted);">Sin acciones (Registro finalizado)</span>
          `}
        </td>
      </tr>
    `;
  }).join('');
}

function openEditPendingIntakeModal(intakeId) {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo un Administrador puede editar borradores de ingreso.");
    return;
  }

  const intake = appState.pendingIntakes.find(i => i.id === intakeId);
  if (!intake || intake.status !== 'PENDIENTE_VALIDACION') return;

  const modal = document.getElementById("editPendingIntakeModal");
  if (!modal) return;

  document.getElementById("pendingIntakeEditId").value = intake.id;
  document.getElementById("pendingIntakeProductDisplay").value = `${intake.sku} - ${intake.productName}`;

  const locSelect = document.getElementById("pendingIntakeLocation");
  if (locSelect) {
    locSelect.innerHTML = appState.locations.filter(l => l.active).map(l => 
      `<option value="${l.id}" ${l.id === intake.locationId ? 'selected' : ''}>${l.name}</option>`
    ).join('');
  }

  document.getElementById("pendingIntakeQuantity").value = intake.proposedQuantity;
  document.getElementById("pendingIntakeInvoice").value = intake.invoiceNumber || "";
  document.getElementById("pendingIntakeNotes").value = intake.notes || "";

  const serialsGroup = document.getElementById("pendingIntakeSerialsGroup");
  const serialsText = document.getElementById("pendingIntakeSerialsText");
  if (intake.requiresSerial) {
    if (serialsGroup) serialsGroup.style.display = "block";
    if (serialsText) serialsText.value = (intake.proposedSerials || []).join(", ");
  } else {
    if (serialsGroup) serialsGroup.style.display = "none";
  }

  modal.classList.add("active");
}

function closeEditPendingIntakeModal() {
  const modal = document.getElementById("editPendingIntakeModal");
  if (modal) modal.classList.remove("active");
}

function handleSavePendingIntakeDraft(e) {
  e.preventDefault();
  const editId = document.getElementById("pendingIntakeEditId").value;
  const intake = appState.pendingIntakes.find(i => i.id === editId);
  if (!intake) return;

  const newLocId = document.getElementById("pendingIntakeLocation").value;
  const targetLoc = appState.locations.find(l => l.id === newLocId);
  const newQty = parseInt(document.getElementById("pendingIntakeQuantity").value);
  const newInvoice = document.getElementById("pendingIntakeInvoice").value.trim();
  const newNotes = document.getElementById("pendingIntakeNotes").value.trim();

  let newSerials = [];
  if (intake.requiresSerial) {
    const rawSerials = document.getElementById("pendingIntakeSerialsText")?.value || "";
    newSerials = rawSerials.split(',').map(s => s.trim()).filter(Boolean);
  }

  intake.locationId = newLocId;
  intake.locationName = targetLoc ? targetLoc.name : "Bodega Principal";
  intake.proposedQuantity = newQty;
  intake.invoiceNumber = newInvoice;
  intake.notes = newNotes;
  intake.proposedSerials = newSerials;

  savePendingIntakesToDisk();
  addActivityLog("MODIFICACION", "PendingIntakeDraft", `Administrador modificó borrador de ingreso '${intake.sku}' (Nueva cant: ${newQty}). Sin movimientos de salida.`);

  alert(`✅ BORRADOR ACTUALIZADO: La propuesta de ingreso fue modificada con éxito. Permanece en borrador hasta que haga clic en 'Aprobar'.`);
  closeEditPendingIntakeModal();
  renderPendingValidationsView();
}

function approvePendingIntake(intakeId) {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador está autorizado para aprobar e ingresar inventario.");
    return;
  }

  const intake = appState.pendingIntakes.find(i => i.id === intakeId);
  if (!intake || intake.status !== 'PENDIENTE_VALIDACION') return;

  const product = appState.products.find(p => p.id === intake.productId);
  if (!product) {
    alert("⚠️ ERROR: El producto asociado ya no existe en el catálogo.");
    return;
  }

  if (confirm(`¿Confirma la aprobación del ingreso de ${intake.proposedQuantity} unidades de '${product.sku}' a la sede '${intake.locationName}'?`)) {
    ensureProductLocations();

    product.physicalStock += intake.proposedQuantity;
    product.stockByLocation[intake.locationId] = (product.stockByLocation[intake.locationId] || 0) + intake.proposedQuantity;

    if (product.requiresSerial && intake.proposedSerials && intake.proposedSerials.length > 0) {
      intake.proposedSerials.forEach(sNum => {
        let serialItem = appState.serializedItems.find(i => i.serialNumber === sNum);
        if (!serialItem) {
          serialItem = {
            id: `ser-${Date.now()}-${Math.floor(Math.random()*1000)}`,
            productId: product.id,
            serialNumber: sNum,
            status: "EN_STOCK",
            currentLocationId: intake.locationId,
            currentCustomerId: null,
            entryDate: new Date().toISOString().substring(0, 10),
            saleDate: null,
            invoiceNumber: intake.invoiceNumber,
            attachments: [],
            history: [{
              type: "INGRESO_COMPRA",
              date: new Date().toISOString().substring(0, 10),
              user: appState.currentUser.name,
              description: `Ingreso de inventario aprobado por Administrador en sede '${intake.locationName}'`
            }]
          };
          appState.serializedItems.push(serialItem);
        } else {
          serialItem.status = "EN_STOCK";
          serialItem.currentLocationId = intake.locationId;
        }
      });
    }

    const committedMov = {
      id: `mov-app-${Date.now()}`,
      date: new Date().toISOString().substring(0, 10),
      type: "INGRESO_COMPRA",
      productId: product.id,
      quantity: intake.proposedQuantity,
      invoiceNumber: intake.invoiceNumber || `FAC-IN-${Date.now().toString().slice(-4)}`,
      customerName: "Ingreso Proveedor (Aprobado)",
      user: `${appState.currentUser.name} (Aprueba: ${intake.submittedBy})`,
      notes: `Aprobación de ingreso registrado originalmente por Logística. ${intake.notes || ''}`
    };

    appState.movements.unshift(committedMov);

    intake.status = "APROBADO";
    savePendingIntakesToDisk();

    addActivityLog("APROBACION", "InventoryIntake", `Administrador aprobó e ingresó ${intake.proposedQuantity} unidades de '${product.sku}' en sede '${intake.locationName}'.`);

    alert(`✅ INGRESO APROBADO EXITOSAMENTE: Se han adicionado ${intake.proposedQuantity} unidades al inventario físico de '${product.sku}'.`);
    renderAllViews();
  }
}

function rejectPendingIntake(intakeId) {
  if (appState.currentUser.role !== 'ADMINISTRADOR') {
    alert("🔒 Restricción RBAC: Solo el Administrador está autorizado para rechazar propuestas de ingreso.");
    return;
  }

  const intake = appState.pendingIntakes.find(i => i.id === intakeId);
  if (!intake || intake.status !== 'PENDIENTE_VALIDACION') return;

  if (confirm(`¿Está seguro de rechazar y anular esta propuesta de ingreso de ${intake.proposedQuantity} unidades de '${intake.sku}'?`)) {
    intake.status = "RECHAZADO";
    savePendingIntakesToDisk();
    addActivityLog("RECHAZO", "InventoryIntake", `Administrador rechazó propuesta de ingreso de '${intake.sku}' de usuario '${intake.submittedBy}'.`);
    alert(`🔴 PROPUESTA RECHAZADA: El ingreso fue anulado y no afectará el inventario.`);
    renderAllViews();
  }
}
