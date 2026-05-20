/** Márgenes de cálculo (mismo criterio que los LISP de AutoCAD). */
const MARGIN_EMITTERS = 1.03;  // 3 % extra para emisores y rollos
const MARGIN_DENSITY  = 1.02;  // 2 % extra para densidad de plantación

// =============================================================================
// ORDEN DE FACETAS POR CATEGORÍA
// Define en qué orden se presentan los selectores en cascada en el cliente.
// Solo se listan los campos que SÍ deben aparecer como filtro; el resto son
// informativos (descripcion, unidad, long_rollo_m, requiere_conector_wr…).
// =============================================================================

const CATEGORY_FACETS = {
  cintas:          ['marca', 'calibre_mil', 'sep_emisor_m'],
  goteros:         ['caudal_lph'],
  microaspersores: ['tipo', 'marca', 'caudal_lph', 'variante']
};

// =============================================================================
// METADATA DE FACETAS
// Controla etiquetas y formato de los valores en el cliente.
// prefix/suffix se aplican al valor visual, no al value del <option>.
// wide = true indica que el select necesita más espacio horizontal.
// =============================================================================

const FACETS_META = {
  tipo:         { label: 'Tipo' },
  marca:        { label: 'Marca' },
  caudal_lph:   { label: 'Caudal',      suffix: ' lph' },
  variante:     { label: 'Modelo',      wide: true },
  calibre_mil:  { label: 'Calibre',     suffix: ' mil' },
  sep_emisor_m: { label: 'Sep. emisor', suffix: ' m' },
  diametro_mm:  { label: 'Diámetro',   prefix: 'Ø', suffix: ' mm' }
};

// =============================================================================
// ACCESORIOS DE MICROASPERSIÓN
// Piezas que se agregan automáticamente por cada microaspersor calculado.
// conector_wr solo se incluye cuando el producto tiene requiere_conector_wr=TRUE.
// =============================================================================

const IRRIGATION_ACCESSORIES = {
  tubing: {
    id:          'tubing_plastro_7mm',
    description: 'TUBING POLIETILENO PLASTRO Ø7 mm x 1 m ®RIVULIS ×INSERCIÓN',
    unit:        'pza',
    per_emitter: 1
  },
  stake: {
    id:          'stake_plastic_5_8',
    description: 'ESTACA PLÁSTICO ø5/8 Pulg ×38 cm',
    unit:        'pza',
    per_emitter: 1
  },
  wr_connector: {
    id:          'wr_connector_7mm',
    description: 'CONECTOR INSERCIÓN (B) Ø7 mm ®WR ×Liso',
    unit:        'pza',
    per_emitter: 1
  }
};

/**
 * Calcula los materiales de un proyecto de emisores de riego.
 * Punto de entrada llamado desde el cliente vía google.script.run.
 *
 * @param {Object} project - Objeto del proyecto enviado por el cliente.
 * @param {Object} project.metadata - Metadatos: { system: 'goteo_linea' | 'microaspersion' | 'goteo_insertable' }
 * @param {Array}  project.sub_areas - Lista de sub-áreas a calcular.
 * @returns {Object} Resultado con resumen y detalle por sub-área.
 */
function calculateProject(project_){

  // Calculamos la lista de materiales
  const calculator = new EmitterCalculator_(project_);
  // throw JSON.stringify(calculator.calculate())
  // Retornamos los respectivo datos
  return calculator.calculate();
}

/**
 * Permite obtener los parámetros de configuración del módulo de emisores
 */
function getEmitterParams() {
  // Variable a retornar
  let paramsObject_ = {
    irrigationArray: [],
    fieldsQualifier: {},
    catalog: null
  };

  // Obtenemos los parámetros globales
  const resources_ = getResources();
  // Referenciamos la hoja de cálculo
  const spreadsheet_ = SpreadsheetApp.openById(resources_.crossingSheetId);

  // Referenciamos las pestañas de sistema de riego y campos del cualificador
  const irrigationSheet_ = spreadsheet_.getSheetByName(resources_.irrigationSheetName);
  const qualifySheet_    = spreadsheet_.getSheetByName(resources_.qualifySheetName);

  // Validamos que exista la pestaña de sistemas de riego
  if (irrigationSheet_) {
    paramsObject_.irrigationArray = consultIrrigationValues_(irrigationSheet_);
  }

  // Validamos que exista la pestaña de campos del cualificador
  if (qualifySheet_) {
    paramsObject_.fieldsQualifier = consultQualifierFields_(qualifySheet_);
  }

  // Obtenemos los datos del catalogo
  paramsObject_.catalog = getCatalogEmitter_(spreadsheet_, resources_);

  // Retornamos los parámetros
  return paramsObject_;
}

/**
 * Construye y devuelve el catálogo completo al cliente.
 * Combina los productos leídos desde la hoja con las facetas y accesorios
 * definidos en el código, para no exponer constantes internas innecesarias.
 */
function getCatalogEmitter_(spreadsheet_, resources_) {
  const catalogObject_ = {};

  // Definimos la lista de catarias
  const categories_ = [
    {sheetName: resources_.thongSheetName, category: "cintas"},
    {sheetName: resources_.microSprinklerSheetName, category: "microaspersores"},
    {sheetName: resources_.dropperSheetName, category: "goteros"},
  ];

  // recorremos cada una de las categorias
  for(let i = 0; i < categories_.length; i++){
    let recordObject_ = categories_[i];

    // Agregamos los datos y agruarlos
    catalogObject_[recordObject_.category] = readCatalogTab_(spreadsheet_, recordObject_.sheetName);
    catalogObject_[recordObject_.category + '_facets'] = CATEGORY_FACETS[recordObject_.category] || [];

  }

  // Variables de las facetas
  catalogObject_.facets_meta = FACETS_META;
  catalogObject_.irrigation_accessories = IRRIGATION_ACCESSORIES;

  return catalogObject_;
}

/**
 * Lee una pestaña de la hoja CATALOG y devuelve un array de objetos.
 * La primera fila se trata como encabezado; las siguientes como datos.
 * Las celdas vacías se omiten (se devuelve el valor tal cual, puede ser '').
 * Los valores numéricos quedan como Number; los booleanos como Boolean.
 */
function readCatalogTab_(spreadsheet_, sheetName_) {
  const sheet_  = spreadsheet_.getSheetByName(sheetName_);

  if (!sheet_) throw new Error('No se encontró la pestaña "' + sheetName_ + '" en la hoja CATALOG.');

  const data    = sheet_.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).trim(); });
  const rows    = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    // Ignorar filas completamente vacías
    if (row.every(function(c) { return c === '' || c == null; })) continue;
    const obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
    });
    rows.push(obj);
  }

  return rows;
}

/**
 * Lee la pestaña "Sistema de riego" y retorna un array de objetos
 * Columnas: ID | Nombre | Descripción
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet_
 * @returns {Array<{id: string, name: string, description: string}>}
 */
function consultIrrigationValues_(sheet_) {
  // Variable a retornar
  const result_ = [];
  // Obtenemos la última fila con datos
  const lastRow_ = sheet_.getLastRow();

  // Validamos que exista más de 1 fila (encabezado + datos)
  if (lastRow_ < 2) return result_;

  // Obtenemos los datos desde la fila 2 hasta la última, en 4 columnas
  const data_ = sheet_.getRange(2, 1, lastRow_ - 1, 4).getValues();

  // Recorremos los datos
  for (let i = 0; i < data_.length; i++) {
    const row_ = data_[i];
    // Saltamos filas donde ID y Nombre estén vacíos
    if (!row_[0] && !row_[1]) continue;

    // Construimos el objeto y lo agregamos al resultado
    result_.push({
      id:          String(row_[0]).trim(), // Columna A: ID del sistema
      name:        String(row_[1]).trim(), // Columna B: Nombre del sistema
      description: String(row_[2]).trim(), // Columna C: Descripción del sistema
      category:    String(row_[3]).trim(), // Columna D: Categoría
    });
  }

  // Retornamos el array de sistemas de riego
  return result_;
}


/**
 * Lee la pestaña "Campos del cualificador" y retorna el objeto EMITTER_FIELDS
 * Columnas: ID | Sistema | Nombre | Tipo | Requerido | Ancho | Ayuda | Opciones
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet_
 * @returns {Object.<string, Array>}
 */
function consultQualifierFields_(sheet_) {
  // Variable a retornar
  const result_ = {};
  // Obtenemos la última fila con datos
  const lastRow_ = sheet_.getLastRow();

  // Validamos que exista más de 1 fila (encabezado + datos)
  if (lastRow_ < 2) return result_;

  // Obtenemos los datos desde la fila 2, en 8 columnas
  const data_ = sheet_.getRange(2, 1, lastRow_ - 1, 8).getValues();

  // Recorremos los datos
  for (let i = 0; i < data_.length; i++) {
    const row_ = data_[i];

    // Mapeamos cada columna a su variable correspondiente
    const fieldId_    = String(row_[0]).trim(); // Columna A: ID del campo
    const sistema_    = String(row_[1]).trim(); // Columna B: Sistema de riego al que pertenece
    const label_      = String(row_[2]).trim(); // Columna C: Nombre visible del campo
    const type_       = String(row_[3]).trim(); // Columna D: Tipo de campo (text | number | select)
    const required_   = row_[4];               // Columna E: Requerido (checkbox TRUE/FALSE o texto)
    const width_      = String(row_[5]).trim(); // Columna F: Ancho del campo (half | full)
    const hint_       = String(row_[6]).trim(); // Columna G: Texto de ayuda opcional
    const optionsRaw_ = String(row_[7]).trim(); // Columna H: Opciones separadas por coma (solo para select)

    // Saltamos filas sin los datos mínimos requeridos
    if (!fieldId_ || !sistema_ || !label_) continue;

    // Inicializamos el array del sistema si aún no existe en el objeto
    if (!result_[sistema_]) {
      result_[sistema_] = [];
    }

    // Construimos el objeto base del campo
    const fieldObj_ = {
      id:       fieldId_,
      label:    label_,
      type:     type_,
      // Soportamos checkbox de GS (TRUE/FALSE) y texto ("true"/"false")
      required: required_ === true || String(required_).toLowerCase() === 'true',
      width:    width_,
    };

    // Agregamos hint solo si tiene valor
    if (hint_) {
      fieldObj_.hint = hint_;
    }

    // Agregamos options solo si el tipo es select y existen valores en la celda
    if (type_ === 'select' && optionsRaw_) {
      fieldObj_.options = optionsRaw_.split(',').map(o => o.trim()).filter(o => o !== '');
    }

    // Agregamos el campo al array del sistema correspondiente
    result_[sistema_].push(fieldObj_);
  }

  // Retornamos el objeto con los campos agrupados por sistema de riego
  return result_;
}

/**
 * @class EmitterCalculator_
 * @classdesc Orquesta el cálculo de materiales para un proyecto de riego.
 * Carga el catálogo una sola vez y procesa cada sub-área según su sistema.
 */
class EmitterCalculator_ {

  /**
   * @param {Object} project - Proyecto recibido del cliente.
   */
  constructor(project) {
    if (!project || !project.sub_areas || project.sub_areas.length === 0) {
      throw new Error('No hay sub-áreas para calcular.');
    }
    if (!project.metadata || !project.metadata.system) {
      throw new Error('No se especificó el sistema de riego.');
    }

    /** @type {string} Sistema activo del proyecto */
    this.system = project.metadata.system;

    /** @type {Array} Lista de sub-áreas */
    this.subAreas = project.sub_areas;

    /** @type {Object} Metadatos del proyecto */
    this.metadata = project.metadata;

    // Catálogo cargado fresco desde la hoja (captura cambios recientes)
    const resources_ = getResources();
    const spreadsheet_ = SpreadsheetApp.openById(resources_.crossingSheetId);
    this.catalog = getCatalogEmitter_(spreadsheet_, resources_);
  }

  /**
   * Ejecuta el cálculo completo del proyecto.
   * @returns {Object} Objeto con detalle por sub-área y resumen de materiales.
   */
  calculate() {
    const detail      = [];
    const accumulator = {};

    this.subAreas.forEach((sa, idx) => {
      const row = this._processSubArea(sa, idx);
      detail.push(row);
      row.materials.forEach(m => this._accumulate(accumulator, m, sa.lot));
    });

    const summary = this._buildSummary(accumulator);
    const totalArea = this.subAreas.reduce((s, sa) => s + (parseFloat(sa.superficie_ha) || 0), 0);

    return {
      system:            this.system,
      metadata:          this.metadata,
      total_area_ha:     totalArea,
      sub_area_detail:   detail,
      materials_summary: summary,
      calculated_at:     new Date().toISOString()
    };
  }

  /**
   * Procesa una sub-área individual y devuelve su fila de detalle.
   * @param {Object} sa  - Sub-área del proyecto.
   * @param {number} idx - Índice base 0.
   * @returns {Object} Fila con métricas y materiales calculados.
   */
  _processSubArea(sa, idx) {
    const areaHa = parseFloat(sa.superficie_ha) || 0;
    const row = {
      index:     idx + 1,
      lot:       sa.lot || '(sin lote)',
      area_ha:   areaHa,
      area_m2:   areaHa * 10000,
      system:    this.system,
      materials: [],
      metrics:   {}
    };

    switch (this.system) {
      case 'goteo_linea':
        this._calcDripLine(sa, row);
        break;
      case 'microaspersion':
        this._calcMicrosprinkler(sa, row);
        break;
      case 'goteo_insertable':
        this._calcInsertableDrip(sa, row);
        break;
      default:
        throw new Error('Sistema no reconocido: ' + this.system);
    }
    return row;
  }

  // --- CÁLCULO POR SISTEMA ------------------------------------------------

  /**
   * Calcula materiales para goteo en línea (cintas/mangueras con gotero).
   * Produce rollos según longitud total con margen del 3 %.
   *
   * @param {Object} sa  - Sub-área: { lot, superficie_ha, sep_surcos, sep_plantas?, regantes_surco, producto_id }
   * @param {Object} row - Fila de resultado que se modifica en lugar.
   */
  _calcDripLine(sa, row) {
    const rowSpacing   = parseFloat(sa.sep_surcos);
    const plantSpacing = parseFloat(sa.sep_plantas) || 0;
    const drippersRow  = parseFloat(sa.regantes_surco);
    const product      = this._findProduct('cintas', sa.producto_id);

    if (!product)            throw new Error(`Lote ${row.lot}: producto de cinta no encontrado.`);
    if (!(rowSpacing > 0))   throw new Error(`Lote ${row.lot}: sep_surcos inválida.`);
    if (!(drippersRow > 0))  throw new Error(`Lote ${row.lot}: regantes_surco inválido.`);

    const lateralSpacing = rowSpacing / drippersRow;
    const emitterArea    = lateralSpacing * product.sep_emisor_m;
    const nEmitters      = this._ceilUp((row.area_m2 / emitterArea) * MARGIN_EMITTERS);
    const totalLengthM   = nEmitters * product.sep_emisor_m;
    const nRolls         = this._ceilUp(totalLengthM / product.long_rollo_m);

    row.metrics = {
      sep_surcos_m:   rowSpacing,
      sep_plantas_m:  plantSpacing || null,
      regantes_surco: drippersRow,
      sep_lateral_m:  this._round2(lateralSpacing),
      sep_emisor_m:   product.sep_emisor_m,
      caudal_lph:     product.caudal_lph,
      n_emisores:     nEmitters,
      long_total_m:   this._round2(totalLengthM),
      long_rollo_m:   product.long_rollo_m,
      n_rollos:       nRolls
    };

    row.materials.push({
      product_id:  product.id,
      description: product.descripcion,
      unit:        product.unidad,
      quantity:    nRolls
    });
  }

  /**
   * Calcula materiales para microaspersión.
   * Incluye automáticamente tubing, estaca y conector WR si aplica.
   *
   * @param {Object} sa  - Sub-área: { lot, superficie_ha, sep_surcos, sep_plantas?, regantes_surco, sep_emisor, producto_id }
   * @param {Object} row - Fila de resultado que se modifica en lugar.
   */
  _calcMicrosprinkler(sa, row) {
    const rowSpacing     = parseFloat(sa.sep_surcos);
    const plantSpacing   = parseFloat(sa.sep_plantas) || 0;
    const drippersRow    = parseFloat(sa.regantes_surco);
    const emitterSpacing = parseFloat(sa.sep_emisor);
    const product        = this._findProduct('microaspersores', sa.producto_id);

    if (!product)               throw new Error(`Lote ${row.lot}: microaspersor no encontrado.`);
    if (!(rowSpacing > 0))      throw new Error(`Lote ${row.lot}: sep_surcos inválida.`);
    if (!(drippersRow > 0))     throw new Error(`Lote ${row.lot}: regantes_surco inválido.`);
    if (!(emitterSpacing > 0))  throw new Error(`Lote ${row.lot}: sep_emisor inválida.`);

    const lateralSpacing = rowSpacing / drippersRow;
    const emitterArea    = lateralSpacing * emitterSpacing;
    const nEmitters      = this._ceilUp((row.area_m2 / emitterArea) * MARGIN_EMITTERS);

    row.metrics = {
      sep_surcos_m:    rowSpacing,
      sep_plantas_m:   plantSpacing || null,
      regantes_surco:  drippersRow,
      sep_lateral_m:   this._round2(lateralSpacing),
      sep_emisor_m:    emitterSpacing,
      caudal_lph:      product.caudal_lph,
      n_emisores:      nEmitters,
      requiere_wr:     !!product.requiere_conector_wr
    };

    const acc = IRRIGATION_ACCESSORIES;
    row.materials.push({ product_id: product.id,         description: product.descripcion,         unit: product.unidad,      quantity: nEmitters });
    row.materials.push({ product_id: acc.tubing.id,      description: acc.tubing.description,      unit: acc.tubing.unit,     quantity: nEmitters });
    row.materials.push({ product_id: acc.stake.id,       description: acc.stake.description,       unit: acc.stake.unit,      quantity: nEmitters });

    if (product.requiere_conector_wr) {
      row.materials.push({ product_id: acc.wr_connector.id, description: acc.wr_connector.description, unit: acc.wr_connector.unit, quantity: nEmitters });
    }
  }

  /**
   * Calcula materiales para goteo insertable (E1000 RIVULIS).
   * Cantidad = densidad de plantas × goteros por planta.
   *
   * @param {Object} sa  - Sub-área: { lot, superficie_ha, sep_surcos, sep_plantas, goteros_planta, producto_id }
   * @param {Object} row - Fila de resultado que se modifica en lugar.
   */
  _calcInsertableDrip(sa, row) {
    const rowSpacing    = parseFloat(sa.sep_surcos);
    const plantSpacing  = parseFloat(sa.sep_plantas);
    const drippersPlant = parseInt(sa.goteros_planta, 10);
    const product       = this._findProduct('goteros', sa.producto_id);

    if (!product)               throw new Error(`Lote ${row.lot}: gotero no encontrado.`);
    if (!(rowSpacing > 0))      throw new Error(`Lote ${row.lot}: sep_surcos inválida.`);
    if (!(plantSpacing > 0))    throw new Error(`Lote ${row.lot}: sep_plantas inválida.`);
    if (!(drippersPlant >= 1))  throw new Error(`Lote ${row.lot}: goteros_planta debe ser ≥ 1.`);

    const plantDensity = this._ceilUp((row.area_m2 / (rowSpacing * plantSpacing)) * MARGIN_DENSITY);
    const nDrippers    = plantDensity * drippersPlant;

    row.metrics = {
      sep_surcos_m:   rowSpacing,
      sep_plantas_m:  plantSpacing,
      densidad_plantas: plantDensity,
      goteros_planta: drippersPlant,
      caudal_lph:     product.caudal_lph,
      n_goteros:      nDrippers
    };

    row.materials.push({ product_id: product.id, description: product.descripcion, unit: product.unidad, quantity: nDrippers });
  }

  // --- HELPERS PRIVADOS ---------------------------------------------------

  /**
   * Busca un producto por ID dentro de una categoría del catálogo.
   * @param {string} category  - 'cintas' | 'microaspersores' | 'goteros'
   * @param {string} productId - ID del producto.
   * @returns {Object|null} Producto encontrado o null.
   */
  _findProduct(category, productId) {
    if (!productId) return null;
    return (this.catalog[category] || []).find(p => p.id === productId) || null;
  }

  /**
   * Acumula cantidades de un material en el objeto acumulador.
   * @param {Object} acc      - Acumulador de materiales.
   * @param {Object} material - Material a acumular.
   * @param {string} lot      - Nombre del lote (para trazabilidad).
   */
  _accumulate(acc, material, lot) {
    if (!acc[material.product_id]) {
      acc[material.product_id] = { description: material.description, unit: material.unit, quantity: 0, lots: new Set() };
    }
    acc[material.product_id].quantity += material.quantity;
    if (lot) acc[material.product_id].lots.add(String(lot));
  }

  /**
   * Construye el array de resumen de materiales ordenado (rollos primero, luego pzas).
   * @param {Object} accumulator - Acumulador con todos los materiales.
   * @returns {Array} Lista ordenada para presentar al usuario.
   */
  _buildSummary(accumulator) {
    return Object.keys(accumulator).map(id => {
      const r = accumulator[id];
      return { product_id: id, description: r.description, unit: r.unit, quantity: r.quantity, lots: Array.from(r.lots).sort() };
    }).sort((a, b) => {
      if (a.unit === 'rollo' && b.unit !== 'rollo') return -1;
      if (b.unit === 'rollo' && a.unit !== 'rollo') return  1;
      return a.description.localeCompare(b.description, 'es');
    });
  }

  /**
   * Redondea hacia arriba corrigiendo errores de punto flotante.
   * @param {number} num
   * @returns {number}
   */
  _ceilUp(num) {
    return Math.abs(num - Math.round(num)) < 1e-8 ? Math.round(num) : Math.ceil(num);
  }

  /** Redondea a 2 decimales. @param {number} num @returns {number} */
  _round2(num) { return Math.round(num * 100) / 100; }
}
