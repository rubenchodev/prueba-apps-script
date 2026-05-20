
/**
 * Clase quue permite generar la lista de materiales requeridos para los proyectos asociadodos
 */
class CruceroSystem {
  constructor(resources_) {
    this.resources = resources_;

    // Parametros de consumibles
    this.CEMENT_CAN_LITERS = 0.48;
    this.CLEANER_PER_LITER_CEMENT = 0.6;
    this.CLEANER_PRESENTATION_L = 0.5;

    // Equivalencias de mm a pulgadas decimales
    this.DIAMETER_EQUIVALENCES = {
      13: 0.5, 19: 0.75, 25: 1, 32: 1.25, 38: 1.5,
      50: 2, 63: 2.5, 75: 3, 100: 4, 150: 6,
      200: 8, 250: 10, 315: 12, 355: 14
    };

    // Diametros comerciales para validacion de saltos
    this.COMMERCIAL_DIAMETERS = [0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4, 6, 8, 10, 12, 14];

    // Dimensionamiento de VAE segun diametro principal
    this.VAE_BY_DIAMETER = {
      1: 1, 1.25: 1, 1.5: 1, 2: 1,
      3: 2, 4: 2, 6: 2,
      8: 3, 10: 3, 12: 4
    };

    // Configuracion de tornilleria por diametro de brida
    this.FLANGE_BOLTS = {
      2: { qty: 4, diameter: '5/8', length: '4-1/2' },
      3: { qty: 4, diameter: '5/8', length: '4-1/2' },
      4: { qty: 8, diameter: '5/8', length: '5' },
      6: { qty: 8, diameter: '3/4', length: '5' },
      8: { qty: 8, diameter: '3/4', length: '5' },
      10: { qty: 8, diameter: '7/8', length: '5' },
      12: { qty: 8, diameter: '7/8', length: '5' }
    };

    // Rendimiento de uniones por cada lata de 0.48L
    this.CEMENT_YIELD = {
      13: 157, 19: 95, 25: 67, 32: 48, 38: 38,
      50: 22, 60: 17, 75: 10, 100: 8, 150: 3, 200: 2
    };

    // Cantidad de uniones cementadas generadas por pieza
    this.JOINTS_BY_PIECE = {
      'CODO 90': 2, 'TEE': 3, 'CRUZ': 4,
      'Reducción Bush': 2, 'Adaptador': 1, 'Brida': 1
    };

    this.CONCEPT_CATALOG = this._loadConceptCatalog();
  }

  /**
   * Carga el catalogo de conceptos desde la hoja de Google Sheets
   */
  _loadConceptCatalog() {
    const spreadsheet = SpreadsheetApp.openById(this.resources.crossingSheetId);
    const sheet = spreadsheet.getSheetByName(this.resources.refConceptSheetName);
    if (!sheet) throw new Error('No existe la hoja de referencia de conceptos');

    const data = sheet.getDataRange().getValues();
    let catalog = {};
    for (let i = 1; i < data.length; i++) {
      const [desc, concept] = data[i];
      if (desc && concept) catalog[desc.trim()] = concept.trim();
    }
    return catalog;
  }

  /**
   * Extrae valores de diametro o longitud de una descripcion usando Regex
   */
  _extractValues(pattern, description) {
    const keys = [];
    const regexStr = pattern
      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      .replace(/\\\{(D2|D|L)\\\}/g, (_, key) => {
        keys.push(key);
        return '([\\d/\\-]+)';
      });

    const match = description.match(new RegExp(`^${regexStr}$`));
    if (!match) return null;

    let values = {};
    keys.forEach((k, i) => { values[k] = match[i + 1]; });
    return values;
  }

  /**
   * Mapea la descripcion tecnica a un concepto comercial del catalogo
   */
  _resolveConcept(description) {
    if (this.CONCEPT_CATALOG[description]) return this.CONCEPT_CATALOG[description];

    for (const pattern in this.CONCEPT_CATALOG) {
      const values = this._extractValues(pattern, description);
      if (values) {
        let conceptTemplate = this.CONCEPT_CATALOG[pattern];
        Object.keys(values).forEach(key => {
          conceptTemplate = conceptTemplate.replace(new RegExp(`\\{${key}\\}`, 'g'), values[key]);
        });
        return conceptTemplate;
      }
    }
    return '';
  }

  /**
   * Metodo principal que procesa la lista de configuraciones y retorna materiales
   */
  generateMaterials(projectConfigs) {
    try {
      let allMaterials = [];
      projectConfigs.forEach(cfg => {
        const mapped = this._mapConfig(cfg);
        allMaterials.push(...this._generateCrossing(mapped));
      });

      const grouped = this._groupMaterials(allMaterials);
      grouped.sort((a, b) => a.type !== b.type ? a.type.localeCompare(b.type) : a.description.localeCompare(b.description));

      grouped.forEach(item => { item.concept = this._resolveConcept(item.description); });

      return { success: true, materials: grouped };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  /**
   * Normaliza los nombres de campos y tipos de datos del objeto de entrada
   */
  _mapConfig(cfg) {
    return {
      crossingType: cfg.crossingType,
      vaeConfig: cfg.vaeConfig,
      valveType: cfg.valveType,
      connectionType: cfg.connectionType || '',
      valveDiameter: parseFloat(cfg.valveDiameter),
      inletDiameterMm: parseInt(cfg.inletValue, 10),
      outletDiameterMm: parseInt(cfg.outletValue, 10),
      trenchDepth: cfg.trenchDepth,
      crossingHeight: cfg.crossingHeight,
      count: cfg.count || 1
    };
  }

  /**
   * Convierte milimetros a pulgadas decimales segun tabla interna
   */
  _mmToInches(mm) { return this.DIAMETER_EQUIVALENCES[mm] || null; }

  /**
   * Obtiene el valor en mm buscando por la pulgada decimal
   */
  _inchesToMm(inches) {
    const mm = Object.keys(this.DIAMETER_EQUIVALENCES).find(k => this.DIAMETER_EQUIVALENCES[k] === inches);
    return mm ? parseInt(mm, 10) : null;
  }

  /**
   * Formatea numeros decimales a representaciones de fraccion comunes
   */
  _formatDiameter(d) {
    const formats = { 0.5: '1/2', 0.75: '3/4', 1.25: '1-1/4', 1.5: '1-1/2', 2.5: '2-1/2' };
    return formats[d] || d.toString();
  }

  /**
   * Valida que los saltos de diametro entre componentes sean tecnicamente viables
   */
  _validateConfig(config) {
    const inletIn = this._mmToInches(config.inletDiameterMm);
    const outletIn = this._mmToInches(config.outletDiameterMm);
    const valveIn = config.valveDiameter;

    if (!this._validateDiameterJumps(inletIn, valveIn, 3, 'down')) {
      return { valid: false, message: `El salto Entrada-Válvula excede 3 tamaños comerciales.` };
    }
    if (!this._validateDiameterJumps(valveIn, outletIn, 2, 'up')) {
      return { valid: false, message: `El salto Válvula-Salida excede 2 tamaños comerciales.` };
    }
    return { valid: true };
  }

  /**
   * Compara posiciones en el arreglo comercial para determinar la magnitud del salto
   */
  _validateDiameterJumps(from, to, max, dir) {
    const idxFrom = this.COMMERCIAL_DIAMETERS.indexOf(from);
    const idxTo = this.COMMERCIAL_DIAMETERS.indexOf(to);
    if (idxFrom === -1 || idxTo === -1) return false;
    const diff = dir === 'down' ? (idxFrom - idxTo) : (idxTo - idxFrom);
    return diff >= 0 && diff <= max;
  }

  /**
   * Coordina la generacion de todas las piezas que componen un crucero
   */
  _generateCrossing(config) {
    const validation = this._validateConfig(config);
    if (!validation.valid) throw new Error(validation.message);

    let materials = [];
    let joints = {};

    const inletIn = this._mmToInches(config.inletDiameterMm);
    const outletIn = this._mmToInches(config.outletDiameterMm);
    const valveIn = config.valveDiameter;

    // Generar pieza de distribucion principal
    const central = this._getCentralPiece(config.crossingType, config.vaeConfig, inletIn);
    materials.push(central);
    this._countJoints(central, joints, config.inletDiameterMm);

    // Calcular tuberia vertical de entrada
    const vLen = (config.trenchDepth + config.crossingHeight) / 100;
    materials.push({
      description: `TUBO PVC HIDRÁULICO C/B ø${this._formatDiameter(inletIn)} Pulg ⒸRD-26`,
      unit: 'MT', qty: vLen, type: 'Tubo'
    });

    // Procesar ramales de valvulas
    const numBranches = config.crossingType === 'Doble' ? 2 : 1;
    for (let i = 0; i < numBranches; i++) {
      const br = this._processBranch(inletIn, valveIn, outletIn, config.valveType, config.connectionType);
      materials.push(...br.materials);
      Object.keys(br.joints).forEach(d => { joints[d] = (joints[d] || 0) + br.joints[d]; });

      // Definir accesorio final segun presencia de VAE
      const finalType = String(config.vaeConfig).includes('S') ? 'TEE' : 'CODO 90';
      materials.push({
        description: `${finalType} PVC HIDRÁULICO CEM ø${this._formatDiameter(outletIn)} Pulg ⒸCED 40`,
        unit: 'PZA', qty: 1, type: finalType
      });
      joints[config.outletDiameterMm] = (joints[config.outletDiameterMm] || 0) + (finalType === 'TEE' ? 3 : 2);
    }

    // Tuberia vertical de salida
    materials.push({
      description: `TUBO PVC HIDRÁULICO C/B ø${this._formatDiameter(outletIn)} Pulg ⒸRD-26`,
      unit: 'MT', qty: vLen * numBranches, type: 'Tubo'
    });

    // Agregar VAEs si la configuracion lo requiere
    if (config.vaeConfig !== 'S/VAE') {
      const vaeRes = this._processVAEs(config.vaeConfig, config.crossingType, inletIn, outletIn);
      materials.push(...vaeRes.materials);
      Object.keys(vaeRes.joints).forEach(d => { joints[d] = (joints[d] || 0) + vaeRes.joints[d]; });
    }

    // Calculo de pegamentos
    materials.push(...this._calculateConsumables(joints));

    return this._multiplyByCount(this._groupMaterials(materials), config.count);
  }

  /**
   * Determina la pieza base (Codo, Tee o Cruz) segun el tipo de crucero
   */
  _getCentralPiece(type, vae, diam) {
    const pType = type === 'Sencillo' ? (vae === 'S/VAE' ? 'CODO 90' : 'TEE') : (vae === 'S/VAE' ? 'TEE' : 'CRUZ');
    return { description: `${pType} PVC HIDRÁULICO CEM ø${this._formatDiameter(diam)} Pulg ⒸCED 40`, unit: 'PZA', qty: 1, type: pType };
  }

  /**
   * Desglosa los componentes internos de un brazo (reducciones, valvulas, adaptadores)
   */
  _processBranch(inletIn, valveIn, outletIn, vType, connType) {
    let materials = [];
    let joints = {};
    const valveMm = this._inchesToMm(valveIn);

    if (inletIn > valveIn) {
      materials.push({
        description: `REDUCCIÓN BUSH PVC HIDRÁULICO CED 40 CEM ø${this._formatDiameter(inletIn)} x ${this._formatDiameter(valveIn)} Pulg`,
        unit: 'PZA', qty: 1, type: 'Reducción Bush'
      });
      joints[this._inchesToMm(inletIn)] = 1;
      joints[valveMm] = 1;
    }

    const tubDesc = `TUBO PVC HIDRÁULICO C/B ø${this._formatDiameter(valveIn)} Pulg ⒸRD-26`;
    materials.push({ description: tubDesc, unit: 'MT', qty: 0.58, type: 'Tubo' }, { description: tubDesc, unit: 'MT', qty: 0.58, type: 'Tubo' });

    const ads = this._getAdapters(vType, valveIn, connType);
    materials.push(...ads.materials);
    Object.keys(ads.joints).forEach(d => { joints[d] = (joints[d] || 0) + ads.joints[d]; });

    materials.push(this._getMainValve(vType, valveIn));

    if (valveIn !== outletIn) {
      const [max, min] = [Math.max(valveIn, outletIn), Math.min(valveIn, outletIn)];
      materials.push({
        description: `REDUCCIÓN BUSH PVC HIDRÁULICO CED 40 CEM ø${this._formatDiameter(max)} x ${this._formatDiameter(min)} Pulg`,
        unit: 'PZA', qty: 1, type: 'Reducción Bush'
      });
      joints[this._inchesToMm(max)] = (joints[this._inchesToMm(max)] || 0) + 1;
      joints[this._inchesToMm(min)] = (joints[this._inchesToMm(min)] || 0) + 1;
    }

    return { materials, joints };
  }

  /**
   * Genera los adaptadores necesarios segun el tipo de conexion de la valvula
   */
  _getAdapters(vType, diamIn, connType) {
    let materials = [];
    let joints = {};
    const dMm = this._inchesToMm(diamIn);
    const dFmt = this._formatDiameter(diamIn);

    const addAdapter = (type) => {
      materials.push({ description: `ADAPTADOR PVC HIDRÁULICO CEM ${type} ø${dFmt} Pulg ⒸCED 40`, unit: 'PZA', qty: 2, type: 'Adaptador' });
      joints[dMm] = 2;
    };

    if (['Angular', 'Hidráulica'].includes(vType)) addAdapter('HEMBRA');
    else if (vType === 'Mariposa') materials.push(...this._buildFlangeSet(diamIn));
    else if (vType === 'Otra') {
      if (connType === 'Rosca Macho') addAdapter('HEMBRA');
      else if (connType === 'Rosca Hembra') addAdapter('MACHO');
      else if (connType === 'Brida') materials.push(...this._buildFlangeSet(diamIn));
      else if (connType === 'Vitaulic') materials.push({ description: `ADAPTADOR VITAULIC ø${dFmt} Pulg`, unit: 'PZA', qty: 2, type: 'Adaptador' });
    }
    return { materials, joints };
  }

  /**
   * Crea el conjunto de brida, tornillos, tuercas y rondanas
   */
  _buildFlangeSet(diamIn) {
    const bolt = this.FLANGE_BOLTS[diamIn] || { qty: 8, diameter: '5/8', length: '5' };
    const dFmt = this._formatDiameter(diamIn);
    return [
      { description: `BRIDA PVC HIDRÁULICO CEM ø${dFmt} Pulg ⒸCED 40`, unit: 'PZA', qty: 2, type: 'Brida' },
      { description: `TORNILLO ACERO AL CARBÓN ☉HEXAGONAL ø${bolt.diameter} Pulg＊Long:${bolt.length} Pulg`, unit: 'PZA', qty: bolt.qty, type: 'Tornillo' },
      { description: `TUERCA GALVANIZADO HEXAGONAL ø${bolt.diameter} Pulg`, unit: 'PZA', qty: bolt.qty, type: 'Tuerca' },
      { description: `RONDANA GALVANIZADO ø${bolt.diameter} Pulg`, unit: 'PZA', qty: bolt.qty, type: 'Rondana' }
    ];
  }

  /**
   * Retorna la descripcion de la valvula principal segun su tipo
   */
  _getMainValve(vType, diamIn) {
    const dFmt = this._formatDiameter(diamIn);
    const descs = { 'Angular': `ANGULAR POLIPROPILENO ø${dFmt} Pulg ×MACHO`, 'Mariposa': `MARIPOSA ALUMINIO ø${dFmt} Pulg ×BRIDADO`, 'Hidráulica': `VÁLVULA HIDRÁULICA ø${dFmt} Pulg ×MACHO` };
    return { description: descs[vType] || `VÁLVULA ESPECIAL ø${dFmt} Pulg`, unit: 'PZA', qty: 1, type: 'Válvula' };
  }

  /**
   * Procesa la adicion de VAEs en los puntos de entrada y salida
   */
  _processVAEs(vaeConfig, crossingType, inletIn, outletIn) {
    let materials = [];
    let joints = {};
    const config = String(vaeConfig);

    if (config.includes('P')) {
      const res = this._buildVAEAssembly(inletIn);
      materials.push(...res.materials);
      Object.keys(res.joints).forEach(d => { joints[d] = (joints[d] || 0) + res.joints[d]; });
    }
    if (config.includes('S')) {
      const count = (crossingType === 'Doble' && config.includes('SS')) ? 2 : 1;
      for (let i = 0; i < count; i++) {
        const res = this._buildVAEAssembly(outletIn);
        materials.push(...res.materials);
        Object.keys(res.joints).forEach(d => { joints[d] = (joints[d] || 0) + res.joints[d]; });
      }
    }
    return { materials, joints };
  }

  /**
   * Construye el ensamble individual de una ventosa (VAE) con sus adaptadores
   */
  _buildVAEAssembly(pipeIn) {
    let materials = [];
    let joints = {};
    const vaeIn = this.VAE_BY_DIAMETER[pipeIn] || 1;
    const [dFmt, vaeFmt] = [this._formatDiameter(pipeIn), this._formatDiameter(vaeIn)];

    if (pipeIn > vaeIn) {
      materials.push({ description: `REDUCCIÓN BUSH PVC HIDRÁULICO CED 40 CEM ø${dFmt} x ${vaeFmt} Pulg`, unit: 'PZA', qty: 1, type: 'Reducción Bush' });
      joints[this._inchesToMm(pipeIn)] = 1;
      joints[this._inchesToMm(vaeIn)] = 1;
    }

    materials.push({ description: `TUBO PVC HIDRÁULICO C/B ø${vaeFmt} Pulg ⒸRD-26`, unit: 'MT', qty: 0.10, type: 'Tubo' });
    const adType = vaeIn <= 1 ? 'HEMBRA' : 'MACHO';
    materials.push({ description: `ADAPTADOR PVC HIDRÁULICO CEM ${adType} ø${vaeFmt} Pulg ⒸCED 40`, unit: 'PZA', qty: 1, type: 'Adaptador' });
    joints[this._inchesToMm(vaeIn)] = (joints[this._inchesToMm(vaeIn)] || 0) + 2;

    const vaeConn = vaeIn <= 1 ? 'MACHO' : 'HEMBRA';
    materials.push({ description: `ADMISIÓN-EXPULSIÓN PLÁSTICO ®EMECK ø${vaeFmt} Pulg ×${vaeConn}`, unit: 'PZA', qty: 1, type: 'VAE' });

    return { materials, joints };
  }

  /**
   * Calcula la cantidad de latas de cemento y limpiador segun uniones totales
   */
  _calculateConsumables(joints) {
    let fractionalCans = 0;
    Object.keys(joints).forEach(dMm => {
      const yield_ = this.CEMENT_YIELD[dMm] || 22;
      fractionalCans += joints[dMm] / yield_;
    });

    const cans = Math.ceil(fractionalCans);
    if (cans <= 0) return [];

    const cleanerL = (cans * this.CEMENT_CAN_LITERS * this.CLEANER_PER_LITER_CEMENT);
    return [
      { description: 'CEMENTO PVC HIDRÁULICO 0.48 L', unit: 'PZA', qty: cans, type: 'Consumible' },
      { description: 'LIMPIADOR PVC HIDRÁULICO 0.5 L', unit: 'PZA', qty: Math.ceil(cleanerL / 0.5), type: 'Consumible' }
    ];
  }

  /**
   * Suma las uniones de una pieza al acumulador de diametros
   */
  _countJoints(piece, joints, dMm) {
    const qty = this.JOINTS_BY_PIECE[piece.type] || 2;
    joints[dMm] = (joints[dMm] || 0) + qty;
  }

  /**
   * Agrupa descripciones duplicadas y suma sus cantidades
   */
  _groupMaterials(list) {
    const grouped = {};
    list.forEach(item => {
      if (grouped[item.description]) grouped[item.description].qty = parseFloat((grouped[item.description].qty + item.qty).toFixed(2));
      else grouped[item.description] = { ...item };
    });
    return Object.values(grouped);
  }

  /**
   * Multiplica las cantidades de la lista por el numero de cruceros iguales
   */
  _multiplyByCount(list, count) {
    return list.map(item => ({ ...item, qty: parseFloat((item.qty * count).toFixed(2)) }));
  }
}