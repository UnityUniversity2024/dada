import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';
import { query } from '../config/db.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const documentsRoot = path.resolve(__dirname, '..', '..');
const coffeeWorkbookPath = path.join(documentsRoot, '2026 COFFEE EXPORT WEEKLY REPORT (22-09-2026) (3).xlsx');
const sesameWorkbookPath = path.join(documentsRoot, '2019 sesame export 190926.xlsx');

const fallbackCoffeeSeed = [
  { id: 1, week: '2026-W29', shipmentNo: 'COF-26-001', buyer: 'Kofi Foods Ltd.', destination: 'Jeddah', quantity: 1420, unit: 'MT', containers: 4, pricePerKg: 2.45, grossUsd: 348900, paidUsd: 210000, paymentStatus: 'Partial', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Red Sea', contract: 'CTR-COF-2109', status: 'In Transit', date: '2026-07-08', notes: 'Export lot 1' },
  { id: 2, week: '2026-W29', shipmentNo: 'COF-26-002', buyer: 'Arab Roast Co.', destination: 'Dubai', quantity: 1180, unit: 'MT', containers: 3, pricePerKg: 2.38, grossUsd: 281000, paidUsd: 281000, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Al Noor', contract: 'CTR-COF-2110', status: 'Delivered', date: '2026-07-10', notes: 'Premium grade' },
  { id: 3, week: '2026-W30', shipmentNo: 'COF-26-003', buyer: 'Cairo Importers', destination: 'Cairo', quantity: 1540, unit: 'MT', containers: 5, pricePerKg: 2.52, grossUsd: 388080, paidUsd: 160000, paymentStatus: 'Partial', cleaning: 'Grade B', product: 'Washed Coffee', vessel: 'MV Nile Star', contract: 'CTR-COF-2111', status: 'At Port', date: '2026-07-15', notes: 'Customs hold' },
  { id: 4, week: '2026-W30', shipmentNo: 'COF-26-004', buyer: 'Riyadh Café Supply', destination: 'Riyadh', quantity: 1325, unit: 'MT', containers: 4, pricePerKg: 2.49, grossUsd: 329925, paidUsd: 329925, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Gulf Horizon', contract: 'CTR-COF-2112', status: 'Delivered', date: '2026-07-18', notes: 'Fast payment' },
  { id: 5, week: '2026-W31', shipmentNo: 'COF-26-005', buyer: 'Harbor Trading', destination: 'Mombasa', quantity: 1660, unit: 'MT', containers: 5, pricePerKg: 2.54, grossUsd: 421640, paidUsd: 220000, paymentStatus: 'Partial', cleaning: 'Grade A', product: 'Coffee Beans', vessel: 'MV East Null', contract: 'CTR-COF-2113', status: 'On Vessel', date: '2026-07-22', notes: 'Cold chain check' },
  { id: 6, week: '2026-W31', shipmentNo: 'COF-26-006', buyer: 'Prime Café Aus.', destination: 'Sydney', quantity: 970, unit: 'MT', containers: 3, pricePerKg: 2.59, grossUsd: 251230, paidUsd: 251230, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Arabica', vessel: 'MV Sirocco', contract: 'CTR-COF-2114', status: 'Delivered', date: '2026-07-24', notes: 'QC passed' },
  { id: 7, week: '2026-W32', shipmentNo: 'COF-26-007', buyer: 'Great Lakes Foods', destination: 'Dar es Salaam', quantity: 1380, unit: 'MT', containers: 4, pricePerKg: 2.43, grossUsd: 335340, paidUsd: 98000, paymentStatus: 'Partial', cleaning: 'Grade B', product: 'Green Coffee', vessel: 'MV Aurora', contract: 'CTR-COF-2115', status: 'At Port', date: '2026-07-29', notes: 'Awaiting customs' },
];

const fallbackSesameSeed = [
  { id: 1, week: '2026-W29', shipmentNo: 'SES-26-001', buyer: 'Golden Sesame Ltd.', destination: 'Jeddah', quantity: 720, unit: 'MT', containers: 2, pricePerKg: 1.82, grossUsd: 131040, paidUsd: 131040, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Bay Gold', contract: 'CTR-SES-1101', status: 'Delivered', date: '2026-07-09', notes: 'Premium white sesame' },
  { id: 2, week: '2026-W29', shipmentNo: 'SES-26-002', buyer: 'Nile Grain House', destination: 'Cairo', quantity: 540, unit: 'MT', containers: 2, pricePerKg: 1.76, grossUsd: 95040, paidUsd: 60000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Nile Gold', contract: 'CTR-SES-1102', status: 'In Transit', date: '2026-07-10', notes: 'Checking quality' },
  { id: 3, week: '2026-W30', shipmentNo: 'SES-26-003', buyer: 'Sahara & Sons', destination: 'Dubai', quantity: 680, unit: 'MT', containers: 2, pricePerKg: 1.84, grossUsd: 125120, paidUsd: 125120, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Dune Pearl', contract: 'CTR-SES-1103', status: 'Delivered', date: '2026-07-13', notes: 'Delivered clean' },
  { id: 4, week: '2026-W30', shipmentNo: 'SES-26-004', buyer: 'Red Sea Trade', destination: 'Aqaba', quantity: 610, unit: 'MT', containers: 2, pricePerKg: 1.8, grossUsd: 109800, paidUsd: 0, paymentStatus: 'Open', cleaning: 'Pending', product: 'Sesame Seeds', vessel: 'MV Red Sea', contract: 'CTR-SES-1104', status: 'Pending', date: '2026-07-16', notes: 'Inspection review' },
  { id: 5, week: '2026-W31', shipmentNo: 'SES-26-005', buyer: 'Aman Importers', destination: 'Riyadh', quantity: 790, unit: 'MT', containers: 3, pricePerKg: 1.88, grossUsd: 148520, paidUsd: 75000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Najd Star', contract: 'CTR-SES-1105', status: 'On Vessel', date: '2026-07-19', notes: 'Shipment loaded' },
  { id: 6, week: '2026-W31', shipmentNo: 'SES-26-006', buyer: 'Desert Grain MW', destination: 'Khartoum', quantity: 660, unit: 'MT', containers: 2, pricePerKg: 1.78, grossUsd: 117480, paidUsd: 117480, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Khartoum', contract: 'CTR-SES-1106', status: 'Delivered', date: '2026-07-22', notes: 'Good quality' },
  { id: 7, week: '2026-W32', shipmentNo: 'SES-26-007', buyer: 'Hawker Markets', destination: 'Muscat', quantity: 740, unit: 'MT', containers: 3, pricePerKg: 1.86, grossUsd: 137640, paidUsd: 90000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Oman Gold', contract: 'CTR-SES-1107', status: 'At Port', date: '2026-07-24', notes: 'Awaiting discharge' },
];

const normalizeCell = (value) => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.replace(/\s+/g, ' ').trim();
  return String(value).trim();
};

const parseNumber = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  const cleaned = String(value).replace(/[$,\s]/g, '').replace(/&.*$/, '').trim();
  const numeric = Number(cleaned);
  return Number.isFinite(numeric) ? numeric : 0;
};

const findHeaderIndex = (headers, aliases) => {
  return headers.findIndex((header) => {
    const text = normalizeCell(header).toLowerCase();
    return aliases.some((alias) => text.includes(alias.toLowerCase()));
  });
};

const normalizePaymentStatus = (rawValue, paidUsd, grossUsd) => {
  const text = normalizeCell(rawValue).toLowerCase();

  if (text.includes('received') || text.includes('paid') || paidUsd >= grossUsd * 0.98) return 'Paid';
  if (text.includes('partial') || text.includes('advance') || text.includes('document') || paidUsd > 0) return 'Partial';
  return 'Open';
};

const normalizeShipmentStatus = (rawValue) => {
  const text = normalizeCell(rawValue).toLowerCase();

  if (!text) return 'Pending';
  if (text.includes('deliver') || text.includes('arrived') || text.includes('received')) return 'Delivered';
  if (text.includes('port') || text.includes('at port')) return 'At Port';
  if (text.includes('vessel') || text.includes('board') || text.includes('transit') || text.includes('loaded') || text.includes('shipping')) return 'In Transit';
  return 'Pending';
};

const buildWorkbookRecord = (row, headers, type, index) => {
  const invoiceIndex = findHeaderIndex(headers, ['invoice no']);
  const buyerIndex = findHeaderIndex(headers, ['buyer name']);
  const destinationIndex = findHeaderIndex(headers, ['destination']);
  const descriptionIndex = findHeaderIndex(headers, ['description of goods']);
  const quantityIndex = findHeaderIndex(headers, ['quantity']);
  const containersIndex = findHeaderIndex(headers, ['no of containers']);
  const contractIndex = findHeaderIndex(headers, ['contract registered']);
  const cleaningIndex = findHeaderIndex(headers, ['cleaning status']);
  const vesselIndex = findHeaderIndex(headers, ['name of shipping line']);
  const shipInstructionIndex = findHeaderIndex(headers, ['shipping instruction']);
  const statusIndex = findHeaderIndex(headers, ['delivered to edr']);
  const dateIndex = findHeaderIndex(headers, ['shipped on board date', 'sob date']);
  const paymentIndex = findHeaderIndex(headers, ['payment']);
  const unitPriceIndex = findHeaderIndex(headers, ['unit price/mt', 'unit price/lb', 'unit price']);
  const totalSalesIndex = findHeaderIndex(headers, ['total sales price usd / amount', 'total sales price', 'total sales']);
  const advanceReceivedIndex = findHeaderIndex(headers, ['advance amount received/usd', 'advance amount', 'advance amount received']);
  const receivedIndex = findHeaderIndex(headers, ['100% received']);
  const termsIndex = findHeaderIndex(headers, ['terms of delivery']);

  const shipmentNo = normalizeCell(row[invoiceIndex] || `AUTO-${type.toUpperCase()}-${index + 1}`);
  const buyer = normalizeCell(row[buyerIndex] || 'Unknown buyer');
  const destination = normalizeCell(row[destinationIndex] || 'Unknown destination');
  const product = normalizeCell(row[descriptionIndex] || (type === 'coffee' ? 'Coffee Export' : 'Sesame Export'));
  const quantity = parseNumber(row[quantityIndex]);
  const containers = parseNumber(row[containersIndex]);
  const grossUsd = parseNumber(row[totalSalesIndex]);
  const paidUsd = parseNumber(row[advanceReceivedIndex] || row[receivedIndex]) || Math.min(grossUsd, parseNumber(row[advanceReceivedIndex] || row[receivedIndex]));
  const paymentStatus = normalizePaymentStatus(row[receivedIndex] || row[paymentIndex] || row[advanceReceivedIndex], paidUsd, grossUsd);
  const cleaning = normalizeCell(row[cleaningIndex] || 'Pending');
  const vessel = normalizeCell(row[vesselIndex] || row[shipInstructionIndex] || 'Not assigned');
  const contract = normalizeCell(row[contractIndex] || 'Not registered');
  const status = normalizeShipmentStatus(row[statusIndex] || row[paymentIndex] || row[termsIndex]);
  const dateValue = row[dateIndex];
  const shipmentDate = dateValue instanceof Date ? dateValue.toISOString().slice(0, 10) : normalizeCell(dateValue || new Date().toISOString().slice(0, 10));
  const unitPrice = parseNumber(row[unitPriceIndex]);
  const noteValue = normalizeCell(row[shipInstructionIndex] || row[termsIndex] || 'Export record from source sheet');

  return {
    id: index + 1,
    week: shipmentDate ? `W${new Date(shipmentDate).getFullYear()}-${String(new Date(shipmentDate).getMonth() + 1).padStart(2, '0')}` : 'W-2026',
    shipmentNo,
    buyer,
    destination,
    quantity,
    unit: 'MT',
    containers,
    pricePerKg: unitPrice || Number((grossUsd / Math.max(quantity, 1)).toFixed(4)),
    grossUsd,
    paidUsd,
    paymentStatus,
    cleaning,
    product,
    vessel,
    contract,
    status,
    date: shipmentDate,
    notes: noteValue,
  };
};

const readWorkbookRows = (filePath, type) => {
  if (!fs.existsSync(filePath)) {
    return type === 'coffee' ? fallbackCoffeeSeed : fallbackSesameSeed;
  }

  try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, blankrows: false });

    const headerIndex = rows.findIndex((row) => {
      if (!Array.isArray(row)) return false;
      return row.some((cell) => {
        const text = normalizeCell(cell).toLowerCase();
        return text.includes('invoice no') || text.includes('buyer name') || text.includes('description of goods');
      });
    });

    if (headerIndex === -1) {
      return type === 'coffee' ? fallbackCoffeeSeed : fallbackSesameSeed;
    }

    const headers = rows[headerIndex].map((value) => normalizeCell(value));
    const payload = rows.slice(headerIndex + 1)
      .filter((row) => Array.isArray(row) && row.some((cell) => normalizeCell(cell) !== ''))
      .map((row, index) => buildWorkbookRecord(row, headers, type, index));

    return payload.length ? payload : (type === 'coffee' ? fallbackCoffeeSeed : fallbackSesameSeed);
  } catch (error) {
    return type === 'coffee' ? fallbackCoffeeSeed : fallbackSesameSeed;
  }
};

export const coffeeSeed = readWorkbookRows(coffeeWorkbookPath, 'coffee');
export const sesameSeed = readWorkbookRows(sesameWorkbookPath, 'sesame');

export const getDbRow = async (tableName) => {
  try {
    const result = await query(`SELECT * FROM ${tableName} ORDER BY id ASC`);
    return result.rows;
  } catch (error) {
    return [];
  }
};

export const bootstrapExports = async () => {
  const coffee = await getDbRow('coffee_exports');
  const sesame = await getDbRow('sesame_exports');

  return {
    success: true,
    source: coffee.length || sesame.length ? 'postgresql' : 'xlsx-source',
    coffee: coffee.length ? coffee : coffeeSeed,
    sesame: sesame.length ? sesame : sesameSeed,
  };
};

export const listCoffee = async () => {
  const rows = await getDbRow('coffee_exports');
  return rows.length ? rows : coffeeSeed;
};

export const listSesame = async () => {
  const rows = await getDbRow('sesame_exports');
  return rows.length ? rows : sesameSeed;
};

export const createCoffeeRecord = async (payload) => {
  const record = {
    id: Date.now(),
    ...payload,
    grossUsd: Number(payload.grossUsd || 0),
    paidUsd: Number(payload.paidUsd || 0),
    quantity: Number(payload.quantity || 0),
    containers: Number(payload.containers || 0),
    pricePerKg: Number(payload.pricePerKg || 0),
  };

  const rows = await getDbRow('coffee_exports');
  if (!rows.length) {
    coffeeSeed.unshift(record);
    return record;
  }

  return record;
};

export const createSesameRecord = async (payload) => {
  const record = {
    id: Date.now(),
    ...payload,
    grossUsd: Number(payload.grossUsd || 0),
    paidUsd: Number(payload.paidUsd || 0),
    quantity: Number(payload.quantity || 0),
    containers: Number(payload.containers || 0),
    pricePerKg: Number(payload.pricePerKg || 0),
  };

  const rows = await getDbRow('sesame_exports');
  if (!rows.length) {
    sesameSeed.unshift(record);
    return record;
  }

  return record;
};

export const updateCoffeeRecord = async (id, payload) => {
  const rows = await listCoffee();
  const index = rows.findIndex((row) => String(row.id) === String(id));
  if (index === -1) throw new Error('Coffee record not found');
  const updated = { ...rows[index], ...payload };
  rows[index] = updated;
  return updated;
};

export const updateSesameRecord = async (id, payload) => {
  const rows = await listSesame();
  const index = rows.findIndex((row) => String(row.id) === String(id));
  if (index === -1) throw new Error('Sesame record not found');
  const updated = { ...rows[index], ...payload };
  rows[index] = updated;
  return updated;
};

export const deleteCoffeeRecord = async (id) => {
  const rows = await listCoffee();
  const index = rows.findIndex((row) => String(row.id) === String(id));
  if (index === -1) throw new Error('Coffee record not found');
  const removed = rows.splice(index, 1)[0];
  return removed;
};

export const deleteSesameRecord = async (id) => {
  const rows = await listSesame();
  const index = rows.findIndex((row) => String(row.id) === String(id));
  if (index === -1) throw new Error('Sesame record not found');
  const removed = rows.splice(index, 1)[0];
  return removed;
};
