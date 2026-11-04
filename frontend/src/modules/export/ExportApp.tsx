import { useEffect, useMemo, useState, type ChangeEvent, type FormEvent } from 'react';
import { createPortal } from 'react-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { coffeeSeed, sesameSeed } from './mockExportData';

type ExportTab = 'coffee' | 'sesame' | 'contract-flow' | 'farm-to-store';

type RolePermission = {
  role: string;
  allowed: string[];
  level: 'View' | 'Edit' | 'Approve' | 'Admin';
};

type WorkflowStep = {
  id: string;
  title: string;
  owner: string;
  status: 'done' | 'active' | 'pending';
  summary: string;
};

type ContractAttachment = {
  id: string;
  name: string;
  type: string;
  size: string;
  uploadedAt: string;
  dataUrl?: string;
};

type ContractPayment = {
  id: string;
  kind: 'LC' | 'TT' | 'Cash' | 'Other';
  amountUsd: number;
  status: 'Open' | 'Partial' | 'Paid';
  reference: string;
  dueDate: string;
  paidDate?: string;
  notes: string;
};

type ContractRecord = {
  id: string;
  contractNo: string;
  title: string;
  buyer: string;
  product: string;
  quantity: string;
  netWeight: string;
  packing: string;
  quality: string;
  price: string;
  terms: string;
  shipment: string;
  shipping: string;
  payment: string;
  insurance: string;
  qualityApproval: string;
  arbitration: string;
  note: string;
  status: 'Draft' | 'Active' | 'In Execution' | 'Settled' | 'Closed';
  approvalTag: string;
  createdAt: string;
  updatedAt: string;
  workflow: WorkflowStep[];
  attachments: ContractAttachment[];
  payments: ContractPayment[];
  roleMatrix: RolePermission[];
};

type ExportRow = {
  id: number;
  week: string;
  shipmentNo: string;
  buyer: string;
  destination: string;
  quantity: number;
  unit: string;
  containers: number;
  pricePerKg: number;
  grossUsd: number;
  paidUsd: number;
  paymentStatus: 'Paid' | 'Partial' | 'Open';
  cleaning: string;
  product: string;
  vessel: string;
  contract: string;
  status: 'Delivered' | 'In Transit' | 'At Port' | 'On Vessel' | 'Pending';
  date: string;
  notes: string;
};

type TrendPoint = {
  label: string;
  value: number;
};

const currency = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  maximumFractionDigits: 2,
});

const compactCurrency = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  notation: 'compact',
  maximumFractionDigits: 1,
});

const dataMap: Record<ExportTab, ExportRow[]> = {
  coffee: coffeeSeed,
  sesame: sesameSeed,
};

const money = (value = 0): string => currency.format(Number(value) || 0);
const compactMoney = (value = 0): string => compactCurrency.format(Number(value) || 0);
const normalizeStatus = (value?: string): string => String(value || '').toLowerCase().replace(/\s+/g, '-');
const formatShort = (value?: string | null): string => value ?? '—';

const defaultRecordForm = (tab: ExportTab): Partial<ExportRow> => ({
  week: `2026-W${Math.max(33, new Date().getUTCDate() % 10 + 30)}`,
  shipmentNo: tab === 'coffee' ? 'COF-26-100' : 'SES-26-100',
  buyer: '',
  destination: '',
  quantity: 1000,
  unit: 'MT',
  containers: 3,
  pricePerKg: 1.8,
  grossUsd: 1800,
  paidUsd: 0,
  paymentStatus: 'Open',
  cleaning: tab === 'coffee' ? 'Grade A' : 'Cleaned',
  product: tab === 'coffee' ? 'Green Coffee' : 'Sesame Seeds',
  vessel: '',
  contract: '',
  status: 'Pending',
  date: new Date().toISOString().slice(0, 10),
  notes: '',
});

const contractModules = [
  { name: 'Exchange agreement', items: ['Overview', 'Terms', 'Documentation', 'Approval'] },
  { name: 'Operations', items: ['Shipment plan', 'Loading schedule', 'Quality check', 'Dispatch'] },
  { name: 'Finance', items: ['LC/Payment', 'Settlement', 'Receivables', 'Reconciliation'] },
  { name: 'Controls', items: ['Role access', 'Audit trail', 'Exceptions', 'Close-out'] },
];

const contractWorkflow: WorkflowStep[] = [
  { id: 'agreement', title: 'Exchange agreement', owner: 'Sales lead', status: 'done', summary: 'Commercial terms signed and confirmed.' },
  { id: 'review', title: 'Review & approval', owner: 'Compliance', status: 'done', summary: 'Sample approval and legal review complete.' },
  { id: 'shipment', title: 'Shipment execution', owner: 'Logistics', status: 'active', summary: 'Bagging, loading and shipping in progress.' },
  { id: 'payment', title: 'Receivable & payment', owner: 'Finance', status: 'pending', summary: 'LC, settlement and cash follow-up pending.' },
  { id: 'close', title: 'Closure', owner: 'Admin', status: 'pending', summary: 'Archive final documents and close the file.' },
];

const contractRoleMatrix: RolePermission[] = [
  { role: 'Sales Manager', allowed: ['Create agreement', 'Issue price confirmation', 'Track shipment status'], level: 'Edit' },
  { role: 'Procurement', allowed: ['Review quality terms', 'Nominate shipping line', 'Confirm arrival schedule'], level: 'Edit' },
  { role: 'Finance Officer', allowed: ['Review LC', 'Approve payment', 'Match receipts'], level: 'Approve' },
  { role: 'Compliance / Legal', allowed: ['Validate contract clauses', 'Approve arbitration terms', 'Verify documents'], level: 'Approve' },
  { role: 'Operations', allowed: ['Track container loading', 'Update dispatch status', 'Escalate delays'], level: 'View' },
  { role: 'Admin', allowed: ['Manage users', 'Configure permissions', 'Close contract file'], level: 'Admin' },
];

const contractDetails = {
  quantity: '320 x 60 KILOS NET BAGS',
  netWeight: '19.2 M.T.',
  packing: 'Grain Pro bags',
  quality: 'ETHIOPIA / ARABICA COFFEE WASHED SIDAMO 2',
  price: '373.0000 US$/LB',
  terms: 'Free On Board Djibouti Net Shipping Weights 0.50 % franchise',
  shipment: 'All month August 2026 at buyer’s call latest 30-07-2026',
  shipping: 'Buyer to nominate shipping line.',
  payment: 'LETTER OF CREDIT AT SIGHT - to be opened prior to shipment.',
  insurance: 'Insurance For Buyer’s Account.',
  qualityApproval: 'Subject Approval Of Samples - Replacement Basis.',
  arbitration: 'If Any, With British Coffee Association (Bca).',
  note: 'This Contract Is Subject To The Terms And Conditions Of The European Standard Contract For Coffee In Force At The Time Of Conclusion Of This Contract And To The Above Particular Conditions Which Override All The Others.',
};

const contractAttachmentsSeed: ContractAttachment[] = [
  { id: 'att-1', name: 'ddd.pdf', type: 'Agreement', size: '2.4 MB', uploadedAt: '2026-08-01' },
  { id: 'att-2', name: 'dd.pdf', type: 'Packing list', size: '1.1 MB', uploadedAt: '2026-08-02' },
  { id: 'att-3', name: 'WendFinance.txt', type: 'Workflow log', size: '32 KB', uploadedAt: '2026-08-03' },
  { id: 'att-4', name: 'WhatsApp Image 2026-10-06 at 3.39.46 PM.jpeg', type: 'Reference image', size: '1.4 MB', uploadedAt: '2026-08-04' },
];

const defaultWorkflow = (): WorkflowStep[] =>
  contractWorkflow.map((step) => ({ ...step }));

const defaultRoleMatrix = (): RolePermission[] =>
  contractRoleMatrix.map((row) => ({ ...row, allowed: [...row.allowed] }));

const seedContracts = (): ContractRecord[] => [
  {
    id: 'ctr-sidamo-2026-08',
    contractNo: 'CTR-SID-2026-08',
    title: 'Sidamo Washed Arabica — August 2026',
    buyer: 'European Roaster Group',
    product: contractDetails.quality,
    quantity: contractDetails.quantity,
    netWeight: contractDetails.netWeight,
    packing: contractDetails.packing,
    quality: contractDetails.quality,
    price: contractDetails.price,
    terms: contractDetails.terms,
    shipment: contractDetails.shipment,
    shipping: contractDetails.shipping,
    payment: contractDetails.payment,
    insurance: contractDetails.insurance,
    qualityApproval: contractDetails.qualityApproval,
    arbitration: contractDetails.arbitration,
    note: contractDetails.note,
    status: 'In Execution',
    approvalTag: 'Approved for shipment',
    createdAt: '2026-07-15',
    updatedAt: '2026-08-20',
    workflow: defaultWorkflow(),
    attachments: contractAttachmentsSeed.map((a) => ({ ...a })),
    payments: [
      {
        id: 'pay-1',
        kind: 'LC',
        amountUsd: 250000,
        status: 'Partial',
        reference: 'LC-DJ-88421',
        dueDate: '2026-08-30',
        paidDate: '2026-08-18',
        notes: 'LC at sight — first draw 60%',
      },
      {
        id: 'pay-2',
        kind: 'TT',
        amountUsd: 80000,
        status: 'Open',
        reference: 'BAL-TT-pending',
        dueDate: '2026-09-15',
        notes: 'Balance after quality release',
      },
    ],
    roleMatrix: defaultRoleMatrix(),
  },
];

const emptyContractForm = (): Partial<ContractRecord> => ({
  contractNo: `CTR-${new Date().getFullYear()}-${String(Date.now()).slice(-4)}`,
  title: '',
  buyer: '',
  product: 'ETHIOPIA / ARABICA COFFEE WASHED SIDAMO 2',
  quantity: '320 x 60 KILOS NET BAGS',
  netWeight: '19.2 M.T.',
  packing: 'Grain Pro bags',
  quality: 'ETHIOPIA / ARABICA COFFEE WASHED SIDAMO 2',
  price: '373.0000 US$/LB',
  terms: 'Free On Board Djibouti Net Shipping Weights 0.50 % franchise',
  shipment: '',
  shipping: 'Buyer to nominate shipping line.',
  payment: 'LETTER OF CREDIT AT SIGHT - to be opened prior to shipment.',
  insurance: "Insurance For Buyer's Account.",
  qualityApproval: 'Subject Approval Of Samples - Replacement Basis.',
  arbitration: 'If Any, With British Coffee Association (Bca).',
  note: '',
  status: 'Draft',
  approvalTag: 'Pending review',
});

const emptyPaymentForm = (): Omit<ContractPayment, 'id'> => ({
  kind: 'LC',
  amountUsd: 0,
  status: 'Open',
  reference: '',
  dueDate: new Date().toISOString().slice(0, 10),
  paidDate: '',
  notes: '',
});

type ExportAppProps = {
  role?: string;
  permissions?: string[];
  canViewContract?: boolean;
  canEditContract?: boolean;
  canAdd?: boolean;
  canEdit?: boolean;
  canDelete?: boolean;
  canUploadData?: boolean;
  canExportExcel?: boolean;
  canExportPdf?: boolean;
  canExportWord?: boolean;
  canExportCsv?: boolean;
  canViewCoffee?: boolean;
  canViewSesame?: boolean;
};


/* ========== NSPF Farm → Store (inlined) ========== */
/** NSPF Coffee & Sesame — Farm → Store pipeline mock data (Ethiopia / ETB + USD) */

type Commodity = 'Sesame' | 'Coffee';

type Supplier = {
  id: string;
  name: string;
  commodity: Commodity;
  origin: string;
  woreda: string;
  contact: string;
  phone: string;
  status: 'Active' | 'Under review' | 'Suspended';
  performanceScore: number;
  notes: string;
};

type SourcingPlan = {
  id: string;
  commodity: Commodity;
  origin: string;
  grade: string;
  targetMt: number;
  targetPriceEtbPerKg: number;
  period: string;
  requiredBy: string;
  status: 'Open' | 'In progress' | 'Closed';
};

type PurchaseReceipt = {
  id: string;
  poNo: string;
  supplierId: string;
  supplierName: string;
  commodity: Commodity;
  origin: string;
  grade: string;
  receiptDate: string;
  grossKg: number;
  netKg: number;
  bags: number;
  moisturePct: number;
  priceEtbPerKg: number;
  deductionsEtb: number;
  status: 'Received' | 'Inspected' | 'Rejected' | 'Settled';
  warehouse: string;
};

type StockLot = {
  id: string;
  lotNo: string;
  commodity: Commodity;
  origin: string;
  grade: string;
  purchaseId: string;
  warehouse: string;
  bin: string;
  netKg: number;
  bags: number;
  qualityStatus: 'Available' | 'Reserved' | 'Quarantine' | 'In process' | 'Dispatched' | 'Rejected';
  receivedAt: string;
  owner: string;
};

type ProcessingBatch = {
  id: string;
  batchNo: string;
  commodity: Commodity;
  inputLotIds: string[];
  processType: string;
  inputKg: number;
  outputKg: number;
  lossKg: number;
  yieldPct: number;
  outputGrade: string;
  labResult: string;
  releaseStatus: 'Hold' | 'Released' | 'Rejected';
  processedAt: string;
  warehouse: string;
};

const suppliersSeed: Supplier[] = [
  { id: 'SUP-SES-01', name: 'Humera Farmers Cooperative', commodity: 'Sesame', origin: 'Humera', woreda: 'Kafta Humera', contact: 'Ato Girmay Haile', phone: '+251914100001', status: 'Active', performanceScore: 92, notes: 'Primary Humera white sesame source' },
  { id: 'SUP-SES-02', name: 'Wollega Highlands Traders', commodity: 'Sesame', origin: 'Wollega', woreda: 'Gimbi', contact: 'W/ro Aster Bekele', phone: '+251911200002', status: 'Active', performanceScore: 88, notes: 'Wollega mixed grade' },
  { id: 'SUP-SES-03', name: 'Metema Collection Center', commodity: 'Sesame', origin: 'Metema', woreda: 'Metema', contact: 'Ato Yusuf Mohammed', phone: '+251918300003', status: 'Under review', performanceScore: 76, notes: 'Seasonal supply' },
  { id: 'SUP-COF-01', name: 'Sidamo Washing Station PLC', commodity: 'Coffee', origin: 'Sidamo', woreda: 'Bensa', contact: 'Ato Tadesse Alemu', phone: '+251911400004', status: 'Active', performanceScore: 95, notes: 'Washed Arabica grade 2' },
  { id: 'SUP-COF-02', name: 'Yirgacheffe Smallholders Union', commodity: 'Coffee', origin: 'Yirgacheffe', woreda: 'Kochere', contact: 'W/ro Hanna Desta', phone: '+251912500005', status: 'Active', performanceScore: 91, notes: 'Specialty lots' },
];

const sourcingPlansSeed: SourcingPlan[] = [
  { id: 'SP-2026-SES-01', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', targetMt: 500, targetPriceEtbPerKg: 118, period: '2026-Q3', requiredBy: '2026-09-30', status: 'In progress' },
  { id: 'SP-2026-SES-02', commodity: 'Sesame', origin: 'Wollega', grade: 'Wollega Mixed', targetMt: 220, targetPriceEtbPerKg: 112, period: '2026-Q3', requiredBy: '2026-10-15', status: 'Open' },
  { id: 'SP-2026-COF-01', commodity: 'Coffee', origin: 'Sidamo', grade: 'Grade 2 Washed', targetMt: 180, targetPriceEtbPerKg: 420, period: '2026-Q3', requiredBy: '2026-09-20', status: 'In progress' },
];

const purchasesSeed: PurchaseReceipt[] = [
  { id: 'PR-SES-2401', poNo: 'PO-SES-2026-014', supplierId: 'SUP-SES-01', supplierName: 'Humera Farmers Cooperative', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', receiptDate: '2026-08-12', grossKg: 52000, netKg: 51200, bags: 1024, moisturePct: 6.2, priceEtbPerKg: 116, deductionsEtb: 18500, status: 'Inspected', warehouse: 'NSPF Main Store – Djibouti Road' },
  { id: 'PR-SES-2402', poNo: 'PO-SES-2026-018', supplierId: 'SUP-SES-02', supplierName: 'Wollega Highlands Traders', commodity: 'Sesame', origin: 'Wollega', grade: 'Wollega Mixed', receiptDate: '2026-08-20', grossKg: 31000, netKg: 30400, bags: 608, moisturePct: 7.1, priceEtbPerKg: 110, deductionsEtb: 9200, status: 'Received', warehouse: 'NSPF Main Store – Djibouti Road' },
  { id: 'PR-SES-2403', poNo: 'PO-SES-2026-021', supplierId: 'SUP-SES-01', supplierName: 'Humera Farmers Cooperative', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', receiptDate: '2026-09-02', grossKg: 48000, netKg: 47250, bags: 945, moisturePct: 6.0, priceEtbPerKg: 119, deductionsEtb: 12000, status: 'Settled', warehouse: 'Cleaning Plant In-Yard' },
  { id: 'PR-COF-1101', poNo: 'PO-COF-2026-009', supplierId: 'SUP-COF-01', supplierName: 'Sidamo Washing Station PLC', commodity: 'Coffee', origin: 'Sidamo', grade: 'Grade 2 Washed', receiptDate: '2026-08-05', grossKg: 42000, netKg: 41200, bags: 687, moisturePct: 11.5, priceEtbPerKg: 415, deductionsEtb: 0, status: 'Inspected', warehouse: 'Coffee Green Store A' },
];

const lotsSeed: StockLot[] = [
  { id: 'LOT-SES-2401-A', lotNo: 'LOT-SES-2401-A', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', purchaseId: 'PR-SES-2401', warehouse: 'NSPF Main Store – Djibouti Road', bin: 'S-A-12', netKg: 25600, bags: 512, qualityStatus: 'Available', receivedAt: '2026-08-12', owner: 'NSPF' },
  { id: 'LOT-SES-2401-B', lotNo: 'LOT-SES-2401-B', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', purchaseId: 'PR-SES-2401', warehouse: 'Cleaning Plant In-Yard', bin: 'PROC-01', netKg: 25600, bags: 512, qualityStatus: 'In process', receivedAt: '2026-08-12', owner: 'NSPF' },
  { id: 'LOT-SES-2402-A', lotNo: 'LOT-SES-2402-A', commodity: 'Sesame', origin: 'Wollega', grade: 'Wollega Mixed', purchaseId: 'PR-SES-2402', warehouse: 'NSPF Main Store – Djibouti Road', bin: 'S-B-03', netKg: 30400, bags: 608, qualityStatus: 'Quarantine', receivedAt: '2026-08-20', owner: 'NSPF' },
  { id: 'LOT-SES-2403-A', lotNo: 'LOT-SES-2403-A', commodity: 'Sesame', origin: 'Humera', grade: 'Humera Whitish', purchaseId: 'PR-SES-2403', warehouse: 'Cleaning Plant In-Yard', bin: 'PROC-02', netKg: 47250, bags: 945, qualityStatus: 'Available', receivedAt: '2026-09-02', owner: 'NSPF' },
  { id: 'LOT-COF-1101-A', lotNo: 'LOT-COF-1101-A', commodity: 'Coffee', origin: 'Sidamo', grade: 'Grade 2 Washed', purchaseId: 'PR-COF-1101', warehouse: 'Coffee Green Store A', bin: 'C-01', netKg: 41200, bags: 687, qualityStatus: 'Reserved', receivedAt: '2026-08-05', owner: 'NSPF' },
];

const processingSeed: ProcessingBatch[] = [
  { id: 'BAT-SES-01', batchNo: 'BAT-SES-2026-01', commodity: 'Sesame', inputLotIds: ['LOT-SES-2401-B'], processType: 'Cleaning + grading', inputKg: 25600, outputKg: 24850, lossKg: 750, yieldPct: 97.1, outputGrade: 'Export cleaned Humera', labResult: 'Moisture 5.8% · Purity 99.2% · Pass', releaseStatus: 'Released', processedAt: '2026-08-18', warehouse: 'Export Ready Store' },
  { id: 'BAT-SES-02', batchNo: 'BAT-SES-2026-02', commodity: 'Sesame', inputLotIds: ['LOT-SES-2403-A'], processType: 'Cleaning', inputKg: 20000, outputKg: 19420, lossKg: 580, yieldPct: 97.1, outputGrade: 'Export cleaned Humera', labResult: 'Pending lab', releaseStatus: 'Hold', processedAt: '2026-09-05', warehouse: 'Cleaning Plant In-Yard' },
  { id: 'BAT-COF-01', batchNo: 'BAT-COF-2026-01', commodity: 'Coffee', inputLotIds: ['LOT-COF-1101-A'], processType: 'Dry milling + grading', inputKg: 15000, outputKg: 14280, lossKg: 720, yieldPct: 95.2, outputGrade: 'Sidamo G2 Export', labResult: 'Cupping 84 · Pass', releaseStatus: 'Released', processedAt: '2026-08-22', warehouse: 'Coffee Export Store' },
];

const pipelineStages = [
  { id: 'plan', title: '1. Plan sourcing', desc: 'Demand, origin, grade, volume, target cost, required date' },
  { id: 'supplier', title: '2. Suppliers / farm sources', desc: 'Register origin suppliers, cooperatives, washing stations' },
  { id: 'purchase', title: '3. Purchase & receive', desc: 'PO, weighment, inspection, goods receipt from farm/collection' },
  { id: 'lot', title: '4. Lot & warehouse stock', desc: 'Unique lots, bins, available/reserved/quarantine balances' },
  { id: 'process', title: '5. Process & quality release', desc: 'Cleaning/milling, yield, lab, hold/release to export store' },
  { id: 'ready', title: '6. Export-ready store', desc: 'Released stock available for contract allocation & shipment' },
];


/** Inline Farm→Store styles (no separate CSS file). */
function FarmToStoreStyles(): JSX.Element {
  return (
    <style>{`
/* ===== NSPF Farm → Store — robust light / white UI ===== */
.farm-to-store {
  position: relative;
  max-width: 1460px;
  margin: 0 auto;
  padding: 18px 16px 36px;
  border-radius: 20px;
  overflow: visible;
  color: #0f2740;
  background:
    radial-gradient(ellipse 80% 50% at 0% 0%, rgba(14, 165, 233, 0.08), transparent 50%),
    radial-gradient(ellipse 60% 40% at 100% 0%, rgba(16, 185, 129, 0.06), transparent 45%),
    linear-gradient(180deg, #f8fbfd 0%, #eef4f9 45%, #f5f8fb 100%);
  border: 1px solid rgba(15, 60, 90, 0.1);
  box-shadow:
    0 1px 0 rgba(255,255,255,0.9) inset,
    0 18px 48px rgba(15, 40, 70, 0.08);
}
.farm-to-store::before {
  content: "";
  pointer-events: none;
  position: absolute;
  inset: 0;
  background-image:
    linear-gradient(rgba(15, 80, 120, 0.035) 1px, transparent 1px),
    linear-gradient(90deg, rgba(15, 80, 120, 0.035) 1px, transparent 1px);
  background-size: 44px 44px;
  mask-image: radial-gradient(ellipse at 50% 0%, black 20%, transparent 80%);
  opacity: 0.7;
}
.farm-to-store > *:not(.fts-modal-backdrop) { position: relative; z-index: 1; }
.farm-to-store .fts-modal-backdrop { position: fixed !important; z-index: 2147483000 !important; }

.fts-hero {
  display: flex; flex-wrap: wrap; justify-content: space-between; gap: 16px;
  margin-bottom: 18px;
  padding: 16px 18px;
  border-radius: 16px;
  background: linear-gradient(135deg, #ffffff 0%, #f3f8fc 100%);
  border: 1px solid rgba(15, 70, 110, 0.12);
  box-shadow: 0 8px 28px rgba(15, 40, 70, 0.06);
}
.fts-hero h2 {
  margin: 6px 0 8px;
  font-size: clamp(1.35rem, 2.5vw, 1.9rem);
  color: #0b2540;
  letter-spacing: -0.02em;
}
.fts-sub { margin: 0; max-width: 740px; color: #4a657c; font-size: 13px; line-height: 1.55; }
.farm-to-store .eyebrow {
  margin: 0;
  text-transform: uppercase;
  letter-spacing: 0.14em;
  font-size: 11px;
  font-weight: 700;
  color: #0284c7;
}
.fts-filters { display: flex; flex-wrap: wrap; gap: 10px; align-items: flex-end; }
.fts-filters label {
  display: flex; flex-direction: column; gap: 4px; font-size: 10px; text-transform: uppercase;
  letter-spacing: 0.08em; color: #3d6b8c; font-weight: 700;
}
.fts-filters select, .fts-filters input {
  min-width: 160px; padding: 9px 12px; border-radius: 11px;
  border: 1px solid rgba(15, 70, 110, 0.18);
  background: #ffffff; color: #0f2740;
  box-shadow: 0 1px 2px rgba(15, 40, 70, 0.04);
}
.fts-kpi-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
  gap: 10px; margin-bottom: 16px;
}
.fts-kpi {
  border-radius: 14px; padding: 12px 14px;
  background: #ffffff;
  border: 1px solid rgba(15, 70, 110, 0.1);
  box-shadow: 0 6px 18px rgba(15, 40, 70, 0.05);
}
.fts-kpi span {
  display: block; font-size: 10px; text-transform: uppercase; letter-spacing: 0.06em;
  color: #5a7a94; margin-bottom: 4px; font-weight: 600;
}
.fts-kpi strong { font-size: 1.15rem; color: #0b2540; font-variant-numeric: tabular-nums; }
.fts-stages {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 8px; margin-bottom: 16px;
}
.fts-stage {
  text-align: left; border-radius: 12px; padding: 10px 12px; cursor: pointer;
  border: 1px solid rgba(15, 70, 110, 0.1);
  background: #ffffff; color: #1e3a50;
  transition: border-color 0.15s, background 0.15s, box-shadow 0.15s, transform 0.12s;
}
.fts-stage:hover {
  border-color: rgba(2, 132, 199, 0.35);
  transform: translateY(-1px);
  box-shadow: 0 6px 16px rgba(2, 132, 199, 0.08);
}
.fts-stage strong { display: block; font-size: 12px; margin-bottom: 2px; color: #0b2540; }
.fts-stage.active {
  border-color: rgba(2, 132, 199, 0.45);
  background: linear-gradient(135deg, #e0f2fe 0%, #f0f9ff 100%);
  box-shadow: 0 0 0 1px rgba(2, 132, 199, 0.15), 0 8px 20px rgba(2, 132, 199, 0.1);
}
.fts-panel {
  border-radius: 18px;
  border: 1px solid rgba(15, 70, 110, 0.1);
  background: #ffffff;
  padding: 16px 16px 20px;
  box-shadow: 0 12px 36px rgba(15, 40, 70, 0.06);
}
.fts-panel-head h3 { margin: 0 0 4px; color: #0b2540; font-size: 1.08rem; }
.fts-panel-head p { margin: 0 0 12px; color: #5a7a94; font-size: 12.5px; }
.farm-to-store .table-wrap {
  overflow: auto; max-height: min(52vh, 520px); border-radius: 14px;
  border: 1px solid rgba(15, 70, 110, 0.12);
  background: #fbfcfe;
}
.farm-to-store table { width: max-content; min-width: 100%; border-collapse: collapse; }
.farm-to-store th, .farm-to-store td {
  padding: 9px 11px; font-size: 12px; text-align: left; white-space: nowrap;
  border-bottom: 1px solid rgba(15, 70, 110, 0.08); color: #1a3348;
}
.farm-to-store th {
  position: sticky; top: 0; z-index: 5;
  background: linear-gradient(180deg, #0f3a56 0%, #144866 100%);
  text-transform: uppercase; letter-spacing: 0.05em; font-size: 10px; color: #e8f4fc;
}
.farm-to-store tbody tr:nth-child(even) td { background: #f5f9fc; }
.farm-to-store tbody tr:hover td { background: #e8f4fc; }
.farm-to-store .pill {
  display: inline-flex; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700;
  background: #e0f2fe; color: #0369a1;
}
.farm-to-store .pill.released, .farm-to-store .pill.available, .farm-to-store .pill.active,
.farm-to-store .pill.settled, .farm-to-store .pill.inspected, .farm-to-store .pill.approved, .farm-to-store .pill.received
{ background: #d1fae5; color: #047857; }
.farm-to-store .pill.hold, .farm-to-store .pill.quarantine, .farm-to-store .pill.under-review,
.farm-to-store .pill.open, .farm-to-store .pill.pending, .farm-to-store .pill.partial, .farm-to-store .pill.in-progress
{ background: #fef3c7; color: #b45309; }
.farm-to-store .pill.rejected, .farm-to-store .pill.suspended
{ background: #fee2e2; color: #b91c1c; }
.farm-to-store .mini-btn {
  border: 1px solid rgba(2, 132, 199, 0.35); background: #f0f9ff;
  color: #0369a1; border-radius: 8px; padding: 5px 10px; font-size: 11px; font-weight: 700;
  cursor: pointer;
}
.farm-to-store .mini-btn:hover { background: #e0f2fe; border-color: #0284c7; }
.fts-ready-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
@media (max-width: 800px) { .fts-ready-grid { grid-template-columns: 1fr; } }
.fts-list { list-style: none; margin: 0; padding: 0; }
.fts-list li {
  padding: 10px 12px; border-radius: 10px; margin-bottom: 8px;
  background: #f5f9fc; border: 1px solid rgba(15, 70, 110, 0.08);
  display: flex; flex-direction: column; gap: 2px; color: #1a3348; font-size: 13px;
}
.fts-list li span { font-size: 11px; color: #5a7a94; }
.fts-trace {
  margin-top: 4px; padding: 14px 16px; border-radius: 14px;
  border: 1px solid rgba(2, 132, 199, 0.18);
  background: linear-gradient(145deg, #f0f9ff, #ffffff);
}
.fts-trace h4 { margin: 0 0 8px; color: #0369a1; }
.fts-trace ol { margin: 0; padding-left: 18px; color: #1e3a50; font-size: 13px; line-height: 1.55; }
.fts-modal-backdrop {
  position: fixed !important; inset: 0 !important; z-index: 2147483000 !important;
  background: rgba(15, 30, 50, 0.4); backdrop-filter: blur(6px);
  display: flex; align-items: center; justify-content: center; padding: 16px;
  isolation: isolate;
}
/* Contract / payment / weekly form modals */
.modal-backdrop {
  position: fixed !important; inset: 0 !important; z-index: 2147483000 !important;
  background: rgba(15, 30, 50, 0.4); backdrop-filter: blur(6px);
  display: flex; align-items: center; justify-content: center; padding: 16px;
  isolation: isolate;
}
.modal-card {
  width: min(720px, 100%); max-height: 90vh; overflow: auto;
  border-radius: 18px; padding: 18px 20px;
  background: #ffffff;
  border: 1px solid rgba(15, 70, 110, 0.14);
  box-shadow: 0 28px 70px rgba(15, 40, 70, 0.22);
  position: relative; z-index: 2147483001;
}
.modal-header {
  display: flex; align-items: flex-start; justify-content: space-between; gap: 12px;
  margin-bottom: 14px;
}
.modal-header h3 { margin: 0; color: #0b2540; }
.close-button {
  border: none; background: #e8f1f8; color: #0b2540; width: 36px; height: 36px;
  border-radius: 10px; font-size: 20px; line-height: 1; cursor: pointer;
}
.fts-modal {
  width: min(720px, 100%); max-height: 90vh; overflow: auto;
  border-radius: 18px; padding: 18px 20px;
  background: #ffffff;
  border: 1px solid rgba(15, 70, 110, 0.14);
  box-shadow: 0 28px 70px rgba(15, 40, 70, 0.18);
}
.fts-modal h3 { margin: 0 0 12px; color: #0b2540; }
.fts-form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.fts-field {
  display: flex; flex-direction: column; gap: 4px; font-size: 11px;
  color: #3d6b8c; text-transform: uppercase; letter-spacing: 0.04em; font-weight: 600;
}
.fts-field input, .fts-field textarea, .fts-field select {
  padding: 8px 10px; border-radius: 8px;
  border: 1px solid rgba(15, 70, 110, 0.18);
  background: #f8fbfd; color: #0f2740; font-size: 13px;
  text-transform: none; letter-spacing: 0; font-weight: 400;
}
.fts-field input:focus, .fts-field textarea:focus, .fts-field select:focus {
  outline: none; border-color: #0284c7;
  box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.15);
  background: #ffffff;
}
.fts-form-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 14px; }
.fts-export-wrap { position: relative; }
.fts-export-menu {
  position: absolute; right: 0; top: 100%; margin-top: 6px; z-index: 50;
  min-width: 220px; border-radius: 12px; padding: 6px;
  background: #ffffff; border: 1px solid rgba(15, 70, 110, 0.14);
  box-shadow: 0 16px 40px rgba(15, 40, 70, 0.12);
}
.fts-export-menu button {
  display: block; width: 100%; text-align: left; border: 0; background: transparent;
  color: #0f2740; padding: 10px 12px; border-radius: 8px; font-size: 12px; font-weight: 600;
}
.fts-export-menu button:hover { background: #e0f2fe; color: #0369a1; }
@media (max-width: 640px) { .fts-form-grid { grid-template-columns: 1fr; } }
`}</style>
  );
}


type StageId = 'plan' | 'supplier' | 'purchase' | 'lot' | 'process' | 'ready';

const etb = (n: number) =>
  `${Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} ETB`;
const kg = (n: number) =>
  `${Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} kg`;
const mt = (n: number) =>
  `${(Number(n || 0) / 1000).toLocaleString(undefined, { maximumFractionDigits: 2 })} MT`;

type Props = {
  canEdit?: boolean;
};


function FarmToStorePanel({ canEdit = true }: { canEdit?: boolean }): JSX.Element {
  type OpTab =
    | 'plan' | 'supplier' | 'purchase' | 'lot' | 'process'
    | 'allocate' | 'docs' | 'logistics' | 'proceeds' | 'nbe' | 'claims' | 'reports';

  type Allocation = { id: string; contractNo: string; lotNo: string; commodity: Commodity; kg: number; status: string; date: string };
  type DocRow = { id: string; shipmentNo: string; docType: string; status: string; issueDate: string; owner: string; notes: string };
  type LogisticsRow = { id: string; shipmentNo: string; contractNo: string; vessel: string; container: string; seal: string; etd: string; eta: string; status: string; dryPort: string };
  type ProceedRow = { id: string; invoiceNo: string; contractNo: string; amountUsd: number; receivedUsd: number; method: string; dueDate: string; status: string };
  type NbeRow = { id: string; ref: string; contractNo: string; deadline: string; action: string; owner: string; evidence: string; status: string };
  type ClaimRow = { id: string; type: string; ref: string; severity: string; owner: string; dueDate: string; impactUsd: number; status: string };

  const [opTab, setOpTab] = useState<OpTab>('supplier');
  const [commodity, setCommodity] = useState<'All' | Commodity>('Sesame');
  const [q, setQ] = useState('');
  const [exportOpen, setExportOpen] = useState(false);
  const [suppliers, setSuppliers] = useState<Supplier[]>(() => [...suppliersSeed]);
  const [plans, setPlans] = useState<SourcingPlan[]>(() => [...sourcingPlansSeed]);
  const [purchases, setPurchases] = useState<PurchaseReceipt[]>(() => [...purchasesSeed]);
  const [lots, setLots] = useState<StockLot[]>(() => [...lotsSeed]);
  const [batches, setBatches] = useState<ProcessingBatch[]>(() => [...processingSeed]);
  const [allocations, setAllocations] = useState<Allocation[]>([
    { id: 'AL-01', contractNo: 'CTR-SES-2026-08', lotNo: 'LOT-SES-2401-A', commodity: 'Sesame', kg: 12000, status: 'Reserved', date: '2026-09-10' },
    { id: 'AL-02', contractNo: 'CTR-SID-2026-08', lotNo: 'LOT-COF-1101-A', commodity: 'Coffee', kg: 15000, status: 'Reserved', date: '2026-08-25' },
  ]);
  const [docs, setDocs] = useState<DocRow[]>([
    { id: 'DOC-01', shipmentNo: 'SESAME/572/07/2026', docType: 'Commercial Invoice', status: 'Approved', issueDate: '2026-07-20', owner: 'Documentation', notes: 'Matches packing list' },
    { id: 'DOC-02', shipmentNo: 'SESAME/572/07/2026', docType: 'Packing List', status: 'Approved', issueDate: '2026-07-20', owner: 'Documentation', notes: '' },
    { id: 'DOC-03', shipmentNo: 'SESAME/572/07/2026', docType: 'Certificate of Origin', status: 'Pending', issueDate: '', owner: 'Documentation', notes: 'ECTA pending' },
    { id: 'DOC-04', shipmentNo: 'COF-26-088', docType: 'Phytosanitary', status: 'Submitted', issueDate: '2026-08-12', owner: 'Documentation', notes: '' },
    { id: 'DOC-05', shipmentNo: 'SESAME/574/08/2026', docType: 'Bill of Lading draft', status: 'Review', issueDate: '2026-08-18', owner: 'Logistics', notes: 'Seal check' },
  ]);
  const [logistics, setLogistics] = useState<LogisticsRow[]>([
    { id: 'LG-01', shipmentNo: 'SESAME/572/07/2026', contractNo: 'CTR-SES-2026-07', vessel: 'CMA CGM', container: 'CMAU1234567', seal: 'SL-9921', etd: '2026-07-28', eta: '2026-08-22', status: 'On vessel', dryPort: 'Modjo' },
    { id: 'LG-02', shipmentNo: 'SESAME/574/08/2026', contractNo: 'CTR-SES-2026-08', vessel: 'NOT GIVEN', container: 'TBD', seal: '—', etd: '2026-08-30', eta: '2026-09-25', status: 'Booking', dryPort: 'Modjo' },
    { id: 'LG-03', shipmentNo: 'COF-26-088', contractNo: 'CTR-SID-2026-08', vessel: 'MSC', container: 'MSCU7654321', seal: 'SL-4410', etd: '2026-08-20', eta: '2026-09-10', status: 'At port', dryPort: 'Modjo' },
  ]);
  const [proceeds, setProceeds] = useState<ProceedRow[]>([
    { id: 'PRC-01', invoiceNo: 'INV-SES-572', contractNo: 'CTR-SES-2026-07', amountUsd: 235600, receivedUsd: 0, method: 'LC', dueDate: '2026-08-15', status: 'Open' },
    { id: 'PRC-02', invoiceNo: 'INV-COF-088', contractNo: 'CTR-SID-2026-08', amountUsd: 330000, receivedUsd: 125000, method: 'LC', dueDate: '2026-08-30', status: 'Partial' },
    { id: 'PRC-03', invoiceNo: 'INV-SES-581', contractNo: 'CTR-SES-2026-08', amountUsd: 171000, receivedUsd: 171000, method: 'TT', dueDate: '2026-09-01', status: 'Received' },
  ]);
  const [nbeRows, setNbeRows] = useState<NbeRow[]>([
    { id: 'NBE-01', ref: 'NBE-SES-572', contractNo: 'CTR-SES-2026-07', deadline: '2026-09-30', action: 'Repatriation evidence', owner: 'Finance/Treasury', evidence: 'Pending bank advice', status: 'Open' },
    { id: 'NBE-02', ref: 'NBE-COF-088', contractNo: 'CTR-SID-2026-08', deadline: '2026-10-15', action: 'Settlement checklist', owner: 'Finance/Treasury', evidence: 'Partial LC draw filed', status: 'In progress' },
  ]);
  const [claims, setClaims] = useState<ClaimRow[]>([
    { id: 'CL-01', type: 'Quality', ref: 'SESAME/572/07/2026', severity: 'Medium', owner: 'Export Manager', dueDate: '2026-09-05', impactUsd: 2500, status: 'Open' },
    { id: 'CL-02', type: 'Document / LC', ref: 'INV-COF-088', severity: 'High', owner: 'Documentation', dueDate: '2026-08-28', impactUsd: 0, status: 'In progress' },
  ]);

  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState('Add record');
  const [form, setForm] = useState<Record<string, string>>({});
  const [formFields, setFormFields] = useState<{ key: string; label: string; type?: string }[]>([]);
  const [editId, setEditId] = useState<string | null>(null);

  const matchCommodity = <T extends { commodity?: Commodity }>(rows: T[]) =>
    rows.filter((r) => commodity === 'All' || !r.commodity || r.commodity === commodity);
  const matchQ = (text: string) => {
    const s = q.trim().toLowerCase();
    return !s || text.toLowerCase().includes(s);
  };

  const schemas: Record<string, { key: string; label: string; type?: string }[]> = {
    plan: [
      { key: 'id', label: 'Plan ID' }, { key: 'commodity', label: 'Commodity' }, { key: 'origin', label: 'Origin' },
      { key: 'grade', label: 'Grade' }, { key: 'targetMt', label: 'Target MT', type: 'number' },
      { key: 'targetPriceEtbPerKg', label: 'Target ETB/kg', type: 'number' }, { key: 'period', label: 'Period' },
      { key: 'requiredBy', label: 'Required by', type: 'date' }, { key: 'status', label: 'Status' },
    ],
    supplier: [
      { key: 'id', label: 'Supplier ID' }, { key: 'name', label: 'Name' }, { key: 'commodity', label: 'Commodity' },
      { key: 'origin', label: 'Origin' }, { key: 'woreda', label: 'Woreda' }, { key: 'contact', label: 'Contact' },
      { key: 'phone', label: 'Phone' }, { key: 'status', label: 'Status' },
      { key: 'performanceScore', label: 'Performance score', type: 'number' }, { key: 'notes', label: 'Notes', type: 'textarea' },
    ],
    purchase: [
      { key: 'id', label: 'Receipt ID' }, { key: 'poNo', label: 'PO No' }, { key: 'supplierId', label: 'Supplier ID' },
      { key: 'supplierName', label: 'Supplier name' }, { key: 'commodity', label: 'Commodity' },
      { key: 'origin', label: 'Origin' }, { key: 'grade', label: 'Grade' },
      { key: 'receiptDate', label: 'Receipt date', type: 'date' },
      { key: 'grossKg', label: 'Gross kg', type: 'number' }, { key: 'netKg', label: 'Net kg', type: 'number' },
      { key: 'bags', label: 'Bags', type: 'number' }, { key: 'moisturePct', label: 'Moisture %', type: 'number' },
      { key: 'priceEtbPerKg', label: 'Price ETB/kg', type: 'number' },
      { key: 'deductionsEtb', label: 'Deductions ETB', type: 'number' },
      { key: 'status', label: 'Status' }, { key: 'warehouse', label: 'Warehouse' },
    ],
    lot: [
      { key: 'id', label: 'ID' }, { key: 'lotNo', label: 'Lot No' }, { key: 'commodity', label: 'Commodity' },
      { key: 'origin', label: 'Origin' }, { key: 'grade', label: 'Grade' }, { key: 'purchaseId', label: 'Purchase ID' },
      { key: 'warehouse', label: 'Warehouse' }, { key: 'bin', label: 'Bin' },
      { key: 'netKg', label: 'Net kg', type: 'number' }, { key: 'bags', label: 'Bags', type: 'number' },
      { key: 'qualityStatus', label: 'Quality status' }, { key: 'receivedAt', label: 'Received at', type: 'date' },
      { key: 'owner', label: 'Owner' },
    ],
    process: [
      { key: 'id', label: 'ID' }, { key: 'batchNo', label: 'Batch No' }, { key: 'commodity', label: 'Commodity' },
      { key: 'inputLotIds', label: 'Input lot IDs (comma)' }, { key: 'processType', label: 'Process type' },
      { key: 'inputKg', label: 'Input kg', type: 'number' }, { key: 'outputKg', label: 'Output kg', type: 'number' },
      { key: 'lossKg', label: 'Loss kg', type: 'number' }, { key: 'yieldPct', label: 'Yield %', type: 'number' },
      { key: 'outputGrade', label: 'Output grade' }, { key: 'labResult', label: 'Lab result', type: 'textarea' },
      { key: 'releaseStatus', label: 'Release status' }, { key: 'processedAt', label: 'Processed at', type: 'date' },
      { key: 'warehouse', label: 'Warehouse' },
    ],
    allocate: [
      { key: 'id', label: 'ID' }, { key: 'contractNo', label: 'Contract No' }, { key: 'lotNo', label: 'Lot No' },
      { key: 'commodity', label: 'Commodity' }, { key: 'kg', label: 'Kg', type: 'number' },
      { key: 'status', label: 'Status' }, { key: 'date', label: 'Date', type: 'date' },
    ],
    docs: [
      { key: 'id', label: 'ID' }, { key: 'shipmentNo', label: 'Shipment No' }, { key: 'docType', label: 'Document type' },
      { key: 'status', label: 'Status' }, { key: 'issueDate', label: 'Issue date', type: 'date' },
      { key: 'owner', label: 'Owner' }, { key: 'notes', label: 'Notes', type: 'textarea' },
    ],
    logistics: [
      { key: 'id', label: 'ID' }, { key: 'shipmentNo', label: 'Shipment No' }, { key: 'contractNo', label: 'Contract No' },
      { key: 'vessel', label: 'Vessel' }, { key: 'container', label: 'Container' }, { key: 'seal', label: 'Seal' },
      { key: 'etd', label: 'ETD', type: 'date' }, { key: 'eta', label: 'ETA', type: 'date' },
      { key: 'status', label: 'Status' }, { key: 'dryPort', label: 'Dry port' },
    ],
    proceeds: [
      { key: 'id', label: 'ID' }, { key: 'invoiceNo', label: 'Invoice No' }, { key: 'contractNo', label: 'Contract No' },
      { key: 'amountUsd', label: 'Amount USD', type: 'number' }, { key: 'receivedUsd', label: 'Received USD', type: 'number' },
      { key: 'method', label: 'Method (LC/TT)' }, { key: 'dueDate', label: 'Due date', type: 'date' },
      { key: 'status', label: 'Status' },
    ],
    nbe: [
      { key: 'id', label: 'ID' }, { key: 'ref', label: 'NBE Ref' }, { key: 'contractNo', label: 'Contract No' },
      { key: 'deadline', label: 'Deadline', type: 'date' }, { key: 'action', label: 'Required action' },
      { key: 'owner', label: 'Owner' }, { key: 'evidence', label: 'Evidence', type: 'textarea' },
      { key: 'status', label: 'Status' },
    ],
    claims: [
      { key: 'id', label: 'ID' }, { key: 'type', label: 'Type' }, { key: 'ref', label: 'Reference' },
      { key: 'severity', label: 'Severity' }, { key: 'owner', label: 'Owner' },
      { key: 'dueDate', label: 'Due date', type: 'date' }, { key: 'impactUsd', label: 'Impact USD', type: 'number' },
      { key: 'status', label: 'Status' },
    ],
  };

  const rowToForm = (row: Record<string, unknown>) => {
    const o: Record<string, string> = {};
    Object.keys(row).forEach((k) => {
      const v = row[k];
      if (Array.isArray(v)) o[k] = v.join(', ');
      else o[k] = v == null ? '' : String(v);
    });
    return o;
  };

  const openAdd = (tab: OpTab) => {
    if (tab === 'reports') return;
    const fields = schemas[tab] || [];
    if (!fields.length) {
      console.warn('[FarmToStore] No schema for tab', tab);
      return;
    }
    setFormFields(fields);
    const defaults: Record<string, string> = {};
    fields.forEach((f) => { defaults[f.key] = ''; });
    if ('commodity' in defaults) defaults.commodity = commodity === 'All' ? 'Sesame' : commodity;
    if ('status' in defaults) defaults.status = tab === 'process' ? 'Hold' : tab === 'lot' ? 'Available' : 'Open';
    if ('receiptDate' in defaults || 'date' in defaults || 'receivedAt' in defaults || 'processedAt' in defaults) {
      const d = new Date().toISOString().slice(0, 10);
      if ('receiptDate' in defaults) defaults.receiptDate = d;
      if ('date' in defaults) defaults.date = d;
      if ('receivedAt' in defaults) defaults.receivedAt = d;
      if ('processedAt' in defaults) defaults.processedAt = d;
    }
    setModalTitle(`Add · ${tab}`);
    setEditId(null);
    setForm(defaults);
    setModalOpen(true);
  };

  const openEdit = (tab: OpTab, id: string, row: Record<string, unknown>) => {
    setFormFields(schemas[tab] || []);
    setModalTitle(`Edit · ${tab}`);
    setEditId(id);
    setForm(rowToForm(row));
    setModalOpen(true);
  };

  const num = (k: string, fallback = 0) => {
    const n = Number(form[k]);
    return Number.isFinite(n) ? n : fallback;
  };

  const saveModal = () => {
    const tab = opTab;
    const idGen = () => `${tab.slice(0, 3).toUpperCase()}-${Date.now().toString(36).toUpperCase()}`;

    if (tab === 'plan') {
      const row: SourcingPlan = {
        id: form.id || editId || idGen(), commodity: (form.commodity as Commodity) || 'Sesame',
        origin: form.origin || '', grade: form.grade || '', targetMt: num('targetMt'),
        targetPriceEtbPerKg: num('targetPriceEtbPerKg'), period: form.period || '',
        requiredBy: form.requiredBy || '', status: (form.status as SourcingPlan['status']) || 'Open',
      };
      setPlans((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'supplier') {
      const row: Supplier = {
        id: form.id || editId || idGen(), name: form.name || '', commodity: (form.commodity as Commodity) || 'Sesame',
        origin: form.origin || '', woreda: form.woreda || '', contact: form.contact || '', phone: form.phone || '',
        status: (form.status as Supplier['status']) || 'Active', performanceScore: num('performanceScore', 80),
        notes: form.notes || '',
      };
      setSuppliers((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'purchase') {
      const row: PurchaseReceipt = {
        id: form.id || editId || idGen(), poNo: form.poNo || '', supplierId: form.supplierId || '',
        supplierName: form.supplierName || '', commodity: (form.commodity as Commodity) || 'Sesame',
        origin: form.origin || '', grade: form.grade || '', receiptDate: form.receiptDate || '',
        grossKg: num('grossKg', num('netKg')), netKg: num('netKg'), bags: num('bags'),
        moisturePct: num('moisturePct'), priceEtbPerKg: num('priceEtbPerKg'), deductionsEtb: num('deductionsEtb'),
        status: (form.status as PurchaseReceipt['status']) || 'Received', warehouse: form.warehouse || '',
      };
      setPurchases((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'lot') {
      const row: StockLot = {
        id: form.id || editId || idGen(), lotNo: form.lotNo || form.id || idGen(),
        commodity: (form.commodity as Commodity) || 'Sesame', origin: form.origin || '', grade: form.grade || '',
        purchaseId: form.purchaseId || '', warehouse: form.warehouse || '', bin: form.bin || '',
        netKg: num('netKg'), bags: num('bags'),
        qualityStatus: (form.qualityStatus as StockLot['qualityStatus']) || 'Available',
        receivedAt: form.receivedAt || '', owner: form.owner || 'NSPF',
      };
      setLots((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'process') {
      const row: ProcessingBatch = {
        id: form.id || editId || idGen(), batchNo: form.batchNo || idGen(),
        commodity: (form.commodity as Commodity) || 'Sesame',
        inputLotIds: (form.inputLotIds || '').split(',').map((s) => s.trim()).filter(Boolean),
        processType: form.processType || '', inputKg: num('inputKg'), outputKg: num('outputKg'),
        lossKg: num('lossKg'), yieldPct: num('yieldPct'), outputGrade: form.outputGrade || '',
        labResult: form.labResult || '', releaseStatus: (form.releaseStatus as ProcessingBatch['releaseStatus']) || 'Hold',
        processedAt: form.processedAt || '', warehouse: form.warehouse || '',
      };
      setBatches((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'allocate') {
      const row: Allocation = {
        id: form.id || editId || idGen(), contractNo: form.contractNo || '', lotNo: form.lotNo || '',
        commodity: (form.commodity as Commodity) || 'Sesame', kg: num('kg'), status: form.status || 'Reserved',
        date: form.date || new Date().toISOString().slice(0, 10),
      };
      setAllocations((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'docs') {
      const row: DocRow = {
        id: form.id || editId || idGen(), shipmentNo: form.shipmentNo || '', docType: form.docType || '',
        status: form.status || 'Pending', issueDate: form.issueDate || '', owner: form.owner || '', notes: form.notes || '',
      };
      setDocs((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'logistics') {
      const row: LogisticsRow = {
        id: form.id || editId || idGen(), shipmentNo: form.shipmentNo || '', contractNo: form.contractNo || '',
        vessel: form.vessel || '', container: form.container || '', seal: form.seal || '',
        etd: form.etd || '', eta: form.eta || '', status: form.status || 'Booking', dryPort: form.dryPort || '',
      };
      setLogistics((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'proceeds') {
      const row: ProceedRow = {
        id: form.id || editId || idGen(), invoiceNo: form.invoiceNo || '', contractNo: form.contractNo || '',
        amountUsd: num('amountUsd'), receivedUsd: num('receivedUsd'), method: form.method || 'LC',
        dueDate: form.dueDate || '', status: form.status || 'Open',
      };
      setProceeds((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'nbe') {
      const row: NbeRow = {
        id: form.id || editId || idGen(), ref: form.ref || '', contractNo: form.contractNo || '',
        deadline: form.deadline || '', action: form.action || '', owner: form.owner || '',
        evidence: form.evidence || '', status: form.status || 'Open',
      };
      setNbeRows((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    } else if (tab === 'claims') {
      const row: ClaimRow = {
        id: form.id || editId || idGen(), type: form.type || '', ref: form.ref || '',
        severity: form.severity || 'Medium', owner: form.owner || '', dueDate: form.dueDate || '',
        impactUsd: num('impactUsd'), status: form.status || 'Open',
      };
      setClaims((prev) => editId ? prev.map((r) => r.id === editId ? row : r) : [...prev, row]);
    }
    setModalOpen(false);
  };

  const removeRow = (tab: OpTab, id: string) => {
    if (!canEdit) return;
    if (!window.confirm('Delete this record?')) return;
    const map: Record<string, (cb: (p: any[]) => any[]) => void> = {
      plan: setPlans as any, supplier: setSuppliers as any, purchase: setPurchases as any,
      lot: setLots as any, process: setBatches as any, allocate: setAllocations as any,
      docs: setDocs as any, logistics: setLogistics as any, proceeds: setProceeds as any,
      nbe: setNbeRows as any, claims: setClaims as any,
    };
    map[tab]?.((p) => p.filter((r) => r.id !== id));
  };

  const currentDataset = (): { title: string; headers: string[]; rows: (string | number)[][] } => {
    if (opTab === 'plan') {
      const rows = matchCommodity(plans).filter((s) => matchQ([s.id, s.origin, s.grade].join(' ')));
      return { title: 'Sourcing Plans', headers: ['Plan ID', 'Commodity', 'Origin', 'Grade', 'Target MT', 'Target ETB/kg', 'Period', 'Required by', 'Status'], rows: rows.map((r) => [r.id, r.commodity, r.origin, r.grade, r.targetMt, r.targetPriceEtbPerKg, r.period, r.requiredBy, r.status]) };
    }
    if (opTab === 'supplier') {
      const rows = matchCommodity(suppliers).filter((s) => matchQ([s.name, s.origin, s.id].join(' ')));
      return { title: 'Suppliers / Farms', headers: ['ID', 'Name', 'Commodity', 'Origin', 'Woreda', 'Contact', 'Phone', 'Score', 'Status', 'Notes'], rows: rows.map((r) => [r.id, r.name, r.commodity, r.origin, r.woreda, r.contact, r.phone, r.performanceScore, r.status, r.notes]) };
    }
    if (opTab === 'purchase') {
      const rows = matchCommodity(purchases).filter((s) => matchQ([s.poNo, s.supplierName].join(' ')));
      return { title: 'Purchase Receipts', headers: ['Receipt', 'PO', 'Supplier ID', 'Supplier', 'Commodity', 'Origin', 'Grade', 'Date', 'Gross kg', 'Net kg', 'Bags', 'Moisture %', 'Price ETB/kg', 'Deductions', 'Status', 'Warehouse'], rows: rows.map((r) => [r.id, r.poNo, r.supplierId, r.supplierName, r.commodity, r.origin, r.grade, r.receiptDate, r.grossKg, r.netKg, r.bags, r.moisturePct, r.priceEtbPerKg, r.deductionsEtb, r.status, r.warehouse]) };
    }
    if (opTab === 'lot') {
      const rows = matchCommodity(lots).filter((s) => matchQ([s.lotNo, s.warehouse].join(' ')));
      return { title: 'Warehouse Lots', headers: ['ID', 'Lot No', 'Commodity', 'Origin', 'Grade', 'Purchase', 'Warehouse', 'Bin', 'Net kg', 'Bags', 'Status', 'Received', 'Owner'], rows: rows.map((r) => [r.id, r.lotNo, r.commodity, r.origin, r.grade, r.purchaseId, r.warehouse, r.bin, r.netKg, r.bags, r.qualityStatus, r.receivedAt, r.owner]) };
    }
    if (opTab === 'process') {
      const rows = matchCommodity(batches).filter((s) => matchQ([s.batchNo, s.processType].join(' ')));
      return { title: 'Processing Batches', headers: ['ID', 'Batch', 'Commodity', 'Process', 'Input lots', 'Input kg', 'Output kg', 'Loss kg', 'Yield %', 'Output grade', 'Lab', 'Release', 'Date', 'Warehouse'], rows: rows.map((r) => [r.id, r.batchNo, r.commodity, r.processType, r.inputLotIds.join('; '), r.inputKg, r.outputKg, r.lossKg, r.yieldPct, r.outputGrade, r.labResult, r.releaseStatus, r.processedAt, r.warehouse]) };
    }
    if (opTab === 'allocate') {
      const rows = matchCommodity(allocations).filter((s) => matchQ([s.contractNo, s.lotNo].join(' ')));
      return { title: 'Stock Allocations', headers: ['ID', 'Contract', 'Lot', 'Commodity', 'Kg', 'Status', 'Date'], rows: rows.map((r) => [r.id, r.contractNo, r.lotNo, r.commodity, r.kg, r.status, r.date]) };
    }
    if (opTab === 'docs') {
      const rows = docs.filter((s) => matchQ([s.shipmentNo, s.docType].join(' ')));
      return { title: 'Export Documents', headers: ['ID', 'Shipment', 'Document type', 'Status', 'Issue date', 'Owner', 'Notes'], rows: rows.map((r) => [r.id, r.shipmentNo, r.docType, r.status, r.issueDate, r.owner, r.notes]) };
    }
    if (opTab === 'logistics') {
      const rows = logistics.filter((s) => matchQ([s.shipmentNo, s.vessel].join(' ')));
      return { title: 'Logistics & Shipment', headers: ['ID', 'Shipment', 'Contract', 'Vessel', 'Container', 'Seal', 'ETD', 'ETA', 'Dry port', 'Status'], rows: rows.map((r) => [r.id, r.shipmentNo, r.contractNo, r.vessel, r.container, r.seal, r.etd, r.eta, r.dryPort, r.status]) };
    }
    if (opTab === 'proceeds') {
      const rows = proceeds.filter((s) => matchQ([s.invoiceNo, s.contractNo].join(' ')));
      return { title: 'Export Proceeds', headers: ['ID', 'Invoice', 'Contract', 'Amount USD', 'Received USD', 'Outstanding', 'Method', 'Due', 'Status'], rows: rows.map((r) => [r.id, r.invoiceNo, r.contractNo, r.amountUsd, r.receivedUsd, r.amountUsd - r.receivedUsd, r.method, r.dueDate, r.status]) };
    }
    if (opTab === 'nbe') {
      const rows = nbeRows.filter((s) => matchQ([s.ref, s.contractNo].join(' ')));
      return { title: 'NBE Compliance', headers: ['ID', 'Ref', 'Contract', 'Deadline', 'Action', 'Owner', 'Evidence', 'Status'], rows: rows.map((r) => [r.id, r.ref, r.contractNo, r.deadline, r.action, r.owner, r.evidence, r.status]) };
    }
    if (opTab === 'claims') {
      const rows = claims.filter((s) => matchQ([s.type, s.ref].join(' ')));
      return { title: 'Claims & Exceptions', headers: ['ID', 'Type', 'Ref', 'Severity', 'Owner', 'Due', 'Impact USD', 'Status'], rows: rows.map((r) => [r.id, r.type, r.ref, r.severity, r.owner, r.dueDate, r.impactUsd, r.status]) };
    }
    // reports pack summary
    return {
      title: 'Operations Summary',
      headers: ['Metric', 'Value'],
      rows: [
        ['Suppliers', suppliers.length], ['Open plans', plans.filter((p) => p.status !== 'Closed').length],
        ['Purchases', purchases.length], ['Lots', lots.length], ['Batches', batches.length],
        ['Allocations', allocations.length], ['Documents', docs.length], ['Logistics', logistics.length],
        ['Proceeds open USD', proceeds.reduce((s, r) => s + (r.amountUsd - r.receivedUsd), 0)],
        ['NBE open', nbeRows.filter((r) => r.status !== 'Closed').length],
        ['Claims open', claims.filter((r) => r.status !== 'Closed').length],
      ],
    };
  };

  const org = 'Nefas Silk Paints Factory PLC';
  const stamp = () => new Date().toISOString().slice(0, 19).replace('T', ' ');

  const runTableExport = async (format: 'csv' | 'excel' | 'pdf' | 'word') => {
    const { title, headers, rows } = currentDataset();
    const gen = stamp();
    const fileBase = `NSPF_${title.replace(/[^\w]+/g, '_')}_${new Date().toISOString().slice(0, 10)}`;
    const colCount = Math.max(headers.length, 1);
    const filterNote = `Commodity filter: ${commodity}  ·  Module: Farm → Store operations`;

    if (format === 'csv') {
      const lines = [
        [org],
        [title + ' (Official Record)'],
        [filterNote],
        [`Generated: ${gen}`],
        [],
        headers,
        ...rows.map((r) => r.map(String)),
        [],
        ['Confidential — internal use only · ' + org],
      ];
      const csv = lines.map((line) => line.map((c) => {
        const s = String(c ?? '');
        return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
      }).join(',')).join('\n');
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `${fileBase}.csv`; a.click();
      URL.revokeObjectURL(url);
      return;
    }

    if (format === 'excel') {
      const wb = new ExcelJS.Workbook();
      wb.creator = org;
      wb.created = new Date();
      const ws = wb.addWorksheet('Report', {
        pageSetup: {
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          paperSize: 9,
          margins: { left: 0.55, right: 0.55, top: 0.6, bottom: 0.6, header: 0.25, footer: 0.25 },
          printTitlesRow: '5:5',
        },
        views: [{ state: 'frozen', ySplit: 5, activeCell: 'B6', showGridLines: false }],
      });
      ws.headerFooter = {
        oddHeader: `&L${org}&C${title}&R${gen}`,
        oddFooter: `&LConfidential — internal use only&CPage &P of &N&R${org}`,
      };

      // Column A left margin gutter
      ws.getColumn(1).width = 3;
      headers.forEach((_, i) => {
        ws.getColumn(i + 2).width = Math.min(30, Math.max(12, String(headers[i]).length + 3));
      });
      ws.getColumn(colCount + 2).width = 3;

      const lastDataCol = colCount + 1; // 1-based with gutter
      const mergeBand = (row: number, value: string, fill: string, font: Partial<ExcelJS.Font>, height: number) => {
        ws.mergeCells(row, 2, row, lastDataCol);
        const cell = ws.getCell(row, 2);
        cell.value = value;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        cell.font = font;
        cell.alignment = { vertical: 'middle', horizontal: row === 2 ? 'center' : 'left' };
        ws.getRow(row).height = height;
        for (let c = 2; c <= lastDataCol; c++) {
          ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        }
      };

      mergeBand(1, org, 'FF0F2837', { bold: true, color: { argb: 'FF7AD9FF' }, size: 10, name: 'Calibri' }, 18);
      mergeBand(2, `${title.toUpperCase()} (Official Record)`, 'FF0F2837', { bold: true, color: { argb: 'FFFFFFFF' }, size: 16, name: 'Calibri' }, 28);
      mergeBand(3, `${filterNote}`, 'FF1A3A4D', { color: { argb: 'FFB8D4E8' }, size: 9, name: 'Calibri' }, 18);
      mergeBand(4, `Generated: ${gen}  ·  Rows: ${rows.length}`, 'FF1A3A4D', { color: { argb: 'FF9EC8E0' }, size: 9, name: 'Calibri', italic: true }, 16);

      const headerRow = ws.getRow(5);
      headerRow.height = 22;
      headers.forEach((h, i) => {
        const cell = headerRow.getCell(i + 2);
        cell.value = h;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF173E5F' } };
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFB0C4D4' } },
          left: { style: 'thin', color: { argb: 'FFB0C4D4' } },
          bottom: { style: 'thin', color: { argb: 'FFB0C4D4' } },
          right: { style: 'thin', color: { argb: 'FFB0C4D4' } },
        };
      });

      rows.forEach((r, ri) => {
        const row = ws.getRow(6 + ri);
        r.forEach((val, ci) => {
          const cell = row.getCell(ci + 2);
          cell.value = val as ExcelJS.CellValue;
          cell.font = { size: 10, name: 'Calibri', color: { argb: 'FF102030' } };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: ri % 2 === 0 ? 'FFFFFFFF' : 'FFF0F7FC' },
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD0DDE8' } },
            left: { style: 'thin', color: { argb: 'FFD0DDE8' } },
            bottom: { style: 'thin', color: { argb: 'FFD0DDE8' } },
            right: { style: 'thin', color: { argb: 'FFD0DDE8' } },
          };
          cell.alignment = { vertical: 'middle', wrapText: true };
        });
      });

      const footRow = 6 + rows.length + 1;
      ws.mergeCells(footRow, 2, footRow, lastDataCol);
      const foot = ws.getCell(footRow, 2);
      foot.value = `End of report · ${org} · Export operations · Confidential`;
      foot.font = { size: 9, italic: true, color: { argb: 'FF4A6070' }, name: 'Calibri' };
      foot.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF3' } };
      foot.alignment = { horizontal: 'center' };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `${fileBase}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      return;
    }

    if (format === 'pdf') {
      const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      const pageWidth = doc.internal.pageSize.getWidth();
      const margin = 40;
      doc.setFillColor(15, 40, 55);
      doc.rect(0, 0, pageWidth, 72, 'F');
      doc.setTextColor(122, 217, 255);
      doc.setFontSize(10);
      doc.text(org, margin, 28);
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(14);
      doc.text(`${title} (Official Record)`, margin, 48);
      doc.setFontSize(9);
      doc.setTextColor(184, 212, 232);
      doc.text(`${filterNote}  ·  Generated: ${gen}`, margin, 64);

      autoTable(doc, {
        startY: 84,
        head: [headers],
        body: rows.map((r) => r.map(String)),
        styles: { fontSize: 8, cellPadding: 4, textColor: [16, 32, 48], lineColor: [176, 196, 212], lineWidth: 0.3 },
        headStyles: { fillColor: [23, 62, 95], textColor: 255, fontStyle: 'bold', cellPadding: 5 },
        alternateRowStyles: { fillColor: [240, 247, 252] },
        margin: { left: margin, right: margin, bottom: 48 },
        didDrawPage: () => {
          const h = doc.internal.pageSize.getHeight();
          doc.setDrawColor(176, 196, 212);
          doc.setLineWidth(0.5);
          doc.line(margin, h - 36, pageWidth - margin, h - 36);
          doc.setFontSize(8);
          doc.setTextColor(90, 110, 125);
          doc.text('Confidential — internal use only', margin, h - 22);
          doc.text(org, pageWidth / 2, h - 22, { align: 'center' });
          doc.text(`Page ${doc.getNumberOfPages()}`, pageWidth - margin, h - 22, { align: 'right' });
        },
      });
      doc.save(`${fileBase}.pdf`);
      return;
    }

    // Word — fully margined official layout
    const th = headers.map((h) => `<th style="background:#173E5F;color:#fff;padding:8px 10px;border:1px solid #B0C4D4;text-align:left">${h}</th>`).join('');
    const body = rows.map((r, i) => {
      const bg = i % 2 === 0 ? '#FFFFFF' : '#F0F7FC';
      return `<tr>${r.map((c) => `<td style="padding:7px 10px;border:1px solid #D0DDE8;background:${bg};color:#102030;font-size:10.5pt">${String(c ?? '')}</td>`).join('')}</tr>`;
    }).join('');
    const html = `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8">
<title>${title}</title>
<style>
  @page { size: A4 landscape; margin: 1.6cm 1.5cm 1.6cm 1.5cm; }
  body { font-family: Calibri, Arial, sans-serif; color: #102030; margin: 0; }
  .band { background: #0F2837; color: #fff; padding: 14px 18px; }
  .band .org { color: #7AD9FF; font-size: 10pt; margin: 0 0 4px; }
  .band h1 { margin: 0; font-size: 16pt; }
  .sub { background: #1A3A4D; color: #B8D4E8; padding: 8px 18px; font-size: 9pt; }
  table { width: 100%; border-collapse: collapse; margin: 16px 0 20px; }
  .foot { margin-top: 18px; padding: 10px 14px; background: #E8EEF3; font-size: 9pt; color: #4A6070; text-align: center; }
</style></head><body>
  <div class="band">
    <p class="org">${org}</p>
    <h1>${title} (Official Record)</h1>
  </div>
  <div class="sub">${filterNote} &nbsp;·&nbsp; Generated: ${gen} &nbsp;·&nbsp; Rows: ${rows.length}</div>
  <table>
    <thead><tr>${th}</tr></thead>
    <tbody>${body}</tbody>
  </table>
  <div class="foot">End of report · ${org} · Export operations · Confidential — internal use only</div>
</body></html>`;
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `${fileBase}.doc`; a.click();
    URL.revokeObjectURL(url);
  };

  const kpis = useMemo(() => {
    const p = matchCommodity(purchases);
    const l = matchCommodity(lots);
    const b = matchCommodity(batches);
    return {
      suppliers: matchCommodity(suppliers).length,
      plans: matchCommodity(plans).filter((x) => x.status !== 'Closed').length,
      netReceived: p.reduce((s, x) => s + x.netKg, 0),
      available: l.filter((x) => x.qualityStatus === 'Available').reduce((s, x) => s + x.netKg, 0),
      released: b.filter((x) => x.releaseStatus === 'Released').reduce((s, x) => s + x.outputKg, 0),
      allocations: matchCommodity(allocations).length,
      docsPending: docs.filter((d) => d.status !== 'Approved').length,
      openProceeds: proceeds.filter((x) => x.status !== 'Received').reduce((s, x) => s + (x.amountUsd - x.receivedUsd), 0),
      openNbe: nbeRows.filter((x) => x.status !== 'Closed').length,
      openClaims: claims.filter((x) => x.status !== 'Closed').length,
    };
  }, [purchases, lots, batches, suppliers, plans, allocations, docs, proceeds, nbeRows, claims, commodity]);

  const etb = (n: number) => `${Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} ETB`;
  const kg = (n: number) => `${Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} kg`;
  const mt = (n: number) => `${(Number(n || 0) / 1000).toLocaleString(undefined, { maximumFractionDigits: 2 })} MT`;
  const usd = (n: number) => `$${Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 2 })}`;

  const tabs: { id: OpTab; label: string }[] = [
    { id: 'plan', label: '1. Sourcing plan' }, { id: 'supplier', label: '2. Suppliers' },
    { id: 'purchase', label: '3. Purchase' }, { id: 'lot', label: '4. Warehouse lots' },
    { id: 'process', label: '5. Processing' }, { id: 'allocate', label: '6–7. Allocate' },
    { id: 'docs', label: '8. Documents' }, { id: 'logistics', label: '9. Logistics' },
    { id: 'proceeds', label: '10. Proceeds' }, { id: 'nbe', label: '11. NBE' },
    { id: 'claims', label: 'Claims' }, { id: 'reports', label: 'Reports' },
  ];

  const filteredPlans = matchCommodity(plans).filter((s) => matchQ([s.id, s.origin, s.grade].join(' ')));
  const filteredSuppliers = matchCommodity(suppliers).filter((s) => matchQ([s.name, s.origin, s.id, s.contact].join(' ')));
  const filteredPurchases = matchCommodity(purchases).filter((s) => matchQ([s.poNo, s.supplierName, s.origin].join(' ')));
  const filteredLots = matchCommodity(lots).filter((s) => matchQ([s.lotNo, s.warehouse, s.grade].join(' ')));
  const filteredBatches = matchCommodity(batches).filter((s) => matchQ([s.batchNo, s.processType, s.outputGrade].join(' ')));
  const filteredAlloc = matchCommodity(allocations).filter((s) => matchQ([s.contractNo, s.lotNo].join(' ')));
  const filteredDocs = docs.filter((s) => matchQ([s.shipmentNo, s.docType, s.status].join(' ')));
  const filteredLog = logistics.filter((s) => matchQ([s.shipmentNo, s.vessel, s.container].join(' ')));
  const filteredProceeds = proceeds.filter((s) => matchQ([s.invoiceNo, s.contractNo, s.status].join(' ')));
  const filteredNbe = nbeRows.filter((s) => matchQ([s.ref, s.contractNo, s.action].join(' ')));
  const filteredClaims = claims.filter((s) => matchQ([s.type, s.ref, s.owner].join(' ')));

  const field = (key: string, label: string, type = 'text') => (
    <label key={key} className="fts-field">
      {label}
      {type === 'textarea' ? (
        <textarea value={form[key] || ''} onChange={(e) => setForm((f) => ({ ...f, [key]: e.target.value }))} rows={2} />
      ) : (
        <input type={type} value={form[key] || ''} onChange={(e) => setForm((f) => ({ ...f, [key]: e.target.value }))} />
      )}
    </label>
  );

  const ExportMenu = () => (
    <div className="fts-export-wrap">
      <button type="button" className="mini-btn" onClick={() => setExportOpen((o) => !o)}>Export ▾</button>
      {exportOpen && (
        <div className="fts-export-menu">
          <button type="button" onClick={() => { setExportOpen(false); void runTableExport('excel'); }}>Excel spreadsheet</button>
          <button type="button" onClick={() => { setExportOpen(false); void runTableExport('pdf'); }}>PDF document</button>
          <button type="button" onClick={() => { setExportOpen(false); void runTableExport('word'); }}>Word document</button>
          <button type="button" onClick={() => { setExportOpen(false); void runTableExport('csv'); }}>CSV (raw data)</button>
        </div>
      )}
    </div>
  );

  const actions = (tab: OpTab, id: string, row: Record<string, unknown>) => (
    <td>
      {canEdit && <button type="button" className="mini-btn" onClick={() => openEdit(tab, id, row)}>Edit</button>}
      {' '}
      {canEdit && <button type="button" className="mini-btn" onClick={() => removeRow(tab, id)}>Del</button>}
    </td>
  );

  return (
    <div className="farm-to-store">
      <div className="fts-hero">
        <div>
          <p className="eyebrow">NSPF · PRS end-to-end</p>
          <h2>Farm → store → contract → shipment → proceeds → NBE</h2>
          <p className="fts-sub">
            Full coffee &amp; sesame operations with complete field CRUD and the same export formats as the main board
            (Excel · PDF · Word · CSV with official header).
          </p>
        </div>
        <div className="fts-filters">
          <label>Commodity
            <select value={commodity} onChange={(e) => setCommodity(e.target.value as 'All' | Commodity)}>
              <option value="All">All</option>
              <option value="Sesame">Sesame</option>
              <option value="Coffee">Coffee</option>
            </select>
          </label>
          <label>Search
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search tables…" />
          </label>
        </div>
      </div>

      <div className="fts-kpi-grid">
        <div className="fts-kpi"><span>Open plans</span><strong>{kpis.plans}</strong></div>
        <div className="fts-kpi"><span>Suppliers</span><strong>{kpis.suppliers}</strong></div>
        <div className="fts-kpi"><span>Received</span><strong>{mt(kpis.netReceived)}</strong></div>
        <div className="fts-kpi"><span>Available stock</span><strong>{mt(kpis.available)}</strong></div>
        <div className="fts-kpi"><span>Released</span><strong>{mt(kpis.released)}</strong></div>
        <div className="fts-kpi"><span>Allocations</span><strong>{kpis.allocations}</strong></div>
        <div className="fts-kpi"><span>Docs pending</span><strong>{kpis.docsPending}</strong></div>
        <div className="fts-kpi"><span>Open proceeds</span><strong>{usd(kpis.openProceeds)}</strong></div>
        <div className="fts-kpi"><span>NBE open</span><strong>{kpis.openNbe}</strong></div>
        <div className="fts-kpi"><span>Claims open</span><strong>{kpis.openClaims}</strong></div>
      </div>

      <div className="fts-stages">
        {tabs.map((s) => (
          <button key={s.id} type="button" className={`fts-stage ${opTab === s.id ? 'active' : ''}`} onClick={() => { setOpTab(s.id); setExportOpen(false); }}>
            <strong>{s.label}</strong>
          </button>
        ))}
      </div>

      <div className="fts-panel">
        <div className="fts-panel-head" style={{ display: 'flex', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
          <div>
            <h3>{tabs.find((x) => x.id === opTab)?.label}</h3>
            <p>All fields editable · Export matches main Coffee/Sesame report format</p>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            {opTab !== 'reports' && (
              <button
                type="button"
                className="mini-btn"
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  openAdd(opTab);
                }}
              >
                + Add
              </button>
            )}
            <ExportMenu />
          </div>
        </div>

        {opTab === 'plan' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Plan</th><th>Commodity</th><th>Origin</th><th>Grade</th><th>Target MT</th><th>ETB/kg</th><th>Period</th><th>Required</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody>{filteredPlans.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.commodity}</td><td>{r.origin}</td><td>{r.grade}</td><td>{r.targetMt}</td><td>{r.targetPriceEtbPerKg}</td><td>{r.period}</td><td>{r.requiredBy}</td><td><span className="pill">{r.status}</span></td>{actions('plan', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'supplier' && (
          <div className="table-wrap"><table>
            <thead><tr><th>ID</th><th>Name</th><th>Commodity</th><th>Origin</th><th>Woreda</th><th>Contact</th><th>Phone</th><th>Score</th><th>Status</th><th>Notes</th><th>Actions</th></tr></thead>
            <tbody>{filteredSuppliers.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.name}</td><td>{r.commodity}</td><td>{r.origin}</td><td>{r.woreda}</td><td>{r.contact}</td><td>{r.phone}</td><td>{r.performanceScore}</td><td><span className="pill">{r.status}</span></td><td>{r.notes}</td>{actions('supplier', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'purchase' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Receipt</th><th>PO</th><th>Supplier</th><th>Commodity</th><th>Origin/Grade</th><th>Date</th><th>Gross</th><th>Net kg</th><th>Bags</th><th>Moist%</th><th>Price</th><th>Deduct</th><th>Value</th><th>Status</th><th>Warehouse</th><th>Actions</th></tr></thead>
            <tbody>{filteredPurchases.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.poNo}</td><td>{r.supplierName}</td><td>{r.commodity}</td><td>{r.origin} · {r.grade}</td><td>{r.receiptDate}</td><td>{kg(r.grossKg)}</td><td>{kg(r.netKg)}</td><td>{r.bags}</td><td>{r.moisturePct}</td><td>{r.priceEtbPerKg}</td><td>{etb(r.deductionsEtb)}</td><td>{etb(r.netKg * r.priceEtbPerKg - r.deductionsEtb)}</td><td><span className="pill">{r.status}</span></td><td>{r.warehouse}</td>{actions('purchase', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'lot' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Lot</th><th>Commodity</th><th>Origin/Grade</th><th>Purchase</th><th>Warehouse</th><th>Bin</th><th>Net kg</th><th>Bags</th><th>Status</th><th>Received</th><th>Owner</th><th>Actions</th></tr></thead>
            <tbody>{filteredLots.map((r) => (
              <tr key={r.id}><td>{r.lotNo}</td><td>{r.commodity}</td><td>{r.origin} · {r.grade}</td><td>{r.purchaseId}</td><td>{r.warehouse}</td><td>{r.bin}</td><td>{kg(r.netKg)}</td><td>{r.bags}</td><td><span className="pill">{r.qualityStatus}</span></td><td>{r.receivedAt}</td><td>{r.owner}</td>{actions('lot', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'process' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Batch</th><th>Commodity</th><th>Process</th><th>Input lots</th><th>In</th><th>Out</th><th>Loss</th><th>Yield</th><th>Grade</th><th>Lab</th><th>Release</th><th>Date</th><th>Store</th><th>Actions</th></tr></thead>
            <tbody>{filteredBatches.map((r) => (
              <tr key={r.id}><td>{r.batchNo}</td><td>{r.commodity}</td><td>{r.processType}</td><td>{r.inputLotIds.join(', ')}</td><td>{kg(r.inputKg)}</td><td>{kg(r.outputKg)}</td><td>{kg(r.lossKg)}</td><td>{r.yieldPct.toFixed(1)}%</td><td>{r.outputGrade}</td><td>{r.labResult}</td><td><span className="pill">{r.releaseStatus}</span></td><td>{r.processedAt}</td><td>{r.warehouse}</td>{actions('process', r.id, { ...r, inputLotIds: r.inputLotIds.join(', ') })}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'allocate' && (
          <div className="table-wrap"><table>
            <thead><tr><th>ID</th><th>Contract</th><th>Lot</th><th>Commodity</th><th>Kg</th><th>Status</th><th>Date</th><th>Actions</th></tr></thead>
            <tbody>{filteredAlloc.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.contractNo}</td><td>{r.lotNo}</td><td>{r.commodity}</td><td>{kg(r.kg)}</td><td><span className="pill">{r.status}</span></td><td>{r.date}</td>{actions('allocate', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'docs' && (
          <div className="table-wrap"><table>
            <thead><tr><th>ID</th><th>Shipment</th><th>Document</th><th>Status</th><th>Issue</th><th>Owner</th><th>Notes</th><th>Actions</th></tr></thead>
            <tbody>{filteredDocs.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.shipmentNo}</td><td>{r.docType}</td><td><span className="pill">{r.status}</span></td><td>{r.issueDate || '—'}</td><td>{r.owner}</td><td>{r.notes}</td>{actions('docs', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'logistics' && (
          <div className="table-wrap"><table>
            <thead><tr><th>ID</th><th>Shipment</th><th>Contract</th><th>Vessel</th><th>Container</th><th>Seal</th><th>ETD</th><th>ETA</th><th>Dry port</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody>{filteredLog.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.shipmentNo}</td><td>{r.contractNo}</td><td>{r.vessel}</td><td>{r.container}</td><td>{r.seal}</td><td>{r.etd}</td><td>{r.eta}</td><td>{r.dryPort}</td><td><span className="pill">{r.status}</span></td>{actions('logistics', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'proceeds' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Invoice</th><th>Contract</th><th>Amount</th><th>Received</th><th>Outstanding</th><th>Method</th><th>Due</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody>{filteredProceeds.map((r) => (
              <tr key={r.id}><td>{r.invoiceNo}</td><td>{r.contractNo}</td><td>{usd(r.amountUsd)}</td><td>{usd(r.receivedUsd)}</td><td>{usd(r.amountUsd - r.receivedUsd)}</td><td>{r.method}</td><td>{r.dueDate}</td><td><span className="pill">{r.status}</span></td>{actions('proceeds', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'nbe' && (
          <div className="table-wrap"><table>
            <thead><tr><th>Ref</th><th>Contract</th><th>Deadline</th><th>Action</th><th>Owner</th><th>Evidence</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody>{filteredNbe.map((r) => (
              <tr key={r.id}><td>{r.ref}</td><td>{r.contractNo}</td><td>{r.deadline}</td><td>{r.action}</td><td>{r.owner}</td><td>{r.evidence}</td><td><span className="pill">{r.status}</span></td>{actions('nbe', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'claims' && (
          <div className="table-wrap"><table>
            <thead><tr><th>ID</th><th>Type</th><th>Ref</th><th>Severity</th><th>Owner</th><th>Due</th><th>Impact</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody>{filteredClaims.map((r) => (
              <tr key={r.id}><td>{r.id}</td><td>{r.type}</td><td>{r.ref}</td><td>{r.severity}</td><td>{r.owner}</td><td>{r.dueDate}</td><td>{usd(r.impactUsd)}</td><td><span className="pill">{r.status}</span></td>{actions('claims', r.id, r as unknown as Record<string, unknown>)}</tr>
            ))}</tbody>
          </table></div>
        )}
        {opTab === 'reports' && (
          <div className="fts-ready-grid">
            <div className="fts-trace">
              <h4>Cross-functional reports</h4>
              <ul className="fts-list">
                <li><strong>Stock coverage</strong><span>Available {mt(kpis.available)} · Released {mt(kpis.released)}</span></li>
                <li><strong>Proceeds</strong><span>Open {usd(kpis.openProceeds)}</span></li>
                <li><strong>Compliance</strong><span>NBE open {kpis.openNbe} · Claims {kpis.openClaims} · Docs pending {kpis.docsPending}</span></li>
              </ul>
              <p style={{ marginTop: 12, color: '#9ec8e0', fontSize: 12 }}>Use <strong>Export ▾</strong> for Excel / PDF / Word / CSV of the summary or switch tabs to export each register with full columns.</p>
            </div>
            <div className="fts-trace">
              <h4>Traceability chain</h4>
              <ol>
                <li>Supplier / farm</li><li>Purchase receipt</li><li>Warehouse lot</li>
                <li>Processing & lab release</li><li>Contract allocation</li>
                <li>Documents & logistics</li><li>Proceeds & NBE closure</li>
              </ol>
            </div>
          </div>
        )}
      </div>

      {modalOpen && createPortal(
        <div
          data-uu-export-modal="1"
          className="fts-modal-backdrop"
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            inset: 0,
            zIndex: 2147483646,
            background: 'rgba(15, 30, 50, 0.55)',
            backdropFilter: 'blur(8px)',
            WebkitBackdropFilter: 'blur(8px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 16,
            margin: 0,
            pointerEvents: 'auto',
          }}
          onClick={() => setModalOpen(false)}
          role="presentation"
        >
          <div
            className="fts-modal"
            style={{
              width: 'min(720px, calc(100vw - 24px))',
              maxHeight: 'min(90vh, 900px)',
              overflow: 'auto',
              borderRadius: 18,
              padding: '18px 20px',
              background: '#ffffff',
              border: '1px solid rgba(15, 70, 110, 0.14)',
              boxShadow: '0 28px 70px rgba(15, 40, 70, 0.35)',
              position: 'relative',
              zIndex: 2147483647,
              pointerEvents: 'auto',
            }}
            onClick={(e) => e.stopPropagation()}
            onMouseDown={(e) => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <h3 style={{ margin: '0 0 12px', color: '#0b2540' }}>{modalTitle}</h3>
            <div className="fts-form-grid">
              {formFields.map((f) => field(f.key, f.label, f.type || 'text'))}
            </div>
            <div className="fts-form-actions">
              <button type="button" className="mini-btn" onClick={() => setModalOpen(false)}>Cancel</button>
              <button type="button" className="mini-btn" onClick={saveModal}>Save</button>
            </div>
          </div>
        </div>,
        document.body,
      )}
    </div>
  );
}


/* ========== end Farm → Store ========== */

export default function ExportApp({
  canViewContract = true,
  canEditContract = true,
  canAdd = true,
  canEdit = true,
  canDelete = true,
  canUploadData = true,
  canExportExcel = true,
  canExportPdf = true,
  canExportWord = true,
  canExportCsv = true,
  canViewCoffee = true,
  canViewSesame = true,
}: ExportAppProps = {}): JSX.Element {
  const canExportAny = canExportExcel || canExportPdf || canExportWord || canExportCsv;
  const initialTab: ExportTab = canViewCoffee ? 'coffee' : canViewSesame ? 'sesame' : canViewContract ? 'contract-flow' : 'coffee';
  const [activeTab, setActiveTab] = useState<ExportTab>(initialTab);
  const [rows, setRows] = useState<ExportRow[]>(dataMap.coffee);
  const [selectedRow, setSelectedRow] = useState<ExportRow | null>(dataMap.coffee[0] ?? null);
  const [apiState, setApiState] = useState<'local' | 'postgres'>('local');
  const [search, setSearch] = useState('');
  const [buyerFilter, setBuyerFilter] = useState('All');
  const [destinationFilter, setDestinationFilter] = useState('All');
  const [paymentFilter, setPaymentFilter] = useState('All');
  const [cleaningFilter, setCleaningFilter] = useState('All');
  const [chartType, setChartType] = useState<'line' | 'bar' | 'pie'>('pie');
  const [showForm, setShowForm] = useState(false);
  const [editingRow, setEditingRow] = useState<ExportRow | null>(null);
  const [form, setForm] = useState<Partial<ExportRow>>(defaultRecordForm('coffee'));
  const [saving, setSaving] = useState(false);
  const [notice, setNotice] = useState('');

  /* ── Contract flow state (isolated from coffee/sesame) ── */
  const [contracts, setContracts] = useState<ContractRecord[]>(() => seedContracts());
  const [selectedContractId, setSelectedContractId] = useState<string | null>(() => seedContracts()[0]?.id ?? null);
  const [showContractForm, setShowContractForm] = useState(false);
  const [editingContract, setEditingContract] = useState<ContractRecord | null>(null);
  const [contractForm, setContractForm] = useState<Partial<ContractRecord>>(emptyContractForm());
  const [contractSaving, setContractSaving] = useState(false);
  const [showPaymentForm, setShowPaymentForm] = useState(false);
  const [paymentForm, setPaymentForm] = useState(emptyPaymentForm());
  const [contractNotice, setContractNotice] = useState('');
  const [uploadBusy, setUploadBusy] = useState(false);
  const [dataUploadBusy, setDataUploadBusy] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [notifications, setNotifications] = useState<
    Array<{ id: string; tone: 'info' | 'success' | 'warn' | 'error'; text: string; at: string; read: boolean }>
  >([]);
  const [bellOpen, setBellOpen] = useState(false);

  const playNotifySound = () => {
    try {
      const Ctx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
      const ctx = new Ctx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(880, ctx.currentTime);
      osc.frequency.exponentialRampToValueAtTime(440, ctx.currentTime + 0.15);
      gain.gain.setValueAtTime(0.12, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start();
      osc.stop(ctx.currentTime + 0.36);
    } catch {
      /* autoplay / unsupported */
    }
  };

  const pushNotify = (text: string, tone: 'info' | 'success' | 'warn' | 'error' = 'info') => {
    const item = {
      id: `n-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      tone,
      text,
      at: new Date().toLocaleString(),
      read: false,
    };
    setNotifications((prev) => [item, ...prev].slice(0, 40));
    playNotifySound();
  };

  const unreadCount = notifications.filter((n) => !n.read).length;
  const markAllRead = () => setNotifications((prev) => prev.map((n) => ({ ...n, read: true })));
  const clearNotifications = () => setNotifications([]);


  const selectedContract = useMemo(
    () => contracts.find((c) => c.id === selectedContractId) ?? contracts[0] ?? null,
    [contracts, selectedContractId],
  );

  const contractPaymentTotals = useMemo(() => {
    if (!selectedContract) return { total: 0, paid: 0, open: 0 };
    const total = selectedContract.payments.reduce((s, p) => s + Number(p.amountUsd || 0), 0);
    const paid = selectedContract.payments
      .filter((p) => p.status === 'Paid')
      .reduce((s, p) => s + Number(p.amountUsd || 0), 0);
    const partial = selectedContract.payments
      .filter((p) => p.status === 'Partial')
      .reduce((s, p) => s + Number(p.amountUsd || 0) * 0.5, 0);
    return { total, paid: paid + partial, open: Math.max(0, total - paid - partial) };
  }, [selectedContract]);

  const openCreateContract = (): void => {
    setEditingContract(null);
    setContractForm(emptyContractForm());
    setShowContractForm(true);
    setContractNotice('');
  };

  const openEditContract = (row: ContractRecord): void => {
    setEditingContract(row);
    setContractForm({ ...row });
    setShowContractForm(true);
    setContractNotice('');
  };

  const closeContractForm = (): void => {
    setShowContractForm(false);
    setEditingContract(null);
    setContractForm(emptyContractForm());
  };

  const handleContractField = (
    event: ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>,
  ): void => {
    const { name, value } = event.target;
    setContractForm((prev) => ({ ...prev, [name]: value }));
  };

  const saveContract = async (event: FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    setContractSaving(true);
    try {
      const now = new Date().toISOString().slice(0, 10);
      if (editingContract) {
        const updated: ContractRecord = {
          ...editingContract,
          ...contractForm,
          id: editingContract.id,
          contractNo: String(contractForm.contractNo || editingContract.contractNo),
          title: String(contractForm.title || editingContract.title),
          buyer: String(contractForm.buyer || editingContract.buyer),
          updatedAt: now,
          workflow: editingContract.workflow,
          attachments: editingContract.attachments,
          payments: editingContract.payments,
          roleMatrix: editingContract.roleMatrix,
        } as ContractRecord;
        setContracts((prev) => prev.map((c) => (c.id === updated.id ? updated : c)));
        setSelectedContractId(updated.id);
        setContractNotice('Contract updated.');
        pushNotify(`Contract ${updated.contractNo} updated`, 'success');
      } else {
        const created: ContractRecord = {
          id: `ctr-${Date.now()}`,
          contractNo: String(contractForm.contractNo || `CTR-${Date.now()}`),
          title: String(contractForm.title || 'New contract'),
          buyer: String(contractForm.buyer || 'Buyer TBD'),
          product: String(contractForm.product || ''),
          quantity: String(contractForm.quantity || ''),
          netWeight: String(contractForm.netWeight || ''),
          packing: String(contractForm.packing || ''),
          quality: String(contractForm.quality || ''),
          price: String(contractForm.price || ''),
          terms: String(contractForm.terms || ''),
          shipment: String(contractForm.shipment || ''),
          shipping: String(contractForm.shipping || ''),
          payment: String(contractForm.payment || ''),
          insurance: String(contractForm.insurance || ''),
          qualityApproval: String(contractForm.qualityApproval || ''),
          arbitration: String(contractForm.arbitration || ''),
          note: String(contractForm.note || ''),
          status: (contractForm.status as ContractRecord['status']) || 'Draft',
          approvalTag: String(contractForm.approvalTag || 'Pending review'),
          createdAt: now,
          updatedAt: now,
          workflow: defaultWorkflow().map((s, i) => ({
            ...s,
            status: i === 0 ? 'active' : 'pending',
          })),
          attachments: [],
          payments: [],
          roleMatrix: defaultRoleMatrix(),
        };
        setContracts((prev) => [created, ...prev]);
        setSelectedContractId(created.id);
        setContractNotice('Contract created.');
        pushNotify(`Contract ${created.contractNo} created`, 'success');
      }
      closeContractForm();
    } catch (err) {
      setContractNotice(err instanceof Error ? err.message : 'Save failed');
    } finally {
      setContractSaving(false);
    }
  };

  const deleteContract = (row: ContractRecord): void => {
    if (!window.confirm(`Delete contract ${row.contractNo}?`)) return;
    setContracts((prev) => {
      const next = prev.filter((c) => c.id !== row.id);
      setSelectedContractId(next[0]?.id ?? null);
      return next;
    });
    setContractNotice('Contract deleted.');
    pushNotify(`Contract ${row.contractNo} deleted`, 'warn');
  };

  const advanceWorkflowStep = (stepId: string): void => {
    if (!selectedContract) return;
    setContracts((prev) =>
      prev.map((c) => {
        if (c.id !== selectedContract.id) return c;
        const steps = c.workflow.map((s) => ({ ...s }));
        const idx = steps.findIndex((s) => s.id === stepId);
        if (idx < 0) return c;
        steps[idx] = { ...steps[idx], status: 'done' };
        if (idx + 1 < steps.length && steps[idx + 1].status === 'pending') {
          steps[idx + 1] = { ...steps[idx + 1], status: 'active' };
        }
        const allDone = steps.every((s) => s.status === 'done');
        return {
          ...c,
          workflow: steps,
          status: allDone ? 'Settled' : c.status === 'Draft' ? 'Active' : c.status,
          updatedAt: new Date().toISOString().slice(0, 10),
        };
      }),
    );
    setContractNotice('Workflow step updated.');
    pushNotify('Contract workflow step advanced', 'info');
  };

  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    if (!selectedContract || !event.target.files?.length) return;
    setUploadBusy(true);
    try {
      const files = Array.from(event.target.files);
      const uploaded: ContractAttachment[] = [];
      for (const file of files) {
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result || ''));
          reader.onerror = () => reject(new Error('Read failed'));
          reader.readAsDataURL(file);
        });
        const sizeKb = Math.max(1, Math.round(file.size / 1024));
        uploaded.push({
          id: `att-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          name: file.name,
          type: file.type || 'File',
          size: sizeKb > 1024 ? `${(sizeKb / 1024).toFixed(1)} MB` : `${sizeKb} KB`,
          uploadedAt: new Date().toISOString().slice(0, 10),
          dataUrl,
        });
      }
      setContracts((prev) =>
        prev.map((c) =>
          c.id === selectedContract.id
            ? { ...c, attachments: [...uploaded, ...c.attachments], updatedAt: new Date().toISOString().slice(0, 10) }
            : c,
        ),
      );
      setContractNotice(`${uploaded.length} file(s) uploaded.`);
      pushNotify(`${uploaded.length} attachment(s) uploaded`, 'success');
    } catch (err) {
      setContractNotice(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      setUploadBusy(false);
      event.target.value = '';
    }
  };

  const handleDataExcelUpload = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    if (!canUploadData) {
      setNotice('Your role cannot upload data files.');
      event.target.value = '';
      return;
    }
    if (activeTab === 'contract-flow') {
      setNotice('Switch to Coffee or Sesame to import shipment data.');
      event.target.value = '';
      return;
    }
    const file = event.target.files?.[0];
    if (!file) return;
    setDataUploadBusy(true);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
      const mapped: ExportRow[] = json.map((row, idx) => {
        const g = (keys: string[]) => {
          for (const k of keys) {
            const hit = Object.keys(row).find(
              (h) => h.toLowerCase().replace(/\s+/g, '') === k.toLowerCase().replace(/\s+/g, ''),
            );
            if (hit != null && row[hit] !== '') return row[hit];
          }
          return '';
        };
        const num = (v: unknown) => Number(String(v).replace(/[^0-9.-]/g, '')) || 0;
        const payRaw = String(g(['paymentStatus', 'Payment Status', 'Status']) || 'Open');
        const pay = (/paid/i.test(payRaw) ? 'Paid' : /partial/i.test(payRaw) ? 'Partial' : 'Open') as ExportRow['paymentStatus'];
        return {
          id: Date.now() + idx,
          week: String(g(['week', 'Week']) || ''),
          shipmentNo: String(g(['shipmentNo', 'Shipment No', 'Shipment']) || `IMP-${idx + 1}`),
          buyer: String(g(['buyer', 'Buyer']) || ''),
          destination: String(g(['destination', 'Destination']) || ''),
          quantity: num(g(['quantity', 'Quantity', 'MT'])),
          unit: String(g(['unit', 'Unit']) || 'MT'),
          containers: num(g(['containers', 'Containers'])),
          pricePerKg: num(g(['pricePerKg', 'Price'])),
          grossUsd: num(g(['grossUsd', 'Gross USD', 'Value'])),
          paidUsd: num(g(['paidUsd', 'Paid USD', 'Paid'])),
          paymentStatus: pay,
          cleaning: String(g(['cleaning', 'Cleaning', 'Grade']) || ''),
          product: String(g(['product', 'Product']) || (activeTab === 'coffee' ? 'Green Coffee' : 'Sesame Seeds')),
          vessel: String(g(['vessel', 'Vessel']) || ''),
          contract: String(g(['contract', 'Contract']) || ''),
          status: 'Pending' as ExportRow['status'],
          date: String(g(['date', 'Date']) || new Date().toISOString().slice(0, 10)),
          notes: String(g(['notes', 'Notes']) || `Imported from ${file.name}`),
        };
      });
      if (!mapped.length) throw new Error('No rows found in the spreadsheet.');
      setRows(mapped);
      setSelectedRow(mapped[0] ?? null);
      setNotice(`Imported ${mapped.length} row(s) into ${activeTab} from ${file.name}`);
      pushNotify(`Data upload: ${mapped.length} ${activeTab} rows`, 'success');
    } catch (err) {
      setNotice(err instanceof Error ? err.message : 'Excel import failed');
      pushNotify('Data upload failed', 'error');
    } finally {
      setDataUploadBusy(false);
      event.target.value = '';
    }
  };

  const removeAttachment = (attId: string): void => {
    if (!selectedContract) return;
    setContracts((prev) =>
      prev.map((c) =>
        c.id === selectedContract.id
          ? { ...c, attachments: c.attachments.filter((a) => a.id !== attId) }
          : c,
      ),
    );
    setContractNotice('Attachment removed.');
    pushNotify('Attachment removed', 'warn');
  };

  const savePayment = (event: FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    if (!selectedContract) return;
    const pay: ContractPayment = {
      id: `pay-${Date.now()}`,
      kind: paymentForm.kind,
      amountUsd: Number(paymentForm.amountUsd || 0),
      status: paymentForm.status,
      reference: paymentForm.reference || '—',
      dueDate: paymentForm.dueDate || new Date().toISOString().slice(0, 10),
      paidDate: paymentForm.paidDate || undefined,
      notes: paymentForm.notes || '',
    };
    setContracts((prev) =>
      prev.map((c) =>
        c.id === selectedContract.id
          ? { ...c, payments: [pay, ...c.payments], updatedAt: new Date().toISOString().slice(0, 10) }
          : c,
      ),
    );
    setPaymentForm(emptyPaymentForm());
    setShowPaymentForm(false);
    setContractNotice('Payment line added.');
    pushNotify('Payment line added to contract', 'success');
  };

  const markPaymentPaid = (payId: string): void => {
    if (!selectedContract) return;
    setContracts((prev) =>
      prev.map((c) => {
        if (c.id !== selectedContract.id) return c;
        return {
          ...c,
          payments: c.payments.map((p) =>
            p.id === payId
              ? { ...p, status: 'Paid' as const, paidDate: new Date().toISOString().slice(0, 10) }
              : p,
          ),
          updatedAt: new Date().toISOString().slice(0, 10),
        };
      }),
    );
    setContractNotice('Payment marked as paid.');
    pushNotify('Payment marked as paid', 'success');
  };

  const removePayment = (payId: string): void => {
    if (!selectedContract) return;
    setContracts((prev) =>
      prev.map((c) =>
        c.id === selectedContract.id
          ? { ...c, payments: c.payments.filter((p) => p.id !== payId) }
          : c,
      ),
    );
    setContractNotice('Payment line removed.');
    pushNotify('Payment line removed', 'warn');
  };


  const apiBaseUrl =
    (import.meta.env.VITE_API_BASE_URL as string | undefined)?.replace(/\/$/, '') ||
    (typeof window !== 'undefined' ? `http://${window.location.hostname}:3002` : 'http://localhost:3002');

  useEffect(() => {
    if (activeTab === 'contract-flow') {
      setSelectedRow(null);
      setApiState('local');
      return;
    }

    const fetchBootstrap = async (): Promise<void> => {
      try {
        const response = await fetch(`${apiBaseUrl}/api/bootstrap`);
        if (!response.ok) throw new Error('Backend unavailable');

        const payload = (await response.json()) as { coffee?: ExportRow[]; sesame?: ExportRow[] };
        const coffee = Array.isArray(payload.coffee) ? payload.coffee : coffeeSeed;
        const sesame = Array.isArray(payload.sesame) ? payload.sesame : sesameSeed;

        const nextRows = activeTab === 'coffee' ? coffee : sesame;
        setRows(nextRows);
        setSelectedRow(nextRows[0] ?? null);
        setApiState('postgres');
      } catch {
        const nextRows = dataMap[activeTab === 'coffee' ? 'coffee' : 'sesame'];
        setRows(nextRows);
        setSelectedRow(nextRows[0] ?? null);
        setApiState('local');
      }
    };

    void fetchBootstrap();
  }, [activeTab, apiBaseUrl]);

  useEffect(() => {
    if (activeTab === 'contract-flow') {
      setSearch('');
      setBuyerFilter('All');
      setDestinationFilter('All');
      setPaymentFilter('All');
      setCleaningFilter('All');
      return;
    }

    const nextRows = dataMap[activeTab === 'coffee' ? 'coffee' : 'sesame'];
    setRows(nextRows);
    setSelectedRow(nextRows[0] ?? null);
    setSearch('');
    setBuyerFilter('All');
    setDestinationFilter('All');
    setPaymentFilter('All');
    setCleaningFilter('All');
  }, [activeTab]);

  const buyers = useMemo<string[]>(() => {
    const unique = new Set(rows.map((row) => row.buyer));
    return ['All', ...Array.from(unique)];
  }, [rows]);

  const destinations = useMemo<string[]>(() => {
    const unique = new Set(rows.map((row) => row.destination));
    return ['All', ...Array.from(unique)];
  }, [rows]);

  const paymentOptions = useMemo<string[]>(() => {
    const unique = new Set(rows.map((row) => row.paymentStatus));
    return ['All', ...Array.from(unique)];
  }, [rows]);

  const cleaningOptions = useMemo<string[]>(() => {
    const unique = new Set(rows.map((row) => row.cleaning));
    return ['All', ...Array.from(unique)];
  }, [rows]);

  const filteredRows = useMemo<ExportRow[]>(() => {
    return rows.filter((row) => {
      const query = search.trim().toLowerCase();
      const matchesSearch =
        !query ||
        [row.shipmentNo, row.buyer, row.destination, row.vessel, row.status, row.product]
          .join(' ')
          .toLowerCase()
          .includes(query);

      const matchesBuyer = buyerFilter === 'All' || row.buyer === buyerFilter;
      const matchesDestination = destinationFilter === 'All' || row.destination === destinationFilter;
      const matchesPayment = paymentFilter === 'All' || row.paymentStatus === paymentFilter;
      const matchesCleaning = cleaningFilter === 'All' || row.cleaning === cleaningFilter;

      return matchesSearch && matchesBuyer && matchesDestination && matchesPayment && matchesCleaning;
    });
  }, [rows, search, buyerFilter, destinationFilter, paymentFilter, cleaningFilter]);

  useEffect(() => {
    if (!filteredRows.length) {
      setSelectedRow(null);
      return;
    }

    const firstMatch = filteredRows[0];
    setSelectedRow((current) => (current && filteredRows.some((row) => row.id === current.id) ? current : firstMatch));
  }, [filteredRows]);

  const totals = useMemo(() => {
    const shipmentCount = filteredRows.length;
    const totalMt = filteredRows.reduce((sum, row) => sum + Number(row.quantity || 0), 0);
    const totalContainers = filteredRows.reduce((sum, row) => sum + Number(row.containers || 0), 0);
    const gross = filteredRows.reduce((sum, row) => sum + Number(row.grossUsd || 0), 0);
    const paid = filteredRows.reduce((sum, row) => sum + Number(row.paidUsd || 0), 0);
    const open = gross - paid;

    return { shipmentCount, totalMt, totalContainers, gross, paid, open };
  }, [filteredRows]);

  const openCreateForm = (): void => {
    const tab: ExportTab =
      activeTab === 'coffee' || activeTab === 'sesame' ? activeTab : 'coffee';
    setEditingRow(null);
    setForm(defaultRecordForm(tab));
    setShowForm(true);
    setNotice('');
    try {
      console.log('[ExportApp] openCreateForm', tab);
    } catch (_) {}
  };

  const openEditForm = (row: ExportRow): void => {
    setEditingRow(row);
    setForm({ ...row });
    setShowForm(true);
    setNotice('');
  };

  const closeForm = (): void => {
    setShowForm(false);
    setEditingRow(null);
    setForm(defaultRecordForm(activeTab));
  };

  const handleFormChange = (event: ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>): void => {
    const { name, value } = event.target;
    setForm((previous) => ({
      ...previous,
      [name]: name === 'quantity' || name === 'containers' || name === 'pricePerKg' || name === 'grossUsd' || name === 'paidUsd'
        ? Number(value || 0)
        : value,
    }));
  };

  const saveRecord = async (event: FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    setSaving(true);

    try {
      const nextPayload = {
        ...form,
        id: editingRow?.id ?? Date.now(),
        week: form.week || `2026-W${new Date().getUTCDate()}`,
        shipmentNo: form.shipmentNo || `${activeTab === 'coffee' ? 'COF' : 'SES'}-26-${String(Date.now()).slice(-3)}`,
        buyer: form.buyer || 'New Buyer',
        destination: form.destination || 'New Market',
        product: form.product || (activeTab === 'coffee' ? 'Green Coffee' : 'Sesame Seeds'),
        quantity: Number(form.quantity || 0),
        unit: form.unit || 'MT',
        containers: Number(form.containers || 0),
        pricePerKg: Number(form.pricePerKg || 0),
        grossUsd: Number(form.grossUsd || Number(form.quantity || 0) * Number(form.pricePerKg || 0)),
        paidUsd: Number(form.paidUsd || 0),
        paymentStatus: form.paymentStatus || 'Open',
        cleaning: form.cleaning || (activeTab === 'coffee' ? 'Grade A' : 'Cleaned'),
        vessel: form.vessel || 'TBD',
        contract: form.contract || 'TBD',
        status: form.status || 'Pending',
        date: form.date || new Date().toISOString().slice(0, 10),
        notes: form.notes || 'Added from dashboard',
      };

      const method = editingRow ? 'PUT' : 'POST';
      const url = editingRow
        ? `${apiBaseUrl}/api/${activeTab}/${editingRow.id}`
        : `${apiBaseUrl}/api/${activeTab}`;

      const response = await fetch(url, {
        method,
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer demo-token',
        },
        body: JSON.stringify(nextPayload),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ message: 'Unable to save record' }));
        throw new Error(errorData.message || 'Unable to save record');
      }

      const result = await response.json().catch(() => nextPayload);
      const persisted = result?.record ?? result?.data ?? result ?? nextPayload;
      const mergedRecord = { ...nextPayload, ...persisted, id: persisted?.id ?? nextPayload.id };

      setRows((previous) => {
        if (editingRow) {
          return previous.map((row) => (String(row.id) === String(editingRow.id) ? { ...row, ...mergedRecord } : row));
        }
        return [mergedRecord as ExportRow, ...previous];
      });

      const updatedSelection = editingRow ? mergedRecord as ExportRow : mergedRecord as ExportRow;
      setSelectedRow(updatedSelection);
      closeForm();
      setNotice(editingRow ? 'Record updated successfully.' : 'Record created successfully.');
      pushNotify(editingRow ? `Shipment ${mergedRecord.shipmentNo} updated` : `Shipment ${mergedRecord.shipmentNo} created`, 'success');
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'An unexpected error occurred.');
    } finally {
      setSaving(false);
    }
  };

  const deleteRecord = async (row: ExportRow): Promise<void> => {
    const confirmed = window.confirm(`Delete ${row.shipmentNo}?`);
    if (!confirmed) return;

    try {
      const response = await fetch(`${apiBaseUrl}/api/${activeTab}/${row.id}`, {
        method: 'DELETE',
        headers: {
          Authorization: 'Bearer demo-token',
        },
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ message: 'Unable to delete record' }));
        throw new Error(errorData.message || 'Unable to delete record');
      }

      setRows((previous) => previous.filter((item) => String(item.id) !== String(row.id)));
      setSelectedRow((previous) => (previous && String(previous.id) === String(row.id) ? null : previous));
      setNotice('Record deleted successfully.');
      pushNotify(`Shipment ${row.shipmentNo} deleted`, 'warn');
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Delete failed.');
    }
  };

  const destinationMix = useMemo(() => {
    const totalsByDestination = new Map<string, number>();
    filteredRows.forEach((row) => {
      totalsByDestination.set(row.destination, (totalsByDestination.get(row.destination) ?? 0) + row.grossUsd);
    });

    const totalGross = [...totalsByDestination.values()].reduce((sum, value) => sum + value, 0) || 1;
    return [...totalsByDestination.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([label, value]) => ({
        label,
        value,
        percent: Math.max(12, (value / totalGross) * 100),
      }));
  }, [filteredRows]);

  const weeklyTrend = useMemo<TrendPoint[]>(() => {
    const totalsByWeek = new Map<string, number>();
    filteredRows.forEach((row) => {
      totalsByWeek.set(row.week, (totalsByWeek.get(row.week) ?? 0) + row.grossUsd);
    });

    return [...totalsByWeek.entries()].map(([label, value]) => ({ label, value }));
  }, [filteredRows]);

  const maxTrendValue = weeklyTrend.reduce((max, point) => Math.max(max, point.value), 0) || 1;
  const pieGradient = useMemo(() => {
    if (!destinationMix.length) return 'conic-gradient(#4ec9ff 0 100%)';

    let start = 0;
    const stops = destinationMix.map((item) => {
      const end = start + (item.value / Math.max(1, destinationMix.reduce((sum, current) => sum + current.value, 0))) * 100;
      const color = ['#67d7ff', '#7a8cff', '#82e0a1', '#ffd56a', '#ff8bb4'][destinationMix.indexOf(item) % 5];
      const segment = `${color} ${start}% ${end}%`;
      start = end;
      return segment;
    });

    return `conic-gradient(${stops.join(', ')})`;
  }, [destinationMix]);


  const stamp = () => new Date().toLocaleString();
  const reportOrg = 'Export Operations';

  /** Full workbook dataset: coffee + sesame + contracts for consolidated reports */
  const buildFullFlowSnapshot = () => {
    const coffeeRows = activeTab === 'coffee' ? filteredRows : (dataMap.coffee || []);
    const sesameRows = activeTab === 'sesame' ? filteredRows : (dataMap.sesame || []);
    // When on a trade tab use filtered for that tab; always include both seeds for "complete flow"
    const coffee = activeTab === 'coffee' ? filteredRows : coffeeSeed;
    const sesame = activeTab === 'sesame' ? filteredRows : sesameSeed;
    return {
      generated: stamp(),
      title: 'Complete Export Operations Report',
      subtitle: 'Coffee · Sesame · Contract flow (shipments, payments, workflow)',
      coffee: activeTab === 'coffee' ? filteredRows : coffeeSeed,
      sesame: activeTab === 'sesame' ? filteredRows : sesameSeed,
      // Always attach current filtered view as "active sheet" context
      activeTab,
      filteredShipments: activeTab === 'contract-flow' ? [] : filteredRows,
      contracts,
      selectedContract,
    };
  };

  const shipmentHeaders = [
    'Shipment No', 'Buyer', 'Destination', 'Product', 'Quantity', 'Unit',
    'Containers', 'Gross USD', 'Paid USD', 'Payment Status', 'Cleaning', 'Vessel', 'Status', 'Week', 'Contract', 'Date', 'Notes',
  ];

  const shipmentToCells = (row: ExportRow) => [
    row.shipmentNo, row.buyer, row.destination, row.product,
    row.quantity, row.unit, row.containers, row.grossUsd, row.paidUsd,
    row.paymentStatus, row.cleaning, row.vessel, row.status, row.week,
    row.contract, row.date, row.notes,
  ];

  const sumShipments = (list: ExportRow[]) => {
    const gross = list.reduce((s, r) => s + Number(r.grossUsd || 0), 0);
    const paid = list.reduce((s, r) => s + Number(r.paidUsd || 0), 0);
    const mt = list.reduce((s, r) => s + Number(r.quantity || 0), 0);
    return { count: list.length, mt, gross, paid, open: gross - paid };
  };

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const escapeCsv = (value: unknown) => {
    const s = String(value ?? '');
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };

  /** CSV — current view or full flow */
  const exportCsv = (scope: 'view' | 'full' = 'view'): void => {
    const lines: string[][] = [];
    lines.push([reportOrg]);
    lines.push([scope === 'full' ? 'Complete Export Flow Report' : `${activeTab.toUpperCase()} Export Report`]);
    lines.push(['Generated', stamp()]);
    lines.push([]);

    if (activeTab === 'contract-flow' || scope === 'full') {
      lines.push(['=== CONTRACTS ===']);
      lines.push(['Contract No', 'Title', 'Buyer', 'Status', 'Approval', 'Quantity', 'Net Weight', 'Price', 'Payment Terms', 'Updated']);
      const list = scope === 'full' || activeTab === 'contract-flow' ? contracts : contracts;
      list.forEach((c) => {
        lines.push([
          c.contractNo, c.title, c.buyer, c.status, c.approvalTag,
          c.quantity, c.netWeight, c.price, c.payment, c.updatedAt,
        ]);
      });
      lines.push([]);
      if (selectedContract) {
        lines.push(['=== PAYMENTS — ' + selectedContract.contractNo + ' ===']);
        lines.push(['Kind', 'Reference', 'Amount USD', 'Status', 'Due', 'Paid', 'Notes']);
        selectedContract.payments.forEach((p) => {
          lines.push([p.kind, p.reference, p.amountUsd, p.status, p.dueDate, p.paidDate || '', p.notes]);
        });
        lines.push([]);
        lines.push(['=== WORKFLOW — ' + selectedContract.contractNo + ' ===']);
        lines.push(['Step', 'Owner', 'Status', 'Summary']);
        selectedContract.workflow.forEach((w) => {
          lines.push([w.title, w.owner, w.status, w.summary]);
        });
        lines.push([]);
      }
    }

    if (activeTab !== 'contract-flow' || scope === 'full') {
      const coffeeList = scope === 'full' ? coffeeSeed : (activeTab === 'coffee' ? filteredRows : []);
      const sesameList = scope === 'full' ? sesameSeed : (activeTab === 'sesame' ? filteredRows : []);
      if (scope === 'full' || activeTab === 'coffee') {
        const s = sumShipments(scope === 'full' ? coffeeSeed : filteredRows);
        lines.push(['=== COFFEE SHIPMENTS ===']);
        lines.push(['Shipments', String(s.count), 'Volume MT', String(s.mt), 'Gross', String(s.gross), 'Paid', String(s.paid), 'Open', String(s.open)]);
        lines.push(shipmentHeaders);
        (scope === 'full' ? coffeeSeed : filteredRows).forEach((r) => lines.push(shipmentToCells(r).map(String)));
        lines.push([]);
      }
      if (scope === 'full' || activeTab === 'sesame') {
        const s = sumShipments(scope === 'full' ? sesameSeed : filteredRows);
        lines.push(['=== SESAME SHIPMENTS ===']);
        lines.push(['Shipments', String(s.count), 'Volume MT', String(s.mt), 'Gross', String(s.gross), 'Paid', String(s.paid), 'Open', String(s.open)]);
        lines.push(shipmentHeaders);
        (scope === 'full' ? sesameSeed : filteredRows).forEach((r) => lines.push(shipmentToCells(r).map(String)));
        lines.push([]);
      }
    }

    lines.push(['Footer', reportOrg, 'Confidential — internal use only']);
    const csv = lines.map((row) => row.map(escapeCsv).join(',')).join('\n');
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), `unity-export-${scope === 'full' ? 'complete' : activeTab}-${Date.now()}.csv`);
  };

  /** Excel — print-ready report (highlighted headers, borders, page setup) via ExcelJS */
  const exportExcel = async (scope: 'view' | 'full' = 'view'): Promise<void> => {
    const gen = stamp();
    const title =
      scope === 'full'
        ? 'Complete Export Operations Report'
        : activeTab === 'contract-flow'
          ? 'Contract Flow Report'
          : `${activeTab === 'coffee' ? 'Coffee' : 'Sesame'} Export Report`;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = reportOrg;
    workbook.created = new Date();
    workbook.modified = new Date();

    const HEADER_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF173E5F' },
    };
    const TITLE_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F2837' },
    };
    const SUB_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1A3A4D' },
    };
    const ALT_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF0F7FC' },
    };
    const FOOT_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE8EEF3' },
    };
    const thin: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FFB0C4D4' } },
      left: { style: 'thin', color: { argb: 'FFB0C4D4' } },
      bottom: { style: 'thin', color: { argb: 'FFB0C4D4' } },
      right: { style: 'thin', color: { argb: 'FFB0C4D4' } },
    };
    const whiteBold: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    const titleFont: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 16, name: 'Calibri' };
    const bodyFont: Partial<ExcelJS.Font> = { size: 10, name: 'Calibri', color: { argb: 'FF102030' } };

    const applyPrint = (ws: ExcelJS.Worksheet, cols: number) => {
      ws.pageSetup = {
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize: 9, // A4
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
        printTitlesRow: '5:5',
      };
      ws.headerFooter = {
        oddHeader: `&L${reportOrg}&C${title}&R${gen}`,
        oddFooter: `&LConfidential — internal use only&C&P / &N&R${reportOrg}`,
      };
      ws.views = [{ state: 'frozen', ySplit: 5, activeCell: 'A6', showGridLines: false }];
    };

    const writeTitleBand = (ws: ExcelJS.Worksheet, section: string, colCount: number) => {
      const last = colCount;
      ws.mergeCells(1, 1, 1, last);
      ws.mergeCells(2, 1, 2, last);
      ws.mergeCells(3, 1, 3, last);
      ws.mergeCells(4, 1, 4, last);

      const r1 = ws.getRow(1);
      r1.getCell(1).value = reportOrg;
      r1.getCell(1).font = { bold: true, color: { argb: 'FF7AD9FF' }, size: 10, name: 'Calibri' };
      r1.getCell(1).fill = TITLE_FILL;
      r1.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
      r1.height = 18;

      const r2 = ws.getRow(2);
      r2.getCell(1).value = title;
      r2.getCell(1).font = titleFont;
      r2.getCell(1).fill = TITLE_FILL;
      r2.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
      r2.height = 28;

      const r3 = ws.getRow(3);
      r3.getCell(1).value = `Section: ${section}  ·  Coffee · Sesame · Contract flow  ·  Generated: ${gen}`;
      r3.getCell(1).font = { color: { argb: 'FFB8D4E8' }, size: 9, name: 'Calibri' };
      r3.getCell(1).fill = SUB_FILL;
      r3.height = 18;

      const r4 = ws.getRow(4);
      r4.getCell(1).value = '';
      r4.getCell(1).fill = SUB_FILL;
      r4.height = 6;

      for (let c = 1; c <= last; c += 1) {
        [1, 2, 3, 4].forEach((rn) => {
          const cell = ws.getRow(rn).getCell(c);
          cell.fill = rn <= 2 ? TITLE_FILL : SUB_FILL;
          cell.border = thin;
        });
      }
    };

    const writeColHeaders = (ws: ExcelJS.Worksheet, rowNum: number, headers: string[]) => {
      const row = ws.getRow(rowNum);
      headers.forEach((h, i) => {
        const cell = row.getCell(i + 1);
        cell.value = h;
        cell.font = whiteBold;
        cell.fill = HEADER_FILL;
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = thin;
      });
      row.height = 22;
      row.commit();
    };

    const writeDataRows = (ws: ExcelJS.Worksheet, startRow: number, rows: (string | number | null | undefined)[][]) => {
      rows.forEach((data, idx) => {
        const row = ws.getRow(startRow + idx);
        data.forEach((val, i) => {
          const cell = row.getCell(i + 1);
          cell.value = val ?? '';
          cell.font = bodyFont;
          cell.border = thin;
          cell.alignment = { vertical: 'middle', wrapText: true };
          if (idx % 2 === 1) cell.fill = ALT_FILL;
        });
        row.height = 18;
      });
      return startRow + rows.length;
    };

    const writeFooter = (ws: ExcelJS.Worksheet, rowNum: number, colCount: number, label: string) => {
      ws.mergeCells(rowNum, 1, rowNum, colCount);
      const cell = ws.getRow(rowNum).getCell(1);
      cell.value = `— End of section: ${label} —  ·  ${reportOrg}  ·  Confidential — internal use only  ·  ${gen}`;
      cell.font = { italic: true, size: 9, color: { argb: 'FF5A7080' }, name: 'Calibri' };
      cell.fill = FOOT_FILL;
      cell.border = thin;
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getRow(rowNum).height = 20;
    };

    const setWidths = (ws: ExcelJS.Worksheet, widths: number[]) => {
      widths.forEach((w, i) => {
        ws.getColumn(i + 1).width = w;
      });
    };

    // ——— COVER ———
    {
      const ws = workbook.addWorksheet('01 Cover', {
        properties: { defaultRowHeight: 18 },
      });
      const coffeeList = scope === 'full' || activeTab === 'coffee' ? (scope === 'full' ? coffeeSeed : filteredRows) : [];
      const sesameList = scope === 'full' || activeTab === 'sesame' ? (scope === 'full' ? sesameSeed : filteredRows) : [];
      const cSum = sumShipments(coffeeList);
      const sSum = sumShipments(sesameList);
      writeTitleBand(ws, 'Cover & executive summary', 6);
      writeColHeaders(ws, 5, ['Metric', 'Value', 'Detail', 'Scope', 'Module', 'Notes']);
      const kpis: (string | number)[][] = [
        ['Coffee shipments', cSum.count, `${cSum.mt} MT`, scope, 'Coffee', 'Volume'],
        ['Coffee gross USD', cSum.gross, `Paid ${cSum.paid}`, scope, 'Coffee', `Open ${cSum.open}`],
        ['Sesame shipments', sSum.count, `${sSum.mt} MT`, scope, 'Sesame', 'Volume'],
        ['Sesame gross USD', sSum.gross, `Paid ${sSum.paid}`, scope, 'Sesame', `Open ${sSum.open}`],
        ['Contracts', contracts.length, selectedContract?.contractNo || '—', scope, 'Contract', selectedContract?.status || ''],
        ['Report title', title, gen, scope === 'full' ? 'Complete flow' : activeTab, 'All', reportOrg],
      ];
      const next = writeDataRows(ws, 6, kpis);
      writeColHeaders(ws, next + 1, ['#', 'Contents', '', '', '', '']);
      writeDataRows(ws, next + 2, [
        ['1', 'Cover & executive summary', '', '', '', ''],
        ['2', 'Coffee shipments (detail)', '', '', '', ''],
        ['3', 'Sesame shipments (detail)', '', '', '', ''],
        ['4', 'Contract register', '', '', '', ''],
        ['5', 'Contract payments, workflow & attachments', '', '', '', ''],
      ]);
      writeFooter(ws, next + 8, 6, 'Cover');
      setWidths(ws, [22, 18, 22, 14, 12, 28]);
      applyPrint(ws, 6);
    }

    const addShipSheet = (name: string, label: string, list: ExportRow[]) => {
      const ws = workbook.addWorksheet(name, { properties: { defaultRowHeight: 18 } });
      const s = sumShipments(list);
      const cols = shipmentHeaders.length;
      writeTitleBand(ws, label, cols);
      // summary strip
      const sumHeaders = ['Shipments', 'Volume MT', 'Gross USD', 'Paid USD', 'Open USD', 'Generated'];
      writeColHeaders(ws, 5, [...sumHeaders, ...Array(Math.max(0, cols - sumHeaders.length)).fill('')]);
      writeDataRows(ws, 6, [[s.count, s.mt, s.gross, s.paid, s.open, gen, ...Array(Math.max(0, cols - 6)).fill('')]]);
      writeColHeaders(ws, 8, shipmentHeaders);
      writeDataRows(ws, 9, list.map((r) => shipmentToCells(r)));
      writeFooter(ws, 9 + list.length + 1, cols, label);
      setWidths(ws, shipmentHeaders.map((h) => Math.min(22, Math.max(12, h.length + 2))));
      applyPrint(ws, cols);
      ws.pageSetup.printTitlesRow = '8:8';
    };

    if (scope === 'full' || activeTab === 'coffee') {
      addShipSheet('02 Coffee', 'Coffee shipments', scope === 'full' ? coffeeSeed : filteredRows);
    }
    if (scope === 'full' || activeTab === 'sesame') {
      addShipSheet('03 Sesame', 'Sesame shipments', scope === 'full' ? sesameSeed : filteredRows);
    }

    if (scope === 'full' || activeTab === 'contract-flow') {
      const ws = workbook.addWorksheet('04 Contracts', { properties: { defaultRowHeight: 18 } });
      const headers = [
        'Contract No', 'Title', 'Buyer', 'Status', 'Approval', 'Quantity', 'Net Weight',
        'Quality', 'Price', 'Terms', 'Shipment', 'Payment', 'Updated',
      ];
      writeTitleBand(ws, 'Contract register', headers.length);
      writeColHeaders(ws, 5, headers);
      writeDataRows(
        ws,
        6,
        contracts.map((c) => [
          c.contractNo, c.title, c.buyer, c.status, c.approvalTag,
          c.quantity, c.netWeight, c.quality, c.price, c.terms, c.shipment, c.payment, c.updatedAt,
        ]),
      );
      writeFooter(ws, 6 + contracts.length + 1, headers.length, 'Contracts');
      setWidths(ws, headers.map((h) => Math.min(24, Math.max(12, h.length + 2))));
      applyPrint(ws, headers.length);

      if (selectedContract) {
        const d = workbook.addWorksheet('05 Contract Detail', { properties: { defaultRowHeight: 18 } });
        writeTitleBand(d, `Contract detail — ${selectedContract.contractNo}`, 7);
        writeColHeaders(d, 5, ['Field', 'Value', '', '', '', '', '']);
        let r = writeDataRows(d, 6, [
          ['Title', selectedContract.title, '', '', '', '', ''],
          ['Buyer', selectedContract.buyer, '', '', '', '', ''],
          ['Status', selectedContract.status, '', '', '', '', ''],
          ['Approval', selectedContract.approvalTag, '', '', '', '', ''],
          ['Quantity', selectedContract.quantity, '', '', '', '', ''],
          ['Net weight', selectedContract.netWeight, '', '', '', '', ''],
          ['Quality', selectedContract.quality, '', '', '', '', ''],
          ['Price', selectedContract.price, '', '', '', '', ''],
          ['Terms', selectedContract.terms, '', '', '', '', ''],
          ['Shipment', selectedContract.shipment, '', '', '', '', ''],
          ['Payment terms', selectedContract.payment, '', '', '', '', ''],
          ['Insurance', selectedContract.insurance, '', '', '', '', ''],
          ['Note', selectedContract.note, '', '', '', '', ''],
        ]);
        r += 1;
        writeColHeaders(d, r, ['Kind', 'Reference', 'Amount USD', 'Status', 'Due', 'Paid', 'Notes']);
        r = writeDataRows(
          d,
          r + 1,
          selectedContract.payments.map((p) => [
            p.kind, p.reference, p.amountUsd, p.status, p.dueDate, p.paidDate || '', p.notes,
          ]),
        );
        r += 1;
        writeColHeaders(d, r, ['Step', 'Owner', 'Status', 'Summary', '', '', '']);
        r = writeDataRows(
          d,
          r + 1,
          selectedContract.workflow.map((w) => [w.title, w.owner, w.status, w.summary, '', '', '']),
        );
        r += 1;
        writeColHeaders(d, r, ['Attachment', 'Type', 'Size', 'Uploaded', '', '', '']);
        r = writeDataRows(
          d,
          r + 1,
          selectedContract.attachments.map((a) => [a.name, a.type, a.size, a.uploadedAt, '', '', '']),
        );
        writeFooter(d, r + 1, 7, selectedContract.contractNo);
        setWidths(d, [18, 28, 14, 12, 12, 12, 22]);
        applyPrint(d, 7);
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      `unity-export-report-${scope === 'full' ? 'complete' : activeTab}-${Date.now()}.xlsx`,
    );
    pushNotify(`Excel report downloaded (${scope === 'full' ? 'complete flow' : activeTab})`, 'success');
  };

  /** PDF — formal header / body / footer */
  const exportPdf = (scope: 'view' | 'full' = 'view'): void => {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 40;
    let y = 0;

    const drawHeader = (title: string, subtitle: string) => {
      doc.setFillColor(15, 40, 55);
      doc.rect(0, 0, pageWidth, 56, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.text(title, margin, 24);
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.text(subtitle, margin, 40);
      doc.text(stamp(), pageWidth - margin, 24, { align: 'right' });
      y = 72;
    };

    const drawFooter = () => {
      const pages = doc.getNumberOfPages();
      for (let i = 1; i <= pages; i += 1) {
        doc.setPage(i);
        doc.setDrawColor(180, 200, 220);
        doc.line(margin, pageHeight - 28, pageWidth - margin, pageHeight - 28);
        doc.setFontSize(8);
        doc.setTextColor(90, 110, 125);
        doc.text(reportOrg, margin, pageHeight - 14);
        doc.text(`Page ${i} of ${pages}`, pageWidth - margin, pageHeight - 14, { align: 'right' });
      }
    };

    const title =
      scope === 'full'
        ? 'Complete Export Operations Report'
        : activeTab === 'contract-flow'
          ? 'Contract Flow Report'
          : `${activeTab === 'coffee' ? 'Coffee' : 'Sesame'} Export Report`;

    drawHeader(title, 'Coffee · Sesame · Contract flow — Export Operations');

    const addShipTable = (label: string, list: ExportRow[]) => {
      const s = sumShipments(list);
      doc.setTextColor(20, 40, 55);
      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      doc.text(label, margin, y);
      y += 14;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text(
        `Records: ${s.count}  ·  Volume: ${s.mt} MT  ·  Gross: ${money(s.gross)}  ·  Paid: ${money(s.paid)}  ·  Open: ${money(s.open)}`,
        margin,
        y,
      );
      y += 8;
      autoTable(doc, {
        startY: y,
        head: [['Shipment', 'Buyer', 'Destination', 'Product', 'Qty', 'Gross', 'Paid', 'Payment', 'Status']],
        body: list.map((r) => [
          r.shipmentNo, r.buyer, r.destination, r.product,
          `${r.quantity} ${r.unit}`, money(r.grossUsd), money(r.paidUsd), r.paymentStatus, r.status,
        ]),
        theme: 'grid',
        styles: { fontSize: 7, cellPadding: 3 },
        headStyles: { fillColor: [23, 66, 95], textColor: 255, fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [240, 247, 252] },
        margin: { left: margin, right: margin },
      });
      y = ((doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable?.finalY || y) + 18;
    };

    if (scope === 'full' || activeTab === 'coffee') {
      addShipTable('Coffee shipments', scope === 'full' ? coffeeSeed : filteredRows);
    }
    if (scope === 'full' || activeTab === 'sesame') {
      if (y > pageHeight - 120) {
        doc.addPage();
        drawHeader(title, 'Continued');
      }
      addShipTable('Sesame shipments', scope === 'full' ? sesameSeed : filteredRows);
    }

    if (scope === 'full' || activeTab === 'contract-flow') {
      if (y > pageHeight - 120) {
        doc.addPage();
        drawHeader(title, 'Contract flow');
      }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.setTextColor(20, 40, 55);
      doc.text('Contracts', margin, y);
      y += 10;
      autoTable(doc, {
        startY: y,
        head: [['Contract No', 'Title', 'Buyer', 'Status', 'Price', 'Updated']],
        body: contracts.map((c) => [c.contractNo, c.title, c.buyer, c.status, c.price, c.updatedAt]),
        theme: 'grid',
        styles: { fontSize: 7, cellPadding: 3 },
        headStyles: { fillColor: [23, 66, 95], textColor: 255 },
        margin: { left: margin, right: margin },
      });
      y = ((doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable?.finalY || y) + 16;

      if (selectedContract) {
        autoTable(doc, {
          startY: y,
          head: [['Payment kind', 'Reference', 'Amount', 'Status', 'Due']],
          body: selectedContract.payments.map((p) => [
            p.kind, p.reference, money(p.amountUsd), p.status, p.dueDate,
          ]),
          theme: 'grid',
          styles: { fontSize: 7, cellPadding: 3 },
          headStyles: { fillColor: [40, 90, 70], textColor: 255 },
          margin: { left: margin, right: margin },
        });
      }
    }

    drawFooter();
    doc.save(`unity-export-${scope === 'full' ? 'complete' : activeTab}-${Date.now()}.pdf`);
  };

  /** Word-compatible HTML document (.doc) */
  const exportWord = (scope: 'view' | 'full' = 'view'): void => {
    const gen = stamp();
    const title =
      scope === 'full'
        ? 'Complete Export Operations Report'
        : activeTab === 'contract-flow'
          ? 'Contract Flow Report'
          : `${activeTab === 'coffee' ? 'Coffee' : 'Sesame'} Export Report`;

    const tableHtml = (headers: string[], rows: (string | number)[][]) => `
      <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:11px;">
        <thead><tr style="background:#173e5f;color:#fff;">${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead>
        <tbody>
          ${rows.map((r, i) => `<tr style="background:${i % 2 ? '#f0f7fc' : '#fff'};">${r.map((c) => `<td>${c ?? ''}</td>`).join('')}</tr>`).join('')}
        </tbody>
      </table>`;

    let body = '';
    if (scope === 'full' || activeTab === 'coffee') {
      const list = scope === 'full' ? coffeeSeed : filteredRows;
      const s = sumShipments(list);
      body += `<h2>Coffee shipments</h2><p>Records: ${s.count} · Volume: ${s.mt} MT · Gross: ${money(s.gross)} · Paid: ${money(s.paid)}</p>`;
      body += tableHtml(
        ['Shipment', 'Buyer', 'Destination', 'Product', 'Qty', 'Gross', 'Paid', 'Status'],
        list.map((r) => [r.shipmentNo, r.buyer, r.destination, r.product, `${r.quantity} ${r.unit}`, money(r.grossUsd), money(r.paidUsd), r.status]),
      );
    }
    if (scope === 'full' || activeTab === 'sesame') {
      const list = scope === 'full' ? sesameSeed : filteredRows;
      const s = sumShipments(list);
      body += `<h2>Sesame shipments</h2><p>Records: ${s.count} · Volume: ${s.mt} MT · Gross: ${money(s.gross)} · Paid: ${money(s.paid)}</p>`;
      body += tableHtml(
        ['Shipment', 'Buyer', 'Destination', 'Product', 'Qty', 'Gross', 'Paid', 'Status'],
        list.map((r) => [r.shipmentNo, r.buyer, r.destination, r.product, `${r.quantity} ${r.unit}`, money(r.grossUsd), money(r.paidUsd), r.status]),
      );
    }
    if (scope === 'full' || activeTab === 'contract-flow') {
      body += `<h2>Contracts</h2>`;
      body += tableHtml(
        ['Contract No', 'Title', 'Buyer', 'Status', 'Price', 'Updated'],
        contracts.map((c) => [c.contractNo, c.title, c.buyer, c.status, c.price, c.updatedAt]),
      );
      if (selectedContract) {
        body += `<h3>Payments — ${selectedContract.contractNo}</h3>`;
        body += tableHtml(
          ['Kind', 'Reference', 'Amount', 'Status', 'Due'],
          selectedContract.payments.map((p) => [p.kind, p.reference, money(p.amountUsd), p.status, p.dueDate]),
        );
        body += `<h3>Workflow</h3>`;
        body += tableHtml(
          ['Step', 'Owner', 'Status', 'Summary'],
          selectedContract.workflow.map((w) => [w.title, w.owner, w.status, w.summary]),
        );
      }
    }

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title></head>
<body style="font-family:Calibri,Arial,sans-serif;color:#102030;">
  <div style="border-bottom:3px solid #173e5f;padding-bottom:12px;margin-bottom:16px;">
    <div style="font-size:12px;letter-spacing:0.12em;text-transform:uppercase;color:#3d7a9c;">${reportOrg}</div>
    <h1 style="margin:6px 0;font-size:22px;">${title}</h1>
    <div style="color:#5a7080;font-size:12px;">Coffee · Sesame · Contract flow · Generated ${gen}</div>
  </div>
  ${body}
  <div style="margin-top:28px;border-top:1px solid #c5d5e0;padding-top:10px;font-size:11px;color:#5a7080;">
    ${reportOrg} · Confidential — internal use only · End of report
  </div>
</body></html>`;

    downloadBlob(new Blob(['\ufeff' + html], { type: 'application/msword' }), `unity-export-${scope === 'full' ? 'complete' : activeTab}-${Date.now()}.doc`);
  };

  const runExport = (format: 'csv' | 'excel' | 'pdf' | 'word', scope: 'view' | 'full') => {
    const allowed =
      (format === 'excel' && canExportExcel) ||
      (format === 'pdf' && canExportPdf) ||
      (format === 'word' && canExportWord) ||
      (format === 'csv' && canExportCsv);
    if (!allowed) {
      setNotice(`Your role cannot export as ${format.toUpperCase()}.`);
      pushNotify(`Export ${format} denied`, 'warn');
      setExportMenuOpen(false);
      return;
    }
    setExportMenuOpen(false);
    if (format === 'csv') {
      exportCsv(scope);
      pushNotify(`CSV export ready (${scope === 'full' ? 'complete flow' : activeTab})`, 'success');
    } else if (format === 'excel') {
      void exportExcel(scope);
    } else if (format === 'pdf') {
      exportPdf(scope);
      pushNotify(`PDF report ready (${scope === 'full' ? 'complete flow' : activeTab})`, 'success');
    } else {
      exportWord(scope);
      pushNotify(`Word report ready (${scope === 'full' ? 'complete flow' : activeTab})`, 'success');
    }
  };

  const contractFlowView = activeTab === 'contract-flow';
  const farmToStoreView = activeTab === 'farm-to-store';

  useEffect(() => {
    if (activeTab === 'contract-flow' && !canViewContract) {
      setActiveTab('coffee');
    }
  }, [activeTab, canViewContract]);

  return (
    <div className="app-shell">
      <style data-uu-export-modal-global>{`
        .rbac-topbar, .rbac-topbar-sticky, .rbac-module-nav, .topbar {
          z-index: 40 !important;
        }
        body > [data-uu-export-modal],
        body > .fts-modal-backdrop,
        body > .modal-backdrop {
          position: fixed !important;
          inset: 0 !important;
          z-index: 2147483646 !important;
        }
      `}</style>
      <header className="topbar">
        <div className="brand-block">
          <small className="eyebrow">Export operations</small>
          <h1>{contractFlowView ? 'Exchange agreement workflow' : activeTab === 'farm-to-store' ? 'Farm → Store pipeline' : activeTab === 'coffee' ? 'Coffee' : 'Sesame'} export overview</h1>
          <div className="brand-meta">
            <span className="tiny-chip">Weekly tracking</span>
            <span className="tiny-chip">Finance view</span>
          </div>
        </div>

        <div className="header-tools">
          <div style={{ position: 'relative' }}>
            <button
              type="button"
              className="ghost-button"
              aria-label="Notifications"
              onClick={() => {
                setBellOpen((o) => !o);
                if (!bellOpen) markAllRead();
              }}
              style={{ position: 'relative', minWidth: 44, minHeight: 40 }}
            >
              <span style={{ fontSize: 18 }} role="img" aria-hidden>🔔</span>
              {unreadCount > 0 && (
                <span
                  style={{
                    position: 'absolute',
                    top: 2,
                    right: 2,
                    minWidth: 18,
                    height: 18,
                    borderRadius: 999,
                    background: '#ef4444',
                    color: '#fff',
                    fontSize: 10,
                    fontWeight: 800,
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    padding: '0 4px',
                  }}
                >
                  {unreadCount > 9 ? '9+' : unreadCount}
                </span>
              )}
            </button>
            {bellOpen && (
              <div
                style={{
                  position: 'absolute',
                  right: 0,
                  top: '110%',
                  zIndex: 60,
                  width: 340,
                  maxHeight: 380,
                  overflow: 'auto',
                  background: 'rgba(9, 22, 31, 0.98)',
                  border: '1px solid rgba(130, 188, 251, 0.28)',
                  borderRadius: 14,
                  boxShadow: '0 18px 40px rgba(0,0,0,0.4)',
                  padding: 10,
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <strong style={{ color: '#e8f6ff', fontSize: 13 }}>Notifications</strong>
                  <button type="button" className="ghost-button" style={{ padding: '4px 8px', fontSize: 11 }} onClick={clearNotifications}>
                    Clear
                  </button>
                </div>
                {notifications.length === 0 ? (
                  <div className="empty-state compact">No alerts yet. Changes will appear here.</div>
                ) : (
                  notifications.map((n) => (
                    <div
                      key={n.id}
                      style={{
                        padding: '10px 12px',
                        marginBottom: 6,
                        borderRadius: 10,
                        background:
                          n.tone === 'success'
                            ? 'rgba(16,185,129,0.12)'
                            : n.tone === 'warn'
                              ? 'rgba(245,158,11,0.12)'
                              : n.tone === 'error'
                                ? 'rgba(239,68,68,0.12)'
                                : 'rgba(56,189,248,0.1)',
                        border: '1px solid rgba(130,188,251,0.12)',
                      }}
                    >
                      <div style={{ color: '#f0fbff', fontSize: 12, fontWeight: 600 }}>{n.text}</div>
                      <div style={{ color: '#8ecbee', fontSize: 10, marginTop: 4 }}>{n.at}</div>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
          <span className={`badge ${apiState === 'postgres' ? 'postgres' : 'local'}`}>
            {apiState === 'postgres' ? 'Live workbook' : 'Local fallback'}
          </span>
          {!contractFlowView && (
            <>
              <button
                className="primary-button"
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  openCreateForm();
                }}
              >
                + Add Record
              </button>
              <button className="ghost-button" type="button" onClick={() => window.location.reload()}>
                Refresh
              </button>
            </>
          )}
        </div>
      </header>

      <section className="tab-row">
        <div className="tab-switch">
          {canViewCoffee && (
            <button
              type="button"
              className={activeTab === 'coffee' ? 'tab-button active' : 'tab-button'}
              onClick={() => setActiveTab('coffee')}
            >
              Coffee
            </button>
          )}
          {canViewSesame && (
            <button
              type="button"
              className={activeTab === 'sesame' ? 'tab-button active' : 'tab-button'}
              onClick={() => setActiveTab('sesame')}
            >
              Sesame
            </button>
          )}
          <button
            type="button"
            className={activeTab === 'farm-to-store' ? 'tab-button active' : 'tab-button'}
            onClick={() => setActiveTab('farm-to-store')}
          >
            Farm → Store
          </button>
          {canViewContract && (
            <button
              type="button"
              className={activeTab === 'contract-flow' ? 'tab-button active' : 'tab-button'}
              onClick={() => setActiveTab('contract-flow')}
            >
              Contract flow
            </button>
          )}
        </div>

        <div className="export-actions" style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
          {canUploadData && !contractFlowView && (
            <label className="primary-button secondary-variant" style={{ cursor: dataUploadBusy ? 'wait' : 'pointer', margin: 0 }}>
              {dataUploadBusy ? 'Uploading…' : '↑ Upload Excel'}
              <input type="file" accept=".xlsx,.xls,.csv" hidden disabled={dataUploadBusy} onChange={handleDataExcelUpload} />
            </label>
          )}
          {!contractFlowView && (
            <button
              type="button"
              className="primary-button"
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                openCreateForm();
              }}
            >
              + Add Record
            </button>
          )}
          {canExportAny && (
          <div style={{ position: 'relative' }}>
            <button
              type="button"
              className="primary-button secondary-variant"
              onClick={() => setExportMenuOpen((o) => !o)}
            >
              Export ▾
            </button>
            {exportMenuOpen && (
              <div
                style={{
                  position: 'absolute',
                  right: 0,
                  top: '110%',
                  zIndex: 50,
                  minWidth: 260,
                  background: 'rgba(9, 22, 31, 0.98)',
                  border: '1px solid rgba(130, 188, 251, 0.25)',
                  borderRadius: 14,
                  boxShadow: '0 18px 40px rgba(0,0,0,0.35)',
                  padding: 8,
                }}
              >
                <div style={{ padding: '8px 12px', fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase', color: '#8ecbee' }}>
                  Current view ({activeTab})
                </div>
                {(['excel', 'pdf', 'word', 'csv'] as const)
                  .filter((fmt) =>
                    (fmt === 'excel' && canExportExcel) ||
                    (fmt === 'pdf' && canExportPdf) ||
                    (fmt === 'word' && canExportWord) ||
                    (fmt === 'csv' && canExportCsv),
                  )
                  .map((fmt) => (
                  <button
                    key={`v-${fmt}`}
                    type="button"
                    className="secondary-button"
                    style={{ width: '100%', marginBottom: 4, textAlign: 'left' }}
                    onClick={() => runExport(fmt, 'view')}
                  >
                    {fmt === 'excel' ? 'Excel (.xlsx)' : fmt === 'pdf' ? 'PDF' : fmt === 'word' ? 'Word (.doc)' : 'CSV'} — this tab
                  </button>
                ))}
                <div style={{ padding: '8px 12px', fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase', color: '#8ecbee' }}>
                  Complete flow (Coffee + Sesame + Contracts)
                </div>
                {(['excel', 'pdf', 'word', 'csv'] as const)
                  .filter((fmt) =>
                    (fmt === 'excel' && canExportExcel) ||
                    (fmt === 'pdf' && canExportPdf) ||
                    (fmt === 'word' && canExportWord) ||
                    (fmt === 'csv' && canExportCsv),
                  )
                  .map((fmt) => (
                  <button
                    key={`f-${fmt}`}
                    type="button"
                    className="secondary-button"
                    style={{ width: '100%', marginBottom: 4, textAlign: 'left' }}
                    onClick={() => runExport(fmt, 'full')}
                  >
                    {fmt === 'excel' ? 'Excel (.xlsx)' : fmt === 'pdf' ? 'PDF' : fmt === 'word' ? 'Word (.doc)' : 'CSV'} — full report
                  </button>
                ))}
              </div>
            )}
          </div>
          )}
        </div>
      </section>

      {farmToStoreView ? (
        <>
          <FarmToStoreStyles />
          <FarmToStorePanel canEdit={Boolean(canEdit || canAdd)} />
        </>
      ) : contractFlowView ? (
        <section className="contract-shell">
          {/* ── Contract list sidebar ── */}
          <div className="contract-sidebar">
            <div className="contract-panel panel-soft">
              <div className="panel-title">Contracts</div>
              <button
                type="button"
                className="primary-button compact-action"
                style={{ width: '100%', marginBottom: 12 }}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  openCreateContract();
                }}
              >
                + New contract
              </button>
              <div className="attachment-list">
                {contracts.length === 0 ? (
                  <div className="empty-state compact">No contracts yet</div>
                ) : (
                  contracts.map((c) => (
                    <div
                      key={c.id}
                      className="attachment-row"
                      style={{
                        cursor: 'pointer',
                        borderColor: selectedContract?.id === c.id ? 'rgba(127, 203, 255, 0.45)' : undefined,
                      }}
                      onClick={() => setSelectedContractId(c.id)}
                    >
                      <div>
                        <strong>{c.contractNo}</strong>
                        <small>{c.buyer} · {c.status}</small>
                      </div>
                      <span>{c.approvalTag}</span>
                    </div>
                  ))
                )}
              </div>
            </div>

            <div className="contract-panel panel-soft">
              <div className="panel-title">Modules</div>
              {contractModules.map((section) => (
                <div className="module-group" key={section.name}>
                  <div className="module-name">{section.name}</div>
                  <ul>
                    {section.items.map((item) => (
                      <li key={`${section.name}-${item}`}>{item}</li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </div>

          <div className="contract-main">
            {!selectedContract ? (
              <div className="panel-soft" style={{ padding: 24 }}>
                <div className="empty-state">Create or select a contract to manage workflow, payments and uploads.</div>
              </div>
            ) : (
              <>
                {contractNotice && <div className="notice-bar">{contractNotice}</div>}

                <div className="contract-header-panel panel-soft">
                  <div>
                    <small className="eyebrow">Contract file</small>
                    <h2>{selectedContract.title || selectedContract.contractNo}</h2>
                    <p className="approval-tag">{selectedContract.approvalTag} · {selectedContract.status}</p>
                  </div>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {canEditContract && (
                      <button type="button" className="secondary-button compact-action" onClick={() => openEditContract(selectedContract)}>
                        Edit
                      </button>
                    )}
                    {canDelete && (
                      <button type="button" className="secondary-button compact-action danger-button" onClick={() => deleteContract(selectedContract)}>
                        Delete
                      </button>
                    )}
                  </div>
                </div>

                {/* KPI payments */}
                <section className="kpi-grid contract-kpi-grid" style={{ marginBottom: 0 }}>
                  <div className="kpi-card">
                    <span>Payment total</span>
                    <strong title={money(contractPaymentTotals.total)}>{money(contractPaymentTotals.total)}</strong>
                    <small>Contract lines</small>
                  </div>
                  <div className="kpi-card">
                    <span>Collected</span>
                    <strong title={money(contractPaymentTotals.paid)}>{money(contractPaymentTotals.paid)}</strong>
                    <small>Paid + partial</small>
                  </div>
                  <div className="kpi-card">
                    <span>Open receivable</span>
                    <strong title={money(contractPaymentTotals.open)}>{money(contractPaymentTotals.open)}</strong>
                    <small>Outstanding</small>
                  </div>
                  <div className="kpi-card">
                    <span>Attachments</span>
                    <strong>{selectedContract.attachments.length}</strong>
                    <small>Files</small>
                  </div>
                </section>

                {/* Commercial fields */}
                <div className="contract-panel panel-soft contract-table-wrap">
                  <div className="table-header-row">
                    <div className="table-header-left">
                      <h3>Commercial terms</h3>
                    </div>
                  </div>
                  <div className="contract-grid">
                    {[
                      ['Contract No', selectedContract.contractNo],
                      ['Buyer', selectedContract.buyer],
                      ['Product / Quality', selectedContract.quality],
                      ['Quantity', selectedContract.quantity],
                      ['Net weight', selectedContract.netWeight],
                      ['Packing', selectedContract.packing],
                      ['Price', selectedContract.price],
                      ['Terms', selectedContract.terms],
                      ['Shipment window', selectedContract.shipment],
                      ['Shipping', selectedContract.shipping],
                      ['Payment terms', selectedContract.payment],
                      ['Insurance', selectedContract.insurance],
                      ['Quality approval', selectedContract.qualityApproval],
                      ['Arbitration', selectedContract.arbitration],
                    ].map(([label, value]) => (
                      <div className="contract-field" key={label}>
                        <label>{label}</label>
                        <strong>{value || '—'}</strong>
                      </div>
                    ))}
                  </div>
                  {selectedContract.note && (
                    <div className="contract-note">{selectedContract.note}</div>
                  )}
                </div>

                {/* Workflow */}
                <div className="contract-panel panel-soft">
                  <div className="panel-title">Workflow</div>
                  <div className="workflow-row">
                    {selectedContract.workflow.map((step) => (
                      <div key={step.id} className={`workflow-card ${step.status}`}>
                        <div className="workflow-dot" />
                        <div className="workflow-title">{step.title}</div>
                        <div className="workflow-owner">{step.owner}</div>
                        <p>{step.summary}</p>
                        <small style={{ textTransform: 'uppercase', letterSpacing: '0.08em', color: '#8ecbee' }}>{step.status}</small>
                        {step.status !== 'done' && (
                          <button
                            type="button"
                            className="primary-button compact-action"
                            style={{ marginTop: 8 }}
                            onClick={() => advanceWorkflowStep(step.id)}
                          >
                            {step.status === 'active' ? 'Mark done' : 'Activate / complete'}
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>

                <div className="contract-lower-grid">
                  {/* Attachments + upload */}
                  <div className="panel-soft attachments-panel">
                    <div className="panel-title">Attachments</div>
                    <label className="primary-button compact-action" style={{ display: 'inline-block', marginBottom: 12, cursor: uploadBusy ? 'wait' : 'pointer' }}>
                      {uploadBusy ? 'Uploading…' : '+ Upload file(s)'}
                      <input
                        type="file"
                        multiple
                        style={{ display: 'none' }}
                        disabled={uploadBusy}
                        onChange={handleFileUpload}
                      />
                    </label>
                    <div className="attachment-list">
                      {selectedContract.attachments.length === 0 ? (
                        <div className="empty-state compact">No files yet</div>
                      ) : (
                        selectedContract.attachments.map((item) => (
                          <div className="attachment-row" key={item.id}>
                            <div>
                              <strong>{item.name}</strong>
                              <small>{item.type} · {item.uploadedAt}</small>
                            </div>
                            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                              <span>{item.size}</span>
                              {item.dataUrl && (
                                <a href={item.dataUrl} download={item.name} className="table-action-button edit" style={{ textDecoration: 'none' }}>
                                  Download
                                </a>
                              )}
                              <button type="button" className="table-action-button delete" onClick={() => removeAttachment(item.id)}>
                                Remove
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  {/* Role matrix (read from contract, same fields) */}
                  <div className="panel-soft roles-panel">
                    <div className="panel-title">Role access</div>
                    <div className="role-table">
                      {selectedContract.roleMatrix.map((entry) => (
                        <div className="role-row" key={entry.role}>
                          <div>
                            <strong>{entry.role}</strong>
                            <small>{entry.level}</small>
                          </div>
                          <ul>
                            {entry.allowed.map((action) => (
                              <li key={`${entry.role}-${action}`}>{action}</li>
                            ))}
                          </ul>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Payments CRUD */}
                <div className="contract-panel panel-soft">
                  <div className="table-header-row">
                    <div className="table-header-left">
                      <h3>Payments & receivables</h3>
                    </div>
                    <button type="button" className="primary-button compact-action" onClick={() => setShowPaymentForm(true)}>
                      + Add payment
                    </button>
                  </div>
                  <div className="table-wrap">
                    {selectedContract.payments.length === 0 ? (
                      <div className="empty-state compact">No payment lines</div>
                    ) : (
                      <table>
                        <thead>
                          <tr>
                            <th>Kind</th>
                            <th>Reference</th>
                            <th>Amount USD</th>
                            <th>Status</th>
                            <th>Due</th>
                            <th>Paid date</th>
                            <th>Notes</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {selectedContract.payments.map((p) => (
                            <tr key={p.id}>
                              <td>{p.kind}</td>
                              <td>{p.reference}</td>
                              <td>{money(p.amountUsd)}</td>
                              <td>
                                <span className={`status-pill ${normalizeStatus(p.status)}`}>{p.status}</span>
                              </td>
                              <td>{p.dueDate}</td>
                              <td>{p.paidDate || '—'}</td>
                              <td>{p.notes || '—'}</td>
                              <td className="table-actions">
                                {p.status !== 'Paid' && (
                                  <button type="button" className="table-action-button edit" onClick={() => markPaymentPaid(p.id)}>
                                    Mark paid
                                  </button>
                                )}
                                <button type="button" className="table-action-button delete" onClick={() => removePayment(p.id)}>
                                  Remove
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                </div>
              </>
            )}
          </div>

          {/* Contract form modal */}
          {showContractForm && createPortal(
            <div data-uu-export-modal="1" className="modal-backdrop" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 2147483646, background: 'rgba(15,30,50,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, pointerEvents: 'auto' }} onClick={closeContractForm}>
              <div className="modal-card" onClick={(event) => event.stopPropagation()}>
                <div className="modal-header">
                  <div>
                    <small className="eyebrow">{editingContract ? 'Edit contract' : 'New contract'}</small>
                    <h3>{editingContract ? 'Edit Contract' : 'Add Contract'}</h3>
                  </div>
                  <button type="button" className="close-button" onClick={closeContractForm} aria-label="Close">
                    ×
                  </button>
                </div>
                <form onSubmit={saveContract} className="record-form">
                  <div className="form-grid">
                    <label>
                      Contract No
                      <input name="contractNo" value={contractForm.contractNo || ''} onChange={handleContractField} required />
                    </label>
                    <label>
                      Title
                      <input name="title" value={contractForm.title || ''} onChange={handleContractField} required />
                    </label>
                    <label>
                      Buyer
                      <input name="buyer" value={contractForm.buyer || ''} onChange={handleContractField} required />
                    </label>
                    <label>
                      Status
                      <select name="status" value={contractForm.status || 'Draft'} onChange={handleContractField}>
                        <option value="Draft">Draft</option>
                        <option value="Active">Active</option>
                        <option value="In Execution">In Execution</option>
                        <option value="Settled">Settled</option>
                        <option value="Closed">Closed</option>
                      </select>
                    </label>
                    <label>
                      Quantity
                      <input name="quantity" value={contractForm.quantity || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Net weight
                      <input name="netWeight" value={contractForm.netWeight || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Packing
                      <input name="packing" value={contractForm.packing || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Quality
                      <input name="quality" value={contractForm.quality || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Price
                      <input name="price" value={contractForm.price || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Terms
                      <input name="terms" value={contractForm.terms || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Shipment
                      <input name="shipment" value={contractForm.shipment || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Shipping
                      <input name="shipping" value={contractForm.shipping || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Payment terms
                      <input name="payment" value={contractForm.payment || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Insurance
                      <input name="insurance" value={contractForm.insurance || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Quality approval
                      <input name="qualityApproval" value={contractForm.qualityApproval || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Arbitration
                      <input name="arbitration" value={contractForm.arbitration || ''} onChange={handleContractField} />
                    </label>
                    <label>
                      Approval tag
                      <input name="approvalTag" value={contractForm.approvalTag || ''} onChange={handleContractField} />
                    </label>
                    <label className="full-span">
                      Notes
                      <textarea name="note" value={contractForm.note || ''} rows={3} onChange={handleContractField} />
                    </label>
                  </div>
                  <div className="form-actions">
                    <button type="button" className="secondary-button" onClick={closeContractForm}>Cancel</button>
                    <button type="submit" className="primary-button" disabled={contractSaving}>
                      {contractSaving ? 'Saving…' : editingContract ? 'Save changes' : 'Create contract'}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          , document.body
          )}

          {/* Payment form modal */}
          {showPaymentForm && createPortal(
            <div data-uu-export-modal="1" className="modal-backdrop" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 2147483646, background: 'rgba(15,30,50,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, pointerEvents: 'auto' }} onClick={() => setShowPaymentForm(false)}>
              <div className="modal-card" onClick={(event) => event.stopPropagation()}>
                <div className="modal-header">
                  <div>
                    <small className="eyebrow">Receivable</small>
                    <h3>Add payment line</h3>
                  </div>
                  <button type="button" className="close-button" onClick={() => setShowPaymentForm(false)}>×</button>
                </div>
                <form onSubmit={savePayment} className="record-form">
                  <div className="form-grid">
                    <label>
                      Kind
                      <select
                        value={paymentForm.kind}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, kind: e.target.value as ContractPayment['kind'] }))}
                      >
                        <option value="LC">LC</option>
                        <option value="TT">TT</option>
                        <option value="Cash">Cash</option>
                        <option value="Other">Other</option>
                      </select>
                    </label>
                    <label>
                      Amount USD
                      <input
                        type="number"
                        step="0.01"
                        value={paymentForm.amountUsd}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, amountUsd: Number(e.target.value || 0) }))}
                        required
                      />
                    </label>
                    <label>
                      Status
                      <select
                        value={paymentForm.status}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, status: e.target.value as ContractPayment['status'] }))}
                      >
                        <option value="Open">Open</option>
                        <option value="Partial">Partial</option>
                        <option value="Paid">Paid</option>
                      </select>
                    </label>
                    <label>
                      Reference
                      <input
                        value={paymentForm.reference}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, reference: e.target.value }))}
                      />
                    </label>
                    <label>
                      Due date
                      <input
                        type="date"
                        value={paymentForm.dueDate}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, dueDate: e.target.value }))}
                      />
                    </label>
                    <label>
                      Paid date
                      <input
                        type="date"
                        value={paymentForm.paidDate || ''}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, paidDate: e.target.value }))}
                      />
                    </label>
                    <label className="full-span">
                      Notes
                      <textarea
                        rows={2}
                        value={paymentForm.notes}
                        onChange={(e) => setPaymentForm((p) => ({ ...p, notes: e.target.value }))}
                      />
                    </label>
                  </div>
                  <div className="form-actions">
                    <button type="button" className="secondary-button" onClick={() => setShowPaymentForm(false)}>Cancel</button>
                    <button type="submit" className="primary-button">Add payment</button>
                  </div>
                </form>
              </div>
            </div>
          , document.body
          )}
        </section>
      ) : (
        <div className="trade-dashboard">
          <section className="kpi-grid">
            <div className="kpi-card">
              <span>Shipments</span>
              <strong>{totals.shipmentCount}</strong>
              <small>{activeTab === 'coffee' ? 'Coffee lots' : 'Sesame lots'}</small>
            </div>
            <div className="kpi-card">
              <span>Volume</span>
              <strong>{totals.totalMt.toLocaleString()}</strong>
              <small>MT</small>
            </div>
            <div className="kpi-card">
              <span>Containers</span>
              <strong>{totals.totalContainers}</strong>
              <small>Loaded</small>
            </div>
            <div className="kpi-card">
              <span>Gross</span>
              <strong>{money(totals.gross)}</strong>
              <small>USD total</small>
            </div>
            <div className="kpi-card">
              <span>Paid</span>
              <strong>{money(totals.paid)}</strong>
              <small>Collected</small>
            </div>
            <div className="kpi-card">
              <span>Open</span>
              <strong>{money(totals.open)}</strong>
              <small>Receivable</small>
            </div>
          </section>

          <section className="charts-grid">
            <div className="panel chart-panel">
              <div className="panel-header">
                <div>
                  <small>Trend</small>
                  <h3>Weekly value</h3>
                </div>
                <div className="chart-picker">
                  <button type="button" className={chartType === 'line' ? 'chart-toggle active' : 'chart-toggle'} onClick={() => setChartType('line')}>Line</button>
                  <button type="button" className={chartType === 'bar' ? 'chart-toggle active' : 'chart-toggle'} onClick={() => setChartType('bar')}>Bar</button>
                  <button type="button" className={chartType === 'pie' ? 'chart-toggle active' : 'chart-toggle'} onClick={() => setChartType('pie')}>Pie</button>
                </div>
              </div>

              {chartType === 'pie' ? (
                <div className="pie-wrap">
                  <div className="pie-chart" style={{ background: pieGradient }}>
                    <div className="pie-inner">
                      <strong>{compactMoney(totals.gross)}</strong>
                    </div>
                  </div>
                  <div className="pie-legend">
                    {destinationMix.map((entry, index) => (
                      <div key={entry.label} className="legend-item">
                        <span className="legend-swatch" style={{ background: ['#67d7ff', '#7a8cff', '#82e0a1', '#ffd56a', '#ff8bb4'][index % 5] }} />
                        <span>{entry.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              ) : chartType === 'bar' ? (
                <div className="bar-chart" role="img" aria-label="Weekly export value bar chart">
                  {weeklyTrend.map((point) => (
                    <div key={point.label} className="bar-group">
                      <div className="bar" style={{ height: `${Math.max(14, (point.value / maxTrendValue) * 100)}%` }} />
                      <span>{point.label.replace('2026-W', 'W')}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <svg viewBox="0 0 500 180" className="trend-chart" role="img" aria-label="Weekly export value chart">
                  {weeklyTrend.length > 0 && (
                    <>
                      <defs>
                        <linearGradient id="areaFill" x1="0" x2="0" y1="0" y2="1">
                          <stop offset="0%" stopColor="#80d8ff" stopOpacity="0.45" />
                          <stop offset="100%" stopColor="#80d8ff" stopOpacity="0.02" />
                        </linearGradient>
                      </defs>

                      {weeklyTrend.map((point, index) => {
                        const x = 24 + index * 100;
                        const y = 150 - (point.value / maxTrendValue) * 110;
                        const next = weeklyTrend[index + 1];
                        const nextX = next ? 24 + (index + 1) * 100 : x;
                        const nextY = next ? 150 - (next.value / maxTrendValue) * 110 : y;

                        return (
                          <g key={point.label}>
                            <line x1={x} x2={x} y1="18" y2="150" stroke="rgba(166,205,255,0.08)" />
                            <circle cx={x} cy={y} r="4" fill="#7ad9ff" />
                            {next && <line x1={x} x2={nextX} y1={y} y2={nextY} stroke="#7ad9ff" strokeWidth="2.5" />}
                          </g>
                        );
                      })}
                    </>
                  )}
                </svg>
              )}

              {chartType !== 'pie' && (
                <div className="trend-labels">
                  {weeklyTrend.map((point) => (
                    <span key={point.label}>{point.label.replace('2026-W', 'W')}</span>
                  ))}
                </div>
              )}
            </div>

            <div className="panel chart-panel">
              <div className="panel-header">
                <div>
                  <small>Mix</small>
                  <h3>Destination value</h3>
                </div>
                <strong>{destinationMix.length} markets</strong>
              </div>

              <div className="market-list">
                {destinationMix.length ? (
                  destinationMix.map((entry) => (
                    <div className="market-row" key={entry.label}>
                      <div className="market-topline">
                        <span>{entry.label}</span>
                        <strong>{money(entry.value)}</strong>
                      </div>
                      <div className="meter">
                        <span style={{ width: `${entry.percent}%` }} />
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="empty-state compact">No destination data</div>
                )}
              </div>
            </div>
          </section>

          {notice && <div className="notice-bar">{notice}</div>}

          <div className="layout">
            <section className="panel table-panel">
              <div className="controls">
                <button type="button" className="primary-button compact-action" onClick={openCreateForm}>
                  + Add Record
                </button>
                <input
                  type="text"
                  placeholder="Search buyer, vessel, destination..."
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                />

                <select value={buyerFilter} onChange={(event) => setBuyerFilter(event.target.value)}>
                  {buyers.map((buyer) => (
                    <option key={buyer} value={buyer}>{buyer === 'All' ? 'All Buyers' : buyer}</option>
                  ))}
                </select>

                <select value={destinationFilter} onChange={(event) => setDestinationFilter(event.target.value)}>
                  {destinations.map((destination) => (
                    <option key={destination} value={destination}>{destination === 'All' ? 'All Destinations' : destination}</option>
                  ))}
                </select>

                <select value={paymentFilter} onChange={(event) => setPaymentFilter(event.target.value)}>
                  {paymentOptions.map((value) => (
                    <option key={value} value={value}>{value === 'All' ? 'All Payment' : value}</option>
                  ))}
                </select>

                <select value={cleaningFilter} onChange={(event) => setCleaningFilter(event.target.value)}>
                  {cleaningOptions.map((value) => (
                    <option key={value} value={value}>{value === 'All' ? 'All Cleaning' : value}</option>
                  ))}
                </select>
              </div>

              <div className="table-wrap">
                {filteredRows.length === 0 ? (
                  <div className="empty-state">No matching rows for this export view.</div>
                ) : (
                  <table>
                    <thead>
                      <tr>
                        <th>Shipment</th>
                        <th>Buyer</th>
                        <th>Destination</th>
                        <th>Product</th>
                        <th>Qty</th>
                        <th>Gross</th>
                        <th>Paid</th>
                        <th>Payment</th>
                        <th>Cleaning</th>
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row) => (
                        <tr key={row.id} className={selectedRow?.id === row.id ? 'selected-row' : ''} onClick={() => setSelectedRow(row)}>
                          <td>{row.shipmentNo}</td>
                          <td>{row.buyer}</td>
                          <td>{row.destination}</td>
                          <td>{row.product}</td>
                          <td>{`${row.quantity} ${row.unit}`}</td>
                          <td>{money(row.grossUsd)}</td>
                          <td>{money(row.paidUsd)}</td>
                          <td>
                            <span className={`status-pill ${normalizeStatus(row.paymentStatus)}`}>
                              {row.paymentStatus}
                            </span>
                          </td>
                          <td>
                            <span className={`status-pill ${normalizeStatus(row.cleaning)}`}>
                              {row.cleaning}
                            </span>
                          </td>
                          <td>
                            <span className={`status-pill ${normalizeStatus(row.status)}`}>
                              {row.status}
                            </span>
                          </td>
                          <td className="table-actions">
                            <button
                              type="button"
                              className="table-action-button edit"
                              onClick={(event) => {
                                event.stopPropagation();
                                openEditForm(row);
                              }}
                            >
                              Edit
                            </button>
                            <button
                              type="button"
                              className="table-action-button delete"
                              onClick={(event) => {
                                event.stopPropagation();
                                void deleteRecord(row);
                              }}
                            >
                              Delete
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </section>

            <aside className="panel sidebar-panel">
              {selectedRow ? (
                <div className="dossier">
                  <div className="panel-header compact-header">
                    <div>
                      <small>Shipment details</small>
                      <h3>{selectedRow.shipmentNo}</h3>
                    </div>
                    <span className={`status-pill ${normalizeStatus(selectedRow.status)}`}>{selectedRow.status}</span>
                  </div>

                  <div className="meta-grid">
                    <div className="meta-item">
                      <label>Buyer</label>
                      <strong>{formatShort(selectedRow.buyer)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Destination</label>
                      <strong>{formatShort(selectedRow.destination)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Contract</label>
                      <strong>{formatShort(selectedRow.contract)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Vessel</label>
                      <strong>{formatShort(selectedRow.vessel)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Gross USD</label>
                      <strong>{money(selectedRow.grossUsd)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Paid</label>
                      <strong>{money(selectedRow.paidUsd)}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Payment</label>
                      <strong>{selectedRow.paymentStatus}</strong>
                    </div>
                    <div className="meta-item">
                      <label>Cleaning</label>
                      <strong>{selectedRow.cleaning}</strong>
                    </div>
                  </div>

                  <div className="notes-box">
                    <div><strong>Week:</strong> {selectedRow.week}</div>
                    <div><strong>Quantity:</strong> {selectedRow.quantity} {selectedRow.unit}</div>
                    <div><strong>Containers:</strong> {selectedRow.containers}</div>
                    <div><strong>Notes:</strong> {selectedRow.notes}</div>
                  </div>

                  <div className="sidebar-actions">
                    {canEdit && (
                      <button type="button" className="primary-button compact-action" onClick={() => openEditForm(selectedRow)}>
                        Edit Record
                      </button>
                    )}
                    {canDelete && (
                      <button type="button" className="secondary-button compact-action danger-button" onClick={() => deleteRecord(selectedRow)}>
                        Delete Record
                      </button>
                    )}
                  </div>
                </div>
              ) : (
                <div className="empty-state">Select a record to view its dossier.</div>
              )}
            </aside>
          </div>

          {showForm && createPortal(
            <div data-uu-export-modal="1" className="modal-backdrop" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 2147483646, background: 'rgba(15,30,50,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, pointerEvents: 'auto' }} onClick={closeForm}>
              <div className="modal-card" onClick={(event) => event.stopPropagation()}>
                <div className="modal-header">
                  <div>
                    <small className="eyebrow">{editingRow ? 'Edit shipment' : 'Add shipment'}</small>
                    <h3>{editingRow ? 'Edit Record' : 'Add Record'}</h3>
                  </div>
                  <button type="button" className="close-button" onClick={closeForm} aria-label="Close form">
                    ×
                  </button>
                </div>

                <form onSubmit={saveRecord} className="record-form">
                  <div className="form-grid">
                    <label>
                      Shipment No
                      <input name="shipmentNo" value={form.shipmentNo || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Week
                      <input name="week" value={form.week || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Buyer
                      <input name="buyer" value={form.buyer || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Destination
                      <input name="destination" value={form.destination || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Product
                      <input name="product" value={form.product || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Vessel
                      <input name="vessel" value={form.vessel || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Quantity
                      <input type="number" name="quantity" value={Number(form.quantity || 0)} onChange={handleFormChange} />
                    </label>
                    <label>
                      Unit
                      <select name="unit" value={form.unit || 'MT'} onChange={handleFormChange}>
                        <option value="MT">MT</option>
                        <option value="KG">KG</option>
                        <option value="Bags">Bags</option>
                      </select>
                    </label>
                    <label>
                      Containers
                      <input type="number" name="containers" value={Number(form.containers || 0)} onChange={handleFormChange} />
                    </label>
                    <label>
                      Price / KG
                      <input type="number" step="0.01" name="pricePerKg" value={Number(form.pricePerKg || 0)} onChange={handleFormChange} />
                    </label>
                    <label>
                      Gross USD
                      <input type="number" step="0.01" name="grossUsd" value={Number(form.grossUsd || 0)} onChange={handleFormChange} />
                    </label>
                    <label>
                      Paid USD
                      <input type="number" step="0.01" name="paidUsd" value={Number(form.paidUsd || 0)} onChange={handleFormChange} />
                    </label>
                    <label>
                      Payment Status
                      <select name="paymentStatus" value={form.paymentStatus || 'Open'} onChange={handleFormChange}>
                        <option value="Paid">Paid</option>
                        <option value="Partial">Partial</option>
                        <option value="Open">Open</option>
                      </select>
                    </label>
                    <label>
                      Cleaning
                      <select name="cleaning" value={form.cleaning || 'Cleaned'} onChange={handleFormChange}>
                        <option value="Grade A">Grade A</option>
                        <option value="Grade B">Grade B</option>
                        <option value="Cleaned">Cleaned</option>
                        <option value="Pending">Pending</option>
                      </select>
                    </label>
                    <label>
                      Contract
                      <input name="contract" value={form.contract || ''} onChange={handleFormChange} />
                    </label>
                    <label>
                      Status
                      <select name="status" value={form.status || 'Pending'} onChange={handleFormChange}>
                        <option value="Delivered">Delivered</option>
                        <option value="In Transit">In Transit</option>
                        <option value="At Port">At Port</option>
                        <option value="On Vessel">On Vessel</option>
                        <option value="Pending">Pending</option>
                      </select>
                    </label>
                    <label>
                      Date
                      <input type="date" name="date" value={form.date || new Date().toISOString().slice(0, 10)} onChange={handleFormChange} />
                    </label>
                    <label className="full-span">
                      Notes
                      <textarea name="notes" value={form.notes || ''} rows={3} onChange={handleFormChange} />
                    </label>
                  </div>

                  <div className="form-actions">
                    <button type="button" className="secondary-button" onClick={closeForm}>Cancel</button>
                    <button type="submit" className="primary-button" disabled={saving}>
                      {saving ? 'Saving...' : editingRow ? 'Save Changes' : 'Add Record'}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          , document.body
          )}
        </div>
      )}
    </div>
  );
}
