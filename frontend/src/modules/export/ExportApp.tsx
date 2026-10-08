import { useEffect, useMemo, useState, type ChangeEvent, type FormEvent } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { coffeeSeed, sesameSeed } from './mockExportData';

type ExportTab = 'coffee' | 'sesame' | 'contract-flow';

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

export default function ExportApp(): JSX.Element {
  const [activeTab, setActiveTab] = useState<ExportTab>('coffee');
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
    setEditingRow(null);
    setForm(defaultRecordForm(activeTab));
    setShowForm(true);
    setNotice('');
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

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-block">
          <small className="eyebrow">Export operations</small>
          <h1>{contractFlowView ? 'Exchange agreement workflow' : activeTab === 'coffee' ? 'Coffee' : 'Sesame'} export overview</h1>
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
              <button className="primary-button" type="button" onClick={openCreateForm}>
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
          <button
            type="button"
            className={activeTab === 'coffee' ? 'tab-button active' : 'tab-button'}
            onClick={() => setActiveTab('coffee')}
          >
            Coffee
          </button>
          <button
            type="button"
            className={activeTab === 'sesame' ? 'tab-button active' : 'tab-button'}
            onClick={() => setActiveTab('sesame')}
          >
            Sesame
          </button>
          <button
            type="button"
            className={activeTab === 'contract-flow' ? 'tab-button active' : 'tab-button'}
            onClick={() => setActiveTab('contract-flow')}
          >
            Contract flow
          </button>
        </div>

        <div className="export-actions" style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
          {!contractFlowView && (
            <button type="button" className="primary-button" onClick={openCreateForm}>
              + Add Record
            </button>
          )}
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
                {(['excel', 'pdf', 'word', 'csv'] as const).map((fmt) => (
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
                {(['excel', 'pdf', 'word', 'csv'] as const).map((fmt) => (
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
        </div>
      </section>

      {contractFlowView ? (
        <section className="contract-shell">
          {/* ── Contract list sidebar ── */}
          <div className="contract-sidebar">
            <div className="contract-panel panel-soft">
              <div className="panel-title">Contracts</div>
              <button type="button" className="primary-button compact-action" style={{ width: '100%', marginBottom: 12 }} onClick={openCreateContract}>
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
                    <button type="button" className="secondary-button compact-action" onClick={() => openEditContract(selectedContract)}>
                      Edit
                    </button>
                    <button type="button" className="secondary-button compact-action danger-button" onClick={() => deleteContract(selectedContract)}>
                      Delete
                    </button>
                  </div>
                </div>

                {/* KPI payments */}
                <section className="kpi-grid" style={{ marginBottom: 0 }}>
                  <div className="kpi-card">
                    <span>Payment total</span>
                    <strong>{money(contractPaymentTotals.total)}</strong>
                    <small>Contract lines</small>
                  </div>
                  <div className="kpi-card">
                    <span>Collected</span>
                    <strong>{money(contractPaymentTotals.paid)}</strong>
                    <small>Paid + partial</small>
                  </div>
                  <div className="kpi-card">
                    <span>Open receivable</span>
                    <strong>{money(contractPaymentTotals.open)}</strong>
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
          {showContractForm && (
            <div className="modal-backdrop" onClick={closeContractForm}>
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
          )}

          {/* Payment form modal */}
          {showPaymentForm && (
            <div className="modal-backdrop" onClick={() => setShowPaymentForm(false)}>
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
                    <button type="button" className="primary-button compact-action" onClick={() => openEditForm(selectedRow)}>
                      Edit Record
                    </button>
                    <button type="button" className="secondary-button compact-action danger-button" onClick={() => deleteRecord(selectedRow)}>
                      Delete Record
                    </button>
                  </div>
                </div>
              ) : (
                <div className="empty-state">Select a record to view its dossier.</div>
              )}
            </aside>
          </div>

          {showForm && (
            <div className="modal-backdrop" onClick={closeForm}>
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
          )}
        </div>
      )}
    </div>
  );
}
