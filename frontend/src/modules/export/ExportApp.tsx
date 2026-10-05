import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

type ExportTab = 'coffee' | 'sesame';

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

const coffeeSeed: ExportRow[] = [
  { id: 1, week: '2026-W29', shipmentNo: 'COF-26-001', buyer: 'Kofi Foods Ltd.', destination: 'Jeddah', quantity: 1420, unit: 'MT', containers: 4, pricePerKg: 2.45, grossUsd: 348900, paidUsd: 210000, paymentStatus: 'Partial', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Red Sea', contract: 'CTR-COF-2109', status: 'In Transit', date: '2026-07-08', notes: 'Export lot 1' },
  { id: 2, week: '2026-W29', shipmentNo: 'COF-26-002', buyer: 'Arab Roast Co.', destination: 'Dubai', quantity: 1180, unit: 'MT', containers: 3, pricePerKg: 2.38, grossUsd: 281000, paidUsd: 281000, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Al Noor', contract: 'CTR-COF-2110', status: 'Delivered', date: '2026-07-10', notes: 'Premium grade' },
  { id: 3, week: '2026-W30', shipmentNo: 'COF-26-003', buyer: 'Cairo Importers', destination: 'Cairo', quantity: 1540, unit: 'MT', containers: 5, pricePerKg: 2.52, grossUsd: 388080, paidUsd: 160000, paymentStatus: 'Partial', cleaning: 'Grade B', product: 'Washed Coffee', vessel: 'MV Nile Star', contract: 'CTR-COF-2111', status: 'At Port', date: '2026-07-15', notes: 'Customs hold' },
  { id: 4, week: '2026-W30', shipmentNo: 'COF-26-004', buyer: 'Riyadh Café Supply', destination: 'Riyadh', quantity: 1325, unit: 'MT', containers: 4, pricePerKg: 2.49, grossUsd: 329925, paidUsd: 329925, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Green Coffee', vessel: 'MV Gulf Horizon', contract: 'CTR-COF-2112', status: 'Delivered', date: '2026-07-18', notes: 'Fast payment' },
  { id: 5, week: '2026-W31', shipmentNo: 'COF-26-005', buyer: 'Harbor Trading', destination: 'Mombasa', quantity: 1660, unit: 'MT', containers: 5, pricePerKg: 2.54, grossUsd: 421640, paidUsd: 220000, paymentStatus: 'Partial', cleaning: 'Grade A', product: 'Coffee Beans', vessel: 'MV East Null', contract: 'CTR-COF-2113', status: 'On Vessel', date: '2026-07-22', notes: 'Cold chain check' },
  { id: 6, week: '2026-W31', shipmentNo: 'COF-26-006', buyer: 'Prime Café Aus.', destination: 'Sydney', quantity: 970, unit: 'MT', containers: 3, pricePerKg: 2.59, grossUsd: 251230, paidUsd: 251230, paymentStatus: 'Paid', cleaning: 'Grade A', product: 'Arabica', vessel: 'MV Sirocco', contract: 'CTR-COF-2114', status: 'Delivered', date: '2026-07-24', notes: 'QC passed' },
  { id: 7, week: '2026-W32', shipmentNo: 'COF-26-007', buyer: 'Great Lakes Foods', destination: 'Dar es Salaam', quantity: 1380, unit: 'MT', containers: 4, pricePerKg: 2.43, grossUsd: 335340, paidUsd: 98000, paymentStatus: 'Partial', cleaning: 'Grade B', product: 'Green Coffee', vessel: 'MV Aurora', contract: 'CTR-COF-2115', status: 'At Port', date: '2026-07-29', notes: 'Awaiting customs' },
];

const sesameSeed: ExportRow[] = [
  { id: 1, week: '2026-W29', shipmentNo: 'SES-26-001', buyer: 'Golden Sesame Ltd.', destination: 'Jeddah', quantity: 720, unit: 'MT', containers: 2, pricePerKg: 1.82, grossUsd: 131040, paidUsd: 131040, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Bay Gold', contract: 'CTR-SES-1101', status: 'Delivered', date: '2026-07-09', notes: 'Premium white sesame' },
  { id: 2, week: '2026-W29', shipmentNo: 'SES-26-002', buyer: 'Nile Grain House', destination: 'Cairo', quantity: 540, unit: 'MT', containers: 2, pricePerKg: 1.76, grossUsd: 95040, paidUsd: 60000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Nile Gold', contract: 'CTR-SES-1102', status: 'In Transit', date: '2026-07-10', notes: 'Checking quality' },
  { id: 3, week: '2026-W30', shipmentNo: 'SES-26-003', buyer: 'Sahara & Sons', destination: 'Dubai', quantity: 680, unit: 'MT', containers: 2, pricePerKg: 1.84, grossUsd: 125120, paidUsd: 125120, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Dune Pearl', contract: 'CTR-SES-1103', status: 'Delivered', date: '2026-07-13', notes: 'Delivered clean' },
  { id: 4, week: '2026-W30', shipmentNo: 'SES-26-004', buyer: 'Red Sea Trade', destination: 'Aqaba', quantity: 610, unit: 'MT', containers: 2, pricePerKg: 1.8, grossUsd: 109800, paidUsd: 0, paymentStatus: 'Open', cleaning: 'Pending', product: 'Sesame Seeds', vessel: 'MV Red Sea', contract: 'CTR-SES-1104', status: 'Pending', date: '2026-07-16', notes: 'Inspection review' },
  { id: 5, week: '2026-W31', shipmentNo: 'SES-26-005', buyer: 'Aman Importers', destination: 'Riyadh', quantity: 790, unit: 'MT', containers: 3, pricePerKg: 1.88, grossUsd: 148520, paidUsd: 75000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Najd Star', contract: 'CTR-SES-1105', status: 'On Vessel', date: '2026-07-19', notes: 'Shipment loaded' },
  { id: 6, week: '2026-W31', shipmentNo: 'SES-26-006', buyer: 'Desert Grain MW', destination: 'Khartoum', quantity: 660, unit: 'MT', containers: 2, pricePerKg: 1.78, grossUsd: 117480, paidUsd: 117480, paymentStatus: 'Paid', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Khartoum', contract: 'CTR-SES-1106', status: 'Delivered', date: '2026-07-22', notes: 'Good quality' },
  { id: 7, week: '2026-W32', shipmentNo: 'SES-26-007', buyer: 'Hawker Markets', destination: 'Muscat', quantity: 740, unit: 'MT', containers: 3, pricePerKg: 1.86, grossUsd: 137640, paidUsd: 90000, paymentStatus: 'Partial', cleaning: 'Cleaned', product: 'Sesame Seeds', vessel: 'MV Oman Gold', contract: 'CTR-SES-1107', status: 'At Port', date: '2026-07-24', notes: 'Awaiting discharge' },
];

const dataMap: Record<ExportTab, ExportRow[]> = {
  coffee: coffeeSeed,
  sesame: sesameSeed,
};

const money = (value = 0): string => currency.format(Number(value) || 0);
const compactMoney = (value = 0): string => compactCurrency.format(Number(value) || 0);
const normalizeStatus = (value?: string): string => String(value || '').toLowerCase().replace(/\s+/g, '-');
const formatShort = (value?: string | null): string => value ?? '—';

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
  const [chartType, setChartType] = useState<'line' | 'bar' | 'pie'>('line');

  const apiBaseUrl = typeof window !== 'undefined' ? `http://${window.location.hostname}:3002` : 'http://localhost:3002';

  useEffect(() => {
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
        const nextRows = dataMap[activeTab];
        setRows(nextRows);
        setSelectedRow(nextRows[0] ?? null);
        setApiState('local');
      }
    };

    void fetchBootstrap();
  }, [activeTab, apiBaseUrl]);

  useEffect(() => {
    const nextRows = dataMap[activeTab];
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

  const buildReportRows = () => filteredRows.map((row) => ({
    Shipment: row.shipmentNo,
    Buyer: row.buyer,
    Destination: row.destination,
    Product: row.product,
    Quantity: `${row.quantity} ${row.unit}`,
    Containers: row.containers,
    'Gross USD': row.grossUsd,
    'Paid USD': row.paidUsd,
    'Payment Status': row.paymentStatus,
    Cleaning: row.cleaning,
    Vessel: row.vessel,
    Status: row.status,
  }));

  const exportCsv = (): void => {
    if (!filteredRows.length) return;

    const headers = [
      'Shipment No',
      'Buyer',
      'Destination',
      'Product',
      'Quantity',
      'Containers',
      'Gross USD',
      'Paid USD',
      'Payment Status',
      'Cleaning',
      'Vessel',
      'Status',
    ];

    const csvRows = filteredRows.map((row) => [
      row.shipmentNo,
      row.buyer,
      row.destination,
      row.product,
      `${row.quantity} ${row.unit}`,
      row.containers,
      row.grossUsd,
      row.paidUsd,
      row.paymentStatus,
      row.cleaning,
      row.vessel,
      row.status,
    ]);

    const csv = [headers, ...csvRows].map((line) => line.join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${activeTab}-export.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const exportExcel = (): void => {
    if (!filteredRows.length) return;

    const workbook = XLSX.utils.book_new();
    const reportRows = buildReportRows();
    const reportTitle = activeTab === 'coffee' ? 'Coffee' : 'Sesame';
    const summaryRows = [
      ['Report', `${reportTitle} Export Overview`],
      ['Generated', new Date().toLocaleString()],
      ['Shipment Count', totals.shipmentCount],
      ['Volume (MT)', totals.totalMt],
      ['Gross USD', totals.gross],
      ['Paid USD', totals.paid],
      ['Open USD', totals.open],
      [],
      ['Shipment', 'Buyer', 'Destination', 'Product', 'Quantity', 'Containers', 'Gross USD', 'Paid USD', 'Payment Status', 'Cleaning', 'Vessel', 'Status'],
      ...reportRows.map((row) => [
        row.Shipment,
        row.Buyer,
        row.Destination,
        row.Product,
        row.Quantity,
        row.Containers,
        row['Gross USD'],
        row['Paid USD'],
        row['Payment Status'],
        row.Cleaning,
        row.Vessel,
        row.Status,
      ]),
      [],
      ['Footer', 'Prepared by Export Operations'],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(summaryRows);
    worksheet['!cols'] = [
      { wch: 18 }, { wch: 25 }, { wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 22 }, { wch: 14 },
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, `${reportTitle} Export`);
    XLSX.writeFile(workbook, `${activeTab}-export-report.xlsx`);
  };

  const exportPdf = (): void => {
    if (!filteredRows.length) return;

    const doc = new jsPDF({ orientation: 'landscape' });
    const title = `${activeTab === 'coffee' ? 'Coffee' : 'Sesame'} Export Overview`;

    doc.setFontSize(18);
    doc.setTextColor(19, 41, 57);
    doc.text(title, 14, 18);

    doc.setFontSize(10);
    doc.setTextColor(80, 96, 110);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 26);
    doc.text(`Shipments: ${totals.shipmentCount}   Volume: ${totals.totalMt} MT   Gross: ${money(totals.gross)}   Paid: ${money(totals.paid)}   Open: ${money(totals.open)}`, 14, 32);

    autoTable(doc, {
      head: [[
        'Shipment',
        'Buyer',
        'Destination',
        'Product',
        'Quantity',
        'Containers',
        'Gross USD',
        'Paid USD',
        'Payment',
        'Cleaning',
        'Status',
      ]],
      body: filteredRows.map((row) => [
        row.shipmentNo,
        row.buyer,
        row.destination,
        row.product,
        `${row.quantity} ${row.unit}`,
        String(row.containers),
        String(row.grossUsd),
        String(row.paidUsd),
        row.paymentStatus,
        row.cleaning,
        row.status,
      ]),
      startY: 40,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 3 },
      headStyles: { fillColor: [23, 66, 95], textColor: [255, 255, 255], fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [240, 247, 252] },
      margin: { left: 14, right: 14 },
    });

    const finalY = (doc as any).lastAutoTable.finalY ?? 150;
    doc.setFontSize(10);
    doc.setTextColor(74, 92, 104);
    doc.text('Prepared by Export Operations', 14, finalY + 12);
    doc.text('End of report', 260, finalY + 12);

    doc.save(`${activeTab}-export-report.pdf`);
  };

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-block">
          <small className="eyebrow">Export operations</small>
          <h1>{activeTab === 'coffee' ? 'Coffee' : 'Sesame'} export overview</h1>
          <div className="brand-meta">
            <span className="tiny-chip">Weekly tracking</span>
            <span className="tiny-chip">Finance view</span>
          </div>
        </div>

        <div className="header-tools">
          <span className={`badge ${apiState === 'postgres' ? 'postgres' : 'local'}`}>
            {apiState === 'postgres' ? 'Live workbook' : 'Local fallback'}
          </span>
          <button className="ghost-button" type="button" onClick={() => window.location.reload()}>
            Refresh
          </button>
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
        </div>

        <div className="export-actions">
          <button type="button" className="primary-button" onClick={exportCsv}>
            Export CSV
          </button>
          <button type="button" className="secondary-button" onClick={exportExcel}>
            Export Excel
          </button>
          <button type="button" className="secondary-button" onClick={exportPdf}>
            Export PDF
          </button>
        </div>
      </section>

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

      <div className="layout">
        <section className="panel table-panel">
          <div className="controls">
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
            </div>
          ) : (
            <div className="empty-state">Select a record to view its dossier.</div>
          )}
        </aside>
      </div>
    </div>
  );
}
