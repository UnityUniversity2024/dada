/** Mock coffee & sesame shipment rows for ExportApp */

export type ExportRowSeed = {
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

export const coffeeSeed: ExportRowSeed[] = [
  {
    id: 1,
    week: '2026-W32',
    shipmentNo: 'COF-26-088',
    buyer: 'European Roaster Group',
    destination: 'Djibouti / Rotterdam',
    quantity: 320,
    unit: 'MT',
    containers: 16,
    pricePerKg: 3.73,
    grossUsd: 1193600,
    paidUsd: 716160,
    paymentStatus: 'Partial',
    cleaning: 'Grade A',
    product: 'Sidamo Washed Arabica',
    vessel: 'MSC DJIBOUTI',
    contract: 'CTR-SID-2026-08',
    status: 'On Vessel',
    date: '2026-08-12',
    notes: 'LC first draw 60%',
  },
  {
    id: 2,
    week: '2026-W30',
    shipmentNo: 'COF-26-071',
    buyer: 'Nordic Specialty AB',
    destination: 'Gothenburg',
    quantity: 96,
    unit: 'MT',
    containers: 5,
    pricePerKg: 4.1,
    grossUsd: 393600,
    paidUsd: 393600,
    paymentStatus: 'Paid',
    cleaning: 'Grade A',
    product: 'Yirgacheffe Gr.1',
    vessel: 'MAERSK ADEN',
    contract: 'CTR-YRG-2026-07',
    status: 'Delivered',
    date: '2026-07-22',
    notes: 'Settled',
  },
  {
    id: 3,
    week: '2026-W34',
    shipmentNo: 'COF-26-095',
    buyer: 'Tokyo Coffee Co',
    destination: 'Yokohama',
    quantity: 60,
    unit: 'MT',
    containers: 3,
    pricePerKg: 4.25,
    grossUsd: 255000,
    paidUsd: 0,
    paymentStatus: 'Open',
    cleaning: 'Grade B',
    product: 'Limu Natural',
    vessel: 'NYK PEARL',
    contract: 'CTR-LIM-2026-08',
    status: 'At Port',
    date: '2026-08-20',
    notes: 'Awaiting phytosanitary',
  },
];

export const sesameSeed: ExportRowSeed[] = [
  {
    id: 1,
    week: '2026-W28',
    shipmentNo: 'SESAME/572/07/2026',
    buyer: 'XIAMEN',
    destination: 'CHINA',
    quantity: 190,
    unit: 'MT',
    containers: 10,
    pricePerKg: 1.24,
    grossUsd: 235600,
    paidUsd: 0,
    paymentStatus: 'Open',
    cleaning: 'COMPLETED',
    product: 'sesame Wollega',
    vessel: 'CMA CGM',
    contract: 'CTR-SES-2026-07',
    status: 'On Vessel',
    date: '2026-07-20',
    notes: 'GIVEN',
  },
  {
    id: 2,
    week: '2026-W31',
    shipmentNo: 'SESAME/574/08/2026',
    buyer: 'INVICTUS',
    destination: 'SLOVENIA',
    quantity: 114,
    unit: 'MT',
    containers: 6,
    pricePerKg: 1.33,
    grossUsd: 151620,
    paidUsd: 50000,
    paymentStatus: 'Partial',
    cleaning: 'PENDING',
    product: 'sesame humera',
    vessel: 'NOT GIVEN',
    contract: 'CTR-SES-2026-08',
    status: 'Pending',
    date: '2026-08-05',
    notes: 'NOT GIVEN',
  },
  {
    id: 3,
    week: '2026-W33',
    shipmentNo: 'SESAME/581/08/2026',
    buyer: 'ITOCHU',
    destination: 'JAPAN',
    quantity: 114,
    unit: 'MT',
    containers: 6,
    pricePerKg: 1.5,
    grossUsd: 171000,
    paidUsd: 171000,
    paymentStatus: 'Paid',
    cleaning: 'COMPLETED',
    product: 'sesame humera',
    vessel: 'ONE HARMONY',
    contract: 'CTR-SES-2026-08B',
    status: 'Delivered',
    date: '2026-08-18',
    notes: 'Settled TT',
  },
];
