export interface FunnelOrderItem {
  itemId: string;
  itemName: string;
  hsnCode?: string;
  quantity: number;
  unit: string;
  rate: number;
  amount: number;
  tripNo?: string;
  vehicleNo?: string;
  invoiceNo?: string;
  invoiceDate?: string;
  status?: "Delivered" | "In Transit" | "Pending" | "Assigned";
}

export interface FunnelTripItem {
  tripId: string;
  tripNo: string;
  vehicleNo: string;
  driverName: string;
  driverPhone?: string;
  transporter: string;
  lrNo: string;
  dispatchDate: string;
  deliveryDate?: string;
  status: "Assigned" | "In Transit" | "Delivered" | "Pending";
  quantityBags?: number;
  tonnageMT?: number;
  invoiceNo?: string;
}

export interface FunnelInvoiceItem {
  invoiceId: string;
  invoiceNo: string;
  invoiceDate: string;
  amount: number;
  status: "Generated" | "Dispatched" | "Pending";
  taxAmount?: number;
  tripNo?: string;
}

export interface PurchaseDeliveryFunnelItem {
  id: string;
  sNo: number;
  orderId: string;
  orderDate: string;
  retailerName: string;
  city: string;
  state: string;
  itemName?: string;
  items?: FunnelOrderItem[];
  trips: FunnelTripItem[];
  invoices: FunnelInvoiceItem[];
  balanceToSettle: number;
  totalAmount: number;
  paidAmount: number;
  batch: string;
  status: "Completed" | "Not Completed";
  statusReason?: string;
  remarks?: string;
}

/**
 * Computes status strictly based on column completeness:
 * If order has assigned trip(s) (count > 0),
 * converted into invoice(s) (count > 0),
 * payment balance settled (balanceToSettle === 0),
 * and batch is assigned, status is "Completed".
 * Otherwise "Not Completed".
 */
export const computeItemStatus = (
  item: Omit<PurchaseDeliveryFunnelItem, "status" | "statusReason">
): { status: "Completed" | "Not Completed"; statusReason: string } => {
  const missingStages: string[] = [];

  if (!item.trips || item.trips.length === 0) {
    missingStages.push("Trip not assigned");
  }
  if (!item.invoices || item.invoices.length === 0) {
    missingStages.push("Invoice not converted");
  }
  if (item.balanceToSettle > 0) {
    missingStages.push("Balance pending");
  }

  if (missingStages.length === 0) {
    return {
      status: "Completed",
      statusReason: "All stages completed & balance settled",
    };
  }

  return {
    status: "Not Completed",
    statusReason: missingStages.join(", "),
  };
};

/* ================= COMPREHENSIVE DUMMY DATASET ================= */

export const DUMMY_PURCHASE_DELIVERY_DATA: PurchaseDeliveryFunnelItem[] = [
  // 1. Matches Excel Row 1: Trips 'xy' & 'tr2' (2 trips), Invoice 'tr1', payment pending, Not Completed
  {
    id: "pd-1",
    sNo: 1,
    orderId: "ORD-2026-901",
    orderDate: "10/09/2026",
    retailerName: "AMMAN TRADERS",
    city: "MADURAI",
    state: "TAMIL NADU",
    itemName: "SUPERFINE SONA MASOORI RICE 25KG",
    items: [
      {
        itemId: "pd1-it-1",
        itemName: "SUPERFINE SONA MASOORI RICE 25KG",
        quantity: 80,
        unit: "Bags",
        rate: 1350,
        amount: 108000,
        tripNo: "xy",
        vehicleNo: "TN 59 AB 1234",
        invoiceNo: "tr1",
        invoiceDate: "12/09/2026",
        status: "Delivered",
      },
      {
        itemId: "pd1-it-2",
        itemName: "PONNI BOILED RICE CLASSIC 25KG",
        quantity: 55,
        unit: "Bags",
        rate: 1400,
        amount: 77000,
        tripNo: "tr2",
        vehicleNo: "TN 58 CZ 8890",
        invoiceNo: "tr1",
        invoiceDate: "12/09/2026",
        status: "In Transit",
      },
    ],
    trips: [
      {
        tripId: "T-01",
        tripNo: "xy",
        vehicleNo: "TN 59 AB 1234",
        driverName: "Murugan",
        driverPhone: "+91 98421 11223",
        transporter: "SRI BALAJI TRANSPORT",
        lrNo: "LR-9821",
        dispatchDate: "11/09/2026",
        deliveryDate: "13/09/2026",
        status: "Delivered",
        quantityBags: 300,
        tonnageMT: 7.5,
        invoiceNo: "tr1",
      },
      {
        tripId: "T-02",
        tripNo: "tr2",
        vehicleNo: "TN 58 CZ 8890",
        driverName: "Karthik",
        driverPhone: "+91 98421 55667",
        transporter: "METRO FREIGHT CARRIERS",
        lrNo: "LR-9835",
        dispatchDate: "12/09/2026",
        deliveryDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 300,
        tonnageMT: 7.5,
        invoiceNo: "tr1",
      },
    ],
    invoices: [
      {
        invoiceId: "INV-01",
        invoiceNo: "tr1",
        invoiceDate: "12/09/2026",
        amount: 185000,
        status: "Generated",
        taxAmount: 9250,
        tripNo: "xy, tr2",
      },
    ],
    balanceToSettle: 85000,
    totalAmount: 185000,
    paidAmount: 100000,
    batch: "BATCH-TR1",
    status: "Not Completed",
    statusReason: "Balance (₹ 85,000)",
    remarks: "Second trip in transit, payment awaiting confirmation",
  },

  // 2. Matches Excel Row 2: PO 'zz', Trip '1', Invoice 'abc', payment '17P' (fully settled), Completed
  {
    id: "pd-2",
    sNo: 2,
    orderId: "zz",
    orderDate: "08/09/2026",
    retailerName: "VENKATESHWARA AGRO",
    city: "SALEM",
    state: "TAMIL NADU",
    itemName: "ROYAL GOLD TOOR DAL PREMIUM",
    items: [
      {
        itemId: "pd2-it-1",
        itemName: "ROYAL GOLD TOOR DAL PREMIUM",
        quantity: 150,
        unit: "Bags",
        rate: 900,
        amount: 135000,
        tripNo: "1",
        vehicleNo: "TN 30 BD 5541",
        invoiceNo: "abc",
        invoiceDate: "09/09/2026",
        status: "Delivered",
      },
      {
        itemId: "pd2-it-2",
        itemName: "CHANA DAL REGULAR 30KG",
        quantity: 100,
        unit: "Bags",
        rate: 800,
        amount: 80000,
        tripNo: "1",
        vehicleNo: "TN 30 BD 5541",
        invoiceNo: "abc",
        invoiceDate: "09/09/2026",
        status: "Delivered",
      },
    ],
    trips: [
      {
        tripId: "T-03",
        tripNo: "1",
        vehicleNo: "TN 30 BD 5541",
        driverName: "Ramesh Kumar",
        driverPhone: "+91 94432 77112",
        transporter: "VRL LOGISTICS",
        lrNo: "LR-44210",
        dispatchDate: "09/09/2026",
        deliveryDate: "10/09/2026",
        status: "Delivered",
        quantityBags: 450,
        tonnageMT: 13.5,
        invoiceNo: "abc",
      },
    ],
    invoices: [
      {
        invoiceId: "INV-02",
        invoiceNo: "abc",
        invoiceDate: "09/09/2026",
        amount: 215000,
        status: "Dispatched",
        taxAmount: 10750,
        tripNo: "1",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 215000,
    paidAmount: 215000,
    batch: "17P",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Delivered on schedule and cleared in full via NEFT (Ref: 17P)",
  },

  // 3. Matches Excel Row 3: PO 'sss', Trip '1', '2', '3' (3 trips!), no invoice, Not Completed
  {
    id: "pd-3",
    sNo: 3,
    orderId: "sss",
    orderDate: "14/09/2026",
    retailerName: "MAHALAKSHMI ENTERPRISES",
    city: "COIMBATORE",
    state: "TAMIL NADU",
    itemName: "PREMIUM URAD DAL GOTA 50KG",
    items: [
      {
        itemId: "pd3-it-1",
        itemName: "PREMIUM URAD DAL GOTA 50KG",
        quantity: 100,
        unit: "Bags",
        rate: 1200,
        amount: 120000,
        tripNo: "1",
        vehicleNo: "TN 38 BX 4412",
        invoiceNo: "Unassigned",
        status: "In Transit",
      },
      {
        itemId: "pd3-it-2",
        itemName: "MOONG WHOLE SPECIAL 50KG",
        quantity: 100,
        unit: "Bags",
        rate: 1200,
        amount: 120000,
        tripNo: "2",
        vehicleNo: "TN 38 CA 9981",
        invoiceNo: "Unassigned",
        status: "In Transit",
      },
      {
        itemId: "pd3-it-3",
        itemName: "TOOR DAL DESI BULK 50KG",
        quantity: 100,
        unit: "Bags",
        rate: 1200,
        amount: 120000,
        tripNo: "3",
        vehicleNo: "TN 37 DY 1109",
        invoiceNo: "Unassigned",
        status: "Assigned",
      },
    ],
    trips: [
      {
        tripId: "T-04",
        tripNo: "1",
        vehicleNo: "TN 38 BX 4412",
        driverName: "Selvam P",
        driverPhone: "+91 97891 22334",
        transporter: "ROYAL EXPRESS CARGO",
        lrNo: "LR-78101",
        dispatchDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 200,
        tonnageMT: 10.0,
      },
      {
        tripId: "T-05",
        tripNo: "2",
        vehicleNo: "TN 38 CA 9981",
        driverName: "Govindasamy",
        driverPhone: "+91 97891 66778",
        transporter: "ROYAL EXPRESS CARGO",
        lrNo: "LR-78102",
        dispatchDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 200,
        tonnageMT: 10.0,
      },
      {
        tripId: "T-06",
        tripNo: "3",
        vehicleNo: "TN 37 DY 1109",
        driverName: "Anand Babu",
        driverPhone: "+91 97891 88990",
        transporter: "SRI BALAJI TRANSPORT",
        lrNo: "LR-78103",
        dispatchDate: "15/09/2026",
        status: "Assigned",
        quantityBags: 200,
        tonnageMT: 10.0,
      },
    ],
    invoices: [],
    balanceToSettle: 360000,
    totalAmount: 360000,
    paidAmount: 0,
    batch: "BCH-SSS-9",
    status: "Not Completed",
    statusReason: "Invoice not converted, Balance pending",
    remarks: "3 trips dispatched from central warehouse; bill generation pending gate check",
  },

  // 4. Order with NO trip assigned yet -> Not Completed
  {
    id: "pd-4",
    sNo: 4,
    orderId: "ORD-2026-1044",
    orderDate: "15/09/2026",
    retailerName: "KAVERI GRAIN TRADERS",
    city: "ERODE",
    state: "TAMIL NADU",
    itemName: "CHANA DAL REGULAR 30KG",
    items: [
      {
        itemId: "pd4-it-1",
        itemName: "CHANA DAL REGULAR 30KG",
        quantity: 100,
        unit: "Bags",
        rate: 850,
        amount: 85000,
        tripNo: "Unassigned",
        invoiceNo: "Unassigned",
        status: "Pending",
      },
      {
        itemId: "pd4-it-2",
        itemName: "MASOOR DAL SPLIT 30KG",
        quantity: 50,
        unit: "Bags",
        rate: 1100,
        amount: 55000,
        tripNo: "Unassigned",
        invoiceNo: "Unassigned",
        status: "Pending",
      },
    ],
    trips: [],
    invoices: [],
    balanceToSettle: 140000,
    totalAmount: 140000,
    paidAmount: 0,
    batch: "BCH-KAV-01",
    status: "Not Completed",
    statusReason: "Trip not assigned, Invoice not converted",
    remarks: "Fresh purchase order placed today, awaiting vehicle assignment",
  },

  // 5. Order with 2 Trips, 2 Invoices, balance fully settled -> Completed
  {
    id: "pd-5",
    sNo: 5,
    orderId: "PO-AGRO-208",
    orderDate: "06/09/2026",
    retailerName: "SRI KRISHNA TRADING CO",
    city: "CHENNAI",
    state: "TAMIL NADU",
    itemName: "PONNI BOILED RICE CLASSIC 25KG",
    items: [
      {
        itemId: "pd5-it-1",
        itemName: "PONNI BOILED RICE CLASSIC 25KG",
        quantity: 100,
        unit: "Bags",
        rate: 1400,
        amount: 140000,
        tripNo: "TRP-201",
        vehicleNo: "TN 04 AA 7712",
        invoiceNo: "INV-SK-01",
        invoiceDate: "08/09/2026",
        status: "Delivered",
      },
      {
        itemId: "pd5-it-2",
        itemName: "JEERA SAMBA RICE PREMIUM 25KG",
        quantity: 80,
        unit: "Bags",
        rate: 1750,
        amount: 140000,
        tripNo: "TRP-202",
        vehicleNo: "TN 05 BG 3349",
        invoiceNo: "INV-SK-02",
        invoiceDate: "08/09/2026",
        status: "Delivered",
      },
    ],
    trips: [
      {
        tripId: "T-07",
        tripNo: "TRP-201",
        vehicleNo: "TN 04 AA 7712",
        driverName: "Balan",
        driverPhone: "+91 94440 12345",
        transporter: "CHENNAI FREIGHTWAYS",
        lrNo: "LR-11029",
        dispatchDate: "07/09/2026",
        deliveryDate: "08/09/2026",
        status: "Delivered",
        quantityBags: 350,
        tonnageMT: 8.75,
        invoiceNo: "INV-SK-01",
      },
      {
        tripId: "T-08",
        tripNo: "TRP-202",
        vehicleNo: "TN 05 BG 3349",
        driverName: "Dhanapal",
        driverPhone: "+91 94440 67890",
        transporter: "CHENNAI FREIGHTWAYS",
        lrNo: "LR-11035",
        dispatchDate: "07/09/2026",
        deliveryDate: "08/09/2026",
        status: "Delivered",
        quantityBags: 350,
        tonnageMT: 8.75,
        invoiceNo: "INV-SK-02",
      },
    ],
    invoices: [
      {
        invoiceId: "INV-03",
        invoiceNo: "INV-SK-01",
        invoiceDate: "08/09/2026",
        amount: 140000,
        status: "Dispatched",
        tripNo: "TRP-201",
      },
      {
        invoiceId: "INV-04",
        invoiceNo: "INV-SK-02",
        invoiceDate: "08/09/2026",
        amount: 140000,
        status: "Dispatched",
        tripNo: "TRP-202",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 280000,
    paidAmount: 280000,
    batch: "BCH-SKT-77",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Both consignments delivered and verified. Account reconciled.",
  },

  // 6. Order with 1 trip, 1 invoice, partial balance pending -> Not Completed
  {
    id: "pd-6",
    sNo: 6,
    orderId: "ORD-2026-1180",
    orderDate: "11/09/2026",
    retailerName: "ANNAPOORNA STORES",
    city: "TIRUCHIRAPPALLI",
    state: "TAMIL NADU",
    itemName: "MOONG WHOLE PREMIUM 25KG",
    trips: [
      {
        tripId: "T-09",
        tripNo: "TR-89",
        vehicleNo: "TN 45 AX 8921",
        driverName: "Senthil Nathan",
        driverPhone: "+91 98432 33445",
        transporter: "DELTA ROADWAYS",
        lrNo: "LR-6231",
        dispatchDate: "12/09/2026",
        deliveryDate: "13/09/2026",
        status: "Delivered",
        quantityBags: 200,
        tonnageMT: 5.0,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-05",
        invoiceNo: "INV-AP-401",
        invoiceDate: "13/09/2026",
        amount: 98000,
        status: "Generated",
      },
    ],
    balanceToSettle: 38000,
    totalAmount: 98000,
    paidAmount: 60000,
    batch: "BATCH-2026-AP",
    status: "Not Completed",
    statusReason: "Balance (₹ 38,000)",
    remarks: "Part payment received; balance due by 20/09/2026",
  },

  // 7. Order with 1 trip, 1 invoice, balance settled -> Completed
  {
    id: "pd-7",
    sNo: 7,
    orderId: "PO-7762",
    orderDate: "05/09/2026",
    retailerName: "THIRUMALAI FOODS",
    city: "VELLORE",
    state: "TAMIL NADU",
    itemName: "RAW RICE GRADE A 50KG",
    trips: [
      {
        tripId: "T-10",
        tripNo: "TR-104",
        vehicleNo: "TN 23 BK 9012",
        driverName: "Mani",
        driverPhone: "+91 99441 55667",
        transporter: "VELLORE LOGISTICS",
        lrNo: "LR-90112",
        dispatchDate: "06/09/2026",
        deliveryDate: "07/09/2026",
        status: "Delivered",
        quantityBags: 250,
        tonnageMT: 12.5,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-06",
        invoiceNo: "INV-TM-77",
        invoiceDate: "07/09/2026",
        amount: 195000,
        status: "Dispatched",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 195000,
    paidAmount: 195000,
    batch: "BCH-TF-12",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Closed with zero variance.",
  },

  // 8. Order with 2 Trips assigned, but NO invoice yet -> Not Completed
  {
    id: "pd-8",
    sNo: 8,
    orderId: "ORD-2026-1299",
    orderDate: "13/09/2026",
    retailerName: "PADMAVATHI WHOLESALE",
    city: "TIRUPUR",
    state: "TAMIL NADU",
    itemName: "MASOOR DAL SPLIT 30KG",
    trips: [
      {
        tripId: "T-11",
        tripNo: "TRP-A",
        vehicleNo: "TN 39 DF 4510",
        driverName: "Vijay",
        driverPhone: "+91 98422 99001",
        transporter: "KONGU CARRIERS",
        lrNo: "LR-33100",
        dispatchDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 200,
        tonnageMT: 6.0,
      },
      {
        tripId: "T-12",
        tripNo: "TRP-B",
        vehicleNo: "TN 39 DF 4511",
        driverName: "Shankar",
        driverPhone: "+91 98422 99002",
        transporter: "KONGU CARRIERS",
        lrNo: "LR-33101",
        dispatchDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 200,
        tonnageMT: 6.0,
      },
    ],
    invoices: [],
    balanceToSettle: 162000,
    totalAmount: 162000,
    paidAmount: 0,
    batch: "BATCH-PW-02",
    status: "Not Completed",
    statusReason: "Invoice not converted, Balance pending",
    remarks: "Vehicles en route to godown, weighment pending",
  },

  // 9. Order with 1 trip, 1 invoice, paid in full -> Completed
  {
    id: "pd-9",
    sNo: 9,
    orderId: "PO-9914",
    orderDate: "03/09/2026",
    retailerName: "SRI MURUGAN OIL & GRAIN MILLS",
    city: "DINDIGUL",
    state: "TAMIL NADU",
    itemName: "FRIED GRAM (POTTUKADALAI) 25KG",
    trips: [
      {
        tripId: "T-13",
        tripNo: "TR-55",
        vehicleNo: "TN 57 BH 6701",
        driverName: "Rajendran",
        driverPhone: "+91 94420 88991",
        transporter: "DINDIGUL EXPRESS",
        lrNo: "LR-88219",
        dispatchDate: "04/09/2026",
        deliveryDate: "05/09/2026",
        status: "Delivered",
        quantityBags: 300,
        tonnageMT: 7.5,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-07",
        invoiceNo: "INV-DGL-09",
        invoiceDate: "05/09/2026",
        amount: 112500,
        status: "Dispatched",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 112500,
    paidAmount: 112500,
    batch: "18X-DGL",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Quality inspected and approved. Payment fully settled.",
  },

  // 10. Order with 4 trips, 3 invoices, balance pending -> Not Completed
  {
    id: "pd-10",
    sNo: 10,
    orderId: "ORD-2026-0810",
    orderDate: "02/09/2026",
    retailerName: "BHARATHI AGRO COMMODITIES",
    city: "THANJAVUR",
    state: "TAMIL NADU",
    itemName: "DELUXE JEERA SAMBA RICE",
    trips: [
      {
        tripId: "T-14",
        tripNo: "T1",
        vehicleNo: "TN 49 AZ 1001",
        driverName: "Prabu",
        transporter: "CAUVERY TRANSPORTS",
        lrNo: "LR-5511",
        dispatchDate: "03/09/2026",
        deliveryDate: "04/09/2026",
        status: "Delivered",
        quantityBags: 250,
        tonnageMT: 6.25,
      },
      {
        tripId: "T-15",
        tripNo: "T2",
        vehicleNo: "TN 49 AZ 1002",
        driverName: "Chandran",
        transporter: "CAUVERY TRANSPORTS",
        lrNo: "LR-5512",
        dispatchDate: "03/09/2026",
        deliveryDate: "04/09/2026",
        status: "Delivered",
        quantityBags: 250,
        tonnageMT: 6.25,
      },
      {
        tripId: "T-16",
        tripNo: "T3",
        vehicleNo: "TN 49 AZ 1003",
        driverName: "Prakash",
        transporter: "CAUVERY TRANSPORTS",
        lrNo: "LR-5513",
        dispatchDate: "04/09/2026",
        deliveryDate: "05/09/2026",
        status: "Delivered",
        quantityBags: 250,
        tonnageMT: 6.25,
      },
      {
        tripId: "T-17",
        tripNo: "T4",
        vehicleNo: "TN 49 AZ 1004",
        driverName: "Murugesh",
        transporter: "CAUVERY TRANSPORTS",
        lrNo: "LR-5514",
        dispatchDate: "05/09/2026",
        status: "In Transit",
        quantityBags: 250,
        tonnageMT: 6.25,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-08",
        invoiceNo: "INV-TNJ-01",
        invoiceDate: "04/09/2026",
        amount: 125000,
        status: "Dispatched",
      },
      {
        invoiceId: "INV-09",
        invoiceNo: "INV-TNJ-02",
        invoiceDate: "04/09/2026",
        amount: 125000,
        status: "Dispatched",
      },
      {
        invoiceId: "INV-10",
        invoiceNo: "INV-TNJ-03",
        invoiceDate: "05/09/2026",
        amount: 125000,
        status: "Generated",
      },
    ],
    balanceToSettle: 125000,
    totalAmount: 500000,
    paidAmount: 375000,
    batch: "BCH-TNJ-99",
    status: "Not Completed",
    statusReason: "4th invoice pending, Balance (₹ 1,25,000)",
    remarks: "Final trip arriving today; final invoice will be raised upon arrival",
  },

  // 11. Completed Order: 2 trips, 2 invoices, settled
  {
    id: "pd-11",
    sNo: 11,
    orderId: "PO-4412",
    orderDate: "01/09/2026",
    retailerName: "GOKUL AGRO PRODUCTS",
    city: "NAGAPATTINAM",
    state: "TAMIL NADU",
    itemName: "SUPER BRAN FLOUR 50KG",
    trips: [
      {
        tripId: "T-18",
        tripNo: "TRP-81",
        vehicleNo: "TN 51 C 2201",
        driverName: "Kandasamy",
        driverPhone: "+91 94433 11223",
        transporter: "EAST COAST LOGISTICS",
        lrNo: "LR-1190",
        dispatchDate: "02/09/2026",
        deliveryDate: "03/09/2026",
        status: "Delivered",
        quantityBags: 200,
        tonnageMT: 10.0,
      },
      {
        tripId: "T-19",
        tripNo: "TRP-82",
        vehicleNo: "TN 51 C 2202",
        driverName: "Subramani",
        driverPhone: "+91 94433 11224",
        transporter: "EAST COAST LOGISTICS",
        lrNo: "LR-1191",
        dispatchDate: "02/09/2026",
        deliveryDate: "03/09/2026",
        status: "Delivered",
        quantityBags: 200,
        tonnageMT: 10.0,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-11",
        invoiceNo: "INV-GK-01",
        invoiceDate: "03/09/2026",
        amount: 110000,
        status: "Dispatched",
      },
      {
        invoiceId: "INV-12",
        invoiceNo: "INV-GK-02",
        invoiceDate: "03/09/2026",
        amount: 110000,
        status: "Dispatched",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 220000,
    paidAmount: 220000,
    batch: "22-GKL",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Full consignment delivered and verified at destination godown.",
  },

  // 12. Not Completed: 1 trip, 1 invoice, balance pending
  {
    id: "pd-12",
    sNo: 12,
    orderId: "ORD-2026-1502",
    orderDate: "12/09/2026",
    retailerName: "SHANMUGA TRADERS",
    city: "CUDDALORE",
    state: "TAMIL NADU",
    itemName: "GREEN GRAM POLISHED 25KG",
    trips: [
      {
        tripId: "T-20",
        tripNo: "TRP-90",
        vehicleNo: "TN 31 AF 8820",
        driverName: "Velu M",
        driverPhone: "+91 98421 99881",
        transporter: "SOUTH ROADWAYS",
        lrNo: "LR-6621",
        dispatchDate: "13/09/2026",
        deliveryDate: "14/09/2026",
        status: "Delivered",
        quantityBags: 250,
        tonnageMT: 6.25,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-13",
        invoiceNo: "INV-ST-12",
        invoiceDate: "14/09/2026",
        amount: 165000,
        status: "Generated",
      },
    ],
    balanceToSettle: 65000,
    totalAmount: 165000,
    paidAmount: 100000,
    batch: "BCH-CDL-05",
    status: "Not Completed",
    statusReason: "Balance (₹ 65,000)",
    remarks: "Advance received, balance under 7-day credit terms",
  },

  // 13. Completed Order: 3 trips, 3 invoices, fully settled
  {
    id: "pd-13",
    sNo: 13,
    orderId: "PO-8820",
    orderDate: "04/09/2026",
    retailerName: "SRI RAM GRAIN DEPOT",
    city: "TIRUNELVELI",
    state: "TAMIL NADU",
    itemName: "IDLI RICE SPECIAL QUALITY 30KG",
    trips: [
      {
        tripId: "T-21",
        tripNo: "TRP-T1",
        vehicleNo: "TN 72 BX 9001",
        driverName: "Pandian",
        transporter: "NELLAI EXPRESS FREIGHT",
        lrNo: "LR-4401",
        dispatchDate: "05/09/2026",
        deliveryDate: "06/09/2026",
        status: "Delivered",
        quantityBags: 300,
        tonnageMT: 9.0,
      },
      {
        tripId: "T-22",
        tripNo: "TRP-T2",
        vehicleNo: "TN 72 BX 9002",
        driverName: "Kumaravel",
        transporter: "NELLAI EXPRESS FREIGHT",
        lrNo: "LR-4402",
        dispatchDate: "05/09/2026",
        deliveryDate: "06/09/2026",
        status: "Delivered",
        quantityBags: 300,
        tonnageMT: 9.0,
      },
      {
        tripId: "T-23",
        tripNo: "TRP-T3",
        vehicleNo: "TN 72 BX 9003",
        driverName: "Sundar",
        transporter: "NELLAI EXPRESS FREIGHT",
        lrNo: "LR-4403",
        dispatchDate: "06/09/2026",
        deliveryDate: "07/09/2026",
        status: "Delivered",
        quantityBags: 300,
        tonnageMT: 9.0,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-14",
        invoiceNo: "INV-SRG-01",
        invoiceDate: "06/09/2026",
        amount: 145000,
        status: "Dispatched",
      },
      {
        invoiceId: "INV-15",
        invoiceNo: "INV-SRG-02",
        invoiceDate: "06/09/2026",
        amount: 145000,
        status: "Dispatched",
      },
      {
        invoiceId: "INV-16",
        invoiceNo: "INV-SRG-03",
        invoiceDate: "07/09/2026",
        amount: 145000,
        status: "Dispatched",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 435000,
    paidAmount: 435000,
    batch: "19-TNV",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Tri-consignment batch cleared and settled through RTGS.",
  },

  // 14. Not Completed: 0 trips, 0 invoices (New purchase booking)
  {
    id: "pd-14",
    sNo: 14,
    orderId: "ORD-2026-1601",
    orderDate: "15/09/2026",
    retailerName: "VASANTHAM ENTERPRISES",
    city: "KANCHEEPURAM",
    state: "TAMIL NADU",
    itemName: "BENGAL GRAM GRADE 1",
    trips: [],
    invoices: [],
    balanceToSettle: 180000,
    totalAmount: 180000,
    paidAmount: 0,
    batch: "BCH-KCH-88",
    status: "Not Completed",
    statusReason: "Trip not assigned, Invoice not converted",
    remarks: "Awaiting transport indent allocation.",
  },

  // 15. Completed Order: 1 trip, 1 invoice, paid
  {
    id: "pd-15",
    sNo: 15,
    orderId: "PO-7731",
    orderDate: "07/09/2026",
    retailerName: "SARAVANA COMMODITIES",
    city: "POLLACHI",
    state: "TAMIL NADU",
    itemName: "BLACK MATPE SPLIT 30KG",
    trips: [
      {
        tripId: "T-24",
        tripNo: "TRP-P1",
        vehicleNo: "TN 41 E 7710",
        driverName: "Ganesan",
        driverPhone: "+91 97890 44551",
        transporter: "POLLACHI CARRIERS",
        lrNo: "LR-9920",
        dispatchDate: "08/09/2026",
        deliveryDate: "09/09/2026",
        status: "Delivered",
        quantityBags: 220,
        tonnageMT: 6.6,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-17",
        invoiceNo: "INV-SC-81",
        invoiceDate: "09/09/2026",
        amount: 158000,
        status: "Dispatched",
      },
    ],
    balanceToSettle: 0,
    totalAmount: 158000,
    paidAmount: 158000,
    batch: "BCH-POL-04",
    status: "Completed",
    statusReason: "All stages completed & balance settled",
    remarks: "Closed order, zero pending liability.",
  },

  // 16. Not Completed: 2 trips, 1 invoice, balance pending
  {
    id: "pd-16",
    sNo: 16,
    orderId: "ORD-2026-1770",
    orderDate: "13/09/2026",
    retailerName: "KRISHNA PULSES & SPICES",
    city: "KARUR",
    state: "TAMIL NADU",
    itemName: "TURMERIC FINGER BULK 50KG",
    trips: [
      {
        tripId: "T-25",
        tripNo: "TR-K1",
        vehicleNo: "TN 47 B 5001",
        driverName: "Dhanush",
        transporter: "KARUR LORRY SERVICE",
        lrNo: "LR-3001",
        dispatchDate: "14/09/2026",
        deliveryDate: "15/09/2026",
        status: "Delivered",
        quantityBags: 150,
        tonnageMT: 7.5,
      },
      {
        tripId: "T-26",
        tripNo: "TR-K2",
        vehicleNo: "TN 47 B 5002",
        driverName: "Vicky",
        transporter: "KARUR LORRY SERVICE",
        lrNo: "LR-3002",
        dispatchDate: "14/09/2026",
        status: "In Transit",
        quantityBags: 150,
        tonnageMT: 7.5,
      },
    ],
    invoices: [
      {
        invoiceId: "INV-18",
        invoiceNo: "INV-KP-01",
        invoiceDate: "15/09/2026",
        amount: 140000,
        status: "Generated",
      },
    ],
    balanceToSettle: 140000,
    totalAmount: 280000,
    paidAmount: 140000,
    batch: "BCH-KRR-21",
    status: "Not Completed",
    statusReason: "Second invoice pending, Balance (₹ 1,40,000)",
    remarks: "Part consignment received, remaining consignment arriving today.",
  },
];

/* ================= HELPER CALCULATIONS ================= */

export interface FunnelSummaryMetrics {
  totalOrders: number;
  ordersWithTrips: number;
  totalTripsAssigned: number;
  tripConversionRate: number;
  ordersWithInvoices: number;
  totalInvoicesGenerated: number;
  invoiceConversionRate: number;
  ordersSettled: number;
  settledConversionRate: number;
  completedOrders: number;
  notCompletedOrders: number;
  completionRate: number;
  totalOrderAmount: number;
  totalPaidAmount: number;
  totalBalanceToSettle: number;
}

export const getFunnelMetrics = (
  data: PurchaseDeliveryFunnelItem[]
): FunnelSummaryMetrics => {
  const totalOrders = data.length;
  if (totalOrders === 0) {
    return {
      totalOrders: 0,
      ordersWithTrips: 0,
      totalTripsAssigned: 0,
      tripConversionRate: 0,
      ordersWithInvoices: 0,
      totalInvoicesGenerated: 0,
      invoiceConversionRate: 0,
      ordersSettled: 0,
      settledConversionRate: 0,
      completedOrders: 0,
      notCompletedOrders: 0,
      completionRate: 0,
      totalOrderAmount: 0,
      totalPaidAmount: 0,
      totalBalanceToSettle: 0,
    };
  }

  let totalTripsAssigned = 0;
  let ordersWithTrips = 0;
  let totalInvoicesGenerated = 0;
  let ordersWithInvoices = 0;
  let ordersSettled = 0;
  let completedOrders = 0;
  let notCompletedOrders = 0;
  let totalOrderAmount = 0;
  let totalPaidAmount = 0;
  let totalBalanceToSettle = 0;

  data.forEach((item) => {
    totalOrderAmount += item.totalAmount;
    totalPaidAmount += item.paidAmount;
    totalBalanceToSettle += item.balanceToSettle;

    const tripsCount = item.trips?.length || 0;
    const invoiceCount = item.invoices?.length || 0;

    totalTripsAssigned += tripsCount;
    if (tripsCount > 0) ordersWithTrips++;

    totalInvoicesGenerated += invoiceCount;
    if (invoiceCount > 0) ordersWithInvoices++;

    if (item.balanceToSettle === 0 && item.totalAmount > 0) {
      ordersSettled++;
    }

    if (item.status === "Completed") {
      completedOrders++;
    } else {
      notCompletedOrders++;
    }
  });

  return {
    totalOrders,
    ordersWithTrips,
    totalTripsAssigned,
    tripConversionRate: Math.round((ordersWithTrips / totalOrders) * 100),
    ordersWithInvoices,
    totalInvoicesGenerated,
    invoiceConversionRate: Math.round((ordersWithInvoices / totalOrders) * 100),
    ordersSettled,
    settledConversionRate: Math.round((ordersSettled / totalOrders) * 100),
    completedOrders,
    notCompletedOrders,
    completionRate: Math.round((completedOrders / totalOrders) * 100),
    totalOrderAmount,
    totalPaidAmount,
    totalBalanceToSettle,
  };
};

/**
 * Returns detailed items for an order.
 * If the order already defines `items`, returns them directly.
 * Otherwise, generates item breakdown correlated with trips and invoices.
 */
export const getOrderItems = (row: PurchaseDeliveryFunnelItem): FunnelOrderItem[] => {
  if (row.items && row.items.length > 0) {
    return row.items;
  }

  // If order has multiple trips, map an item to each trip:
  if (row.trips && row.trips.length > 1) {
    const count = row.trips.length;
    const baseAmt = Math.floor(row.totalAmount / count);
    return row.trips.map((t, idx) => {
      const isLast = idx === count - 1;
      const amt = isLast ? row.totalAmount - baseAmt * (count - 1) : baseAmt;
      const bags = t.quantityBags || 120;
      const inv = row.invoices[idx] || row.invoices[0];
      return {
        itemId: `${row.id}-item-${idx + 1}`,
        itemName: idx === 0 ? (row.itemName || "Agro Commodity Grade A") : `Premium Grain Lot ${idx + 1}`,
        quantity: bags,
        unit: "Bags",
        rate: Math.round(amt / bags),
        amount: amt,
        tripNo: t.tripNo,
        vehicleNo: t.vehicleNo,
        invoiceNo: inv ? inv.invoiceNo : "Unassigned",
        invoiceDate: inv ? inv.invoiceDate : undefined,
        status: t.status,
      };
    });
  }

  // Single or zero trips: split into 2 items for rich multi-item demonstration
  const pTrip = row.trips && row.trips[0];
  const pInv = row.invoices && row.invoices[0];
  const totalBags = pTrip?.quantityBags || 150;
  const b1 = Math.max(1, Math.floor(totalBags * 0.6));
  const b2 = Math.max(1, totalBags - b1);
  const a1 = Math.round(row.totalAmount * 0.6);
  const a2 = row.totalAmount - a1;

  return [
    {
      itemId: `${row.id}-item-1`,
      itemName: row.itemName || "Primary Commodity Lot",
      quantity: b1,
      unit: "Bags",
      rate: Math.round(a1 / b1),
      amount: a1,
      tripNo: pTrip ? pTrip.tripNo : "Unassigned",
      vehicleNo: pTrip?.vehicleNo,
      invoiceNo: pInv ? pInv.invoiceNo : "Unassigned",
      invoiceDate: pInv?.invoiceDate,
      status: pTrip ? pTrip.status : "Pending",
    },
    {
      itemId: `${row.id}-item-2`,
      itemName: `${row.itemName ? row.itemName.split(" ")[0] : "Agro"} Standard Pulses / Secondary Lot`,
      quantity: b2,
      unit: "Bags",
      rate: Math.round(a2 / b2),
      amount: a2,
      tripNo: pTrip ? pTrip.tripNo : "Unassigned",
      vehicleNo: pTrip?.vehicleNo,
      invoiceNo: pInv ? pInv.invoiceNo : "Unassigned",
      invoiceDate: pInv?.invoiceDate,
      status: pTrip ? pTrip.status : "Pending",
    },
  ];
};
