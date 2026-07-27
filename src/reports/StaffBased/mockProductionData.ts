export interface MockProductionDetail {
  id: string;
  Date: string;
  Category: "Inwards" | "Process" | "Outwards";
  VouType: "Pur / Ret" | "Int Trf" | "Adj" | "Atty" | "Wt.Check" | "Cleaning" | "Sales";
  Godown: string;
  StaffName: string;
  ItemName: string;
  JournalNo: string;
  Qty: number; // Tonnage or quantity
  Role: "Atten by" | "Created" | "Printed By" | "Taken" | "Check" | "Check 1" | "Delivery";
}

// Generate a set of realistic mock production details
export const mockProductionDetails: MockProductionDetail[] = [
  // --- July 25, 2026 (Today) ---
  {
    id: "m1",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Aravind Swamy",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-01",
    Qty: 12.5,
    Role: "Atten by"
  },
  {
    id: "m2",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Balaji Vignesh",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-01",
    Qty: 12.5,
    Role: "Created"
  },
  {
    id: "m3",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Chaitanya Roy",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-01",
    Qty: 12.5,
    Role: "Printed By"
  },
  {
    id: "m4",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Devendra Nath",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0725-02",
    Qty: 8.4,
    Role: "Taken"
  },
  {
    id: "m5",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Elangovan P.",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0725-02",
    Qty: 8.4,
    Role: "Check"
  },
  {
    id: "m6",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Faizal Khan",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0725-02",
    Qty: 8.4,
    Role: "Check 1"
  },
  {
    id: "m7",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Gowtham Raj",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0725-02",
    Qty: 8.4,
    Role: "Delivery"
  },
  {
    id: "m8",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Int Trf",
    Godown: "Cold Storage",
    StaffName: "Hari Prasad",
    ItemName: "Premium Ponni Rice - 25kg",
    JournalNo: "JRN-0725-03",
    Qty: 25.0,
    Role: "Atten by"
  },
  {
    id: "m9",
    Date: "2026-07-25",
    Category: "Inwards",
    VouType: "Int Trf",
    Godown: "Cold Storage",
    StaffName: "Irfan Pathan",
    ItemName: "Premium Ponni Rice - 25kg",
    JournalNo: "JRN-0725-03",
    Qty: 25.0,
    Role: "Check"
  },
  {
    id: "m10",
    Date: "2026-07-25",
    Category: "Process",
    VouType: "Adj",
    Godown: "Main Godown",
    StaffName: "Aravind Swamy",
    ItemName: "Brown Rice Organic - 10kg",
    JournalNo: "JRN-0725-04",
    Qty: 5.2,
    Role: "Created"
  },
  {
    id: "m11",
    Date: "2026-07-25",
    Category: "Process",
    VouType: "Atty",
    Godown: "Cold Storage",
    StaffName: "Chaitanya Roy",
    ItemName: "Broken Rice Grade B - 50kg",
    JournalNo: "JRN-0725-05",
    Qty: 18.0,
    Role: "Check 1"
  },
  {
    id: "m12",
    Date: "2026-07-25",
    Category: "Process",
    VouType: "Wt.Check",
    Godown: "Main Godown",
    StaffName: "Devendra Nath",
    ItemName: "Premium Ponni Rice - 25kg",
    JournalNo: "JRN-0725-06",
    Qty: 14.5,
    Role: "Check"
  },
  {
    id: "m13",
    Date: "2026-07-25",
    Category: "Process",
    VouType: "Cleaning",
    Godown: "Cold Storage",
    StaffName: "Faizal Khan",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-07",
    Qty: 11.2,
    Role: "Atten by"
  },
  {
    id: "m14",
    Date: "2026-07-25",
    Category: "Outwards",
    VouType: "Sales",
    Godown: "Main Godown",
    StaffName: "Gowtham Raj",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-08",
    Qty: 30.0,
    Role: "Delivery"
  },
  {
    id: "m15",
    Date: "2026-07-25",
    Category: "Outwards",
    VouType: "Sales",
    Godown: "Main Godown",
    StaffName: "Irfan Pathan",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0725-08",
    Qty: 30.0,
    Role: "Check"
  },
  {
    id: "m16",
    Date: "2026-07-25",
    Category: "Outwards",
    VouType: "Int Trf",
    Godown: "Secondary Warehouse",
    StaffName: "Balaji Vignesh",
    ItemName: "Brown Rice Organic - 10kg",
    JournalNo: "JRN-0725-09",
    Qty: 10.0,
    Role: "Taken"
  },

  // --- July 24, 2026 (Yesterday) ---
  {
    id: "m17",
    Date: "2026-07-24",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Balaji Vignesh",
    ItemName: "Premium Ponni Rice - 25kg",
    JournalNo: "JRN-0724-01",
    Qty: 14.0,
    Role: "Atten by"
  },
  {
    id: "m18",
    Date: "2026-07-24",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Main Godown",
    StaffName: "Chaitanya Roy",
    ItemName: "Premium Ponni Rice - 25kg",
    JournalNo: "JRN-0724-01",
    Qty: 14.0,
    Role: "Created"
  },
  {
    id: "m19",
    Date: "2026-07-24",
    Category: "Process",
    VouType: "Wt.Check",
    Godown: "Main Godown",
    StaffName: "Hari Prasad",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0724-02",
    Qty: 22.1,
    Role: "Check"
  },
  {
    id: "m20",
    Date: "2026-07-24",
    Category: "Outwards",
    VouType: "Sales",
    Godown: "Cold Storage",
    StaffName: "Elangovan P.",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0724-03",
    Qty: 19.5,
    Role: "Delivery"
  },
  {
    id: "m21",
    Date: "2026-07-24",
    Category: "Outwards",
    VouType: "Int Trf",
    Godown: "Main Godown",
    StaffName: "Faizal Khan",
    ItemName: "Broken Rice Grade B - 50kg",
    JournalNo: "JRN-0724-04",
    Qty: 15.0,
    Role: "Taken"
  },

  // --- July 23, 2026 ---
  {
    id: "m22",
    Date: "2026-07-23",
    Category: "Inwards",
    VouType: "Pur / Ret",
    Godown: "Secondary Warehouse",
    StaffName: "Devendra Nath",
    ItemName: "Brown Rice Organic - 10kg",
    JournalNo: "JRN-0723-01",
    Qty: 6.0,
    Role: "Atten by"
  },
  {
    id: "m23",
    Date: "2026-07-23",
    Category: "Process",
    VouType: "Cleaning",
    Godown: "Main Godown",
    StaffName: "Gowtham Raj",
    ItemName: "Rice Sona Masuri - 25kg",
    JournalNo: "JRN-0723-02",
    Qty: 11.5,
    Role: "Created"
  },
  {
    id: "m24",
    Date: "2026-07-23",
    Category: "Outwards",
    VouType: "Sales",
    Godown: "Main Godown",
    StaffName: "Irfan Pathan",
    ItemName: "Basmati Rice Premium - 5kg",
    JournalNo: "JRN-0723-03",
    Qty: 15.2,
    Role: "Delivery"
  }
];
