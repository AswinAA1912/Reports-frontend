import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* ================= TYPES ================= */

export interface InvolvedStaff {
  Id: string;
  Do_Id: string;
  Emp_Id: number;
  Emp_Type_Id: number;
  Emp_Name: string;
  Involved_Emp_Type: string;
}

export interface StockDetail {
  Do_Date: string;
  Delivery_Order_Id: string;
  Item_Id: number;
  S_No: number;
  Godown_Name: string;
  Product_Name: string;
  Product_Image_Name: string;
  UOM: string;
  Brand_Name: string;
  Bill_Qty: number;
  Act_Qty: number;
  Alt_Act_Qty: number;
  unitValue: number;
  itemRate: number;
  billedRate: number;
  quantityDifference: number;
}

export interface ShetSheetItem {
  Do_Id: string;
  Do_Inv_No: string;
  Voucher_Type: string;
  voucherTypeGet: string;
  Do_Date: string;
  Retailer_Id: number;
  Delivery_Status: string;
  Delivery_Status_Id: number;
  retailerNameGet: string;
  Branch_Id: number;
  branchNameGet: string;
  Total_Invoice_value: number;
  Cancel_status: string;
  Created_by: string;
  Created_on: string;
  staffInvolvedStatus: number;
  createdOn: string;
  Narration: string;
  Created_BY_Name: string;
  Pending_Bags: number | null;
  Status: string;
  Remarks: string | null;
  Follow_Up: string | null;
  LR_Uploaded_By: string | null;
  LR_Uploaded_By_Name: string;
  trip_no: string;
  Id: string | null;
  Imageurl: string;
  Image_Name: string;
  imageStatus: string;
  involvedStaffs?: InvolvedStaff[];
  stockDetails?: StockDetail[];
}

export interface ShetSheetReportParams {
  fromDate?: string;
  toDate?: string;
  Fromdate?: string;
  Todate?: string;
  FromDate?: string;
  ToDate?: string;
}

/* ================= SERVICE ================= */

export const ShetSheetReportService = {
  getShetSheetReport: (params?: ShetSheetReportParams) => {
    const from = params?.fromDate || params?.Fromdate || params?.FromDate;
    const to = params?.toDate || params?.Todate || params?.ToDate;

    return axios.get<{ success: boolean; data: ShetSheetItem[] }>(
      `${getBaseURL()}api/reports/shetSheetReport`,
      {
        params: {
          ...params,
          fromDate: from,
          Todate: to,
          Fromdate: from,
          toDate: to,
        },
      }
    );
  },
};
