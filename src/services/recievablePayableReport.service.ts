import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* ================= TYPES ================= */

export interface RecievablePayableItem {
  Acc_Id: string;
  Account_name: string;
  invoice_date: string;
  invoice_no: string;
  Bal_Amount: number;
}

/* ================= API SERVICES ================= */

export const RecievablePayableReportService = {
  getReceivables: (params?: { Fromdate?: string }) =>
    axios.get<{ success: boolean; data: RecievablePayableItem[] }>(
      `${getBaseURL()}api/reports/externalAPI/recievable`,
      { params }
    ),

  getPayables: (params?: { Fromdate?: string }) =>
    axios.get<{ success: boolean; data: RecievablePayableItem[] }>(
      `${getBaseURL()}api/reports/externalAPI/payable`,
      { params }
    ),
};
