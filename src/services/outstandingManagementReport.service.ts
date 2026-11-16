import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* ================= TYPES ================= */

export interface OutstandingManagementItem {
    Acc_Id: string;
    Account_name: string;
    invoice_date: string;
    invoice_no: string;
    Bal_Amount: number;
    CR_DR: "CR" | "DR" | string;
    Group_Name?: string;
}

/* ================= API SERVICE ================= */

export const OutstandingManagementReportService = {
    getRecievablePayableMgt: (params?: { Todate?: string }) =>
        axios.get<{ success: boolean; data: OutstandingManagementItem[] } | OutstandingManagementItem[]>(
            `${getBaseURL()}api/reports/externalAPI/recievablepayablemgt`,
            { params }
        ),
};
