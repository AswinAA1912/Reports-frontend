import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* =========================
   Types
========================= */

export interface OnlinePurchaseReport {
    Ledger_Date: string;
    Retailer_Name: string;
    voucher_name: string;
    Ref_Brokers: string;
    Party_Location: string;
    Party_District: string;
    invoice_no: string;
    Total_Invoice_value: string;
    Item_Count: String;
    Product_Name: string;
    Bill_Qty: string;
    Rate: string;
    Amount: string;
}

export interface SalesGroupConfig {
    filterType?: string;
    tableId?: number;
    columnName: string;
    tableName?: string;
    aliasName?: string;
    valueColumn?: string;

    FilterLevel?: number;
    Level_Id?: number;

    isGroupFilter?: boolean;
    listTypes?: string;

    displayName?: string;
    groupOrder?: number;

    options?: {
        value: string;
        label: string;
    }[];
}

/* =========================
   Service
========================= */

export const OnlinePurchaseReportService = {
    getReports: (params?: { Fromdate?: string; Todate?: string; invoice_no?: string }) =>
        axios.get<{ success: boolean; data: OnlinePurchaseReport[] }>(
            `${getBaseURL()}api/reports/externalAPI/onlinePurchaseReport`,
            // `http://192.168.1.92:9001/api/reports/externalAPI/onlinePurchaseReport`,
            { params }
        ),
};

export const OnlinePurchaseReportItemService = {
    getReportsitem: (params?: { Fromdate?: string; Todate?: string; invoice_no?: string }) =>
        axios.get<{ success: boolean; data: OnlinePurchaseReport[] }>(
            `${getBaseURL()}api/reports/externalAPI/onlinePurchaseReportItem`,
            //  `http://192.168.1.92:9001/api/reports/externalAPI/onlinePurchaseReportItem`,
            { params }
        ),
};

