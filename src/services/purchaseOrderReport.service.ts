import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface PurchaseOrderItem {
    invoice_no: string;
    Ledger_date: string;
    Product_name: string;
    Bill_Qty: number;
    Rate: number;
    Amount: number;
    Retailer_Name: string;
    Total_Invoice_Value: number;
}

export interface PurchaseOrderResponse {
    data: PurchaseOrderItem[]
}

export const PurchaseOrderReport = {
    getPurchaseOrder: (params?: {Fromdate:string; Todate: string}) =>
        axios.get<{success:boolean; data: PurchaseOrderResponse}>(
            // `http://192.168.1.92:9001/api/reports/externalAPI/PurchaseOrderReport`,
              `${getBaseURL()}api/reports/externalAPI/PurchaseOrderReport`,
            {params}
        )
}

export const PurchaseOrderReportItem = {
    getPurchaseOrderItem: (params?: {Fromdate:string; Todate: string}) =>
        axios.get<{success:boolean; data: PurchaseOrderResponse}>(
            //  `http://192.168.1.92:9001/api/reports/externalAPI/PurchaseOrderReportItem`,
              `${getBaseURL()}api/reports/externalAPI/PurchaseOrderReportItem`,
            {params}
        )
}