import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface BatchStockItem {
    Product_Id: string;
    Batch_No: string | null;
    OB_Act_Qty: number;
    OB_Bal_Qty: number;
    IN_Act_Qty: number;
    IN_Qty: number;
    OUT_Act_Qty: number;
    OUT_Qty: number;
    Act_Bal_Qty: number;
    Bal_Qty: number;
    stock_item_name: string;
    Godown_Id?: string;
    Godown_Name?: string;
    [key: string]: any;
}

export interface GodownMasterItem {
    Godown_Id: string;
    Godown_Name: string;
    [key: string]: any;
}

export interface GodownMasterResponse {
    data: GodownMasterItem[];
    success?: boolean;
    message?: string;
}

export interface BatchStockResponse {
    data: BatchStockItem[];
    message?: string;
    success?: boolean;
    others?: any;
}

export const BatchStockReportService = {
    getBatchStockReport: (params?: { Fromdate?: string; Todate?: string }) =>
        axios.get<BatchStockResponse>(
            `${getBaseURL()}api/reports/externalAPI/batchStockReport`,
            { params }
        ),
    getBatchStockReportGodown: (params?: { Fromdate?: string; Todate?: string; Godown_Id?: string | number }) =>
        axios.get<BatchStockResponse>(
            `${getBaseURL()}api/reports/externalAPI/batchStockReportGodown`,
            { params }
        ),
    getGodownMaster: () =>
        axios.get<GodownMasterResponse>(
            `${getBaseURL()}api/masters/godown`
        ),
};

