import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface StockDifferenceItem {
    SD_Id: string;
    Stock_Date: string;
    Godown_Id: string;
    Product_Id: string;
    Item_Name: string;
    Stock_Group: string;
    Grade_Item_Group: string;
    System_Qty: number;
    Actual_Qty: number;
    Difference_Qty: number;
    Remarks: string;
    Created_By: string;
    Created_By_Name: string;
    Created_At: string;
    [key: string]: any;
}

export interface GodownMasterItem {
    Godown_Id: string | number;
    Godown_Name: string;
    [key: string]: any;
}

export interface GodownMasterResponse {
    data: GodownMasterItem[];
    success?: boolean;
    message?: string;
}

export interface StockDifferenceResponse {
    data: StockDifferenceItem[];
    success?: boolean;
    message?: string;
    [key: string]: any;
}

export const StockDifferenceService = {
    getReport: (params: {
        Fromdate: string;
        Todate: string;
        Godown_Id?: string | number;
    }) =>
        axios.get<StockDifferenceResponse>(
            `${getBaseURL()}api/reports/stockDifference/report`,
            { params }
        ),

    getGodowns: () =>
        axios.get<GodownMasterResponse>(
            `${getBaseURL()}api/masters/godown`
        ),
};
