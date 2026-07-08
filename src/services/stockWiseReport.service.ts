import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface stockWiseReport {
    Product_Id: string;
    stock_item_name: string;
    Trans_Date: string;
    Group_Name: string;
    OB_Act_Qty: string;
    Pur_Act_Qty: string;
    Sal_Act_Qty: string;
    Bal_Act_Qty: string;
    OB_Bal_Qty: string;
    Pur_Qty: string;
    Sal_Qty: string;
    Bal_Qty: string;
    Brand: string;
    Group_ST: string;
    Bag: string;
    Stock_Group: string;
    S_Sub_Group_1: string;
    Grade_Item_Group: string;
    Item_Name_Modified: string;
    Date_Added: string;
    POS_Group: string;
    Active: string;
    POS_Item_Name: string;
    Product_Rate: string;
    Stock_Item: string;
    Godown_Name: string;
    [key: string]: string | number | undefined;
}

/* ---------------- GROUP CONFIG ---------------- */

export interface StockGroupConfig {
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

export interface Level2Filter {
    Group_ST: string;
    Group_ST_Name: string;
}


/* ---------------- GROUPING API ---------------- */

export const stockGroupingService = {
    getGroupingConfig: (reportName: string) =>
        axios.get<{ success: boolean; data: StockGroupConfig[] }>(
            `${getBaseURL()}api/sales/salesFilterDropdown`,
            { params: { reportName } }
        ),
};

/* ---------------- DATA APIs ---------------- */

export const itemwisestockreportservice = {
    getItemwiseReports: (params?: { Fromdate?: string; Todate?: string }) =>
        axios.get<{ success: boolean; data: stockWiseReport[] }>(
            `${getBaseURL()}api/reports/storageStock/itemWiseMobile`,
            { params }
        ),
};

export const godownwisestockreportservice = {
    getGodownwiseReports: (params?: {
        Godown_Id: number | string;
        Fromdate?: string;
        Todate?: string;
        filter1?: string;
        filter2?: string;
        filter3?: string;
        groupFilter1?: string;
        groupFilter2?: string;
        groupFilter3?: string;
    }) =>
        axios.get<{ success: boolean; data: stockWiseReport[] }>(
            `${getBaseURL()}api/reports/storageStock/godownitemWise`,
            { params }
        ),
};

/* ---------------- TRANSACTION APIs ---------------- */

export const itemTransactionService = {
    getItemTransactions: (params: {
        fromDate: string;
        toDate: string;
        Product_Id: number;
    }) =>
        axios.get<{ success: boolean; data: stockWiseReport[] }>(
            `${getBaseURL()}api/reports/itemexpenseReport`,
            { params }
        ),
};

export const godownItemTransactionService = {
    getGodownItemTransactions: (params: {
        fromDate: string;
        toDate: string;
        Product_Id: number;
        Godown_Id: number;
    }) =>
        axios.get<{ success: boolean; data: stockWiseReport[] }>(
            `${getBaseURL()}api/reports/godownexpenseReport`,
            { params }
        ),
};

