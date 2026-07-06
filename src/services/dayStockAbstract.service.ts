import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* ----------------  DATASET ---------------- */

export interface StockAbstractData1 {
    sales_voucher_type_id: number,
    voucher_name: string,
    Act_Sal_Qty: number,
    Act_Alt_Sal_Qty: number,
    Bill_Alt_Sal_Qty: number,
    Bill_Sal_Qty: number
}


export interface StockAbstractData2 {
    purchase_voucher_type_id: number,
    voucher_name: string,
    Act_Pur_Qty: number,
    Act_Alt_Pur_Qty: number,
    Bill_Alt_Pur_Qty: number,
    Bill_Pur_Qty: number
}

export interface StockAbstractData3 {
    stock_journal_type_id: number,
    voucher_name: string,
    IN_Qty: number,
    Out_Qty: number,
    ACt_In_Qty: number,
    ACt_Out_Qty: number,
    Group_Name: string
}

export interface StockAbstractData4 {
    godown_id: string;
    godown_name: string;
    parent_godown_name: string;
    IN_Qty: number;
    Out_Qty: number;
    ACt_In_Qty: number;
    ACt_Out_Qty: number;
    OB_Qty: number;
    ACt_OB_Qty: number;
    CL_QTY: number;
    CL_ACt_QTY: number;
}

export interface StockAbstractData5 {
    Sal_Out_Qty: number;
    Sal_ACt_Out_Qty: number;
    OWSG_Out_Qty: number;
    OWSG_ACt_Out_Qty: number;
    TSG_Out_Qty: number;
    TSG_ACt_Out_Qty: number;
    Out_Qty: number;
    ACt_Out_Qty: number;
    Bal_Qty: number;
    Bal_Act_Qty: number;
}

export interface StockAbstractData6 {
    ST_In_Qty: number;
    ST_ACt_In_Qty: number;
    OWSG_In_Qty: number;
    OWSG_ACt_In_Qty: number;
    TSG_In_Qty: number;
    TSG_ACt_In_Qty: number;
    In_Qty: number;
    ACt_In_Qty: number;
    PSG_In_Qty: number;
    PSG_ACt_In_Qty: number;
    Bal_Qty: number;
    Bal_Act_Qty: number;
}
export interface StockAbstractData7 {
    OWSG_Out_Qty: number;
    OWSG_ACt_Out_Qty: number;
    TSG_Out_Qty: number;
    TSG_ACt_Out_Qty: number;
    Out_Qty: number;
    ACt_Out_Qty: number;
    SG_Out_Qty: number;
    SG_ACt_Out_Qty: number;
    Bal_Qty: number;
    Bal_Act_Qty: number;
}

export interface StockAbstractData8 {
    Pur_IN_Qty: number;
    Pur_ACt_IN_Qty: number;
    IWSG_In_Qty: number;
    INSG_ACt_In_Qty: number;
    Bal_Qty: number;
    Bal_Act_Qty: number;
}

export interface StockAbstractData9 {
    Bal_Qty: number;
    Bal_Act_Qty: number;
    Trans_Type: string;
}

/* ---------------- FINAL RESPONSE TYPE ---------------- */

export interface StockAbstractReportResponse {
    Data1: StockAbstractData1[];
    Data2: StockAbstractData2[];
    Data3: StockAbstractData3[];
    Data4: StockAbstractData4[];
    Data5: StockAbstractData5[];
    Data6: StockAbstractData6[];
    Data7: StockAbstractData7[];
    Data8: StockAbstractData8[];
    Data9: StockAbstractData9[];
}

/* ---------------- STOCK ABSTRACT REPORT SERVICE ---------------- */

export const StockAbstractReportService = {
    getStockAbstractReport: async (params?: {
        Predate?: string;
        Fromdate?: string;
        Todate?: string;
    }): Promise<StockAbstractReportResponse> => {
        const finalParams = { ...params };
        if (!finalParams.Predate && finalParams.Fromdate) {
            const d = new Date(finalParams.Fromdate);
            d.setDate(d.getDate() - 1);
            finalParams.Predate = d.toISOString().split("T")[0];
        }
        const res = await axios.get<{
            success: boolean;
            data: any;
        }>(
            // 'http://192.168.1.5:9001/api/reports/externalAPI/dayStockAbstract',
            `${getBaseURL()}api/reports/externalAPI/dayStockAbstract`,
            {
                params: finalParams,
            }
        );

        const data = res.data.data || {};

        return {
            Data1:
                data.Data1 ||
                data["Data1"] ||
                [],

            Data2:
                data.Data2 ||
                data["Data2"] ||
                [],

            Data3:
                data.Data3 ||
                data["Data3"] ||
                [],

            Data4:
                data.Data4 ||
                data["Data4"] ||
                [],

            Data5:
                data.Data5 ||
                data["Data5"] ||
                [],

            Data6:
                data.Data6 ||
                data["Data6"] ||
                [],

            Data7:
                data.Data7 ||
                data["Data7"] ||
                [],

            Data8:
                data.Data8 ||
                data["Data8"] ||
                [],

            Data9:
                data.Data9 ||
                data["Data9"] ||
                [],
        };
    },
};