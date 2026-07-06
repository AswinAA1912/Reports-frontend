import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

/* ================= TYPES ================= */

export interface CostCenter {
  Cost_Center_Id: number;
  Cost_Center_Name: string;
}


export interface StaffBasedReport {
  ST_Inv_Id: string;
  Branch_Id: number;
  Journal_no: string;
  Stock_Journal_date: string;
  Month_No: number;
  Invoice_Month: string;
  Invoice_Year: number;
  Month_Year: string;
  Stock_Journal_Bill_type: string;
  Stock_Journal_Voucher_type: string;
  Invoice_no: string;
  Narration: string;
  Product_Id: string;
  Product_Name: string;
  Godown_Id: number;
  Batch_No: string;
  Qty: number;
  Act_Qty: number;
  Rate: number;
  Amt: number;
  Unit: string;
  Trans_Id: string;
  Stock_Item: string;
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
  Brokerage: number | null;
  Coolie: number | null;
  Empty_Cost: string;
  Primary_Cost: string;
  Transporter_Name: string;
  Broker_Name: string;
  Load_Man: string;
  Others1: string;
  Others2: string;
  Others3: string;
  Checker: string;
  Delivery_Man: string;
  Others4: string;
  Others5: string;
  Others6: string;
  Driver: string;
  Godown_Name: string;
}

export interface StaffBasedCountReport {
  ST_Inv_Id: string;
  Journal_no: string;
  Stock_Journal_date: string;
  Month_No: string;
  Invoice_Month: string;
  Invoice_Year: string;
  Month_Year: string;
  Bill_type: string;
  Stock_Journal_Voucher_type: string;
  Invoice_no: string;
  Narration: string;
  Created_on: string;
  Created_By: string;
}

/* ================= API SERVICES ================= */

export const costCenterListService = {
  getStaff: () =>
    axios.get<{ success: boolean; data: CostCenter[] }>(
      //  ` http://192.168.1.5:9001/api/reports/externalAPI/costCenter`,
      `${getBaseURL()}api/reports/externalAPI/costCenter`
    ),
};

export const staffBasedReportService = {
  getStaffBasedReport: (params?: { Fromdate?: string; Todate?: string }) =>
    axios.get<{ success: boolean; data: StaffBasedReport[] }>(
      //  ` http://192.168.1.5:9001/api/reports/externalAPI/staffbased`,
      `${getBaseURL()}api/reports/externalAPI/staffbased`,
      { params }
    ),
};

export const StaffBasedCountReportService = {
  getStaffBasedCountReport: (params?: { Fromdate?: string; Todate?: string}) =>
    axios.get<{ success: boolean; data: StaffBasedCountReport[]}>(
      //  ` http://192.168.1.5:9001/api/reports/externalAPI/staffbasedCount`,
      `${getBaseURL()}api/reports/externalAPI/staffbasedCount`,
      { params }
    ),
};











