import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface SalesDeliveryItem {
  Metric: string;
  SalesOrder: number;
  SalesInvoice: number;
  Printed: number;
  Others1: number; // Taken
  Others2: number; // Check
  Dispatch: number;
  Delivery: number;
  ShedSheet: number;
}

export const SalesDeliveryReportService = {
  getSalesDeliveryCumulative: (params?: { Fromdate?: string; Todate?: string }) =>
    axios.get<{ success: boolean; data: SalesDeliveryItem[] }>(
      `${getBaseURL()}api/reports/externalAPI/salesDeliveryCummulative`,
      { params }
    ),
};
