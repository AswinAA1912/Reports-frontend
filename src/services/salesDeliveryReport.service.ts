import axios from "axios";
import { getBaseURL } from "../config/portalBaseURL";

export interface SalesDeliveryItem {
  ReportDate?: string;
  Metric: string;
  SalesOrder: string;
  SalesInvoice: string;
  Printed: string;
  Taken: string;
  Check: string;
  Dispatch: string;
  Delivery: string;
  ShedSheet: string;
}

export const SalesDeliveryReportService = {
  getSalesDeliveryCumulative: (params?: { Fromdate?: string; Todate?: string }) =>
    axios.get<{ success: boolean; data: SalesDeliveryItem[] }>(
      `${getBaseURL()}api/reports/externalAPI/salesDeliveryCummulative`,
      { params }
    ),
  getSalesDeliveryDaywise: (params?: { Fromdate?: string; Todate?: string }) =>
    axios.get<{ success: boolean; data: SalesDeliveryItem[] }>(
      `${getBaseURL()}api/reports/externalAPI/salesDeliveryDaywise`,
      { params }
    ),
};
