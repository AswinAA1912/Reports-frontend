/* eslint-disable @typescript-eslint/no-explicit-any */
import dayjs from "dayjs";
import {
    OnlinePurchaseReportService,
    OnlinePurchaseReportItemService,
    OnlinePurchaseReport,
    PurchaseOrderPaymentItem,
} from "./OnlinePurchaseReport.service";
import {
    PurchaseOrderReport,
    PurchaseOrderReportItem,
    PurchaseOrderItem,
    PurchaseOrderTripItem,
    PurchaseOrderTripItemService,
} from "./purchaseOrderReport.service";
import {
    PurchasePaymentItem,
    PurchaseOrderItemDetail,
} from "./purchasePayment.service";
import {
    PurchaseItemOrder,
    ArrivedInvoiceDetail,
    cleanItemName,
} from "./purchaseItemPayment.service";
export { cleanItemName } from "./purchaseItemPayment.service";
import {
    PurchaseDeliveryItem,
} from "./purchaseDelivery.service";
import {
    TodayArrivalItem,
    WholeOrderDetails,
    DeliveryStatus,
    PurchaseReportFilterType,
} from "./purchaseReport.service";

/* =========================================================================
   HELPER UTILITIES
========================================================================= */

const cleanStr = (val: any): string => (val !== undefined && val !== null ? String(val).trim() : "");

const parseNum = (val: any): number => {
    if (val === undefined || val === null || val === "") return 0;
    const n = typeof val === "number" ? val : parseFloat(String(val).replace(/,/g, "").trim());
    return isNaN(n) ? 0 : n;
};

const normalizeName = (name: string): string => {
    return cleanItemName(cleanStr(name))
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
};

export const formatKgQuantity = (qtyInKg: number | string): string => {
    const num = typeof qtyInKg === "number" ? qtyInKg : parseFloat(String(qtyInKg || 0)) || 0;
    if (num <= 0) return "0 KG";
    if (num >= 1000) {
        const tons = (num / 1000).toLocaleString("en-IN", {
            minimumFractionDigits: 0,
            maximumFractionDigits: 2,
        });
        const kgStr = Math.round(num).toLocaleString("en-IN");
        return `${tons} T (${kgStr} KG)`;
    }
    return `${Math.round(num).toLocaleString("en-IN")} KG`;
};

export const formatKgQuantityCompact = (qtyInKg: number | string): string => {
    const num = typeof qtyInKg === "number" ? qtyInKg : parseFloat(String(qtyInKg || 0)) || 0;
    if (num <= 0) return "0 KG";
    if (num >= 1000) {
        const tons = (num / 1000).toLocaleString("en-IN", {
            minimumFractionDigits: 0,
            maximumFractionDigits: 2,
        });
        return `${tons} T (${Math.round(num).toLocaleString("en-IN")} KG)`;
    }
    return `${Math.round(num).toLocaleString("en-IN")} KG`;
};

export const deriveStockGroup = (productName: string): string => {
    const p = (productName || "").toUpperCase();
    if (p.includes("CHANA") || p.includes("MALDA") || p.includes("DESI")) return "2.Desi cp";
    if (p.includes("BEAN") || p.includes("RAJMA") || p.includes("VARI")) return "1.Beans";
    if (p.includes("DHAL") || p.includes("TOOR") || p.includes("URAD") || p.includes("GRAM") || p.includes("MOONG")) {
        return "3.Dhall";
    }
    if (p.includes("OIL") || p.includes("GHEE")) return "4.Oils";
    return "General Items";
};

/* =========================================================================
   RAW DATA FETCHING
========================================================================= */

const extractArray = (res: any): any[] => {
    if (!res) return [];
    if (Array.isArray(res)) return res;
    if (Array.isArray(res.data)) return res.data;
    if (Array.isArray(res.data?.data)) return res.data.data;
    if (Array.isArray(res.data?.data?.data)) return res.data.data.data;
    if (Array.isArray(res.result)) return res.result;
    return [];
};

export interface RawPurchaseDatasets {
    purchaseOrders: PurchaseOrderItem[];
    purchaseInvoices: OnlinePurchaseReport[];
    tripItems: PurchaseOrderTripItem[];
}

export const fetchRawPurchaseDatasets = async (params?: {
    Fromdate?: string;
    Todate?: string;
}): Promise<RawPurchaseDatasets> => {
    const from = params?.Fromdate || dayjs().startOf("month").format("YYYY-MM-DD");
    const to = params?.Todate || dayjs().format("YYYY-MM-DD");

    let poList: PurchaseOrderItem[] = [];
    let invList: OnlinePurchaseReport[] = [];
    let tripList: PurchaseOrderTripItem[] = [];

    const [poItemRes, poRes, invItemRes, invRes, tripRes] = await Promise.allSettled([
        PurchaseOrderReportItem.getPurchaseOrderItem({ Fromdate: from, Todate: to }),
        PurchaseOrderReport.getPurchaseOrder({ Fromdate: from, Todate: to }),
        OnlinePurchaseReportItemService.getReportsitem({ Fromdate: from, Todate: to }),
        OnlinePurchaseReportService.getReports({ Fromdate: from, Todate: to }),
        PurchaseOrderTripItemService.getPurchaseOrderTripItemDetails({ Fromdate: from, Todate: to }),
    ]);

    // 1. Extract Purchase Orders (Prefer Item-level endpoint from PurchaseOrderReportItem)
    if (poItemRes.status === "fulfilled" && poItemRes.value) {
        const raw = extractArray(poItemRes.value.data) || extractArray(poItemRes.value);
        if (raw.length > 0) {
            poList = raw;
        }
    }
    if (poList.length === 0 && poRes.status === "fulfilled" && poRes.value) {
        const raw = extractArray(poRes.value.data) || extractArray(poRes.value);
        if (raw.length > 0) {
            poList = raw;
        }
    }

    // 2. Extract Purchase Invoices (Prefer Item-level endpoint)
    if (invItemRes.status === "fulfilled" && invItemRes.value) {
        const raw = extractArray(invItemRes.value.data) || extractArray(invItemRes.value);
        if (raw.length > 0) {
            invList = raw;
        }
    }
    if (invList.length === 0 && invRes.status === "fulfilled" && invRes.value) {
        const raw = extractArray(invRes.value.data) || extractArray(invRes.value);
        if (raw.length > 0) {
            invList = raw;
        }
    }

    // 3. Extract Purchase Order Trip Items (Trip & Inward Arrival Details with Batch Details)
    if (tripRes.status === "fulfilled" && tripRes.value) {
        const raw = extractArray(tripRes.value.data) || extractArray(tripRes.value);
        if (raw.length > 0) {
            tripList = raw;
        }
    }

    return { purchaseOrders: poList, purchaseInvoices: invList, tripItems: tripList };
};

/* =========================================================================
   DATASET BUILDERS (Pure live API data, no dummy values, empty fields if missing)
========================================================================= */

/**
 * 1. Transform raw datasets into PurchasePaymentItem[] (for PurchasePaymentReport)
 */
export const buildPurchasePaymentDataset = (
    poRows: PurchaseOrderItem[],
    invRows: OnlinePurchaseReport[],
    tripRows: PurchaseOrderTripItem[] = [],
    paymentsMap?: Map<string, PurchaseOrderPaymentItem>
): PurchasePaymentItem[] => {
    // If tripRows are available from PurchaseOrderTripItemService, use rich live dataset
    if (tripRows && tripRows.length > 0) {
        const itemOrders = buildPurchaseItemPaymentDataset(poRows, invRows, tripRows, paymentsMap);
        if (itemOrders && itemOrders.length > 0) {
            return itemOrders.map((ord, idx) => ({
                id: ord.id,
                sNo: idx + 1,
                purOrderNo: ord.purOrderNo,
                transId: ord.transId,
                transIds: ord.transIds,
                supplierName: ord.supplierName,
                orderDate: ord.orderDate,
                dueDate: ord.dueDate,
                items: ord.items,
                orderedTons: ord.orderedTons,
                arrivedTons: ord.arrivedTons,
                pendingTons: ord.pendingTons,
                deliveryPercentage: ord.deliveryPercentage,
                deliveryStatus: ord.deliveryStatus,
                totalPayment: ord.totalValue,
                paidPayment: ord.paidAmount,
                pendingPayment: ord.pendingPayment,
                paymentStatus: ord.paymentStatus,
                paymentNo: ord.paymentNo || (ord.arrivedInvoices && ord.arrivedInvoices.length > 0 ? ord.arrivedInvoices[0].invoiceNo : ""),
                purInvNo: ord.purInvNo,
                inwardJouNo: ord.inwardJouNo || (ord.arrivedInvoices && ord.arrivedInvoices.length > 0 ? ord.arrivedInvoices.length : ""),
                deliveryLocation: ord.deliveryLocation,
                notes: ord.notes || "",
            }));
        }
    }

    // If no data returned from either API, return empty array (NO DUMMY DATA)
    if ((!poRows || poRows.length === 0) && (!invRows || invRows.length === 0)) {
        return [];
    }

    const result: PurchasePaymentItem[] = [];
    let sNoCounter = 1;

    // Group PO items by invoice_no (PO Number)
    const poGroupMap: Record<string, PurchaseOrderItem[]> = {};
    (poRows || []).forEach((row) => {
        const key = cleanStr(
            (row as any).invoice_no ||
            (row as any).Invoice_No ||
            (row as any).voucher_no ||
            (row as any).Voucher_No ||
            (row as any).order_no ||
            (row as any).Order_No ||
            (row as any).Ref_Brokers
        ) || "PO-UNKNOWN";
        if (!poGroupMap[key]) {
            poGroupMap[key] = [];
        }
        poGroupMap[key].push(row);
    });

    // 2. Group raw invoice items by unique invoice_no FIRST so multi-item invoices are matched as a whole
    const invGroupMap: Record<string, OnlinePurchaseReport[]> = {};
    (invRows || []).forEach((inv) => {
        const invNo = cleanStr((inv as any).invoice_no || (inv as any).Invoice_No);
        if (!invNo) return;
        if (!invGroupMap[invNo]) invGroupMap[invNo] = [];
        invGroupMap[invNo].push(inv);
    });
    const uniqueInvNos = Object.keys(invGroupMap);

    // Track which invoices were matched to POs
    const poKeys = Object.keys(poGroupMap);
    const poInvoiceMap: Record<string, OnlinePurchaseReport[]> = {};
    poKeys.forEach((k) => (poInvoiceMap[k] = []));
    const matchedInvoices = new Set<string>();

    // Pass 1: Explicit PO reference in Ref_Brokers, voucher_name, voucher_no, Narration
    uniqueInvNos.forEach((invNo) => {
        const rows = invGroupMap[invNo];
        const ref = rows
            .map((r) =>
                cleanStr(
                    (r as any).Ref_Brokers ||
                    (r as any).ref_brokers ||
                    (r as any).voucher_name ||
                    (r as any).voucher_no ||
                    (r as any).Narration ||
                    (r as any).narration
                )
            )
            .filter(Boolean)
            .join(" ");

        for (const poNo of poKeys) {
            if (ref && poNo && ref.toLowerCase().includes(poNo.toLowerCase())) {
                poInvoiceMap[poNo].push(...rows);
                matchedInvoices.add(invNo);
                break;
            }
        }
    });

    // Pass 2: Best supplier + amount + product + date match for remaining invoices
    uniqueInvNos.forEach((invNo) => {
        if (matchedInvoices.has(invNo)) return;
        const rows = invGroupMap[invNo];
        const first = rows[0];
        const invSupplier = cleanStr(
            (first as any).Retailer_Name ||
            (first as any).Party_Name ||
            (first as any).Vendor_Name
        );
        const invAmount =
            parseNum((first as any).Total_Invoice_value ?? (first as any).Total_Invoice_Value) ||
            rows.reduce((s, r) => s + parseNum((r as any).Amount ?? (r as any).amount), 0);
        const invQty = rows.reduce(
            (s, r) =>
                s +
                parseNum(
                    (r as any).Bill_Qty ??
                    (r as any).Bill_qty ??
                    (r as any).Weight ??
                    (r as any).weight
                ),
            0
        );
        const invProducts = rows
            .map((r) =>
                cleanStr(
                    (r as any).Product_Name ||
                    (r as any).Product_name ||
                    (r as any).Stock_Item_Name ||
                    (r as any).Item_Name
                )
            )
            .filter(Boolean);

        let bestPo: string | null = null;
        let highestScore = 0;

        for (const poNo of poKeys) {
            const poItems = poGroupMap[poNo];
            const firstPo = poItems[0];
            const poSupplier = cleanStr(
                (firstPo as any).Retailer_Name ||
                (firstPo as any).Party_Name ||
                (firstPo as any).Vendor_Name
            );
            if (!invSupplier || !poSupplier) continue;
            if (
                !invSupplier.toLowerCase().includes(poSupplier.toLowerCase()) &&
                !poSupplier.toLowerCase().includes(invSupplier.toLowerCase())
            ) {
                continue;
            }

            let score = 10;
            const poTotalAmt = poItems.reduce((s, it) => s + parseNum((it as any).Amount), 0);
            const poTotalWt = poItems.reduce(
                (s, it) => s + parseNum((it as any).Weight ?? (it as any).weight ?? (it as any).Bill_Qty),
                0
            );

            if (Math.abs(poTotalAmt - invAmount) < 2) score += 100;
            if (Math.abs(poTotalWt - invQty) < 2) score += 80;

            const hasProdMatch = poItems.some((poIt) => {
                const poPName = cleanStr(
                    (poIt as any).Product_Name ||
                    (poIt as any).Stock_Item ||
                    (poIt as any).Item_Name
                ).toLowerCase();
                return invProducts.some(
                    (invP) =>
                        invP.toLowerCase().includes(poPName) ||
                        poPName.includes(invP.toLowerCase())
                );
            });
            if (hasProdMatch) score += 50;

            const poDate = cleanStr((firstPo as any).Ledger_date || (firstPo as any).Ledger_Date || (firstPo as any).Date);
            const invDate = cleanStr((first as any).Ledger_Date || (first as any).Ledger_date);
            if (poDate && invDate) {
                const diffDays = Math.abs((new Date(invDate).getTime() - new Date(poDate).getTime()) / (1000 * 60 * 60 * 24));
                if (diffDays <= 7) score += 20;
                else if (diffDays <= 30) score += 10;
            }

            if (score > highestScore) {
                highestScore = score;
                bestPo = poNo;
            }
        }

        if (bestPo && highestScore >= 50) {
            poInvoiceMap[bestPo].push(...rows);
            matchedInvoices.add(invNo);
        }
    });

    Object.entries(poGroupMap).forEach(([poNo, items]) => {
        const first = items[0];
        const supplierName = cleanStr(
            (first as any).Retailer_Name ||
            (first as any).Retailer_name ||
            (first as any).Reatailer_Name ||
            (first as any).reatailer_name ||
            (first as any).Party_Name ||
            (first as any).Party_name ||
            (first as any).Supplier_Name ||
            (first as any).Vendor_Name
        );
        const orderDate = cleanStr(
            (first as any).Ledger_date ||
            (first as any).Ledger_Date ||
            (first as any).Date ||
            (first as any).date ||
            (first as any).Voucher_Date
        );
        const dueDate = orderDate && dayjs(orderDate).isValid()
            ? dayjs(orderDate).add(30, "day").format("YYYY-MM-DD")
            : "";

        // Get pre-matched invoices for this PO
        const matchingInvoices: OnlinePurchaseReport[] = poInvoiceMap[poNo] || [];

        // Extract Trans_Ids for this PO
        const poTransIds = new Set<string>();
        items.forEach((it: any) => {
            const rawVal = it.Trans_Id ?? it.trans_id ?? it.Id ?? it.id ?? it.OrderId ?? it.order_id;
            if (Array.isArray(rawVal)) {
                rawVal.forEach((v) => {
                    const s = cleanStr(v);
                    if (s) poTransIds.add(s);
                });
            } else {
                const s = cleanStr(rawVal);
                if (s) poTransIds.add(s);
            }
        });
        const transIdList = Array.from(poTransIds);
        const primaryTransId = transIdList[0] || "";

        // Build item details from PurchaseOrderReportItem
        let totalOrderedTons = 0;
        let totalArrivedTons = 0;
        let totalPOValue = 0;

        const itemDetails: PurchaseOrderItemDetail[] = items.map((it, idx) => {
            const rawItemName = cleanStr(
                (it as any).Product_name ||
                (it as any).Product_Name ||
                (it as any).Stock_Item_Name ||
                (it as any).Item_Name ||
                (it as any).item_name ||
                (it as any).Product ||
                (it as any).product
            );
            const itOrderedTons = parseNum(
                (it as any).Weight ??
                (it as any).weight ??
                (it as any).Bill_Qty ??
                (it as any).Bill_qty ??
                (it as any).Billed_Qty ??
                (it as any).billed_qty ??
                (it as any).Qty ??
                (it as any).qty ??
                (it as any).Quantity ??
                (it as any).quantity
            );
            const itRate = parseNum(
                (it as any).Rate ??
                (it as any).rate ??
                (it as any).Price ??
                (it as any).price
            );
            const itAmount = parseNum(
                (it as any).Amount ??
                (it as any).amount ??
                (it as any).Total_Amount ??
                (it as any).total_amount ??
                (it as any).Item_Amount ??
                (it as any).Total_Invoice_Value ??
                (it as any).Total_Invoice_value
            ) || Math.round(itOrderedTons * itRate);

            totalOrderedTons += itOrderedTons;
            totalPOValue += itAmount;

            // Find invoices matching this specific item
            const normItName = normalizeName(rawItemName);
            const itInvoices = matchingInvoices.filter((inv) => {
                const normInvItem = normalizeName(
                    (inv as any).Product_Name ||
                    (inv as any).Product_name ||
                    (inv as any).Stock_Item_Name ||
                    (inv as any).Item_Name ||
                    (inv as any).item_name
                );
                return normInvItem && (normInvItem.includes(normItName) || normItName.includes(normInvItem));
            });

            const itArrivedTons = itInvoices.reduce((sum, inv) => {
                return sum + parseNum(
                    (inv as any).Bill_Qty ??
                    (inv as any).Bill_qty ??
                    (inv as any).Weight ??
                    (inv as any).weight ??
                    (inv as any).Billed_Qty ??
                    (inv as any).Qty ??
                    (inv as any).Quantity
                );
            }, 0);
            totalArrivedTons += itArrivedTons;

            const pendingTons = Math.max(0, itOrderedTons - itArrivedTons);
            const deliveryPercentage = itOrderedTons > 0
                ? Math.min(100, Math.round((itArrivedTons / itOrderedTons) * 100))
                : (itArrivedTons > 0 ? 100 : 0);

            const arrivalStatus: "COMPLETED" | "PARTIALLY_ARRIVED" | "PENDING" =
                deliveryPercentage >= 100
                    ? "COMPLETED"
                    : deliveryPercentage > 0
                        ? "PARTIALLY_ARRIVED"
                        : "PENDING";

            const latestInv = itInvoices[itInvoices.length - 1];
            // Leave batch empty per user requirement (do not show Voucher_Type)
            const lastArrivalDate = latestInv ? cleanStr((latestInv as any).Ledger_Date || (latestInv as any).Ledger_date) : "";

            return {
                itemId: `item-${poNo}-${idx + 1}`,
                itemName: rawItemName,
                stockGroup: cleanStr((it as any).Stock_Group || (it as any).stockGroup) || deriveStockGroup(rawItemName),
                orderedTons: itOrderedTons,
                arrivedTons: itArrivedTons,
                pendingTons,
                deliveryPercentage,
                ratePerTon: itRate,
                itemAmount: itAmount,
                arrivalStatus,
                inwardBatch: undefined,
                lastArrivalDate: lastArrivalDate || undefined,
            };
        });

        // Summary calculations
        const pendingTons = Math.max(0, totalOrderedTons - totalArrivedTons);
        const deliveryPercentage = totalOrderedTons > 0
            ? Math.min(100, Math.round((totalArrivedTons / totalOrderedTons) * 100))
            : (totalArrivedTons > 0 ? 100 : 0);

        const deliveryStatus: "COMPLETED" | "PARTIAL" | "PENDING" =
            deliveryPercentage >= 100 ? "COMPLETED" : deliveryPercentage > 0 ? "PARTIAL" : "PENDING";

        const uniqueMatchingInvNos = Array.from(new Set(matchingInvoices.map((inv) => cleanStr(inv.invoice_no)).filter(Boolean)));
        const purInvNo = uniqueMatchingInvNos.join(", ");
        const primaryInv = matchingInvoices[0];
        const deliveryLocation = primaryInv
            ? (cleanStr(primaryInv.Party_Location) || cleanStr(primaryInv.Party_District))
            : "";

        // Payment details: Match with paymentsMap if provided
        let totalPaidPayment = 0;
        let pendingPayment = totalPOValue;
        let paymentStatus: "PAID" | "PARTIALLY_PAID" | "PENDING" | "OVERDUE" = "PENDING";
        let paymentNo = "";

        if (paymentsMap && paymentsMap.size > 0) {
            let matchedCredit = 0;
            let matchedDebit = 0;
            let matchedBal = 0;
            let hasMatched = false;

            uniqueMatchingInvNos.forEach((invNo) => {
                const k = invNo.toLowerCase();
                const p = paymentsMap.get(k) || paymentsMap.get(k.replace(/^0+/, ""));
                if (p) {
                    hasMatched = true;
                    matchedCredit += parseNum(p.Credit_Amt);
                    matchedDebit += parseNum(p.Debit_Amt);
                    matchedBal += parseNum(p.Bal_Amount);
                    if (!paymentNo) paymentNo = p.invoice_no;
                }
            });

            if (!hasMatched && poNo && poNo !== "PO-UNKNOWN") {
                const k = poNo.toLowerCase();
                const p = paymentsMap.get(k) || paymentsMap.get(k.replace(/^0+/, ""));
                if (p) {
                    hasMatched = true;
                    matchedCredit += parseNum(p.Credit_Amt);
                    matchedDebit += parseNum(p.Debit_Amt);
                    matchedBal += parseNum(p.Bal_Amount);
                    if (!paymentNo) paymentNo = p.invoice_no;
                }
            }

            if (hasMatched) {
                const effectiveTotal = matchedCredit > 0 ? matchedCredit : totalPOValue;
                totalPaidPayment = matchedDebit;
                pendingPayment = matchedBal !== 0 ? Math.abs(matchedBal) : Math.max(0, effectiveTotal - totalPaidPayment);
                paymentStatus = pendingPayment <= 0 && totalPaidPayment > 0
                    ? "PAID"
                    : totalPaidPayment > 0
                        ? "PARTIALLY_PAID"
                        : (dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING");
            } else {
                paymentStatus = dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING";
            }
        } else {
            paymentStatus = dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING";
        }

        const row: PurchasePaymentItem = {
            id: `po-${poNo}`,
            sNo: sNoCounter++,
            purOrderNo: poNo,
            transId: primaryTransId,
            transIds: transIdList,
            supplierName,
            orderDate,
            dueDate,
            orderedTons: totalOrderedTons,
            arrivedTons: totalArrivedTons,
            pendingTons,
            deliveryPercentage,
            deliveryStatus,
            totalPayment: totalPOValue,
            paidPayment: totalPaidPayment,
            pendingPayment,
            paymentStatus,
            paymentNo: "", // Leave empty if no payment data in API
            purInvNo: purInvNo || "",
            inwardJouNo: uniqueMatchingInvNos.length > 0 ? uniqueMatchingInvNos.length : "",
            deliveryLocation,
            notes: "",
            items: itemDetails,
        };
        (row as any).matchingInvoices = matchingInvoices;
        result.push(row);
    });

    // If there are standalone inward invoices not matched to any PO, include them as well (grouped by unique invoice_no)
    uniqueInvNos.forEach((invNo, invIdx) => {
        if (!matchedInvoices.has(invNo)) {
            const rows = invGroupMap[invNo];
            const first = rows[0];
            const refPO = rows.map((r) => cleanStr((r as any).Ref_Brokers || (r as any).ref_brokers || (r as any).voucher_name)).filter(Boolean)[0] || "";
            const invSupplier = cleanStr((first as any).Retailer_Name || (first as any).Party_Name || (first as any).Vendor_Name);
            const invDate = cleanStr((first as any).Ledger_Date || (first as any).Ledger_date);
            const location = cleanStr((first as any).Party_Location || (first as any).Party_District) || "";

            let totalQty = 0;
            let totalAmount = 0;
            const seenKeys = new Set<string>();
            const itemDetails: PurchaseOrderItemDetail[] = [];

            rows.forEach((r, rIdx) => {
                const rawProduct = cleanStr(
                    (r as any).Product_Name ||
                    (r as any).Product_name ||
                    (r as any).Stock_Item_Name ||
                    (r as any).Item_Name
                );
                const q = parseNum(
                    (r as any).Bill_Qty ??
                    (r as any).Bill_qty ??
                    (r as any).Weight ??
                    (r as any).weight ??
                    (r as any).Billed_Qty ??
                    (r as any).Qty ??
                    (r as any).Quantity
                );
                const rate = parseNum((r as any).Rate ?? (r as any).rate ?? (r as any).Price);
                const amt = parseNum(
                    (r as any).Amount ??
                    (r as any).amount ??
                    (r as any).Total_Invoice_value ??
                    (r as any).Total_Invoice_Value
                ) || Math.round(q * rate);

                const itemKey = `${rawProduct}||${q}||${amt}||${(r as any).Trans_Id || ""}`;
                if (!seenKeys.has(itemKey)) {
                    seenKeys.add(itemKey);
                    totalQty += q;
                    totalAmount += amt;

                    itemDetails.push({
                        itemId: `item-inv-${invIdx + 1}-${rIdx + 1}`,
                        itemName: rawProduct,
                        stockGroup: cleanStr((r as any).Stock_Group) || deriveStockGroup(rawProduct),
                        orderedTons: q,
                        arrivedTons: q,
                        pendingTons: 0,
                        deliveryPercentage: 100,
                        ratePerTon: rate,
                        itemAmount: amt,
                        arrivalStatus: "COMPLETED",
                        inwardBatch: undefined,
                        lastArrivalDate: invDate || undefined,
                    });
                }
            });

            const standaloneRow: PurchasePaymentItem = {
                id: `inv-row-${invIdx + 1}`,
                sNo: sNoCounter++,
                purOrderNo: refPO,
                supplierName: invSupplier,
                orderDate: invDate,
                dueDate: "",
                orderedTons: totalQty,
                arrivedTons: totalQty,
                pendingTons: 0,
                deliveryPercentage: 100,
                deliveryStatus: "COMPLETED",
                totalPayment: totalAmount,
                paidPayment: 0,
                pendingPayment: totalAmount,
                paymentStatus: "PENDING",
                paymentNo: "", // Empty
                purInvNo: invNo,
                inwardJouNo: "",
                deliveryLocation: location,
                notes: "",
                items: itemDetails,
            };
            (standaloneRow as any).matchingInvoices = rows;
            result.push(standaloneRow);
        }
    });

    return result;
};

/**
 * 2. Transform raw datasets into PurchaseItemOrder[] (for PurchaseItemPaymentReport)
 * Maps purchase invoices / trip items from purchaseOrderTripItemDetails to PurchaseOrder by Id in PurchaseOrder Trans_Id
 * with full BatchDetails (Batch_No, BatchLocation)
 */
export const buildPurchaseItemPaymentDataset = (
    poRows: PurchaseOrderItem[],
    invRows: OnlinePurchaseReport[],
    tripRows: PurchaseOrderTripItem[] = [],
    paymentsMap?: Map<string, PurchaseOrderPaymentItem>
): PurchaseItemOrder[] => {
    // If no data returned from any API, return empty array (NO DUMMY DATA)
    if ((!poRows || poRows.length === 0) && (!invRows || invRows.length === 0) && (!tripRows || tripRows.length === 0)) {
        return [];
    }

    // PRIMARY PATH: If tripRows from purchaseOrderTripItemDetails are available, map by OrderId === Trans_Id
    if (tripRows && tripRows.length > 0) {
        // Group non-canceled trips by OrderId (string)
        const tripGroupMap = new Map<string, PurchaseOrderTripItem[]>();
        tripRows.forEach((t) => {
            if (cleanStr(t.TripStatus).toLowerCase() === "canceled") return;
            const oId = cleanStr(t.OrderId);
            if (!oId) return;
            if (!tripGroupMap.has(oId)) {
                tripGroupMap.set(oId, []);
            }
            tripGroupMap.get(oId)!.push(t);
        });

        // Group PO items by invoice_no (PO Number)
        const poGroupMap: Record<string, PurchaseOrderItem[]> = {};
        (poRows || []).forEach((row) => {
            const key = cleanStr(
                row.invoice_no ||
                (row as any).Invoice_No ||
                (row as any).voucher_no ||
                (row as any).order_no ||
                (row as any).Order_No
            ) || "PO-UNKNOWN";
            if (!poGroupMap[key]) {
                poGroupMap[key] = [];
            }
            poGroupMap[key].push(row);
        });

        const poKeys = Object.keys(poGroupMap);
        let sNoCounter = 1;

        return poKeys.map((poNo) => {
            const poItems = poGroupMap[poNo];
            const firstPo = poItems[0];

            // Extract all Trans_Ids / Ids / PO Numbers associated with this PO
            const poTransIds = new Set<string>();
            if (poNo && poNo !== "PO-UNKNOWN") {
                poTransIds.add(cleanStr(poNo));
            }
            poItems.forEach((it: any) => {
                const candidates = [
                    it.Trans_Id,
                    it.trans_id,
                    it.Id,
                    it.id,
                    it.OrderId,
                    it.order_id,
                    it.order_no,
                    it.Order_No,
                    it.invoice_no,
                    it.Invoice_No,
                    it.voucher_no,
                ];
                candidates.forEach((rawVal) => {
                    if (Array.isArray(rawVal)) {
                        rawVal.forEach((v) => {
                            const s = cleanStr(v);
                            if (s) poTransIds.add(s);
                        });
                    } else {
                        const s = cleanStr(rawVal);
                        if (s) poTransIds.add(s);
                    }
                });
            });

            // Find all trip items mapped to this PO by OrderId === Trans_Id
            const matchedTrips: PurchaseOrderTripItem[] = [];
            poTransIds.forEach((tid) => {
                if (tripGroupMap.has(tid)) {
                    matchedTrips.push(...tripGroupMap.get(tid)!);
                }
            });

            const supplierName = cleanStr(
                (firstPo as any).Retailer_Name ||
                (firstPo as any).PartyName ||
                (firstPo as any).Party_Name ||
                (firstPo as any).Supplier_Name ||
                (firstPo as any).Vendor_Name
            );
            const orderDate = cleanStr(
                (firstPo as any).Ledger_date ||
                (firstPo as any).Ledger_Date ||
                (firstPo as any).TradeConfirmDate ||
                (firstPo as any).Date
            );
            const dueDate = orderDate && dayjs(orderDate).isValid()
                ? dayjs(orderDate).add(30, "day").format("YYYY-MM-DD")
                : "";

            let totalOrderedTons = 0;
            let totalPOValue = 0;

            // Map PO Items and match item-level batches and arrived qty
            const itemDetails: PurchaseOrderItemDetail[] = poItems.map((it, itIdx) => {
                const rawItemName = cleanStr(
                    (it as any).Product_name ||
                    (it as any).Product_Name ||
                    (it as any).Stock_Item ||
                    (it as any).Stock_Item_Name ||
                    (it as any).Item_Name
                );
                const itProdId = (it as any).Product_Id;
                const itOrderedTons = parseNum(
                    (it as any).Weight ??
                    (it as any).Bill_Qty ??
                    (it as any).Billed_Qty ??
                    (it as any).Qty ??
                    (it as any).Quantity
                );
                const itRate = parseNum((it as any).Rate ?? (it as any).Price);
                const itAmount = parseNum(
                    (it as any).Amount ??
                    (it as any).Total_Invoice_Value ??
                    (it as any).Total_Invoice_value
                ) || Math.round(itOrderedTons * itRate);

                totalOrderedTons += itOrderedTons;
                totalPOValue += itAmount;

                // Match trip items for this specific product (by Product_Id or normalized name)
                const normItName = normalizeName(rawItemName);
                const matchingItemTrips = matchedTrips.filter((t) => {
                    if (itProdId && t.Product_Id && Number(itProdId) === Number(t.Product_Id)) {
                        return true;
                    }
                    const normTripProd = normalizeName(t.Product_Name);
                    return normTripProd && (normTripProd.includes(normItName) || normItName.includes(normTripProd));
                });

                const itArrivedTons = matchingItemTrips.reduce((sum, t) => sum + parseNum(t.QTY ?? t.KGS), 0);
                const pendingTons = Math.max(0, itOrderedTons - itArrivedTons);
                const deliveryPercentage = itOrderedTons > 0
                    ? Math.min(100, Math.round((itArrivedTons / itOrderedTons) * 100))
                    : (itArrivedTons > 0 ? 100 : 0);
                const arrivalStatus: "COMPLETED" | "PARTIALLY_ARRIVED" | "PENDING" =
                    deliveryPercentage >= 100 ? "COMPLETED" : deliveryPercentage > 0 ? "PARTIALLY_ARRIVED" : "PENDING";

                // Unique batches for this item
                const batches = Array.from(new Set(matchingItemTrips.map((t) => cleanStr(t.Batch_No)).filter(Boolean)));
                const latestTrip = matchingItemTrips[matchingItemTrips.length - 1];
                const lastArrivalDate = latestTrip ? cleanStr(latestTrip.Arrival_Date || latestTrip.Trip_Date) : "";

                return {
                    itemId: `item-${poNo}-${itIdx + 1}`,
                    itemName: rawItemName,
                    stockGroup: cleanStr((it as any).Stock_Group || (it as any).stockGroup) || deriveStockGroup(rawItemName),
                    orderedTons: itOrderedTons,
                    arrivedTons: itArrivedTons,
                    pendingTons,
                    deliveryPercentage,
                    ratePerTon: itRate,
                    itemAmount: itAmount,
                    arrivalStatus,
                    inwardBatch: batches.join(", ") || undefined,
                    lastArrivalDate: lastArrivalDate || undefined,
                };
            });

            // Map arrived purchase invoices / trip items with full BatchDetails
            const arrivedInvoices: ArrivedInvoiceDetail[] = matchedTrips.map((t, invIdx) => {
                const invNo = cleanStr(t.TR_INV_ID || t.BillNo || t.Challan_No || `TRIP-${t.Trip_Id}`);
                const invDate = cleanStr(t.Arrival_Date || t.Trip_Date || orderDate);
                const prodName = cleanStr(t.Product_Name || "Item");
                const qty = parseNum(t.QTY ?? t.KGS);
                const unit = cleanStr(t.Units) || "kg";
                const totalVal = parseNum(t.Total_Value ?? t.Taxable_Value);
                const rate = parseNum(t.Gst_Rate ?? (qty > 0 ? totalVal / qty : 0));
                const batchNo = cleanStr(t.Batch_No); // The Batch details!
                const batchLocation = cleanStr(t.BatchLocation);
                const vehicle = cleanStr(t.Vehicle_No);

                return {
                    invoiceId: `trip-${t.Trip_Id}-${t.Product_Id}-${invIdx + 1}`,
                    invoiceNo: invNo,
                    invoiceDate: invDate,
                    itemId: `item-inv-${t.Product_Id || invIdx + 1}`,
                    itemName: prodName,
                    stockGroup: deriveStockGroup(prodName),
                    batchNo: batchNo, // Batch Details
                    batchLocation: batchLocation,
                    arrivedQty: qty,
                    unit: unit,
                    rate: rate,
                    amount: totalVal,
                    paidAmount: 0,
                    pendingPayment: totalVal,
                    paymentStatus: "PENDING",
                    inwardVehicle: vehicle,
                    inwardJouNo: cleanStr(t.TR_INV_ID) || cleanStr(t.BillNo) || (t.T_No ? `MIA/${String(t.T_No).padStart(6, "0")}/26-27` : "-"),
                    TR_INV_ID: cleanStr(t.TR_INV_ID),
                    godown: batchLocation || (t.To_Location ? `Location ${t.To_Location}` : ""),
                    challanNo: cleanStr(t.Challan_No),
                    tripId: t.Trip_Id,
                    narration: cleanStr(t.Narration),
                    tripStatus: cleanStr(t.TripStatus),
                    billType: cleanStr(t.BillType),
                };
            });

            const totalArrivedTons = arrivedInvoices.reduce((s, inv) => s + (inv.arrivedQty || 0), 0);
            const pendingTons = Math.max(0, totalOrderedTons - totalArrivedTons);
            const deliveryPercentage = totalOrderedTons > 0
                ? Math.min(100, Math.round((totalArrivedTons / totalOrderedTons) * 100))
                : (totalArrivedTons > 0 ? 100 : 0);
            const deliveryStatus: "COMPLETED" | "PARTIAL" | "PENDING" =
                deliveryPercentage >= 100 ? "COMPLETED" : deliveryPercentage > 0 ? "PARTIAL" : "PENDING";

            const uniqueInvNos = Array.from(new Set(arrivedInvoices.map((inv) => inv.invoiceNo).filter(Boolean)));
            const purInvNo = uniqueInvNos.join(", ");
            const deliveryLocation = cleanStr(
                matchedTrips[0]?.BatchLocation ||
                (matchedTrips[0]?.To_Location ? `Location ${matchedTrips[0]?.To_Location}` : "") ||
                (firstPo as any).DeliveryLocation ||
                ""
            );

            // Find actual database Trans_Id/OrderId (numeric/non-poNo ID)
            let actualTransId = "";
            for (const it of poItems) {
                const candidates = [
                    (it as any).Trans_Id,
                    (it as any).trans_id,
                    (it as any).Id,
                    (it as any).id,
                    (it as any).OrderId,
                    (it as any).order_id,
                ];
                for (const c of candidates) {
                    if (c === undefined || c === null) continue;
                    let s = "";
                    if (Array.isArray(c)) {
                        const vals = c.map((v) => cleanStr(v)).filter((v) => v && v !== poNo);
                        s = vals[0] || "";
                    } else {
                        s = cleanStr(c);
                        if (s.includes(",")) {
                            const parts = s.split(",").map((p) => cleanStr(p)).filter((p) => p && p !== poNo);
                            s = parts[0] || "";
                        }
                    }
                    if (s && s !== poNo && (!isNaN(Number(s)) || !s.includes("/"))) {
                        actualTransId = s;
                        break;
                    }
                }
                if (actualTransId) break;
            }

            const transIdList = Array.from(poTransIds);
            let primaryTransId = actualTransId || transIdList.find((id) => id !== poNo && !isNaN(Number(id))) || "";
            if (primaryTransId.includes(",")) {
                const parts = primaryTransId.split(",").map((p) => cleanStr(p)).filter((p) => p && p !== poNo);
                primaryTransId = parts[0] || "";
            }

            let paidAmount = 0;
            let pendingPayment = totalPOValue;
            let totalValue = totalPOValue;
            let paymentStatus: "PAID" | "PARTIALLY_PAID" | "PENDING" | "OVERDUE" = "PENDING";
            let paymentNo = "";

            if (paymentsMap && paymentsMap.size > 0) {
                let matchedCredit = 0;
                let matchedDebit = 0;
                let matchedBal = 0;
                let hasMatchedPayment = false;

                // 1. Check all arrived invoices
                arrivedInvoices.forEach((inv) => {
                    const invK = cleanStr(inv.invoiceNo).toLowerCase();
                    const p = invK ? (
                        paymentsMap.get(invK) ||
                        paymentsMap.get(invK.replace(/^0+/, "")) ||
                        paymentsMap.get(invK.replace(/^mia\//, "ps/")) ||
                        paymentsMap.get(invK.replace(/^[a-z]+\//, ""))
                    ) : undefined;
                    if (p) {
                        hasMatchedPayment = true;
                        matchedCredit += parseNum(p.Credit_Amt);
                        matchedDebit += parseNum(p.Debit_Amt);
                        matchedBal += parseNum(p.Bal_Amount);
                        inv.paidAmount = parseNum(p.Debit_Amt);
                        inv.pendingPayment = Math.abs(parseNum(p.Bal_Amount));
                        inv.paymentStatus = inv.pendingPayment <= 0 && inv.paidAmount > 0 ? "PAID" : inv.paidAmount > 0 ? "PARTIALLY_PAID" : "PENDING";
                        (inv as any).paymentNo = p.invoice_no;
                        if (!paymentNo) paymentNo = p.invoice_no;
                    }
                });

                // 2. Check PO No if no invoice match
                if (!hasMatchedPayment && poNo && poNo !== "PO-UNKNOWN") {
                    const poK = poNo.toLowerCase();
                    const p = paymentsMap.get(poK) || paymentsMap.get(poK.replace(/^0+/, ""));
                    if (p) {
                        hasMatchedPayment = true;
                        matchedCredit += parseNum(p.Credit_Amt);
                        matchedDebit += parseNum(p.Debit_Amt);
                        matchedBal += parseNum(p.Bal_Amount);
                        if (!paymentNo) paymentNo = p.invoice_no;
                    }
                }

                // 3. Check transIds if still no match
                if (!hasMatchedPayment && transIdList.length > 0) {
                    for (const tid of transIdList) {
                        const tK = tid.toLowerCase();
                        const p = paymentsMap.get(tK) || paymentsMap.get(tK.replace(/^0+/, ""));
                        if (p) {
                            hasMatchedPayment = true;
                            matchedCredit += parseNum(p.Credit_Amt);
                            matchedDebit += parseNum(p.Debit_Amt);
                            matchedBal += parseNum(p.Bal_Amount);
                            if (!paymentNo) paymentNo = p.invoice_no;
                            break;
                        }
                    }
                }

                if (hasMatchedPayment) {
                    totalValue = matchedCredit > 0 ? matchedCredit : totalPOValue;
                    paidAmount = matchedDebit;
                    pendingPayment = matchedBal !== 0 ? Math.abs(matchedBal) : Math.max(0, totalValue - paidAmount);
                    paymentStatus = pendingPayment <= 0 && paidAmount > 0
                        ? "PAID"
                        : paidAmount > 0
                            ? "PARTIALLY_PAID"
                            : (dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING");

                    // Propagate payment info to arrived invoices if not individually matched
                    if (arrivedInvoices.length > 0) {
                        arrivedInvoices.forEach((inv) => {
                            if (!inv.paidAmount || inv.paidAmount === 0) {
                                const ratio = arrivedInvoices.length === 1 ? 1 : ((inv.amount || 1) / (totalPOValue || 1));
                                inv.paidAmount = Math.round(paidAmount * ratio);
                                inv.pendingPayment = Math.max(0, inv.amount - inv.paidAmount);
                                inv.paymentStatus = inv.pendingPayment <= 0 && inv.paidAmount > 0 ? "PAID" : inv.paidAmount > 0 ? "PARTIALLY_PAID" : "PENDING";
                            }
                            if (!(inv as any).paymentNo && paymentNo) {
                                (inv as any).paymentNo = paymentNo;
                            }
                        });
                    }
                } else {
                    paymentStatus = dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING";
                }
            } else {
                paymentStatus = dueDate && dayjs(dueDate).isBefore(dayjs()) ? "OVERDUE" : "PENDING";
            }

            return {
                id: `po-${poNo}`,
                sNo: sNoCounter++,
                purOrderNo: poNo,
                transId: primaryTransId,
                transIds: transIdList,
                supplierName,
                orderDate,
                dueDate,
                orderedTons: totalOrderedTons,
                arrivedTons: totalArrivedTons,
                pendingTons,
                deliveryPercentage,
                deliveryStatus,
                totalValue,
                paidAmount,
                pendingPayment,
                paymentStatus,
                paymentNo,
                purInvNo,
                inwardJouNo: matchedTrips[0]?.TR_INV_ID || (uniqueInvNos.length > 0 ? uniqueInvNos[0] : "-"),
                deliveryLocation,
                notes: matchedTrips[0]?.Narration || "",
                items: itemDetails,
                arrivedInvoices,
            };
        });
    }

    // FALLBACK: When tripRows are not available, fall back to basePayments logic
    const basePayments = buildPurchasePaymentDataset(poRows, invRows, [], paymentsMap);

    if (!basePayments || basePayments.length === 0) {
        return [];
    }

    return basePayments.map((po, poIdx) => {
        const matchingInvoices: OnlinePurchaseReport[] = (po as any).matchingInvoices || [];
        const invoiceGroupMap: Record<string, OnlinePurchaseReport[]> = {};
        matchingInvoices.forEach((inv) => {
            const invNo = cleanStr((inv as any).invoice_no || (inv as any).Invoice_No);
            if (!invNo) return;
            if (!invoiceGroupMap[invNo]) {
                invoiceGroupMap[invNo] = [];
            }
            invoiceGroupMap[invNo].push(inv);
        });

        const arrivedInvoices: ArrivedInvoiceDetail[] = Object.entries(invoiceGroupMap).map(([invNo, rows], invIdx) => {
            const first = rows[0];
            const invDate = cleanStr((first as any).Ledger_Date || (first as any).Ledger_date) || po.orderDate;

            const itemNames = Array.from(
                new Set(
                    rows.map((r) =>
                        cleanStr(
                            (r as any).Product_Name ||
                            (r as any).Product_name ||
                            (r as any).Stock_Item_Name ||
                            (r as any).Item_Name
                        )
                    ).filter(Boolean)
                )
            ).join(", ");

            const seenItemKeys = new Set<string>();
            let totalInvQty = 0;
            let totalInvAmount = 0;
            let primaryRate = parseNum((first as any).Rate ?? (first as any).rate ?? (first as any).Price);

            rows.forEach((r) => {
                const prod = cleanStr((r as any).Product_Name || (r as any).Item_Name);
                const q = parseNum(
                    (r as any).Bill_Qty ??
                    (r as any).Bill_qty ??
                    (r as any).Weight ??
                    (r as any).weight ??
                    (r as any).Billed_Qty ??
                    (r as any).Qty ??
                    (r as any).Quantity
                );
                const amt = parseNum(
                    (r as any).Amount ??
                    (r as any).amount ??
                    (r as any).Total_Invoice_value ??
                    (r as any).Total_Invoice_Value
                );
                const rRate = parseNum((r as any).Rate ?? (r as any).rate ?? (r as any).Price);
                const itemKey = `${prod}||${q}||${amt}||${(r as any).Trans_Id || ""}`;
                if (!seenItemKeys.has(itemKey)) {
                    seenItemKeys.add(itemKey);
                    totalInvQty += q;
                    totalInvAmount += amt > 0 ? amt : Math.round(q * rRate);
                    if (rRate > 0) primaryRate = rRate;
                }
            });

            return {
                invoiceId: `inv-${po.purOrderNo || "PO"}-${invIdx + 1}`,
                invoiceNo: invNo,
                invoiceDate: invDate,
                itemId: `item-inv-${invIdx + 1}`,
                itemName: itemNames || "Purchase Item",
                stockGroup: cleanStr((first as any).Stock_Group) || deriveStockGroup(itemNames),
                batchNo: "",
                arrivedQty: totalInvQty,
                unit: "Tons",
                rate: primaryRate,
                amount: totalInvAmount,
                paidAmount: 0,
                pendingPayment: totalInvAmount,
                paymentStatus: "PENDING",
                inwardVehicle: cleanStr((first as any).Ref_Brokers || (first as any).ref_brokers) || "",
                godown: cleanStr((first as any).Party_Location || (first as any).Party_District) || "",
            };
        });

        return {
            id: po.id,
            sNo: poIdx + 1,
            purOrderNo: po.purOrderNo,
            transId: po.transId || (po as any).transId || "",
            transIds: po.transIds || (po as any).transIds || [],
            supplierName: po.supplierName,
            orderDate: po.orderDate,
            dueDate: po.dueDate,
            orderedTons: po.orderedTons,
            arrivedTons: po.arrivedTons,
            pendingTons: po.pendingTons,
            deliveryPercentage: po.deliveryPercentage,
            deliveryStatus: po.deliveryStatus,
            totalValue: po.totalPayment,
            paidAmount: po.paidPayment,
            pendingPayment: po.pendingPayment,
            paymentStatus: po.paymentStatus,
            paymentNo: po.paymentNo,
            purInvNo: po.purInvNo,
            inwardJouNo: po.inwardJouNo,
            deliveryLocation: po.deliveryLocation,
            notes: po.notes,
            items: po.items,
            arrivedInvoices,
        };
    });
};

/**
 * 3. Transform raw datasets into PurchaseDeliveryItem[] (for PurchaseDelivery)
 */
export const buildPurchaseDeliveryDataset = (
    poRows: PurchaseOrderItem[],
    invRows: OnlinePurchaseReport[],
    tripRows: PurchaseOrderTripItem[] = [],
    paymentsMap?: Map<string, PurchaseOrderPaymentItem>
): PurchaseDeliveryItem[] => {
    // If tripRows are available, build unified dataset from trips & POs
    if (tripRows && tripRows.length > 0) {
        const itemOrders = buildPurchaseItemPaymentDataset(poRows, invRows, tripRows, paymentsMap);
        if (itemOrders && itemOrders.length > 0) {
            const records: PurchaseDeliveryItem[] = [];
            let sNoCounter = 1;

            itemOrders.forEach((order) => {
                // 1. Arrived Deliveries (COMPLETED)
                if (order.arrivedInvoices && order.arrivedInvoices.length > 0) {
                    order.arrivedInvoices.forEach((inv) => {
                        const qtyText = inv.arrivedQty > 0 ? ` (${formatKgQuantity(inv.arrivedQty)})` : "";
                        const batchText = inv.batchNo ? ` [Batch: ${inv.batchNo}]` : "";
                        const rawName = cleanItemName(inv.itemName);

                        const invKey = cleanStr(inv.invoiceNo).toLowerCase();
                        const pMatch = paymentsMap ? (
                            paymentsMap.get(invKey) ||
                            paymentsMap.get(invKey.replace(/^mia\//, "ps/")) ||
                            paymentsMap.get(invKey.replace(/^[a-z]+\//, ""))
                        ) : undefined;

                        const effectivePaid = inv.paidAmount > 0
                            ? inv.paidAmount
                            : (pMatch && pMatch.Debit_Amt > 0 ? pMatch.Debit_Amt : (order.paidAmount > 0 ? order.paidAmount : ""));

                        const effectivePayNo = (inv as any).paymentNo || (pMatch ? pMatch.invoice_no : "") || order.paymentNo || (inv.paymentStatus === "PAID" || inv.paymentStatus === "PARTIALLY_PAID" ? inv.invoiceNo : "-");

                        records.push({
                            id: `pd-inv-${order.id}-${inv.invoiceId || sNoCounter}`,
                            sNo: sNoCounter++,
                            stockGroup: inv.stockGroup || deriveStockGroup(rawName),
                            inwardBatchWithItemName: `${rawName}${qtyText}${batchText}`,
                            purOrderNo: order.purOrderNo,
                            inwardJouNo: (inv as any).inwardJouNo || cleanStr((inv as any).TR_INV_ID) || order.inwardJouNo || "-",
                            purInvNo: inv.invoiceNo || "-",
                            paymentNo: effectivePayNo,
                            paymentAmt: effectivePaid,
                            status: "COMPLETED",
                            orderDate: inv.invoiceDate || order.orderDate,
                        });
                    });
                }

                // 2. Pending items (NOT COMPLETED)
                if (order.items && order.items.length > 0) {
                    order.items.forEach((it) => {
                        if (it.pendingTons > 0 || it.arrivedTons === 0) {
                            const pendingQty = it.pendingTons > 0 ? it.pendingTons : it.orderedTons;
                            const qtyText = pendingQty > 0 ? ` (${formatKgQuantity(pendingQty)})` : "";
                            const rawName = cleanItemName(it.itemName);

                            records.push({
                                id: `pd-pending-${order.id}-${it.itemId || sNoCounter}`,
                                sNo: sNoCounter++,
                                stockGroup: it.stockGroup || deriveStockGroup(rawName),
                                inwardBatchWithItemName: `${rawName}${qtyText}`,
                                purOrderNo: order.purOrderNo,
                                inwardJouNo: "-",
                                purInvNo: "-",
                                paymentNo: "-",
                                paymentAmt: "",
                                status: "NOT COMPLETED",
                                orderDate: order.orderDate,
                            });
                        }
                    });
                }
            });

            return records;
        }
    }

    // Fallback if no tripRows
    if ((!invRows || invRows.length === 0) && (!poRows || poRows.length === 0)) {
        return [];
    }

    const records: PurchaseDeliveryItem[] = [];
    let sNoCounter = 1;

    // 1. Arrived invoice items (COMPLETED deliveries)
    if (invRows && invRows.length > 0) {
        invRows.forEach((inv) => {
            const rawProduct = cleanStr(
                (inv as any).Product_Name ||
                (inv as any).Product_name ||
                (inv as any).Stock_Item_Name ||
                (inv as any).Item_Name
            );
            const purInvNo = cleanStr((inv as any).invoice_no || (inv as any).Invoice_No);
            const purOrderNo = cleanStr((inv as any).Ref_Brokers || (inv as any).ref_brokers || (inv as any).voucher_name);
            const qty = parseNum(
                (inv as any).Bill_Qty ??
                (inv as any).Bill_qty ??
                (inv as any).Weight ??
                (inv as any).weight ??
                (inv as any).Billed_Qty ??
                (inv as any).Qty
            );
            const qtyText = qty > 0 ? ` (${formatKgQuantity(qty)})` : "";
            const matchedPay = purInvNo ? paymentsMap?.get(purInvNo.toLowerCase()) : undefined;

            records.push({
                id: `pd-inv-${sNoCounter}`,
                sNo: sNoCounter++,
                stockGroup: cleanStr((inv as any).Stock_Group) || deriveStockGroup(rawProduct),
                inwardBatchWithItemName: `${rawProduct}${qtyText}`,
                purOrderNo,
                inwardJouNo: 1,
                purInvNo,
                paymentNo: matchedPay ? (matchedPay.Bal_Amount === 0 ? "PAID" : matchedPay.Debit_Amt > 0 ? "PARTIAL" : "PENDING") : "",
                paymentAmt: matchedPay ? matchedPay.Debit_Amt : "",
                status: "COMPLETED",
                orderDate: cleanStr((inv as any).Ledger_Date || (inv as any).Ledger_date),
            });
        });
    }

    // 2. PO items that are pending delivery (NOT COMPLETED)
    if (poRows && poRows.length > 0) {
        poRows.forEach((po) => {
            const rawProduct = cleanStr(
                (po as any).Product_Name ||
                (po as any).Product_name ||
                (po as any).Stock_Item_Name ||
                (po as any).Item_Name
            );
            const purOrderNo = cleanStr((po as any).invoice_no || (po as any).Invoice_No || (po as any).voucher_no);
            const poQty = parseNum(
                (po as any).Weight ??
                (po as any).weight ??
                (po as any).Bill_Qty ??
                (po as any).Bill_qty ??
                (po as any).Qty
            );
            const qtyText = poQty > 0 ? ` (${formatKgQuantity(poQty)})` : "";

            const hasArrived = (invRows || []).some((inv) => {
                const ref = cleanStr((inv as any).Ref_Brokers || (inv as any).ref_brokers || (inv as any).voucher_name);
                return purOrderNo && ref && ref.toLowerCase().includes(purOrderNo.toLowerCase());
            });

            if (!hasArrived) {
                records.push({
                    id: `pd-po-${sNoCounter}`,
                    sNo: sNoCounter++,
                    stockGroup: cleanStr((po as any).Stock_Group) || deriveStockGroup(rawProduct),
                    inwardBatchWithItemName: `${rawProduct}${qtyText}`,
                    purOrderNo,
                    inwardJouNo: "-",
                    purInvNo: "-",
                    paymentNo: "-",
                    paymentAmt: "",
                    status: "NOT COMPLETED",
                    orderDate: cleanStr((po as any).Ledger_date || (po as any).Ledger_Date),
                });
            }
        });
    }

    return records;
};

/**
 * 4. Transform raw datasets into TodayArrivalItem[] (for PurchaseReport)
 */
export const buildPurchaseReportDataset = (
    poRows: PurchaseOrderItem[],
    invRows: OnlinePurchaseReport[],
    filterType: PurchaseReportFilterType = "today_arrival",
    tripRows: PurchaseOrderTripItem[] = []
): TodayArrivalItem[] => {
    const todayStr = dayjs().format("YYYY-MM-DD");
    const itemPayments = buildPurchaseItemPaymentDataset(poRows, invRows, tripRows);

    if (!itemPayments || itemPayments.length === 0) {
        return [];
    }

    const allArrivals: TodayArrivalItem[] = [];
    let counter = 1;

    itemPayments.forEach((order) => {
        const orderDateFormatted = order.orderDate && dayjs(order.orderDate).isValid()
            ? dayjs(order.orderDate).format("DD/MM/YYYY")
            : order.orderDate;

        const wholeOrder: WholeOrderDetails = {
            orderId: order.purOrderNo,
            purOrderNo: order.purOrderNo,
            orderDate: orderDateFormatted,
            retailerName: order.supplierName,
            city: order.deliveryLocation || "",
            state: "Tamil Nadu",
            totalOrderTonnage: order.orderedTons,
            totalOrderValue: order.totalValue,
            paymentStatus: "Unpaid",
            paidAmount: 0,
            balanceAmount: order.totalValue,
            items: order.items.map((it, idx) => ({
                id: it.itemId,
                itemNo: idx + 1,
                itemName: it.itemName,
                qty: `${it.orderedTons} Tons`,
                tonnage: it.orderedTons,
                rate: it.ratePerTon,
                itemValue: it.itemAmount,
                arrivalDate: it.lastArrivalDate && dayjs(it.lastArrivalDate).isValid()
                    ? dayjs(it.lastArrivalDate).format("DD/MM/YYYY")
                    : "-",
                status: (it.deliveryPercentage >= 100
                    ? (it.lastArrivalDate === todayStr ? "Arrived Today" : "Arrived Earlier")
                    : it.deliveryPercentage > 0
                        ? "In Transit"
                        : "Pending Delivery") as DeliveryStatus,
                tripDetail: undefined, // Leave empty/undefined (NO fake driver or vehicle)
            })),
        };

        // If invoices arrived, create arrival records
        if (order.arrivedInvoices.length > 0) {
            order.arrivedInvoices.forEach((inv) => {
                const invDateFormatted = inv.invoiceDate && dayjs(inv.invoiceDate).isValid()
                    ? dayjs(inv.invoiceDate).format("DD/MM/YYYY")
                    : inv.invoiceDate;

                allArrivals.push({
                    id: inv.invoiceId,
                    sNo: counter++,
                    date: invDateFormatted,
                    orderId: order.purOrderNo,
                    purOrderNo: order.purOrderNo,
                    itemName: inv.itemName,
                    retailerName: order.supplierName,
                    city: inv.godown || order.deliveryLocation || "",
                    state: "Tamil Nadu",
                    qty: `${inv.arrivedQty} Tons`,
                    tonnage: inv.arrivedQty,
                    itemValue: inv.amount,
                    statusTag: inv.invoiceDate === todayStr ? "Arrived Today" : "Arrived Earlier",
                    order: wholeOrder,
                });
            });
        } else {
            // Pending arrival record for primary item
            const firstIt = order.items[0];
            allArrivals.push({
                id: `pending-${order.purOrderNo}`,
                sNo: counter++,
                date: orderDateFormatted,
                orderId: order.purOrderNo,
                purOrderNo: order.purOrderNo,
                itemName: firstIt ? firstIt.itemName : "General Product",
                retailerName: order.supplierName,
                city: order.deliveryLocation || "",
                state: "Tamil Nadu",
                qty: firstIt ? `${firstIt.orderedTons} Tons` : `${order.orderedTons} Tons`,
                tonnage: firstIt ? firstIt.orderedTons : order.orderedTons,
                itemValue: firstIt ? firstIt.itemAmount : order.totalValue,
                statusTag: "Pending Delivery",
                order: wholeOrder,
            });
        }
    });

    // Apply filter tab
    switch (filterType) {
        case "today_arrival":
            return allArrivals.filter((r) => r.statusTag === "Arrived Today" || r.statusTag === "Arrived Earlier");
        case "pending_orders":
            return allArrivals.filter((r) => r.order.items.some((it) => it.status === "Pending Delivery" || it.status === "In Transit"));
        case "pending_payments":
            return allArrivals.filter((r) => r.order.balanceAmount > 0);
        case "completed":
            return allArrivals.filter((r) => r.order.items.every((it) => it.status === "Arrived Today" || it.status === "Arrived Earlier"));
        default:
            return allArrivals;
    }
};
