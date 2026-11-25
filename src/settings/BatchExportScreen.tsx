import React, { useEffect, useState } from "react";
import {
    Box,
    Paper,
    Typography,
    Button,
    IconButton,
    MenuItem,
    Select,
    FormControl,
    InputLabel,
    Checkbox,
    ListItemText,
    Chip,
    TextField,
    CircularProgress,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Tooltip,
    Alert,
} from "@mui/material";
import DeleteOutlineIcon from "@mui/icons-material/DeleteOutline";
import SaveIcon from "@mui/icons-material/Save";
import PictureAsPdfIcon from "@mui/icons-material/PictureAsPdf";
import TableChartIcon from "@mui/icons-material/TableChart";
import BookmarkBorderIcon from "@mui/icons-material/BookmarkBorder";
import LayersIcon from "@mui/icons-material/Layers";
import TuneIcon from "@mui/icons-material/Tune";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import {
    DndContext,
    closestCenter,
    PointerSensor,
    useSensor,
    useSensors,
    DragEndEvent,
} from "@dnd-kit/core";
import {
    SortableContext,
    useSortable,
    verticalListSortingStrategy,
    arrayMove,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import dayjs from "dayjs";
import { toast } from "react-toastify";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";

import Header from "../Layout/Header";
import { MenuService } from "../services/menus.service";
import { SettingsService } from "../services/reportSettings.services";
import { SalesReportLedgerService, SalesReportItemService } from "../services/SalesReport.service";
import { onlinePaymentReportService } from "../services/expenseReport.service";
import { mapExpenseData } from "../reports/Expenses/ExpensesReport";
import { OutstandingManagementReportService } from "../services/outstandingManagementReport.service";
import { AgingScale, calculateDays, getScaleLabelForDays } from "../reports/Expenses/OutstandingManagementReport";
import { RateMasterAdminService } from "../services/rateMasterAdmin.service";
import { RecievablePayableReportService } from "../services/recievablePayableReport.service";
import { cashboxService } from "../services/cashbox.service";
import { OnlineSalesReportService, OnlineSalesReportItemService, onlineSalesReportLOLService, onlineSalesReportItemLOLService } from "../services/OnlineSalesReport.service";
import { OnlinePurchaseReportService, OnlinePurchaseReportItemService } from "../services/OnlinePurchaseReport.service";
import { SaleOrderReport, SaleOrderReportItem, PendingSaleOrderReport, PendingSaleOrderItemReport } from "../services/saleOrderReport.service";
import { PurchaseOrderReport, PurchaseOrderReportItem } from "../services/purchaseOrderReport.service";
import { bankboxService } from "../services/bankbox.service";
import { DayAbstractReportService } from "../services/dayAbstract.service";
import { StockAbstractReportService } from "../services/dayStockAbstract.service";
import { itemwisestockvaluereportservice, godownwisestockvaluereportservice } from "../services/stockValueReport.service";
import { stockValueReportService } from "../services/stockValueRateMaster.service";
import { ShetSheetReportService, ShetSheetItem } from "../services/ShetSheetReport.service";
import { StockDifferenceService, StockDifferenceItem } from "../services/Stock Difference.service";

const formatINR = (v: number) =>
    new Intl.NumberFormat("en-IN", {
        style: "currency",
        currency: "INR",
    }).format(v);

/* ================= TYPES ================= */

export interface ReportDefinition {
    key: string;
    name: string;
    path: string;
    parentReportName: string;
    modes: string[];
    exportTypes: string[];
    fetchData: (params: { from: string; to: string; mode: string; templateId?: number | "default" }) => Promise<any>;
}

export interface AvailableMenuReportItem {
    key: string;
    name: string;
    path: string;
    parentReportName: string;
    definition: ReportDefinition;
}

export interface SelectedReportConfig {
    id: string; // unique instance id
    reportKey: string;
    reportName: string;
    path?: string;

    availableModes: string[];
    selectedModes: string[];

    availableTemplates: Array<{ id: number; name: string }>;
    selectedTemplateId: number | "default";

    availableExportTypes: string[];
    selectedExportType: string;
}

export interface SavedPresetTemplate {
    id: string;
    name: string;
    createdAt: string;
    reports: Array<{
        reportKey: string;
        selectedModes: string[];
        selectedTemplateId: number | "default";
        selectedExportType: string;
    }>;
}

const STORAGE_PRESETS_KEY = "pukal_batch_export_presets";

/* ================= DND SORTABLE REPORT ROW ================= */

interface SortableReportRowProps {
    cfg: SelectedReportConfig;
    idx: number;
    onUpdateReport: (id: string, updates: Partial<SelectedReportConfig>) => void;
    onRemoveReport: (id: string) => void;
}

const SortableReportRow: React.FC<SortableReportRowProps> = ({
    cfg,
    idx,
    onUpdateReport,
    onRemoveReport,
}) => {
    const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
        id: cfg.id,
    });

    const style: React.CSSProperties = {
        transform: CSS.Transform.toString(transform),
        transition,
        opacity: isDragging ? 0.6 : 1,
        backgroundColor: isDragging ? "#e2e8f0" : undefined,
    };

    return (
        <TableRow
            ref={setNodeRef}
            style={style}
            hover
            sx={{ "&:nth-of-type(even)": { bgcolor: "#f8fafc" } }}
        >
            {/* Drag Handle & S.No */}
            <TableCell align="center">
                <Box display="flex" alignItems="center" justifyContent="center" gap={0.5}>
                    <Box
                        {...attributes}
                        {...listeners}
                        sx={{
                            cursor: "grab",
                            display: "flex",
                            alignItems: "center",
                            color: "#94a3b8",
                            "&:hover": { color: "#1E3A8A" },
                            "&:active": { cursor: "grabbing" },
                        }}
                    >
                        <DragIndicatorIcon fontSize="small" />
                    </Box>
                    <Chip
                        label={`#${idx + 1}`}
                        size="small"
                        sx={{
                            bgcolor: "#e2e8f0",
                            color: "#1e293b",
                            fontWeight: 700,
                            height: 22,
                            fontSize: "0.75rem",
                        }}
                    />
                </Box>
            </TableCell>

            {/* Report Name */}
            <TableCell>
                <Typography fontWeight={700} fontSize="0.88rem" color="#0f172a">
                    {cfg.reportName}
                </Typography>
                <Typography variant="caption" color="text.secondary">
                    {cfg.reportKey}
                </Typography>
            </TableCell>

            {/* Multi-select Modes */}
            <TableCell>
                <FormControl size="small" fullWidth>
                    <Select
                        multiple
                        value={cfg.selectedModes}
                        onChange={(e) => {
                            const val = typeof e.target.value === "string" ? e.target.value.split(",") : e.target.value;
                            onUpdateReport(cfg.id, { selectedModes: val.length ? val : [cfg.availableModes[0]] });
                        }}
                        renderValue={(selected) => (
                            <Box sx={{ display: "flex", flexWrap: "wrap", gap: 0.5 }}>
                                {selected.map((v) => (
                                    <Chip key={v} label={v} size="small" color="primary" sx={{ height: 22, fontSize: "0.72rem" }} />
                                ))}
                            </Box>
                        )}
                        sx={{ fontSize: "0.82rem", bgcolor: "#fff" }}
                    >
                        {cfg.availableModes.map((m) => (
                            <MenuItem key={m} value={m}>
                                <Checkbox checked={cfg.selectedModes.indexOf(m) > -1} size="small" />
                                <ListItemText primary={m} primaryTypographyProps={{ fontSize: "0.82rem" }} />
                            </MenuItem>
                        ))}
                    </Select>
                </FormControl>
            </TableCell>

            {/* Saved Template Selector */}
            <TableCell>
                <FormControl size="small" fullWidth>
                    <Select
                        value={cfg.selectedTemplateId}
                        onChange={(e) => onUpdateReport(cfg.id, { selectedTemplateId: e.target.value as any })}
                        sx={{ fontSize: "0.82rem", bgcolor: "#fff" }}
                    >
                        <MenuItem value="default" sx={{ fontSize: "0.82rem" }}>Default Columns</MenuItem>
                        {cfg.availableTemplates.map((t) => (
                            <MenuItem key={t.id} value={t.id} sx={{ fontSize: "0.82rem" }}>
                                {t.name}
                            </MenuItem>
                        ))}
                    </Select>
                </FormControl>
            </TableCell>

            {/* Export Type / Style */}
            <TableCell>
                <FormControl size="small" fullWidth>
                    <Select
                        value={cfg.selectedExportType}
                        onChange={(e) => onUpdateReport(cfg.id, { selectedExportType: e.target.value })}
                        sx={{ fontSize: "0.82rem", bgcolor: "#fff" }}
                    >
                        {cfg.availableExportTypes.map((et) => (
                            <MenuItem key={et} value={et} sx={{ fontSize: "0.82rem" }}>
                                {et}
                            </MenuItem>
                        ))}
                    </Select>
                </FormControl>
            </TableCell>

            {/* Action */}
            <TableCell align="center">
                <Tooltip title="Remove Report">
                    <IconButton size="small" color="error" onClick={() => onRemoveReport(cfg.id)}>
                        <DeleteOutlineIcon fontSize="small" />
                    </IconButton>
                </Tooltip>
            </TableCell>
        </TableRow>
    );
};

/* ================= CASH BOX HELPER CONSTANTS & FUNCTIONS ================= */

interface CashBoxGroupConfig {
    key: string;
    label: string;
    side: "debit" | "credit";
    masterKey: "Cash" | "Bank" | "LedgerGrp" | "DEX" | "IDEX" | "Others";
}

const CASHBOX_DEBIT_GROUPS: CashBoxGroupConfig[] = [
    { key: "cash_paid", label: "Cash Transfer (Paid)", side: "debit", masterKey: "Cash" },
    { key: "bank_dep", label: "Bank Deposit (Contra)", side: "debit", masterKey: "Bank" },
    { key: "ledger_pay", label: "Ledger Groups (Payment)", side: "debit", masterKey: "LedgerGrp" },
    { key: "dex_deb", label: "Direct Expenses", side: "debit", masterKey: "DEX" },
    { key: "idex_deb", label: "In-Direct Expenses", side: "debit", masterKey: "IDEX" },
    { key: "others_deb", label: "Others", side: "debit", masterKey: "Others" }
];

const CASHBOX_CREDIT_GROUPS: CashBoxGroupConfig[] = [
    { key: "cash_rec", label: "Cash Transfer (Received)", side: "credit", masterKey: "Cash" },
    { key: "bank_rec", label: "Bank Received (Contra)", side: "credit", masterKey: "Bank" },
    { key: "ledger_rec", label: "Ledger Groups (Receipts)", side: "credit", masterKey: "LedgerGrp" },
    { key: "dex_cred", label: "Direct Expenses- Income", side: "credit", masterKey: "DEX" },
    { key: "idex_cred", label: "InDirect Expenses- Income", side: "credit", masterKey: "IDEX" },
    { key: "others_cred", label: "Others", side: "credit", masterKey: "Others" }
];

/* ================= GLOBAL EXPORT FORMATTING HELPERS ================= */

let cachedNirmalaBase64: string | null = null;
/**
 * Loads Nirmala TrueType font (with complete Tamil Unicode glyph range support)
 * and returns it as a base64 string for jsPDF embedding.
 */
async function loadNirmalaFont(): Promise<string | null> {
    if (cachedNirmalaBase64) return cachedNirmalaBase64;
    try {
        const res = await fetch("/Nirmala.ttf");
        if (!res.ok) {
            console.warn("Could not fetch /Nirmala.ttf: HTTP", res.status);
            return null;
        }
        const buffer = await res.arrayBuffer();
        let binary = "";
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        const chunkSize = 8192;
        for (let i = 0; i < len; i += chunkSize) {
            binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)));
        }
        cachedNirmalaBase64 = btoa(binary);
        return cachedNirmalaBase64;
    } catch (e) {
        console.warn("Failed to load Nirmala Tamil font:", e);
        return null;
    }
}

/**
 * Checks if a string looks like an ISO date or if column key indicates date.
 * Formats as "DD-MM-YYYY" or "DD-MM-YYYY HH:mm".
 */
export const formatExportDate = (val: any, colKey?: string): string | null => {
    if (val === undefined || val === null || val === "") return null;
    if (typeof val !== "string" && !(val instanceof Date)) return null;

    const str = String(val).trim();
    // Match ISO formats like YYYY-MM-DD, YYYY-MM-DDTHH:mm:ss, YYYY-MM-DDTHH:mm:ss.sssZ, etc.
    const isoDateRegex = /^\d{4}-\d{2}-\d{2}(?:[T\s]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
    const isDateCol = colKey && /(?:date|created_on|updated_on|do_date|ledger_date|payment_date|inv_date|invoice_date)/i.test(colKey);

    if (isoDateRegex.test(str) || (isDateCol && dayjs(str).isValid())) {
        const d = dayjs(str);
        if (!d.isValid()) return str;

        // If time is present and not 00:00:00, format with time, else date only
        const hasTime = str.includes("T") && !str.includes("T00:00:00");
        return hasTime ? d.format("DD-MM-YYYY HH:mm") : d.format("DD-MM-YYYY");
    }
    return null;
};

/**
 * Cleans corrupted / legacy non-Unicode glyphs and symbols from any text field (item name, staff/transport, retailer, etc.).
 * Fallbacks to cleaner fields if current value contains corrupted legacy glyphs.
 * NOTE: Tamil Unicode characters (range \u0B80-\u0BFF) are fully preserved.
 */
export const cleanExportText = (val: any, colKey?: string, row?: any): string => {
    if (val === undefined || val === null) return "";
    let str = String(val).trim();
    if (!str) return "";

    // If key is related to item/product name and row has a cleaner fallback
    const isItemCol = colKey && /(?:item_name|product_name|pos_item_name|stock_item)/i.test(colKey);
    if (isItemCol && row) {
        // Only consider garbled if it contains specific legacy corruption symbols, NOT Tamil characters
        const hasGarble = /[•*µÆ¿©®™§¶±×÷=≠~]/.test(str);
        if (hasGarble) {
            const cleanerCandidate = row.Product_Name || row.Stock_Item || row.Item_Name || row.Item_Name_Modified;
            if (cleanerCandidate && typeof cleanerCandidate === "string" && cleanerCandidate.trim()) {
                const alt = cleanerCandidate.trim();
                if (!/[•*µÆ¿©®™§¶±×÷≠]/.test(alt)) {
                    str = alt;
                }
            }
        }
    }

    // Clean rogue legacy font encoding characters across ALL columns and reports
    // Characters like •, *, µ, Æ, ¿, ©, ®, ™, §, ¶, ±, ×, ÷, ¢, £, ¥, ¬, ¦, «, », etc.
    // Tamil characters (\u0B80-\u0BFF) are strictly kept intact!
    str = str
        .replace(/[•*µÆ¿©®™§¶±×÷≠¢£¥¬¦«»]/g, " ")
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, "")
        .replace(/\s*=\s*/g, " ")
        .replace(/\s{2,}/g, " ")
        .trim();

    return str;
};

/**
 * Computes dynamic column widths and styles for PDF autoTable to guarantee:
 * 1. The table NEVER exceeds the page width (297mm landscape - 20mm margins = 277mm printable).
 * 2. Columns with extensive content (like Retailer/Party, Product, Staff/Transport, Narration)
 *    are assigned sufficient width and wrap onto 2-3 lines (`overflow: 'linebreak'`).
 * 3. The last column is never clipped or pushed off the page edge.
 */
export const getDynamicTableStyles = (
    headers: string[],
    colConfigs: Array<{ key?: string; label?: string }>,
    printableWidth: number = 277
) => {
    const totalCols = headers.length; // includes S.No
    const pdfFontSize = totalCols > 15 ? 8 : totalCols > 12 ? 9 : totalCols > 9 ? 10 : 12;
    const cellPad = totalCols > 15 ? 0.8 : totalCols > 12 ? 1.0 : 1.5;

    // Relative weight calculation
    const weights: number[] = [0.8]; // S.No index 0

    colConfigs.forEach((c) => {
        const k = String(c.key || c.label || "").toLowerCase();
        if (/(?:s_no|sno|idx)/.test(k)) {
            weights.push(0.8);
        } else if (/(?:date|do_date|created_on|inv_date|invoice_date|ledger_date)/.test(k)) {
            weights.push(1.3);
        } else if (/(?:bags|qty|quantity|rate|billedrate|itemrate|amount|total)/.test(k)) {
            weights.push(1.0);
        } else if (/(?:trip_no|status|uom|follow_up)/.test(k)) {
            weights.push(1.2);
        } else if (/(?:do_inv_no|invoice_no|inv_no)/.test(k)) {
            weights.push(1.8);
        } else if (/(?:vouchertypeget|voucher_type|branch)/.test(k)) {
            weights.push(1.8);
        } else if (/(?:godown|godown_name)/.test(k)) {
            weights.push(1.6);
        } else if (/(?:staff_load|loadman|load_man)/.test(k)) {
            weights.push(2.0);
        } else if (/(?:staff_transport|transport)/.test(k)) {
            weights.push(4.0); // Sufficient width to prevent truncation and allow 2-3 line wrapping
        } else if (/(?:retailer|retailernameget|party|party_name|product_name|pos_item_name|item_name|narration)/.test(k)) {
            weights.push(3.5); // Wide columns that wrap across 2-3 lines
        } else if (/(?:staff_)/.test(k)) {
            weights.push(3.0);
        } else {
            weights.push(2.0);
        }
    });

    const sumWeights = weights.reduce((a, b) => a + b, 0);
    const columnStyles: Record<number, { cellWidth: number; halign?: "left" | "center" | "right" }> = {};

    // Distribute printableWidth precisely across all columns
    weights.forEach((w, idx) => {
        const widthMm = (w / sumWeights) * printableWidth;
        const colCfg = idx > 0 ? colConfigs[idx - 1] : null;
        let halign: "left" | "center" | "right" = "left";

        if (idx === 0) {
            halign = "center";
        } else if (colCfg) {
            const k = String(colCfg.key || colCfg.label || "").toLowerCase();
            if (/(?:date|do_date|created_on|trip_no)/.test(k)) {
                halign = "center";
            } else if (/(?:bags|qty|quantity|amount|rate|value|total|balance|bal)/.test(k)) {
                halign = "right";
            }
        }

        columnStyles[idx] = {
            cellWidth: widthMm,
            halign,
        };
    });

    return { pdfFontSize, cellPad, columnStyles };
};

/**
 * Universal cell formatter for table exports (PDF & Excel)
 */
export const formatExportCellValue = (val: any, colKey: string, row: any, forPdf: boolean = false): any => {
    if (val === undefined || val === null || val === "") return "";

    // 1. Date formatting
    const formattedDate = formatExportDate(val, colKey);
    if (formattedDate !== null) return formattedDate;

    // 2. Numeric formatting (single digit after decimal)
    if (typeof val === "number") {
        return forPdf ? val.toFixed(1) : Number(val.toFixed(1));
    }
    const num = Number(val);
    const isNumCol = /(?:rate|qty|quantity|amount|total|value|balance|bal|cogs|margin|percent|pct|bags)/i.test(colKey);
    if (!isNaN(num) && typeof val !== "boolean" && isNumCol) {
        return forPdf ? num.toFixed(1) : Number(num.toFixed(1));
    }

    // 3. Text cleaning
    return cleanExportText(val, colKey, row);
};

const styleCashBoxWorksheet = (ws: XLSX.WorkSheet) => {
    if (!ws || !ws["!ref"]) return;
    const range = XLSX.utils.decode_range(ws["!ref"]);

    const borderStyle = {
        top: { style: "thin", color: { rgb: "CFCFCF" } },
        bottom: { style: "thin", color: { rgb: "CFCFCF" } },
        left: { style: "thin", color: { rgb: "CFCFCF" } },
        right: { style: "thin", color: { rgb: "CFCFCF" } }
    };

    for (let R = range.s.r; R <= range.e.r; ++R) {
        const particularsCellDebit = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
        const particularsCellCredit = ws[XLSX.utils.encode_cell({ r: R, c: 2 })];

        const particularsValDebit = particularsCellDebit ? String(particularsCellDebit.v || "").trim() : "";
        const particularsValCredit = particularsCellCredit ? String(particularsCellCredit.v || "").trim() : "";

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = ws[cellAddress];
            if (!cell) continue;

            cell.s = cell.s || {};
            cell.s.font = { name: "Arial", sz: 10, color: { rgb: "000000" } };
            cell.s.border = borderStyle;

            if (R === 0) {
                cell.s.font = { name: "Arial", sz: 12, bold: true, color: { rgb: "FFFFFF" } };
                cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                cell.s.alignment = { horizontal: "center", vertical: "center" };
            } else if (R === 2) {
                cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                cell.s.alignment = { horizontal: "center", vertical: "center" };
            } else {
                const isDebitOpeningOrClosing = particularsValDebit === "Opening Balance" || particularsValDebit === "Closing Balance";
                const isCreditOpeningOrClosing = particularsValCredit === "Opening Balance" || particularsValCredit === "Closing Balance";
                const isDebitGroupHeader = CASHBOX_DEBIT_GROUPS.some(g => g.label === particularsValDebit);
                const isCreditGroupHeader = CASHBOX_CREDIT_GROUPS.some(g => g.label === particularsValCredit);

                if (isDebitOpeningOrClosing || isCreditOpeningOrClosing) {
                    cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "B45309" } };
                    cell.s.fill = { fgColor: { rgb: "EEEEEE" } };
                } else if (isDebitGroupHeader || isCreditGroupHeader) {
                    cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "1E3A8A" } };
                    cell.s.fill = { fgColor: { rgb: "E2E8F0" } };
                }

                if (C === 1 || C === 3) {
                    cell.s.alignment = { horizontal: "right" };
                    if (typeof cell.v === "number") {
                        cell.s.numFmt = "₹#,##0.00";
                    }
                }
            }
        }
    }
};

const parseCashBoxData = (reportData: any) => {
    if (!reportData) return null;
    const obList = reportData.OB || [];
    const allCashAccIds = new Set((reportData.Cash || []).map((acc: any) => String(acc.Acc_Id).trim()));
    const allTx = reportData.Data1 || [];

    const cashInvoiceKeys = new Set<string>();
    allTx.forEach((tx: any) => {
        const isCreditCash = tx.Credit_Ac_Id && allCashAccIds.has(String(tx.Credit_Ac_Id).trim());
        const isDebitCash = tx.Debit_Ac_Id && allCashAccIds.has(String(tx.Debit_Ac_Id).trim());
        if (isCreditCash || isDebitCash) {
            const key = tx.invoice_no || tx.Trans_Id || "";
            if (key) cashInvoiceKeys.add(key);
        }
    });

    const cleanTx = allTx.filter((tx: any) => {
        const key = tx.invoice_no || tx.Trans_Id || "";
        return key && cashInvoiceKeys.has(key);
    });

    const matchingOpposingAccIds = new Set<string>();
    const nonCashAccIds = new Set<string>();

    (reportData.LedgerGrp || []).forEach((acc: any) => nonCashAccIds.add(String(acc.Acc_Id)));
    (reportData.DEX || []).forEach((acc: any) => nonCashAccIds.add(String(acc.Acc_Id)));
    (reportData.IDEX || []).forEach((acc: any) => nonCashAccIds.add(String(acc.Acc_Id)));

    cleanTx.forEach((tx: any) => {
        const isDebitCash = tx.Debit_Ac_Id && allCashAccIds.has(String(tx.Debit_Ac_Id));
        const isCreditCash = tx.Credit_Ac_Id && allCashAccIds.has(String(tx.Credit_Ac_Id));
        if (isDebitCash && tx.Credit_Ac_Id) matchingOpposingAccIds.add(String(tx.Credit_Ac_Id));
        if (isCreditCash && tx.Debit_Ac_Id) matchingOpposingAccIds.add(String(tx.Debit_Ac_Id));
    });

    const masterMap: Record<string, any[]> = {
        Cash: reportData.Cash || [],
        Bank: (reportData.Bank || []).filter((acc: any) => matchingOpposingAccIds.has(String(acc.Acc_Id)) && !allCashAccIds.has(String(acc.Acc_Id))),
        LedgerGrp: (reportData.LedgerGrp || []).filter((acc: any) => matchingOpposingAccIds.has(String(acc.Acc_Id)) && !allCashAccIds.has(String(acc.Acc_Id))),
        DEX: (reportData.DEX || []).filter((acc: any) => matchingOpposingAccIds.has(String(acc.Acc_Id)) && !allCashAccIds.has(String(acc.Acc_Id))),
        IDEX: (reportData.IDEX || []).filter((acc: any) => matchingOpposingAccIds.has(String(acc.Acc_Id)) && !allCashAccIds.has(String(acc.Acc_Id))),
        Others: []
    };

    let opening = 0;
    obList.forEach((ob: any) => {
        opening += Number(ob.OB_Amount) || 0;
    });

    const getGroupData = (config: CashBoxGroupConfig) => {
        const masters = masterMap[config.masterKey] || [];
        const masterIds = new Set(masters.map((m: any) => String(m.Acc_Id).trim()));

        const matched = cleanTx.filter((tx: any) => {
            if (config.masterKey === "Cash") {
                if (config.side === "debit") {
                    return tx.Credit_Ac_Id && allCashAccIds.has(String(tx.Credit_Ac_Id).trim());
                } else {
                    return tx.Debit_Ac_Id && allCashAccIds.has(String(tx.Debit_Ac_Id).trim());
                }
            }
            if (config.side === "debit") {
                return tx.Debit_Ac_Id && masterIds.has(String(tx.Debit_Ac_Id).trim());
            } else {
                return tx.Credit_Ac_Id && masterIds.has(String(tx.Credit_Ac_Id).trim());
            }
        });

        const subLedgerMap: Record<string, { accId: string; name: string; amount: number }> = {};
        matched.forEach((tx: any) => {
            const accId = config.side === "debit" ? String(tx.Debit_Ac_Id).trim() : String(tx.Credit_Ac_Id).trim();
            const amount = Number(config.side === "debit" ? tx.Dr_Amount : tx.Cr_Amount) || 0;

            if (!subLedgerMap[accId]) {
                const allLists = [
                    ...(reportData.Cash || []),
                    ...(reportData.Bank || []),
                    ...(reportData.LedgerGrp || []),
                    ...(reportData.DEX || []),
                    ...(reportData.IDEX || [])
                ];
                const found = allLists.find((m: any) => String(m.Acc_Id).trim() === accId);
                subLedgerMap[accId] = {
                    accId,
                    name: found ? found.Account_name : (tx.Particulars || `Account (${accId})`),
                    amount: 0
                };
            }
            subLedgerMap[accId].amount += amount;
        });

        const subLedgers = Object.values(subLedgerMap)
            .filter((s) => s.amount !== 0)
            .sort((a, b) => b.amount - a.amount);
        const total = subLedgers.reduce((sum, s) => sum + s.amount, 0);
        return { subLedgers, total };
    };

    const debitGroups = CASHBOX_DEBIT_GROUPS.map((grp) => ({ ...grp, ...getGroupData(grp) }));
    const creditGroups = CASHBOX_CREDIT_GROUPS.map((grp) => ({ ...grp, ...getGroupData(grp) }));
    const totalDebits = debitGroups.reduce((sum, g) => sum + g.total, 0);
    const totalCredits = creditGroups.reduce((sum, g) => sum + g.total, 0);

    let closing = 0;
    const clsList = reportData.Cls || [];
    if (clsList.length > 0) {
        let sum = 0;
        clsList.forEach((c: any) => {
            sum += Number(c.CL_Amount) || 0;
        });
        closing = opening + sum;
    } else {
        closing = opening + totalCredits - totalDebits;
    }

    return { debitGroups, creditGroups, opening, closing };
};

/* ================= REPORT CATALOG DEFINITIONS ================= */

const CUSTOM_CATALOG_OVERRIDES: Record<string, Partial<ReportDefinition>> = {
    // 1. RATE MASTER ADMIN
    "/ratemasteradmin": {
        modes: ["Default"],
        exportTypes: ["Standard"],
        fetchData: async ({ to, templateId }) => {
            const res = await RateMasterAdminService.getReport({ Todate: to });
            let data1: any[] = [];
            let data2: any[] = [];
            if (res.data?.success && res.data.data) {
                data1 = (res.data.data.Data1 || []).map((row: any) => {
                    const listRate = Number(row.Rate ?? 0);
                    const minRate = Number(row.Min_Rate ?? 0);
                    const margin = listRate - minRate;
                    const marginPercent = listRate !== 0 ? (margin / listRate) * 100 : 0;
                    return {
                        ...row,
                        Margin: margin,
                        Margin_Percentage: Number(marginPercent.toFixed(2))
                    };
                });
                data2 = res.data.data.Data2 || [];
            }

            // Updated filter matches mapped Product IDs from Data2
            const mappedIds = new Set(data2.map((d: any) => String(d.Product_Id)));
            const filteredRows = data1.filter((r: any) => mappedIds.has(String(r.Product_Id)));

            // Check if template columns & grouping apply
            let templateCols: any[] = [];
            let groupingCols: string[] = [];
            if (templateId && templateId !== "default") {
                try {
                    const tRes = await SettingsService.getReportEditData({ reportId: Number(templateId), typeId: 1 });
                    if (tRes.data?.success && tRes.data.data?.columns) {
                        templateCols = tRes.data.data.columns.filter((c: any) => c.enabled).sort((a: any, b: any) => a.order - b.order);
                        groupingCols = templateCols
                            .filter((c: any) => c.groupBy && c.groupBy > 0)
                            .sort((a: any, b: any) => (a.groupBy || 0) - (b.groupBy || 0))
                            .map((c: any) => c.key);
                    }
                } catch (e) {
                    console.warn("Could not load template for RateMasterAdmin:", e);
                }
            }

            return {
                type: "rateMasterAdmin",
                title: "Rate Master Admin",
                rows: filteredRows,
                templateCols,
                groupingCols,
                toDate: to
            };
        }
    },

    // 2. EXPENSE REPORT
    "/expense": {
        modes: ["Direct", "Indirect"],
        exportTypes: ["Management", "Normal"],
        fetchData: async ({ from, to, mode }) => {
            const res = await onlinePaymentReportService.getOnlinePaymentReport({ Fromdate: from, Todate: to });
            const mapped = mapExpenseData(
                res.Summary || [],
                res.DirectExpense || [],
                res.IndirectExpense || []
            );
            return {
                type: "expense",
                title: `${mode} Expenses`,
                mode,
                raw: res,
                mapped,
                from,
                to
            };
        }
    },

    // 3. RECIEVABLE PAYABLE
    "/recievablepayable": {
        modes: ["Recievable", "Payable", "Debtors (Receivable)", "Creditors (Payable)"],
        exportTypes: ["Standard"],
        fetchData: async ({ to, mode, templateId }) => {
            // Check if mode is Outstanding Management style (Recievable / Payable / Abstract / Expanded)
            const isOutstandingStyle = mode === "Recievable" || mode === "Payable" || mode === "Abstract" || mode === "Expanded";
            if (isOutstandingStyle) {
                const res = await OutstandingManagementReportService.getRecievablePayableMgt({ Todate: to });
                const list: any[] = Array.isArray(res)
                    ? res
                    : (res.data && Array.isArray(res.data) ? res.data : (res as any)?.data?.data || []);

                const isRecievable = mode === "Recievable" || mode === "Abstract" || mode.toLowerCase().includes("recievable") || mode.toLowerCase().includes("receivable");
                const targetCrDr = isRecievable ? "CR" : "DR";
                const amountColumnTitle = isRecievable ? "Recievable Amt" : "Payable Amt";
                const reportTitle = isRecievable ? "Recievable Payable - Recievable" : "Recievable Payable - Payable";

                let scales: AgingScale[] = [{ id: "s1", from: 0, to: "" }];
                if (templateId && templateId !== "default") {
                    const savedScalesJson = localStorage.getItem(`template_scales_${templateId}`);
                    if (savedScalesJson) {
                        try {
                            const parsed = JSON.parse(savedScalesJson);
                            if (Array.isArray(parsed) && parsed.length > 0) {
                                scales = parsed;
                            }
                        } catch (e) {
                            // ignore
                        }
                    }
                }

                const filtered = list.filter((r) => String(r.CR_DR || "").toUpperCase() === targetCrDr);
                const accountDaysMap = new Map<string, {
                    baseItem: any;
                    days: number;
                    scaleLabel: string;
                    totalAmount: number;
                    invoiceCount: number;
                }>();

                filtered.forEach((r) => {
                    const days = calculateDays(r.invoice_date, to);
                    const scaleLabel = getScaleLabelForDays(days, scales);
                    const key = `${r.Account_name || "Unknown"}__${days}`;

                    if (!accountDaysMap.has(key)) {
                        accountDaysMap.set(key, {
                            baseItem: r,
                            days,
                            scaleLabel,
                            totalAmount: Number(r.Bal_Amount) || 0,
                            invoiceCount: 1,
                        });
                    } else {
                        const existing = accountDaysMap.get(key)!;
                        existing.totalAmount += Number(r.Bal_Amount) || 0;
                        existing.invoiceCount += 1;
                    }
                });

                const processedData: any[] = [];
                let idx = 0;
                accountDaysMap.forEach((val) => {
                    idx++;
                    const { baseItem, days, scaleLabel, totalAmount, invoiceCount } = val;
                    processedData.push({
                        id: `${baseItem.Acc_Id || ""}-${days}-${idx}`,
                        Acc_Id: baseItem.Acc_Id,
                        Account_name: baseItem.Account_name || "Unknown",
                        invoice_date: baseItem.invoice_date,
                        invoice_no: baseItem.invoice_no,
                        Bal_Amount: totalAmount,
                        CR_DR: baseItem.CR_DR,
                        Group_Name: baseItem.Group_Name || "General Group",
                        days,
                        scaleLabel,
                        invoice_count: invoiceCount,
                    });
                });

                const grandTotalAmount = processedData.reduce((acc, row) => acc + (row.Bal_Amount || 0), 0);
                const map = new Map<string, any[]>();
                scales.forEach((sc) => {
                    const fromNum = sc.from === "" ? 0 : Number(sc.from);
                    const toNum = sc.to === "" ? Infinity : Number(sc.to);
                    const label = sc.to === "" || !isFinite(toNum) ? `Above ${fromNum} Days` : `${fromNum} to ${toNum} Days`;
                    map.set(label, []);
                });

                processedData.forEach((row) => {
                    if (!map.has(row.scaleLabel)) {
                        map.set(row.scaleLabel, []);
                    }
                    map.get(row.scaleLabel)!.push(row);
                });

                const groupedData: { label: string; subtotal: number; items: any[] }[] = [];
                map.forEach((items, label) => {
                    const subtotal = items.reduce((sum, it) => sum + it.Bal_Amount, 0);
                    groupedData.push({ label, subtotal, items });
                });

                const lastGroupLevel = scales.length - 1;
                const flattenedRows: any[] = [];
                groupedData.forEach((group, gIdx) => {
                    const isLastLevel = gIdx === lastGroupLevel;
                    flattenedRows.push({
                        __group: true,
                        __key: group.label,
                        __label: group.label,
                        __subtotal: group.subtotal,
                        __count: group.items.length,
                        __level: gIdx,
                        __isLastLevel: isLastLevel,
                    });

                    if (isLastLevel) {
                        const sortedItems = [...group.items].sort((a, b) => {
                            if (a.days !== b.days) return a.days - b.days;
                            return (a.Account_name || "").localeCompare(b.Account_name || "");
                        });
                        sortedItems.forEach((item) => flattenedRows.push(item));
                    }
                });

                return {
                    type: "outstandingManagement",
                    title: reportTitle,
                    amountColumnTitle,
                    grandTotalAmount,
                    flattenedRows,
                    toDate: to
                };
            }

            const isReceivable = mode.toLowerCase().includes("debtor") || mode.toLowerCase().includes("receivable");
            const apiCall = isReceivable
                ? RecievablePayableReportService.getReceivables
                : RecievablePayableReportService.getPayables;

            const res = await apiCall({ Todate: to });
            const rawList = res.data?.data || [];

            const formatted = rawList.map((r: any) => {
                const debit = r.CR_DR === "DR" ? Number(r.Bal_Amount) || 0 : 0;
                const credit = r.CR_DR === "CR" ? Number(r.Bal_Amount) || 0 : 0;
                return {
                    ...r,
                    Debit_Amount: debit,
                    Credit_Amount: credit
                };
            });

            return {
                type: "recievablePayable",
                title: isReceivable ? "Debtors (Receivable) Report" : "Creditors (Payable) Report",
                isReceivable,
                rows: formatted,
                toDate: to
            };
        }
    },

    // 4. CASH BOX
    "/cashbox": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            if (mode === "Abstract") {
                const res = await cashboxService.getCashBoxReport({ Fromdate: from, Todate: to });
                const parsed = parseCashBoxData(res);
                return {
                    type: "cashboxAbstract",
                    title: "Cash Box Transaction",
                    parsed,
                    from,
                    to
                };
            } else {
                const res = await cashboxService.getCashBoxDetailedReport({ Fromdate: from, Todate: to });
                return {
                    type: "cashboxExpanded",
                    title: "Cash Box Detailed Transaction",
                    detailedData: res,
                    from,
                    to
                };
            }
        }
    },

    // 5. ONLINE SALES REPORT
    "/salesreport": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode, templateId }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? OnlineSalesReportItemService.getReportsitem : OnlineSalesReportService.getReports;
            const res = await service({ Fromdate: from, Todate: to });
            const rows = res.data?.data || [];

            let templateCols: any[] = [];
            if (templateId && templateId !== "default") {
                try {
                    const tRes = await SettingsService.getReportEditData({ reportId: Number(templateId), typeId: isExp ? 2 : 1 });
                    if (tRes.data?.success && tRes.data.data?.columns) {
                        templateCols = tRes.data.data.columns.filter((c: any) => c.enabled).sort((a: any, b: any) => a.order - b.order);
                    }
                } catch (e) {
                    console.warn(e);
                }
            }

            return {
                type: "table",
                title: `Online Sales Report (${mode})`,
                rows,
                templateCols
            };
        }
    },

    // 6. ONLINE SALES REPORT LOL
    "/salesreportlol": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode, templateId }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? onlineSalesReportItemLOLService.getReportsItemLOL : onlineSalesReportLOLService.getReportsLOL;
            const res = await service({ Fromdate: from, Todate: to });
            const rawRows = res.data?.data || [];

            let rows = rawRows.map((row: any) => {
                const r = Number(row.Rate) || 0;
                const rm = Number(row.List_Rate ?? row.Rate_Master_Rate ?? row.Rate_Master ?? row.rate_master_rate ?? row.RateMasterRate ?? 0);
                const diff = Number((rm - r).toFixed(2));
                const diffPercent = rm !== 0 ? Number((((rm - r) / rm) * 100).toFixed(2)) : 0;
                return {
                    ...row,
                    Diff: diff,
                    Diff_Percentage: `${diffPercent.toFixed(2)}%`,
                };
            });

            let templateCols: any[] = [];
            if (templateId && templateId !== "default") {
                try {
                    const tRes = await SettingsService.getReportEditData({ reportId: Number(templateId), typeId: isExp ? 2 : 1 });
                    if (tRes.data?.success && tRes.data.data?.columns) {
                        templateCols = tRes.data.data.columns.filter((c: any) => c.enabled).sort((a: any, b: any) => a.order - b.order);
                    }
                } catch (e) {
                    console.warn(e);
                }
            }

            return {
                type: "table",
                title: `Online Sales Report LOL (${mode})`,
                rows,
                templateCols
            };
        }
    },

    // 7. ONLINE PUR REPORT
    "/purchasereport": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode, templateId }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? OnlinePurchaseReportItemService.getReportsitem : OnlinePurchaseReportService.getReports;
            const res = await service({ Fromdate: from, Todate: to });
            const rows = res.data?.data || [];

            let templateCols: any[] = [];
            if (templateId && templateId !== "default") {
                try {
                    const tRes = await SettingsService.getReportEditData({ reportId: Number(templateId), typeId: isExp ? 2 : 1 });
                    if (tRes.data?.success && tRes.data.data?.columns) {
                        templateCols = tRes.data.data.columns.filter((c: any) => c.enabled).sort((a: any, b: any) => a.order - b.order);
                    }
                } catch (e) {
                    console.warn(e);
                }
            }

            return {
                type: "table",
                title: `Online Purchase Report (${mode})`,
                rows,
                templateCols
            };
        }
    },

    // 8. SALES ANALYTICS (SALES REPORT LR)
    "/salesreportlr": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? SalesReportItemService.getReports : SalesReportLedgerService.getReports;
            const res = await service({ Fromdate: from, Todate: to });
            return { type: "table", title: `Sales Analytics (${mode})`, rows: res.data?.data || [] };
        }
    },

    // 9. OUTSTANDING MANAGEMENT
    "/outstandingmanagement": {
        modes: ["Recievable", "Payable"],
        exportTypes: ["Standard"],
        fetchData: async ({ to, mode, templateId }) => {
            const res = await OutstandingManagementReportService.getRecievablePayableMgt({ Todate: to });
            const list: any[] = Array.isArray(res)
                ? res
                : (res.data && Array.isArray(res.data) ? res.data : (res as any)?.data?.data || []);

            const isRecievable = mode === "Recievable" || mode === "Abstract" || mode.toLowerCase().includes("recievable") || mode.toLowerCase().includes("receivable");
            const targetCrDr = isRecievable ? "CR" : "DR";
            const amountColumnTitle = isRecievable ? "Recievable Amt" : "Payable Amt";
            const reportTitle = isRecievable ? "Outstanding Management - Receivable" : "Outstanding Management - Payable";

            // Load scales (default 0 to MAX, or from template)
            let scales: AgingScale[] = [{ id: "s1", from: 0, to: "" }];
            if (templateId && templateId !== "default") {
                const savedScalesJson = localStorage.getItem(`template_scales_${templateId}`);
                if (savedScalesJson) {
                    try {
                        const parsed = JSON.parse(savedScalesJson);
                        if (Array.isArray(parsed) && parsed.length > 0) {
                            scales = parsed;
                        }
                    } catch (e) {
                        // ignore
                    }
                }
                if (scales.length === 1 && scales[0].to === "") {
                    try {
                        const tRes = await SettingsService.getReportEditData({ reportId: Number(templateId), typeId: 1 });
                        const templateCols = tRes.data?.data?.columns || [];
                        const scaleCols = templateCols.filter((c: any) => c.key && c.key.startsWith("SCALE_"));
                        if (scaleCols.length > 0) {
                            scales = scaleCols.map((c: any, sIdx: number) => {
                                const parts = c.key.replace("SCALE_", "").split("_");
                                return {
                                    id: `s-${sIdx + 1}`,
                                    from: parts[0] === "" ? 0 : Number(parts[0]),
                                    to: parts[1] === "" || parts[1] === "MAX" ? "" : Number(parts[1])
                                };
                            });
                        }
                    } catch (e) {
                        // ignore
                    }
                }
            }

            // Filter by targetCrDr
            const filtered = list.filter((r) => String(r.CR_DR || "").toUpperCase() === targetCrDr);

            // Group by Account_name + days to consolidate duplicate invoices
            const accountDaysMap = new Map<string, {
                baseItem: any;
                days: number;
                scaleLabel: string;
                totalAmount: number;
                invoiceCount: number;
            }>();

            filtered.forEach((r) => {
                const days = calculateDays(r.invoice_date, to);
                const scaleLabel = getScaleLabelForDays(days, scales);
                const key = `${r.Account_name || "Unknown"}__${days}`;

                if (!accountDaysMap.has(key)) {
                    accountDaysMap.set(key, {
                        baseItem: r,
                        days,
                        scaleLabel,
                        totalAmount: Number(r.Bal_Amount) || 0,
                        invoiceCount: 1,
                    });
                } else {
                    const existing = accountDaysMap.get(key)!;
                    existing.totalAmount += Number(r.Bal_Amount) || 0;
                    existing.invoiceCount += 1;
                }
            });

            const processedData: any[] = [];
            let idx = 0;
            accountDaysMap.forEach((val) => {
                idx++;
                const { baseItem, days, scaleLabel, totalAmount, invoiceCount } = val;
                processedData.push({
                    id: `${baseItem.Acc_Id || ""}-${days}-${idx}`,
                    Acc_Id: baseItem.Acc_Id,
                    Account_name: baseItem.Account_name || "Unknown",
                    invoice_date: baseItem.invoice_date,
                    invoice_no: baseItem.invoice_no,
                    Bal_Amount: totalAmount,
                    CR_DR: baseItem.CR_DR,
                    Group_Name: baseItem.Group_Name || "General Group",
                    days,
                    scaleLabel,
                    invoice_count: invoiceCount,
                });
            });

            const grandTotalAmount = processedData.reduce((acc, row) => acc + (row.Bal_Amount || 0), 0);

            // Group by scale intervals
            const map = new Map<string, any[]>();
            scales.forEach((sc) => {
                const fromNum = sc.from === "" ? 0 : Number(sc.from);
                const toNum = sc.to === "" ? Infinity : Number(sc.to);
                const label = sc.to === "" || !isFinite(toNum) ? `Above ${fromNum} Days` : `${fromNum} to ${toNum} Days`;
                map.set(label, []);
            });

            processedData.forEach((row) => {
                if (!map.has(row.scaleLabel)) {
                    map.set(row.scaleLabel, []);
                }
                map.get(row.scaleLabel)!.push(row);
            });

            const groupedData: { label: string; subtotal: number; items: any[] }[] = [];
            map.forEach((items, label) => {
                const subtotal = items.reduce((sum, it) => sum + it.Bal_Amount, 0);
                groupedData.push({
                    label,
                    subtotal,
                    items,
                });
            });

            // Flatten rows matching OutstandingManagementReport.tsx:
            // Last group level shows its account names sorted ascending by days and then account name
            const lastGroupLevel = scales.length - 1;
            const flattenedRows: any[] = [];
            groupedData.forEach((group, gIdx) => {
                const isLastLevel = gIdx === lastGroupLevel;
                flattenedRows.push({
                    __group: true,
                    __key: group.label,
                    __label: group.label,
                    __subtotal: group.subtotal,
                    __count: group.items.length,
                    __level: gIdx,
                    __isLastLevel: isLastLevel,
                });

                if (isLastLevel) {
                    const sortedItems = [...group.items].sort((a, b) => {
                        if (a.days !== b.days) return a.days - b.days;
                        return (a.Account_name || "").localeCompare(b.Account_name || "");
                    });
                    sortedItems.forEach((item) => flattenedRows.push(item));
                }
            });

            return {
                type: "outstandingManagement",
                title: reportTitle,
                amountColumnTitle,
                grandTotalAmount,
                flattenedRows,
                toDate: to
            };
        }
    },

    // 10. SALE ORDER REPORT
    "/saleorder": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? SaleOrderReportItem.getSaleOrderItem : SaleOrderReport.getSaleOrder;
            const res = await service({ Fromdate: from, Todate: to });
            const rows = res.data?.data?.data || res.data?.data || [];
            return { type: "table", title: `Sale Order Report (${mode})`, rows: Array.isArray(rows) ? rows : [] };
        }
    },

    // 11. PENDING SALE ORDER
    "/pendingsaleorder": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ to, mode }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? PendingSaleOrderItemReport.getPendingSaleOrderItem : PendingSaleOrderReport.getPendingSaleOrder;
            const res = await service({ Todate: to });
            const rows = res.data?.data?.data || res.data?.data || [];
            return { type: "table", title: `Pending Sale Order (${mode})`, rows: Array.isArray(rows) ? rows : [] };
        }
    },

    // 12. PURCHASE ORDER
    "/purchaseorder": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            const isExp = mode === "Expanded";
            const service = isExp ? PurchaseOrderReportItem.getPurchaseOrderItem : PurchaseOrderReport.getPurchaseOrder;
            const res = await service({ Fromdate: from, Todate: to });
            const rows = res.data?.data?.data || res.data?.data || [];
            return { type: "table", title: `Purchase Order Report (${mode})`, rows: Array.isArray(rows) ? rows : [] };
        }
    },

    // 13. BANK ABSTRACT
    "/bankabstract": {
        modes: ["Abstract", "Expanded"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            if (mode === "Abstract") {
                const res = await bankboxService.getBankBoxReport({ Fromdate: from, Todate: to });
                const parsed = parseCashBoxData(res);
                return { type: "cashboxAbstract", title: "Bank Box Transaction", parsed, from, to };
            } else {
                const res = await bankboxService.getBankBoxDetailedReport({ Fromdate: from, Todate: to });
                return { type: "cashboxExpanded", title: "Bank Box Detailed Transaction", detailedData: res, from, to };
            }
        }
    },

    // 14. DAY ABSTRACT
    "/dayabstract": {
        modes: ["Abstract"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to }) => {
            const res = await DayAbstractReportService.getDayAbstractReport({ Fromdate: from, Todate: to });
            const list = res.Data3 || [];
            return { type: "table", title: "Day Abstract Report", rows: list };
        }
    },

    // 15. STOCK ABSTRACT
    "/stockabstract": {
        modes: ["Abstract"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to }) => {
            const res = await StockAbstractReportService.getStockAbstractReport({ Fromdate: from, Todate: to });
            const list = res.Data4 || [];
            return { type: "table", title: "Stock Abstract Report", rows: list };
        }
    },

    // 16. STOCK VALUE
    "/stockvalue": {
        modes: ["Itemwise", "Godownwise"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, mode }) => {
            const isGodown = mode === "Godownwise";
            const service = isGodown ? godownwisestockvaluereportservice.getGodownwiseReports : itemwisestockvaluereportservice.getItemwiseReports;
            const res = await service({ Fromdate: from, Todate: to });
            const rows = res.data?.data || [];
            return { type: "table", title: `Stock Value (${mode})`, rows };
        }
    },

    // 17. RATE MASTER
    "/ratemaster": {
        modes: ["Default"],
        exportTypes: ["Standard"],
        fetchData: async ({ from }) => {
            const res = await stockValueReportService.getPosRates({ FromDate: from });
            const rows = res.data?.data?.posRateMaster || [];
            return { type: "table", title: "Rate Master", rows };
        }
    },

    // 18. SHET SHEET REPORT
    "/shetSheet": {
        modes: ["Default"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, templateId }) => {
            let isAllDetails = false;
            if (templateId && templateId !== "default") {
                const savedAll = localStorage.getItem(`template_allDetails_${templateId}`);
                if (savedAll !== null) {
                    isAllDetails = savedAll === "true";
                }
            }

            const res = await ShetSheetReportService.getShetSheetReport({
                Fromdate: from,
                Todate: to,
                allDetails: isAllDetails,
            });
            const rawItems: ShetSheetItem[] = res.data?.data || (Array.isArray(res.data) ? (res.data as any) : []);

            let templateCols: any[] = [];
            if (templateId && templateId !== "default") {
                try {
                    const tRes = await SettingsService.getReportEditData({
                        reportId: Number(templateId),
                        typeId: 1,
                    });
                    if (tRes.data?.success && tRes.data.data?.columns) {
                        templateCols = tRes.data.data.columns
                            .filter((c: any) => c.enabled)
                            .sort((a: any, b: any) => (a.order ?? 0) - (b.order ?? 0));
                    }
                } catch (e) {
                    console.warn("Could not load template for ShetSheet:", e);
                }
            }

            // Fallback default columns if template not selected or empty
            const activeCols = templateCols.length > 0 ? templateCols : [
                { key: "Do_Date", label: "Date" },
                { key: "retailerNameGet", label: "Party Name" },
                { key: "Bags", label: "Bags" },
                { key: "Godown_Name", label: "Godown" },
                { key: "Follow_Up", label: "Follow Up" },
                { key: "Status", label: "Status" },
            ];

            const isDetailedStock = activeCols.some((c: any) =>
                ["Product_Name", "Bill_Qty", "Act_Qty", "billedRate", "itemRate", "UOM"].includes(c.key)
            );

            // Pivot Flattening identical to ShetSheetReport.tsx
            const flattened: any[] = [];
            rawItems.forEach((item) => {
                const remarksVal = (item.Remarks && item.Remarks.trim()) || "No remarks";

                const staffMap: Record<string, string> = {};
                if (Array.isArray(item.involvedStaffs)) {
                    item.involvedStaffs.forEach((st) => {
                        if (st.Involved_Emp_Type && st.Emp_Name) {
                            const k = `staff_${st.Involved_Emp_Type.trim()}`;
                            if (staffMap[k]) {
                                const existingList = staffMap[k].split(", ").map((s) => s.trim());
                                if (!existingList.includes(st.Emp_Name.trim())) {
                                    staffMap[k] = `${staffMap[k]}, ${st.Emp_Name.trim()}`;
                                }
                            } else {
                                staffMap[k] = st.Emp_Name.trim();
                            }
                        }
                    });
                }

                const baseRowInfo = {
                    RemarksGroup: remarksVal,
                    Do_Date: item.Do_Date ? dayjs(item.Do_Date).format("DD-MM-YY") : "",
                    retailerNameGet: item.retailerNameGet || "",
                    Follow_Up: item.Follow_Up || "",
                    Status: item.Status || item.Delivery_Status || "Pending",
                    Do_Inv_No: item.Do_Inv_No || "",
                    voucherTypeGet: item.voucherTypeGet || "",
                    branchNameGet: item.branchNameGet || "",
                    Total_Invoice_value: Number(item.Total_Invoice_value || 0),
                    Created_BY_Name: item.Created_BY_Name || "",
                    Created_on: item.Created_on ? dayjs(item.Created_on).format("DD-MM-YYYY HH:mm") : "",
                    Narration: item.Narration || "",
                    trip_no: item.trip_no || "",
                    ...staffMap,
                };

                const stocks = Array.isArray(item.stockDetails) && item.stockDetails.length > 0 ? item.stockDetails : [];
                const totalAltBags = stocks.reduce((sum, s) => sum + Number(s.Alt_Act_Qty || 0), 0);
                const totalActBags = stocks.reduce((sum, s) => sum + Number(s.Act_Qty || 0), 0);
                const invoiceTotalBags =
                    item.Pending_Bags !== null && item.Pending_Bags !== undefined
                        ? Number(item.Pending_Bags)
                        : totalAltBags > 0
                            ? totalAltBags
                            : totalActBags;

                const uniqueGodowns = Array.from(new Set(stocks.map((s) => (s.Godown_Name || "").trim()).filter(Boolean)));
                const godownDisplay = uniqueGodowns.join(", ");

                if (!isDetailedStock) {
                    const uniqueProducts = Array.from(new Set(stocks.map((s) => (s.Product_Name || "").trim()).filter(Boolean)));
                    const sumBillQty = stocks.reduce((sum, s) => sum + Number(s.Bill_Qty || 0), 0);
                    const sumActQty = stocks.reduce((sum, s) => sum + Number(s.Act_Qty || 0), 0);

                    flattened.push({
                        ...baseRowInfo,
                        Bags: invoiceTotalBags,
                        Godown_Name: godownDisplay,
                        Product_Name: uniqueProducts.join(", "),
                        UOM: stocks[0]?.UOM || "",
                        Bill_Qty: sumBillQty,
                        Act_Qty: sumActQty,
                        billedRate: Number(stocks[0]?.billedRate || 0),
                        itemRate: Number(stocks[0]?.itemRate || 0),
                    });
                } else {
                    if (stocks.length === 0) {
                        flattened.push({
                            ...baseRowInfo,
                            Bags: invoiceTotalBags,
                            Godown_Name: "",
                            Product_Name: "",
                            UOM: "",
                            Bill_Qty: 0,
                            Act_Qty: 0,
                            billedRate: 0,
                            itemRate: 0,
                        });
                    } else {
                        const sameGodown = uniqueGodowns.length === 1 ? uniqueGodowns[0] : "";
                        const productGroupsMap = new Map<string, typeof stocks>();

                        stocks.forEach((s) => {
                            const groupKey = sameGodown
                                ? (s.Product_Name || "").trim()
                                : `${(s.Godown_Name || "").trim()}___${(s.Product_Name || "").trim()}`;

                            if (!productGroupsMap.has(groupKey)) {
                                productGroupsMap.set(groupKey, []);
                            }
                            productGroupsMap.get(groupKey)!.push(s);
                        });

                        productGroupsMap.forEach((pStocks) => {
                            const firstStock = pStocks[0];
                            const pName = firstStock.Product_Name || "";
                            const gName = sameGodown || firstStock.Godown_Name || "";

                            const sumAltQty = pStocks.reduce((sum, s) => sum + Number(s.Alt_Act_Qty || 0), 0);
                            const sumActQty = pStocks.reduce((sum, s) => sum + Number(s.Act_Qty || 0), 0);
                            const sumBillQty = pStocks.reduce((sum, s) => sum + Number(s.Bill_Qty || 0), 0);
                            const bagsCount = sumAltQty > 0 ? sumAltQty : sumActQty;

                            flattened.push({
                                ...baseRowInfo,
                                Bags: bagsCount,
                                Godown_Name: gName,
                                Product_Name: pName,
                                UOM: firstStock.UOM || "",
                                Bill_Qty: sumBillQty,
                                Act_Qty: sumActQty,
                                billedRate: Number(firstStock.billedRate || 0),
                                itemRate: Number(firstStock.itemRate || 0),
                            });
                        });
                    }
                }
            });

            // Group by Remarks
            const groupsMap = new Map<string, any[]>();
            flattened.forEach((row) => {
                const grp = row.RemarksGroup;
                if (!groupsMap.has(grp)) groupsMap.set(grp, []);
                groupsMap.get(grp)!.push(row);
            });

            const groupedData = Array.from(groupsMap.entries()).map(([groupName, items]) => ({
                groupName,
                items,
                totalBags: items.reduce((sum, r) => sum + (Number(r.Bags) || 0), 0),
            }));

            return {
                type: "shetSheet",
                title: "Shet Sheet Pending Report",
                activeCols,
                groupedData,
                fromDate: from,
                toDate: to,
            };
        }
    },

    // 19. STOCK DIFFERENCE REPORT
    "/stockdifference": {
        modes: ["Default"],
        exportTypes: ["Standard"],
        fetchData: async ({ from, to, templateId }) => {
            let godownId = 0;
            let templateCols: any[] = [];

            if (templateId && templateId !== "default") {
                const storedGodownId = localStorage.getItem(`template_godown_${templateId}`);
                if (storedGodownId && !isNaN(Number(storedGodownId))) {
                    godownId = Number(storedGodownId);
                }

                try {
                    const tRes = await SettingsService.getReportEditData({
                        reportId: Number(templateId),
                        typeId: 1,
                    });
                    if (tRes.data?.success && tRes.data.data) {
                        const info = tRes.data.data.reportInfo || tRes.data.data;
                        if (!godownId && info?.Abstract_SP && !isNaN(Number(info.Abstract_SP))) {
                            godownId = Number(info.Abstract_SP);
                        }
                        const cols = tRes.data.data.columns || [];
                        if (Array.isArray(cols) && cols.length > 0) {
                            templateCols = cols.filter((c: any) => c.enabled).sort((a: any, b: any) => (a.order ?? 0) - (b.order ?? 0));
                        }
                    }
                } catch (e) {
                    console.warn("Could not load template for Stock Difference Report:", e);
                }
            }

            const res = await StockDifferenceService.getReport({
                Fromdate: from,
                Todate: to,
                Godown_Id: godownId,
            });
            const rawItems: StockDifferenceItem[] = res.data?.data || (Array.isArray(res.data) ? (res.data as any) : []);

            // Default frontend columns matching StockDifferenceReport.tsx exactly:
            const defaultFrontendCols = [
                { key: "Godown_Name", label: "Godown" },
                { key: "Stock_Date", label: "Date" },
                { key: "Grade_Item_Group", label: "Grade Item Group" },
                { key: "Item_Name", label: "Item Name" },
                { key: "System_Qty", label: "System Stock" },
                { key: "Actual_Qty", label: "Manual Stock" },
                { key: "Difference_Qty", label: "Difference" },
                { key: "Remarks", label: "Remarks" },
                { key: "Created_At", label: "Created Date-Time" },
            ];

            const activeCols = templateCols.length > 0 ? templateCols : defaultFrontendCols;

            // Group by Stock_Group exactly like the parent StockDifferenceReport.tsx screen
            const groupMap = new Map<string, StockDifferenceItem[]>();
            rawItems.forEach((item) => {
                const grpName = String(item.Stock_Group || "Others").trim();
                if (!groupMap.has(grpName)) {
                    groupMap.set(grpName, []);
                }
                groupMap.get(grpName)!.push(item);
            });

            const groupedData: Array<{ groupName: string; rows: StockDifferenceItem[] }> = [];
            groupMap.forEach((items, groupName) => {
                groupedData.push({ groupName, rows: items });
            });

            return {
                type: "stockDifference",
                title: "Stock Difference Report",
                activeCols,
                groupedData,
                rawRows: rawItems,
                fromDate: from,
                toDate: to,
            };
        }
    }
};

/* ================= COMPONENT ================= */

const BatchExportScreen: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    /* Date range */
    const [fromDate, setFromDate] = useState<string>(today);
    const [toDate, setToDate] = useState<string>(today);

    /* Selected Reports to Export */
    const [selectedReports, setSelectedReports] = useState<SelectedReportConfig[]>([]);

    /* Available Report Menu Options */
    const [availableMenuReports, setAvailableMenuReports] = useState<AvailableMenuReportItem[]>([]);

    /* Saved Preset Templates */
    const [presetTemplates, setPresetTemplates] = useState<SavedPresetTemplate[]>([]);
    const [selectedPresetId, setSelectedPresetId] = useState<string>("");

    /* Save Preset Dialog */
    const [saveDialogOpen, setSaveDialogOpen] = useState<boolean>(false);
    const [presetNameInput, setPresetNameInput] = useState<string>("");

    /* Loading state */
    const [exporting, setExporting] = useState<boolean>(false);
    const [exportStatusText, setExportStatusText] = useState<string>("" );
    const [loadingPreset, setLoadingPreset] = useState<boolean>(false);

    /* DnD Sensors for Reordering Reports */
    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: { distance: 5 },
        })
    );

    const handleDragEnd = (event: DragEndEvent) => {
        const { active, over } = event;
        if (!over || active.id === over.id) return;

        setSelectedReports((prev) => {
            const oldIndex = prev.findIndex((item) => item.id === active.id);
            const newIndex = prev.findIndex((item) => item.id === over.id);
            if (oldIndex !== -1 && newIndex !== -1) {
                return arrayMove(prev, oldIndex, newIndex);
            }
            return prev;
        });
    };

    /* Initialize available reports & presets */
    useEffect(() => {
        // Load presets from backend database (tbl_ERP_Report) with fallback to localStorage
        SettingsService.getBatchPresets()
            .then((res) => {
                if (res.data?.success && Array.isArray(res.data.data) && res.data.data.length > 0) {
                    setPresetTemplates(res.data.data);
                } else {
                    // Fallback to local storage if DB has no presets yet
                    const saved = localStorage.getItem(STORAGE_PRESETS_KEY);
                    if (saved) {
                        const localPresets = JSON.parse(saved);
                        setPresetTemplates(localPresets);
                    }
                }
            })
            .catch((e) => {
                console.warn("Failed to load export presets from DB, checking localStorage:", e);
                try {
                    const saved = localStorage.getItem(STORAGE_PRESETS_KEY);
                    if (saved) {
                        setPresetTemplates(JSON.parse(saved));
                    }
                } catch (err) {
                    console.error("Failed to load export presets:", err);
                }
            });

        // Fetch menus and list ALL available report submenus dynamically
        Promise.all([
            MenuService.getMenus(),
            SettingsService.getMenuSP().catch(() => ({ data: { success: false, data: [] } }))
        ])
            .then(([resMenus, resSP]) => {
                const menus = resMenus.data?.data || [];
                const spMappings = (resSP.data?.data || []) as any[];

                const found: Array<{
                    key: string;
                    name: string;
                    path: string;
                    parentReportName: string;
                    definition: ReportDefinition;
                }> = [];

                // 1. Scan MenuService menus
                menus.forEach((m: any) => {
                    const subList = m.SubMenu || (m.SubRoutes ? [m] : []);
                    subList.forEach((sub: any) => {
                        const path = sub.rUrl || sub.path || "";
                        const name = sub.name || sub.label || "";
                        if (!path || !name) return;

                        // Check if override exists (normalize key by removing slashes/dashes and lowercasing)
                        const cleanPath = path.toLowerCase().replace(/\/+/g, "/").trim();
                        const overrideKey = Object.keys(CUSTOM_CATALOG_OVERRIDES).find((k) =>
                            cleanPath === k.toLowerCase() ||
                            cleanPath.endsWith(k.toLowerCase()) ||
                            name.trim().toLowerCase() === k.replace("/", "").toLowerCase()
                        );
                        const override = overrideKey ? CUSTOM_CATALOG_OVERRIDES[overrideKey] : undefined;

                        // Find matching SP mapping if available
                        const spMatch = spMappings.find((sp) =>
                            sp.Report_Name?.trim().toLowerCase() === name.trim().toLowerCase()
                        );

                        const modes = override?.modes || (spMatch ? ["Abstract", "Expanded"] : ["Default"]);
                        const exportTypes = override?.exportTypes || ["Standard"];
                        const key = path.replace(/[\/\-_]/g, "").toLowerCase() || name.replace(/[\s\-_]/g, "").toLowerCase();

                        if (!found.some((x) => x.path.toLowerCase() === path.toLowerCase())) {
                            found.push({
                                key,
                                name,
                                path,
                                parentReportName: name,
                                definition: {
                                    key,
                                    name,
                                    path,
                                    parentReportName: name,
                                    modes,
                                    exportTypes,
                                    fetchData: override?.fetchData || (async ({ from, to, mode, templateId }) => {
                                        if (spMatch) {
                                            const spName = mode === "Expanded" && spMatch.Expanded_SP ? spMatch.Expanded_SP : (spMatch.Abstract_SP || spMatch.Expanded_SP);
                                            if (spName) {
                                                const res = await SettingsService.executeSP({
                                                    spName,
                                                    params: { Fromdate: from, Todate: to }
                                                });
                                                const rows = res.data?.data || [];

                                                // If template selected, fetch enabled columns
                                                let templateCols: any[] = [];
                                                if (templateId && templateId !== "default") {
                                                    try {
                                                        const tRes = await SettingsService.getReportEditData({
                                                            reportId: Number(templateId),
                                                            typeId: mode === "Expanded" ? 2 : 1
                                                        });
                                                        if (tRes.data?.success && tRes.data.data?.columns) {
                                                            templateCols = tRes.data.data.columns.filter((c: any) => c.enabled).sort((a: any, b: any) => a.order - b.order);
                                                        }
                                                    } catch (e) {
                                                        console.warn(e);
                                                    }
                                                }

                                                return {
                                                    type: "table",
                                                    title: `${name} (${mode})`,
                                                    rows,
                                                    templateCols
                                                };
                                            }
                                        }
                                        return { type: "table", title: `${name} (${mode})`, rows: [] };
                                    })
                                }
                            });
                        }
                    });
                });

                // 2. Also register any custom reports in CUSTOM_CATALOG_OVERRIDES that might not appear in MenuService
                const customNames: Record<string, string> = {
                    "/ratemasteradmin": "Rate Master Admin",
                    "/expense": "Expenses Report",
                    "/recievablepayable": "Receivable & Payable Report",
                    "/cashbox": "Cash Box Report",
                    "/salesreport": "Online Sales Report",
                    "/salesreportlol": "Online Sales Report LOL",
                    "/purchasereport": "Online Purchase Report",
                    "/salesreportlr": "Sales Analytics Report",
                    "/outstandingmanagement": "Outstanding Management Report",
                    "/saleorder": "Sale Order Report",
                    "/pendingsaleorder": "Pending Sale Order",
                    "/purchaseorder": "Purchase Order Report",
                    "/bankabstract": "Bank Abstract Report",
                    "/dayabstract": "Day Abstract Report",
                    "/stockabstract": "Stock Abstract Report",
                    "/stockvalue": "Stock Value Report",
                    "/ratemaster": "Rate Master",
                    "/shetSheet": "Shet Sheet Report",
                    "/stockdifference": "Stock Difference Report",
                };

                Object.entries(CUSTOM_CATALOG_OVERRIDES).forEach(([cPath, cDef]) => {
                    const normPath = cPath.toLowerCase();
                    if (!found.some((f) => f.path.toLowerCase() === normPath || f.key === normPath.replace(/[\/\-_]/g, ""))) {
                        const defaultName = customNames[normPath] || normPath.replace("/", "").toUpperCase();
                        const key = normPath.replace(/[\/\-_]/g, "");
                        found.push({
                            key,
                            name: defaultName,
                            path: cPath,
                            parentReportName: defaultName,
                            definition: {
                                key,
                                name: defaultName,
                                path: cPath,
                                parentReportName: defaultName,
                                modes: cDef.modes || ["Default"],
                                exportTypes: cDef.exportTypes || ["Standard"],
                                fetchData: cDef.fetchData!
                            }
                        });
                    }
                });

                setAvailableMenuReports(found);
            })
            .catch((err) => {
                console.error("Failed to load menus in BatchExportScreen:", err);
            });
    }, []);

    /* Helper: Fetch templates with dynamic multi-variant & reportList fallback */
    const fetchTemplatesForParent = async (parentReportName: string) => {
        let templates: Array<{ id: number; name: string }> = [];
        const clean = (parentReportName || "").trim();
        const normClean = clean.toLowerCase().replace(/[\s\-_/]/g, "");

        const isRecOrOut = normClean.includes("recievable") || normClean.includes("receivable") || normClean.includes("outstanding");

        const variants = Array.from(
            new Set([
                clean,
                clean.toUpperCase(),
                clean.toLowerCase(),
                clean.replace(/\s+/g, ""),
                clean.replace(/[\/\-_]/g, " ").trim(),
                clean.replace(/[\/\-_]/g, "").trim(),
                // Specific report aliases
                clean === "SHETSHEET REPORT" || clean === "SHET SHEET REPORT" || clean.toLowerCase().includes("shet")
                    ? "SHET SHEET REPORT"
                    : "",
                clean.toLowerCase().includes("stock difference") ? "STOCK DIFFERENCE REPORT" : "",
                clean.toLowerCase().includes("rate master") ? "RATE MASTER ADMIN" : "",
                clean.toLowerCase().includes("overall staff") ? "Overall Staff Report" : "",
                clean.toLowerCase().includes("ledger") && clean.toLowerCase().includes("item") ? "LEDGER WISE ITEM" : "",
                isRecOrOut ? "OUTSTANDING MANAGEMENT" : "",
                isRecOrOut ? "RECIEVABLE PAYABLE" : "",
                isRecOrOut ? "RECEIVABLE PAYABLE" : "",
            ])
        ).filter(Boolean);

        for (const variant of variants) {
            try {
                const res = await SettingsService.getReportsByParent(variant);
                const data = res.data?.data || [];
                if (data.length > 0) {
                    const mapped = data.map((t: any) => ({
                        id: t.Report_Id || t.id,
                        name: t.Report_Name || t.name,
                    }));
                    mapped.forEach((m: any) => {
                        if (!templates.some((ex) => ex.id === m.id)) {
                            templates.push(m);
                        }
                    });
                }
            } catch (e) {
                // Try next variant
            }
        }

        // Dynamic fallback: If no templates found via getReportsByParent, query the full reportList dynamically
        if (templates.length === 0) {
            try {
                const listRes = await SettingsService.getReportList();
                const allData = listRes.data?.data || (Array.isArray(listRes.data) ? listRes.data : {});
                
                // If grouped object { "PARENT_NAME": [ { Report_Id, Report_Name, ... } ] }
                let allReports: any[] = [];
                if (Array.isArray(allData)) {
                    allReports = allData;
                } else if (allData && typeof allData === "object") {
                    Object.entries(allData).forEach(([parentKey, list]: [string, any]) => {
                        if (Array.isArray(list)) {
                            list.forEach((item: any) => {
                                allReports.push({
                                    ...item,
                                    Parent_Report: item.Parent_Report || parentKey
                                });
                            });
                        }
                    });
                }

                if (allReports.length > 0) {
                    const matched = allReports.filter((r: any) => {
                        const rParent = String(r.Parent_Report || "").toLowerCase().replace(/[\s\-_/]/g, "");
                        const rName = String(r.Report_Name || "").toLowerCase().replace(/[\s\-_/]/g, "");
                        if (isRecOrOut) {
                            return rParent.includes("outstanding") || rParent.includes("recievable") || rParent.includes("receivable");
                        }
                        return (
                            (rParent && (rParent === normClean || rParent.includes(normClean) || normClean.includes(rParent))) ||
                            (rName && (rName === normClean || rName.includes(normClean)))
                        );
                    });

                    if (matched.length > 0) {
                        templates = matched.map((t: any) => ({
                            id: t.Report_Id || t.id,
                            name: t.Report_Name || t.name,
                        }));
                    }
                }
            } catch (err) {
                console.warn("Dynamic template discovery fallback error:", err);
            }
        }

        return templates;
    };

    /* Add report to selection */
    const handleAddReport = async (reportKey: string) => {
        const item = availableMenuReports.find((c) => c.key === reportKey);
        if (!item) return;

        const cat = item.definition;

        // Fetch templates for this report
        const templates = await fetchTemplatesForParent(cat.parentReportName);

        const newConfig: SelectedReportConfig = {
            id: `${cat.key}_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
            reportKey: cat.key,
            reportName: cat.name,
            path: cat.path,
            availableModes: cat.modes,
            selectedModes: [cat.modes[0]],
            availableTemplates: templates,
            selectedTemplateId: "default",
            availableExportTypes: cat.exportTypes,
            selectedExportType: cat.exportTypes[0]
        };

        setSelectedReports((prev) => [...prev, newConfig]);
    };

    /* Remove report */
    const handleRemoveReport = (id: string) => {
        setSelectedReports((prev) => prev.filter((r) => r.id !== id));
    };

    /* Update report config */
    const handleUpdateReport = (id: string, updates: Partial<SelectedReportConfig>) => {
        setSelectedReports((prev) =>
            prev.map((r) => (r.id === id ? { ...r, ...updates } : r))
        );
    };

    /* Open Save/Edit Preset Dialog */
    const handleOpenSavePresetDialog = () => {
        if (selectedReports.length === 0) {
            toast.error("Add at least one report before saving preset");
            return;
        }

        const existingPreset = presetTemplates.find((p) => p.id === selectedPresetId);
        if (existingPreset) {
            // Pre-fill with the loaded preset's name so it opens in Edit mode
            setPresetNameInput(existingPreset.name);
        } else {
            setPresetNameInput("");
        }
        setSaveDialogOpen(true);
    };

    /* Save current selection as Preset Template (or update existing) */
    const handleSavePreset = async () => {
        if (!presetNameInput.trim()) {
            toast.error("Please provide a preset template name");
            return;
        }
        if (selectedReports.length === 0) {
            toast.error("Add at least one report before saving preset");
            return;
        }

        const existingPreset = presetTemplates.find((p) => p.id === selectedPresetId);
        const userData = JSON.parse(localStorage.getItem("user") || "{}");
        const createdBy = userData?.id || 0;

        const reportsPayload = selectedReports.map((r) => ({
            reportKey: r.reportKey,
            selectedModes: r.selectedModes,
            selectedTemplateId: r.selectedTemplateId,
            selectedExportType: r.selectedExportType,
        }));

        try {
            if (existingPreset) {
                // Update the existing preset in database
                const res = await SettingsService.saveBatchPreset({
                    id: selectedPresetId,
                    name: presetNameInput.trim(),
                    reports: reportsPayload,
                    createdBy
                });

                const updatedPreset: SavedPresetTemplate = {
                    ...existingPreset,
                    id: res.data?.data?.id || existingPreset.id,
                    name: presetNameInput.trim(),
                    createdAt: res.data?.data?.createdAt || dayjs().format("DD-MM-YYYY HH:mm"),
                    reports: reportsPayload,
                };

                const updatedList = presetTemplates.map((p) =>
                    p.id === selectedPresetId ? updatedPreset : p
                );
                setPresetTemplates(updatedList);
                localStorage.setItem(STORAGE_PRESETS_KEY, JSON.stringify(updatedList));
                setPresetNameInput("");
                setSaveDialogOpen(false);
                toast.success(`Preset "${updatedPreset.name}" updated successfully! ✅`);
            } else {
                // Create a new preset in database
                const res = await SettingsService.saveBatchPreset({
                    name: presetNameInput.trim(),
                    reports: reportsPayload,
                    createdBy
                });

                const newPreset: SavedPresetTemplate = {
                    id: String(res.data?.data?.id || `preset_${Date.now()}`),
                    name: presetNameInput.trim(),
                    createdAt: res.data?.data?.createdAt || dayjs().format("DD-MM-YYYY HH:mm"),
                    reports: reportsPayload,
                };

                const updated = [...presetTemplates, newPreset];
                setPresetTemplates(updated);
                localStorage.setItem(STORAGE_PRESETS_KEY, JSON.stringify(updated));
                setSelectedPresetId(newPreset.id);
                setPresetNameInput("");
                setSaveDialogOpen(false);
                toast.success(`Preset "${newPreset.name}" saved successfully ✅`);
            }
        } catch (saveErr) {
            console.error("Save preset to DB error:", saveErr);
            toast.error("Failed to save preset ❌");
        }
    };

    /* Load Preset */
    const handleApplyPreset = async (presetId: string) => {
        setSelectedPresetId(presetId);
        if (!presetId) {
            return;
        }

        const preset = presetTemplates.find((p) => p.id === presetId);
        if (!preset) return;

        setLoadingPreset(true);
        try {
            const reconstructed: SelectedReportConfig[] = [];

            for (const item of preset.reports) {
                const foundMenu = availableMenuReports.find((c) =>
                    c.key === item.reportKey ||
                    c.path.toLowerCase().replace(/[\/\-_]/g, "") === item.reportKey.toLowerCase().replace(/[\/\-_]/g, "") ||
                    c.name.toLowerCase().replace(/[\s\-_]/g, "") === item.reportKey.toLowerCase().replace(/[\s\-_]/g, "")
                );
                const cat = foundMenu?.definition;
                if (!cat) continue;

                const templates = await fetchTemplatesForParent(cat.parentReportName);

                reconstructed.push({
                    id: `${cat.key}_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
                    reportKey: cat.key,
                    reportName: cat.name,
                    path: cat.path,
                    availableModes: cat.modes,
                    selectedModes: item.selectedModes.length ? item.selectedModes : [cat.modes[0]],
                    availableTemplates: templates,
                    selectedTemplateId: item.selectedTemplateId ?? "default",
                    availableExportTypes: cat.exportTypes,
                    selectedExportType: item.selectedExportType || cat.exportTypes[0]
                });
            }

            setSelectedReports(reconstructed);
            toast.info(`Loaded preset: ${preset.name}`);
        } catch (err) {
            console.error("Failed to apply preset:", err);
            toast.error("Failed to load preset");
        } finally {
            setLoadingPreset(false);
        }
    };

    /* Delete Preset */
    const handleDeletePreset = async (presetId: string) => {
        try {
            if (presetId && !String(presetId).startsWith("preset_")) {
                await SettingsService.deleteBatchPreset(presetId);
            }
        } catch (e) {
            console.error("Delete preset from DB failed:", e);
        }

        const updated = presetTemplates.filter((p) => p.id !== presetId);
        setPresetTemplates(updated);
        localStorage.setItem(STORAGE_PRESETS_KEY, JSON.stringify(updated));
        if (selectedPresetId === presetId) {
            setSelectedPresetId("");
        }
        toast.success("Preset removed");
    };

    /* ================= BATCH EXPORT EXCEL ================= */

    const handleExportSingleExcel = async () => {
        if (!selectedReports.length) {
            toast.warning("Please select at least one report to export.");
            return;
        }

        try {
            setExporting(true);
            setExportStatusText("Initializing Excel Workbook...");

            const wb = XLSX.utils.book_new();

            for (let i = 0; i < selectedReports.length; i++) {
                const cfg = selectedReports[i];
                const item = availableMenuReports.find((c) =>
                    c.key === cfg.reportKey ||
                    c.path.toLowerCase().replace(/[\/\-_]/g, "") === cfg.reportKey.toLowerCase().replace(/[\/\-_]/g, "") ||
                    c.name.toLowerCase().replace(/[\s\-_]/g, "") === cfg.reportKey.toLowerCase().replace(/[\s\-_]/g, "")
                );
                const cat = item?.definition;
                if (!cat) continue;

                for (const mode of cfg.selectedModes) {
                    const templateName = cfg.selectedTemplateId !== "default"
                        ? cfg.availableTemplates.find((t) => String(t.id) === String(cfg.selectedTemplateId))?.name
                        : null;
                    const fullReportTitle = templateName
                        ? `${cfg.reportName} - ${mode} - ${templateName}`
                        : `${cfg.reportName} - ${mode}`;

                    setExportStatusText(`Fetching ${fullReportTitle}...`);

                    const data = await cat.fetchData({
                        from: fromDate,
                        to: toDate,
                        mode,
                        templateId: cfg.selectedTemplateId
                    });

                    // 1. EXPENSE REPORT
                    if (data.type === "expense") {
                        const isDirect = mode === "Direct";
                        const sectionData = data.mapped?.[isDirect ? "Direct" : "Indirect"] || data.raw?.[isDirect ? "Direct" : "Indirect"] || { groups: [] };

                        if (cfg.selectedExportType === "Management") {
                            // Bright Yellow Management Layout
                            const formatDisplayDate = (dStr: string) => {
                                if (!dStr) return "";
                                const d = dayjs(dStr);
                                return d.isValid() ? d.format("DD.MM.YYYY") : dStr;
                            };
                            const dateHeader = fromDate === toDate
                                ? formatDisplayDate(toDate)
                                : `${formatDisplayDate(fromDate)} - ${formatDisplayDate(toDate)}`;

                            const borderAll = {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } }
                            };
                            const headerStyle = {
                                font: { bold: true, sz: 12, name: "Calibri", color: { rgb: "000000" } },
                                fill: { fgColor: { rgb: "FFFF00" } },
                                alignment: { horizontal: "center", vertical: "center" },
                                border: borderAll
                            };
                            const colHeaderStyle = {
                                font: { bold: true, sz: 12, name: "Calibri", color: { rgb: "000000" } },
                                fill: { fgColor: { rgb: "FFFF00" } },
                                alignment: { horizontal: "center", vertical: "center" },
                                border: borderAll
                            };
                            const groupStyle = {
                                font: { bold: true, sz: 11, name: "Calibri", color: { rgb: "000000" } },
                                alignment: { horizontal: "left", vertical: "center" },
                                border: borderAll
                            };
                            const subGroupStyle = {
                                font: { sz: 11, name: "Calibri", color: { rgb: "000000" } },
                                alignment: { horizontal: "left", vertical: "center" },
                                border: borderAll
                            };
                            const amountStyle = {
                                font: { sz: 11, name: "Calibri", color: { rgb: "000000" } },
                                alignment: { horizontal: "right", vertical: "center" },
                                border: borderAll,
                                numFmt: "#,##0.00"
                            };
                            const blankBorderedCell = { border: borderAll };
                            const footerTitleStyle = {
                                font: { bold: true, sz: 12, name: "Calibri", color: { rgb: "000000" } },
                                fill: { fgColor: { rgb: "FFFF00" } },
                                alignment: { horizontal: "center", vertical: "center" },
                                border: borderAll
                            };
                            const footerAmountStyle = {
                                font: { bold: true, sz: 12, name: "Calibri", color: { rgb: "000000" } },
                                fill: { fgColor: { rgb: "FFFF00" } },
                                alignment: { horizontal: "right", vertical: "center" },
                                border: borderAll,
                                numFmt: "#,##0.00"
                            };

                            const titleText = fullReportTitle;
                            const rows: any[][] = [];

                            rows.push([{ v: titleText, s: headerStyle }, { v: "", s: headerStyle }]);
                            rows.push([{ v: dateHeader, s: headerStyle }, { v: "", s: headerStyle }]);
                            rows.push([{ v: "Particulars", s: colHeaderStyle }, { v: "Amount", s: colHeaderStyle }]);

                            let grandTotal = 0;
                            sectionData.groups.forEach((group: any) => {
                                rows.push([{ v: group.name.toUpperCase(), s: groupStyle }, { v: "", s: blankBorderedCell }]);
                                (group.subGroups || []).forEach((sub: any) => {
                                    const amt = Number(sub.balance || 0);
                                    grandTotal += amt;
                                    rows.push([
                                        { v: `    ${sub.name}`, s: subGroupStyle },
                                        { v: amt, t: "n", s: amountStyle }
                                    ]);
                                });
                            });

                            rows.push([
                                { v: "Grand Total", s: footerTitleStyle },
                                { v: grandTotal, t: "n", s: footerAmountStyle }
                            ]);

                            const ws = XLSX.utils.aoa_to_sheet(rows);
                            ws["!merges"] = [
                                { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
                                { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }
                            ];
                            ws["!cols"] = [{ wch: 45 }, { wch: 20 }];
                            ws["!rows"] = [{ hpt: 24 }, { hpt: 20 }, { hpt: 20 }];

                            const sheetName = `${mode}_Expenses`.substring(0, 31);
                            XLSX.utils.book_append_sheet(wb, ws, sheetName);
                        } else {
                            // Normal Flat Detailed List
                            const flatRows: any[] = [];
                            sectionData.groups.forEach((g: any) => {
                                (g.subGroups || []).forEach((sg: any) => {
                                    (sg.ledgers || []).forEach((l: any) => {
                                        flatRows.push({
                                            Group: g.name,
                                            SubGroup: sg.name,
                                            Date: l.payment_date ? dayjs(l.payment_date).format("DD-MM-YYYY") : "",
                                            InvoiceNo: l.payment_invoice_no || "",
                                            Voucher: l.voucher_name || "",
                                            Debit: l.entryType === "DR" ? l.amount : 0,
                                            Credit: l.entryType === "CR" ? l.amount : 0,
                                            Remarks: l.remarks || ""
                                        });
                                    });
                                });
                            });
                            const ws = XLSX.utils.json_to_sheet(flatRows.length ? flatRows : [{ Message: "No records" }]);
                            const sheetName = `${mode}_Exp_Normal`.substring(0, 31);
                            XLSX.utils.book_append_sheet(wb, ws, sheetName);
                        }
                    }
                    // 2. RATE MASTER ADMIN
                    else if (data.type === "rateMasterAdmin") {
                        const rawRows = data.rows || [];
                        const templateCols = data.templateCols || [];
                        const groupingCols = data.groupingCols || [];

                        // Enabled columns projection
                        const enabledCols = templateCols.length > 0
                            ? templateCols
                            : [
                                { key: "Product_Name", label: "Product" },
                                { key: "Min_Rate", label: "Min Rate", isNumeric: true },
                                { key: "Rate", label: "List Rate", isNumeric: true },
                                { key: "Margin", label: "Margin", isNumeric: true },
                                { key: "Margin_Percentage", label: "%", isNumeric: true },
                                { key: "COGS", label: "COGS", isNumeric: true },
                                { key: "GP_Percentage_COGS", label: "COGS %", isNumeric: true },
                                { key: "Max_Rate", label: "Max Rate", isNumeric: true }
                            ];

                        const exportAoa: any[][] = [];
                        const headerRow = ["S.No", ...enabledCols.map((c: any) => c.label)];
                        exportAoa.push(headerRow);

                        let serial = 1;
                        if (groupingCols.length > 0) {
                            // Group by first grouping col
                            const grpKey = groupingCols[0];
                            const groupMap = new Map<string, any[]>();
                            rawRows.forEach((r: any) => {
                                const val = String(r[grpKey] ?? "Others");
                                if (!groupMap.has(val)) groupMap.set(val, []);
                                groupMap.get(val)!.push(r);
                            });

                            groupMap.forEach((grpRows, val) => {
                                const groupRow = [
                                    "",
                                    `Group: ${grpKey} = ${val}`,
                                    ...enabledCols.slice(1).map((c: any) => {
                                        if (c.isNumeric || c.key === "Rate" || c.key === "Min_Rate" || c.key === "Max_Rate" || c.key === "COGS" || c.key === "Margin") {
                                            const sum = grpRows.reduce((s: number, r: any) => s + (Number(r[c.key]) || 0), 0);
                                            return c.key === "GP_Percentage_COGS" || c.key === "Margin_Percentage"
                                                ? (sum / grpRows.length).toFixed(1)
                                                : sum;
                                        }
                                        return "";
                                    })
                                ];
                                exportAoa.push(groupRow);

                                grpRows.forEach((r: any) => {
                                    exportAoa.push([
                                        serial++,
                                        ...enabledCols.map((c: any) => {
                                            return formatExportCellValue(r[c.key], c.key, r, false);
                                        })
                                    ]);
                                });
                            });
                        } else {
                            rawRows.forEach((r: any) => {
                                exportAoa.push([
                                    serial++,
                                    ...enabledCols.map((c: any) => {
                                        return formatExportCellValue(r[c.key], c.key, r, false);
                                    })
                                ]);
                            });
                        }

                        const ws = XLSX.utils.aoa_to_sheet(exportAoa.length ? exportAoa : [["No records found"]]);

                        // Apply yellow header box styling matching user reference photo
                        if (exportAoa.length > 0) {
                            const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
                            const borderAll = {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } }
                            };

                            for (let R = range.s.r; R <= range.e.r; ++R) {
                                for (let C = range.s.c; C <= range.e.c; ++C) {
                                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                    const cell = ws[cellRef];
                                    if (!cell) continue;

                                    if (R === 0) {
                                        cell.s = {
                                            fill: { fgColor: { rgb: "FFFF00" } },
                                            font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } },
                                            alignment: { horizontal: C === 0 ? "center" : "center", vertical: "center" },
                                            border: borderAll
                                        };
                                    } else {
                                        const isNumeric = typeof cell.v === "number";
                                        cell.s = {
                                            font: { name: "Calibri", sz: 11, color: { rgb: "000000" } },
                                            alignment: {
                                                horizontal: C === 0 ? "center" : isNumeric ? "right" : "left",
                                                vertical: "center"
                                            },
                                            border: borderAll
                                        };
                                    }
                                }
                            }

                            const colWidths: Array<{ wch: number }> = [{ wch: 7 }];
                            enabledCols.forEach((c: any) => {
                                let maxLen = (c.label || "").length;
                                rawRows.slice(0, 50).forEach((r: any) => {
                                    const l = String(r[c.key] ?? "").length;
                                    if (l > maxLen) maxLen = l;
                                });
                                colWidths.push({ wch: Math.min(Math.max(maxLen + 3, 10), 38) });
                            });
                            ws["!cols"] = colWidths;
                        }

                        XLSX.utils.book_append_sheet(wb, ws, "Rate_Master_Admin".substring(0, 31));
                    }
                    // 3. RECIEVABLE PAYABLE
                    else if (data.type === "recievablePayable") {
                        const rows = data.rows || [];
                        const groupMap = new Map<string, any[]>();
                        rows.forEach((r: any) => {
                            const grp = r.Group_Name || "No Group";
                            if (!groupMap.has(grp)) groupMap.set(grp, []);
                            groupMap.get(grp)!.push(r);
                        });

                        const exportRows: any[] = [];
                        let totalDebit = 0;
                        let totalCredit = 0;
                        let sNo = 1;

                        groupMap.forEach((items, grpName) => {
                            let subDeb = 0;
                            let subCred = 0;
                            items.forEach((it: any) => {
                                subDeb += Number(it.Debit_Amount) || 0;
                                subCred += Number(it.Credit_Amount) || 0;
                            });
                            totalDebit += subDeb;
                            totalCredit += subCred;

                            exportRows.push({
                                "S.No": "",
                                "Account Name": grpName,
                                "Invoice Date": "",
                                "Invoice No": "",
                                "Recievable": subDeb !== 0 ? subDeb : "",
                                "Payable": subCred !== 0 ? subCred : ""
                            });

                            items.forEach((it: any) => {
                                exportRows.push({
                                    "S.No": sNo++,
                                    "Account Name": it.Account_name || "",
                                    "Invoice Date": it.invoice_date ? dayjs(it.invoice_date).format("DD/MM/YYYY") : "",
                                    "Invoice No": it.invoice_no || "",
                                    "Recievable": it.Debit_Amount || 0,
                                    "Payable": it.Credit_Amount || 0
                                });
                            });
                        });

                        const bal = data.isReceivable ? totalDebit - totalCredit : totalCredit - totalDebit;
                        exportRows.push({
                            "S.No": "",
                            "Account Name": `Grand Total (Bal: ${formatINR(bal)})`,
                            "Invoice Date": "",
                            "Invoice No": "",
                            "Recievable": totalDebit,
                            "Payable": totalCredit
                        });

                        const ws = XLSX.utils.json_to_sheet(exportRows.length ? exportRows : [{ Message: "No records" }]);
                        const safeTitle = (data.isReceivable ? "Debtors_Receivable" : "Creditors_Payable").substring(0, 31);
                        XLSX.utils.book_append_sheet(wb, ws, safeTitle);
                    }
                    // 4. CASH BOX ABSTRACT
                    else if (data.type === "cashboxAbstract") {
                        const parsed = data.parsed;
                        const excelData: any[][] = [];
                        const dateStr = fromDate === toDate
                            ? dayjs(fromDate).format("DD-MM-YYYY")
                            : `${dayjs(fromDate).format("DD-MM-YYYY")} TO ${dayjs(toDate).format("DD-MM-YYYY")}`;

                        excelData.push([`CASH BOX TRANSACTION - ${dateStr}`]);
                        excelData.push([]);
                        excelData.push(["Particulars (Debit)", "Debit Amt", "Particulars (Credit)", "Credit amt"]);
                        excelData.push(["", "", "Opening Balance", parsed?.opening ?? 0]);

                        const debitGroups = parsed?.debitGroups || [];
                        const creditGroups = parsed?.creditGroups || [];
                        const maxGrp = Math.max(debitGroups.length, creditGroups.length);

                        for (let idx = 0; idx < maxGrp; idx++) {
                            const leftGroup = debitGroups[idx] || { label: "", total: "", subLedgers: [] };
                            const rightGroup = creditGroups[idx] || { label: "", total: "", subLedgers: [] };

                            excelData.push([
                                leftGroup.label,
                                leftGroup.total || "",
                                rightGroup.label,
                                rightGroup.total || ""
                            ]);

                            const maxSubRows = Math.max(leftGroup.subLedgers?.length || 0, rightGroup.subLedgers?.length || 0);
                            for (let sIdx = 0; sIdx < maxSubRows; sIdx++) {
                                const leftSub = leftGroup.subLedgers?.[sIdx];
                                const rightSub = rightGroup.subLedgers?.[sIdx];
                                excelData.push([
                                    leftSub ? `  ${sIdx + 1}. ${leftSub.name}` : "",
                                    leftSub ? leftSub.amount : "",
                                    rightSub ? `  ${sIdx + 1}. ${rightSub.name}` : "",
                                    rightSub ? rightSub.amount : ""
                                ]);
                            }
                        }

                        excelData.push(["", "", "Closing Balance", parsed?.closing ?? 0]);

                        const ws = XLSX.utils.aoa_to_sheet(excelData);
                        styleCashBoxWorksheet(ws);
                        ws["!cols"] = [{ wch: 35 }, { wch: 18 }, { wch: 35 }, { wch: 18 }];
                        XLSX.utils.book_append_sheet(wb, ws, "Cash_Box_Abstract".substring(0, 31));
                    }
                    // 5. CASH BOX EXPANDED
                    else if (data.type === "cashboxExpanded") {
                        const detailedData = data.detailedData;
                        const rows = detailedData?.Data || [];
                        const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Message: "No records" }]);
                        XLSX.utils.book_append_sheet(wb, ws, "Cash_Box_Expanded".substring(0, 31));
                    }
                    // 6. OUTSTANDING MANAGEMENT
                    else if (data.type === "outstandingManagement") {
                        const exportAoa: any[][] = [];
                        exportAoa.push([fullReportTitle]);
                        exportAoa.push(["Account Name", "Days", data.amountColumnTitle]);
                        exportAoa.push(["Total", "", data.grandTotalAmount]);

                        const flattenedRows = data.flattenedRows || [];
                        flattenedRows.forEach((row: any) => {
                            if (row.__group) {
                                exportAoa.push([row.__label, "", row.__subtotal]);
                            } else {
                                const accDisplay = row.invoice_count && row.invoice_count > 1
                                    ? `${row.Account_name} (${row.invoice_count})`
                                    : row.Account_name;
                                exportAoa.push([accDisplay, row.days, row.Bal_Amount]);
                            }
                        });

                        const worksheet = XLSX.utils.aoa_to_sheet(exportAoa);
                        worksheet["!merges"] = [
                            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
                        ];

                        // Styling cells with professional report theme (#1E3A8A)
                        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:C1");
                        for (let R = range.s.r; R <= range.e.r; ++R) {
                            for (let C = range.s.c; C <= range.e.c; ++C) {
                                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                const cell = worksheet[cellRef];
                                if (!cell) continue;

                                cell.s = {
                                    font: { name: "Calibri", sz: 11 },
                                    border: {
                                        top: { style: "thin", color: { rgb: "000000" } },
                                        bottom: { style: "thin", color: { rgb: "000000" } },
                                        left: { style: "thin", color: { rgb: "000000" } },
                                        right: { style: "thin", color: { rgb: "000000" } },
                                    },
                                };

                                if (R === 0) {
                                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                                    cell.s.font = { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } };
                                    cell.s.alignment = { horizontal: "center", vertical: "center" };
                                } else if (R === 1) {
                                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                                    cell.s.font = { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } };
                                    cell.s.alignment = { horizontal: C === 0 ? "left" : "center", vertical: "center" };
                                } else if (R === 2) {
                                    cell.s.fill = { fgColor: { rgb: "EFF6FF" } };
                                    cell.s.font = { name: "Calibri", sz: 12, bold: true, color: { rgb: "1E3A8A" } };
                                    cell.s.alignment = { horizontal: C === 0 ? "left" : "right" };
                                } else {
                                    const rowData = flattenedRows[R - 3];
                                    if (rowData && rowData.__group) {
                                        cell.s.fill = { fgColor: { rgb: "F1F5F9" } };
                                        cell.s.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "0F172A" } };
                                        cell.s.alignment = { horizontal: C === 0 ? "left" : "right" };
                                    } else {
                                        if (C === 1) cell.s.alignment = { horizontal: "center" };
                                        if (C === 2) cell.s.alignment = { horizontal: "right" };
                                    }
                                }
                            }
                        }

                        worksheet["!cols"] = [{ wch: 45 }, { wch: 12 }, { wch: 20 }];
                        const safeTitle = `${cfg.reportName.replace(/[^a-zA-Z0-9]/g, "")}_${mode}`.substring(0, 31);
                        XLSX.utils.book_append_sheet(wb, worksheet, safeTitle);
                    }
                    // 7. SHET SHEET REPORT (EXCEL)
                    else if (data.type === "shetSheet") {
                        const activeCols: any[] = data.activeCols || [];
                        const groupedData: any[] = data.groupedData || [];

                        const exportAoa: any[][] = [];
                        exportAoa.push([fullReportTitle]);
                        exportAoa.push(["S.No", ...activeCols.map((c: any) => c.label)]);

                        groupedData.forEach((grp: any) => {
                            exportAoa.push([grp.groupName]);
                            grp.items.forEach((item: any, idx: number) => {
                                const rowData = [
                                    idx + 1,
                                    ...activeCols.map((c: any) => {
                                        return formatExportCellValue(item[c.key], c.key, item, false);
                                    }),
                                ];
                                exportAoa.push(rowData);
                            });
                        });

                        const ws = XLSX.utils.aoa_to_sheet(exportAoa);
                        ws["!merges"] = [
                            { s: { r: 0, c: 0 }, e: { r: 0, c: activeCols.length } },
                        ];

                        const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
                        let currentGroupRow = -1;

                        for (let R = range.s.r; R <= range.e.r; ++R) {
                            const firstCell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
                            const isGroup = firstCell && groupedData.some((g: any) => g.groupName === firstCell.v);
                            if (isGroup) {
                                currentGroupRow = R;
                                if (!ws["!merges"]) ws["!merges"] = [];
                                ws["!merges"].push({ s: { r: R, c: 0 }, e: { r: R, c: activeCols.length } });
                            }

                            for (let C = range.s.c; C <= range.e.c; ++C) {
                                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                const cell = ws[cellRef];
                                if (!cell) continue;

                                cell.s = {
                                    font: { name: "Calibri", sz: 11 },
                                    border: {
                                        top: { style: "thin", color: { rgb: "000000" } },
                                        bottom: { style: "thin", color: { rgb: "000000" } },
                                        left: { style: "thin", color: { rgb: "000000" } },
                                        right: { style: "thin", color: { rgb: "000000" } },
                                    },
                                };

                                if (R === 0) {
                                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                                    cell.s.font = { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } };
                                    cell.s.alignment = { horizontal: "center", vertical: "center" };
                                } else if (R === 1) {
                                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                                    cell.s.font = { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } };
                                    cell.s.alignment = { horizontal: C === 0 || C === 3 ? "center" : "left", vertical: "center" };
                                } else if (R === currentGroupRow) {
                                    cell.s.fill = { fgColor: { rgb: "DBEAFE" } };
                                    cell.s.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "1E3A8A" } };
                                    cell.s.alignment = { horizontal: "left", vertical: "center" };
                                } else {
                                    if (C === 0 || C === 1) {
                                        cell.s.alignment = { horizontal: "center" };
                                    }
                                }
                            }
                        }

                        // Dynamic column widths for Excel
                        const colWidths: Array<{ wch: number }> = [{ wch: 8 }]; // S.No
                        activeCols.forEach((c: any) => {
                            let maxLen = (c.label || "").length;
                            groupedData.forEach((grp: any) => {
                                (grp.items || []).slice(0, 100).forEach((item: any) => {
                                    const l = String(item[c.key] ?? "").length;
                                    if (l > maxLen) maxLen = l;
                                });
                            });
                            colWidths.push({ wch: Math.min(Math.max(maxLen + 3, 10), 38) });
                        });
                        ws["!cols"] = colWidths;
                        const safeTitle = `${cfg.reportName.replace(/[^a-zA-Z0-9]/g, "")}`.substring(0, 31);
                        XLSX.utils.book_append_sheet(wb, ws, safeTitle);
                    }
                    // 8. STOCK DIFFERENCE REPORT (Excel)
                    else if (data.type === "stockDifference") {
                        const activeCols: any[] = data.activeCols || [];
                        const groupedData: any[] = data.groupedData || [];

                        const headers = ["S.No", "Stock Group", ...activeCols.map((c: any) => c.label)];
                        const exportAoa: any[][] = [];
                        exportAoa.push(headers);

                        let serial = 1;
                        const groupRowIndices: number[] = [];

                        groupedData.forEach((grp: any) => {
                            const groupRowIndex = exportAoa.length;
                            groupRowIndices.push(groupRowIndex);

                            const groupRow: any[] = [
                                "",
                                grp.groupName,
                            ];

                            activeCols.forEach((c: any) => {
                                if (/(?:System_Qty|Actual_Qty|Difference_Qty)/i.test(c.key)) {
                                    const sum = (grp.rows || []).reduce(
                                        (s: number, r: any) => s + Number(r[c.key] || 0),
                                        0
                                    );
                                    groupRow.push(Number(sum.toFixed(1)));
                                } else {
                                    groupRow.push(grp.rows.length);
                                }
                            });
                            exportAoa.push(groupRow);

                            (grp.rows || []).forEach((row: any) => {
                                const normalRow: any[] = [serial++, ""];
                                activeCols.forEach((c: any) => {
                                    normalRow.push(formatExportCellValue(row[c.key], c.key, row, false));
                                });
                                exportAoa.push(normalRow);
                            });
                        });

                        const ws = XLSX.utils.aoa_to_sheet(exportAoa.length ? exportAoa : [["No records found"]]);

                        if (exportAoa.length > 0) {
                            const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
                            const borderAll = {
                                top: { style: "thin", color: { rgb: "CBD5E1" } },
                                bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                                left: { style: "thin", color: { rgb: "CBD5E1" } },
                                right: { style: "thin", color: { rgb: "CBD5E1" } },
                            };

                            for (let R = range.s.r; R <= range.e.r; ++R) {
                                const isHeader = R === 0;
                                const isGroup = groupRowIndices.includes(R);

                                for (let C = range.s.c; C <= range.e.c; ++C) {
                                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                    const cell = ws[cellRef];
                                    if (!cell) continue;

                                    if (isHeader) {
                                        cell.s = {
                                            fill: { fgColor: { rgb: "FFFF00" } },
                                            font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } },
                                            alignment: { horizontal: C === 0 ? "center" : "left", vertical: "center" },
                                            border: borderAll,
                                        };
                                    } else if (isGroup) {
                                        cell.s = {
                                            fill: { fgColor: { rgb: "F1F5F9" } },
                                            font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "1E3A8A" } },
                                            alignment: { horizontal: C === 0 ? "center" : typeof cell.v === "number" ? "right" : "left", vertical: "center" },
                                            border: borderAll,
                                        };
                                    } else {
                                        const isNumeric = typeof cell.v === "number";
                                        cell.s = {
                                            font: { name: "Calibri", sz: 11, color: { rgb: "000000" } },
                                            alignment: {
                                                horizontal: C === 0 ? "center" : isNumeric ? "right" : "left",
                                                vertical: "center",
                                            },
                                            border: borderAll,
                                        };
                                    }
                                }
                            }

                            // Dynamic column widths
                            const colWidths: Array<{ wch: number }> = [{ wch: 7 }, { wch: 18 }]; // S.No, Stock Group
                            activeCols.forEach((c) => {
                                let maxLen = (c.label || "").length;
                                (data.rawRows || []).slice(0, 100).forEach((r: any) => {
                                    const l = String(r[c.key] ?? "").length;
                                    if (l > maxLen) maxLen = l;
                                });
                                colWidths.push({ wch: Math.min(Math.max(maxLen + 3, 10), 38) });
                            });
                            ws["!cols"] = colWidths;
                        }

                        const safeTitle = `${cfg.reportName.replace(/[^a-zA-Z0-9]/g, "")}`.substring(0, 31);
                        XLSX.utils.book_append_sheet(wb, ws, safeTitle);
                    }
                    // 9. STANDARD TABLE ROWS (Supports Template Columns & Yellow Header Box Layout)
                    else {
                        const rawRows = data.rows || [];
                        const templateCols = data.templateCols || [];

                        // Determine columns to display
                        let colsToUse: Array<{ key: string; label: string }> = [];
                        if (templateCols.length > 0) {
                            colsToUse = templateCols.map((c: any) => ({
                                key: c.key,
                                label: c.label || c.key,
                            }));
                        } else if (rawRows.length > 0) {
                            const keys = Object.keys(rawRows[0]).slice(0, 15);
                            colsToUse = keys.map((k) => ({
                                key: k,
                                label: k.replace(/_/g, " ").toUpperCase(),
                            }));
                        }

                        // Standard font size 12pt for headers and readable font size for data
                        const headerFontSize = 12;
                        const fontSize = 11;

                        const exportAoa: any[][] = [];
                        const headerRow = ["S.No", ...colsToUse.map((c) => c.label)];
                        exportAoa.push(headerRow);

                        rawRows.forEach((r: any, idx: number) => {
                            const row = [
                                idx + 1,
                                ...colsToUse.map((c) => {
                                    return formatExportCellValue(r[c.key], c.key, r, false);
                                }),
                            ];
                            exportAoa.push(row);
                        });

                        const ws = XLSX.utils.aoa_to_sheet(exportAoa.length ? exportAoa : [["No records found"]]);

                        // Apply yellow header, solid black borders, and alignments
                        if (exportAoa.length > 0) {
                            const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
                            const borderAll = {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } },
                            };

                            for (let R = range.s.r; R <= range.e.r; ++R) {
                                for (let C = range.s.c; C <= range.e.c; ++C) {
                                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                    const cell = ws[cellRef];
                                    if (!cell) continue;

                                    if (R === 0) {
                                        // Header row: Yellow background (#FFFF00), black bold font, bordered box
                                        cell.s = {
                                            fill: { fgColor: { rgb: "FFFF00" } },
                                            font: { name: "Calibri", sz: headerFontSize, bold: true, color: { rgb: "000000" } },
                                            alignment: { horizontal: C === 0 ? "center" : "center", vertical: "center" },
                                            border: borderAll,
                                        };
                                    } else {
                                        // Data row: bordered cell, clean alignment
                                        const isNumeric = typeof cell.v === "number";
                                        cell.s = {
                                            font: { name: "Calibri", sz: fontSize, color: { rgb: "000000" } },
                                            alignment: {
                                                horizontal: C === 0 ? "center" : isNumeric ? "right" : "left",
                                                vertical: "center",
                                            },
                                            border: borderAll,
                                        };
                                    }
                                }
                            }

                            // Dynamic column widths
                            const colWidths: Array<{ wch: number }> = [{ wch: 7 }]; // S.No
                            colsToUse.forEach((c) => {
                                let maxLen = (c.label || "").length;
                                rawRows.slice(0, 100).forEach((r: any) => {
                                    const l = String(r[c.key] ?? "").length;
                                    if (l > maxLen) maxLen = l;
                                });
                                colWidths.push({ wch: Math.min(Math.max(maxLen + 3, 10), 38) });
                            });
                            ws["!cols"] = colWidths;
                        }

                        const safeTitle = `${cfg.reportName.replace(/[^a-zA-Z0-9]/g, "")}_${mode}`.substring(0, 31);
                        XLSX.utils.book_append_sheet(wb, ws, safeTitle);
                    }
                }
            }

            setExportStatusText("Finalizing Excel file...");
            XLSX.writeFile(wb, `Consolidated_Report_${dayjs().format("YYYYMMDD_HHmm")}.xlsx`);
            toast.success("Single Excel Workbook downloaded successfully! 🎉");
        } catch (err) {
            console.error("Batch Excel Export error:", err);
            toast.error("Failed to generate consolidated Excel export.");
        } finally {
            setExporting(false);
            setExportStatusText("");
        }
    };

    /* ================= BATCH EXPORT PDF ================= */

    const handleExportSinglePDF = async () => {
        if (!selectedReports.length) {
            toast.warning("Please select at least one report to export.");
            return;
        }

        try {
            setExporting(true);
            setExportStatusText("Initializing PDF Document & Loading Tamil Font...");

            const fontBase64 = await loadNirmalaFont();
            const doc = new jsPDF("landscape", "mm", "a4");

            if (fontBase64) {
                doc.addFileToVFS("Nirmala.ttf", fontBase64);
                doc.addFont("Nirmala.ttf", "Nirmala", "normal");
                doc.addFont("Nirmala.ttf", "Nirmala", "bold");
                doc.setFont("Nirmala", "normal");
            }
            const activeFontName = fontBase64 ? "Nirmala" : "helvetica";

            let isFirstPage = true;

            for (let i = 0; i < selectedReports.length; i++) {
                const cfg = selectedReports[i];
                const item = availableMenuReports.find((c) =>
                    c.key === cfg.reportKey ||
                    c.path.toLowerCase().replace(/[\/\-_]/g, "") === cfg.reportKey.toLowerCase().replace(/[\/\-_]/g, "") ||
                    c.name.toLowerCase().replace(/[\s\-_]/g, "") === cfg.reportKey.toLowerCase().replace(/[\s\-_]/g, "")
                );
                const cat = item?.definition;
                if (!cat) continue;

                for (const mode of cfg.selectedModes) {
                    const templateName = cfg.selectedTemplateId !== "default"
                        ? cfg.availableTemplates.find((t) => String(t.id) === String(cfg.selectedTemplateId))?.name
                        : null;
                    const fullReportTitle = templateName
                        ? `${cfg.reportName} - ${mode} - ${templateName}`
                        : `${cfg.reportName} - ${mode}`;

                    setExportStatusText(`Fetching & Rendering ${fullReportTitle}...`);

                    if (!isFirstPage) {
                        doc.addPage();
                    }
                    isFirstPage = false;

                    const data = await cat.fetchData({
                        from: fromDate,
                        to: toDate,
                        mode,
                        templateId: cfg.selectedTemplateId
                    });

                    // 1. EXPENSE REPORT
                    if (data.type === "expense") {
                        const isDirect = mode === "Direct";
                        const sectionData = data.mapped?.[isDirect ? "Direct" : "Indirect"] || data.raw?.[isDirect ? "Direct" : "Indirect"] || { groups: [] };

                        if (cfg.selectedExportType === "Management") {
                            // Yellow Management View matching ExpensesReport.tsx
                            const formatDisplayDate = (dStr: string) => {
                                if (!dStr) return "";
                                const d = dayjs(dStr);
                                return d.isValid() ? d.format("DD.MM.YYYY") : dStr;
                            };
                            const dateHeader = fromDate === toDate
                                ? formatDisplayDate(toDate)
                                : `${formatDisplayDate(fromDate)} - ${formatDisplayDate(toDate)}`;
                            const titleText = fullReportTitle;

                            const bodyRows: any[] = [];
                            let grandTotal = 0;

                            sectionData.groups.forEach((g: any) => {
                                bodyRows.push([
                                    { content: g.name.toUpperCase(), styles: { fontStyle: "bold", fontSize: 9.5, textColor: [0, 0, 0] } },
                                    { content: "", styles: { halign: "right", textColor: [0, 0, 0] } }
                                ]);
                                (g.subGroups || []).forEach((sg: any) => {
                                    const b = Number(sg.balance || 0);
                                    grandTotal += b;
                                    bodyRows.push([
                                        { content: `    ${sg.name}`, styles: { fontSize: 9, textColor: [0, 0, 0] } },
                                        { content: b.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }), styles: { halign: "right", fontSize: 9, textColor: [0, 0, 0] } }
                                    ]);
                                });
                            });

                            autoTable(doc, {
                                startY: 15,
                                margin: { left: 20, right: 20 },
                                theme: "grid",
                                styles: {
                                    font: activeFontName,
                                    lineColor: [0, 0, 0],
                                    lineWidth: 0.2,
                                    cellPadding: 2,
                                    textColor: [0, 0, 0]
                                },
                                head: [
                                    [
                                        {
                                            content: titleText,
                                            colSpan: 2,
                                            styles: { font: activeFontName, halign: "center", fontStyle: "bold", fontSize: 12, fillColor: [255, 255, 0], textColor: [0, 0, 0] }
                                        }
                                    ],
                                    [
                                        {
                                            content: dateHeader,
                                            colSpan: 2,
                                            styles: { font: activeFontName, halign: "center", fontStyle: "bold", fontSize: 10, fillColor: [255, 255, 0], textColor: [0, 0, 0] }
                                        }
                                    ],
                                    [
                                        { content: "Particulars", styles: { font: activeFontName, halign: "center", fontStyle: "bold", fontSize: 9.5, fillColor: [255, 255, 255], textColor: [0, 0, 0] } },
                                        { content: "Amount", styles: { font: activeFontName, halign: "center", fontStyle: "bold", fontSize: 9.5, fillColor: [255, 255, 255], textColor: [0, 0, 0] } }
                                    ]
                                ],
                                body: bodyRows,
                                foot: [
                                    [
                                        { content: "Grand Total", styles: { font: activeFontName, halign: "center", fontStyle: "bold", fontSize: 11, fillColor: [255, 255, 0], textColor: [0, 0, 0] } },
                                        { content: grandTotal.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }), styles: { font: activeFontName, halign: "right", fontStyle: "bold", fontSize: 11, fillColor: [255, 255, 0], textColor: [0, 0, 0] } }
                                    ]
                                ],
                                columnStyles: {
                                    0: { cellWidth: 160 },
                                    1: { cellWidth: 80 }
                                }
                            });
                        } else {
                            // Normal PDF format
                            doc.setFillColor(30, 58, 138);
                            doc.rect(0, 0, 297, 18, "F");
                            doc.setTextColor(255, 255, 255);
                            doc.setFontSize(13);
                            doc.text(fullReportTitle, 14, 12);
                            doc.setFontSize(9);
                            doc.text(`Period: ${fromDate} to ${toDate}`, 210, 12);

                            const flat: any[] = [];
                            sectionData.groups.forEach((g: any) => {
                                (g.subGroups || []).forEach((sg: any) => {
                                    (sg.ledgers || []).forEach((l: any, idx: number) => {
                                        flat.push([
                                            idx + 1,
                                            g.name,
                                            sg.name,
                                            l.payment_date ? dayjs(l.payment_date).format("DD-MM-YYYY") : "",
                                            l.voucher_name || "",
                                            l.entryType === "DR" ? formatINR(l.amount) : "-",
                                            l.entryType === "CR" ? formatINR(l.amount) : "-"
                                        ]);
                                    });
                                });
                            });

                            autoTable(doc, {
                                startY: 25,
                                head: [["S.No", "Group", "Account", "Date", "Voucher", "Debit", "Credit"]],
                                body: flat,
                                theme: "grid",
                                styles: {
                                    font: activeFontName,
                                    fontSize: 8,
                                    cellPadding: 2,
                                    lineColor: [0, 0, 0],
                                    lineWidth: 0.2,
                                    textColor: [0, 0, 0]
                                },
                                headStyles: {
                                    font: activeFontName,
                                    fillColor: [255, 255, 0],
                                    textColor: [0, 0, 0],
                                    fontStyle: "bold",
                                    lineColor: [0, 0, 0],
                                    lineWidth: 0.2
                                },
                                didParseCell: (cellData) => {
                                    if (cellData.column.index === 0) {
                                        cellData.cell.styles.halign = "center";
                                    } else if (cellData.column.index === 5 || cellData.column.index === 6) {
                                        cellData.cell.styles.halign = "right";
                                    }
                                }
                            });
                        }
                    }
                    // 2. RATE MASTER ADMIN
                    else if (data.type === "rateMasterAdmin") {
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(fullReportTitle, 14, 12);
                        doc.setFontSize(9);
                        doc.text(`As of Date: ${toDate}`, 220, 12);

                        const rawRows = data.rows || [];
                        const templateCols = data.templateCols || [];
                        const groupingCols = data.groupingCols || [];

                        // Enabled columns projection matching user desired table format
                        const enabledCols = templateCols.length > 0
                            ? templateCols
                            : [
                                { key: "Product_Name", label: "Product" },
                                { key: "Min_Rate", label: "Min Rate", isNumeric: true },
                                { key: "Rate", label: "List Rate", isNumeric: true },
                                { key: "Margin", label: "Margin", isNumeric: true },
                                { key: "Margin_Percentage", label: "%", isNumeric: true }
                            ];

                        const headers = ["S.No", ...enabledCols.map((c: any) => c.label)];
                        let serial = 1;
                        const body: any[] = [];
                        const groupRowIndices: number[] = [];

                        if (groupingCols.length > 0) {
                            const grpKey = groupingCols[0];
                            const groupMap = new Map<string, any[]>();
                            rawRows.forEach((r: any) => {
                                const val = String(r[grpKey] ?? "Others");
                                if (!groupMap.has(val)) groupMap.set(val, []);
                                groupMap.get(val)!.push(r);
                            });

                            groupMap.forEach((grpRows, val) => {
                                groupRowIndices.push(body.length);
                                const groupRow = [
                                    "",
                                    `Group: ${grpKey} = ${val}`,
                                    ...enabledCols.slice(1).map((c: any) => {
                                        if (c.isNumeric || c.key === "Rate" || c.key === "Min_Rate" || c.key === "Max_Rate" || c.key === "COGS" || c.key === "Margin") {
                                            const sum = grpRows.reduce((s: number, r: any) => s + (Number(r[c.key]) || 0), 0);
                                            const finalNum = c.key === "GP_Percentage_COGS" || c.key === "Margin_Percentage"
                                                ? (sum / grpRows.length)
                                                : sum;
                                            return finalNum.toFixed(1);
                                        }
                                        return "";
                                    })
                                ];
                                body.push(groupRow);

                                grpRows.forEach((r: any) => {
                                    body.push([
                                        String(serial++),
                                        ...enabledCols.map((c: any) => {
                                            return formatExportCellValue(r[c.key], c.key, r, true);
                                        })
                                    ]);
                                });
                            });
                        } else {
                            rawRows.forEach((r: any) => {
                                body.push([
                                    String(serial++),
                                    ...enabledCols.map((c: any) => {
                                        return formatExportCellValue(r[c.key], c.key, r, true);
                                    })
                                ]);
                            });
                        }

                        const { pdfFontSize, cellPad, columnStyles } = getDynamicTableStyles(headers, enabledCols, 277);

                        autoTable(doc, {
                            startY: 25,
                            head: [headers],
                            body,
                            theme: "grid",
                            margin: { left: 10, right: 10 },
                            tableWidth: 277,
                            styles: {
                                font: activeFontName,
                                fontSize: pdfFontSize,
                                cellPadding: cellPad,
                                overflow: "linebreak",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0], // Yellow Header (#FFFF00)
                                textColor: [0, 0, 0], // Black Bold text
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            },
                            columnStyles,
                            didParseCell: (cellData) => {
                                if (groupRowIndices.includes(cellData.row.index)) {
                                    cellData.cell.styles.fillColor = [219, 234, 254];
                                    cellData.cell.styles.textColor = [30, 58, 138];
                                    cellData.cell.styles.fontStyle = "bold";
                                }
                            }
                        });
                    }
                    // 3. RECIEVABLE PAYABLE
                    else if (data.type === "recievablePayable") {
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(`${fullReportTitle} - As of ${dayjs(toDate).format("DD/MM/YYYY")}`, 14, 12);

                        const rows = data.rows || [];
                        const groupMap = new Map<string, any[]>();
                        rows.forEach((r: any) => {
                            const grp = r.Group_Name || "No Group";
                            if (!groupMap.has(grp)) groupMap.set(grp, []);
                            groupMap.get(grp)!.push(r);
                        });

                        const tableHeaders = [["S.No", "Account Name", "Invoice Date", "Invoice No", "Recievable", "Payable"]];
                        const tableBody: any[] = [];
                        let totalDebit = 0;
                        let totalCredit = 0;
                        let sNo = 1;
                        const headerIndices = new Set<number>();

                        groupMap.forEach((items, grpName) => {
                            let subDeb = 0;
                            let subCred = 0;
                            items.forEach((it: any) => {
                                subDeb += Number(it.Debit_Amount) || 0;
                                subCred += Number(it.Credit_Amount) || 0;
                            });
                            totalDebit += subDeb;
                            totalCredit += subCred;

                            headerIndices.add(tableBody.length);
                            tableBody.push([
                                "",
                                grpName,
                                "",
                                "",
                                subDeb !== 0 ? subDeb.toFixed(2) : "0.00",
                                subCred !== 0 ? subCred.toFixed(2) : "0.00"
                            ]);

                            items.forEach((it: any) => {
                                tableBody.push([
                                    sNo++,
                                    it.Account_name || "",
                                    it.invoice_date ? dayjs(it.invoice_date).format("DD/MM/YYYY") : "",
                                    it.invoice_no || "",
                                    it.Debit_Amount ? Number(it.Debit_Amount).toFixed(2) : "0.00",
                                    it.Credit_Amount ? Number(it.Credit_Amount).toFixed(2) : "0.00"
                                ]);
                            });
                        });

                        const bal = data.isReceivable ? totalDebit - totalCredit : totalCredit - totalDebit;
                        tableBody.push([
                            "",
                            `Grand Total (Bal: ${formatINR(bal)})`,
                            "",
                            "",
                            totalDebit.toFixed(2),
                            totalCredit.toFixed(2)
                        ]);

                        autoTable(doc, {
                            startY: 25,
                            head: tableHeaders,
                            body: tableBody,
                            theme: "grid",
                            styles: {
                                font: activeFontName,
                                fontSize: 8,
                                cellPadding: 2,
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0],
                                textColor: [0, 0, 0],
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            },
                            didParseCell: (cellData) => {
                                if (cellData.row.index === tableBody.length - 1) {
                                    cellData.cell.styles.fontStyle = "bold";
                                    cellData.cell.styles.fillColor = [241, 245, 249];
                                    if (cellData.column.index === 4 || cellData.column.index === 5) cellData.cell.styles.halign = "right";
                                } else if (headerIndices.has(cellData.row.index)) {
                                    cellData.cell.styles.fontStyle = "bold";
                                    cellData.cell.styles.fillColor = [226, 232, 240];
                                    cellData.cell.styles.textColor = [30, 58, 138];
                                    if (cellData.column.index === 4 || cellData.column.index === 5) cellData.cell.styles.halign = "right";
                                } else if (cellData.column.index === 4 || cellData.column.index === 5) {
                                    cellData.cell.styles.halign = "right";
                                }
                                if (cellData.column.index === 0) {
                                    cellData.cell.styles.halign = "center";
                                }
                            }
                        });
                    }
                    // 4. CASH BOX ABSTRACT
                    else if (data.type === "cashboxAbstract") {
                        const parsed = data.parsed;
                        const dateStr = fromDate === toDate
                            ? dayjs(fromDate).format("DD-MM-YYYY")
                            : `${dayjs(fromDate).format("DD-MM-YYYY")} TO ${dayjs(toDate).format("DD-MM-YYYY")}`;

                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(`${fullReportTitle} - ${dateStr}`, 14, 12);

                        const formatNum = (v: number) => (v === 0 ? "-" : formatINR(v));
                        const pdfBody: any[][] = [];
                        pdfBody.push(["", "", "Opening Balance", formatNum(parsed?.opening ?? 0)]);

                        const debitGroups = parsed?.debitGroups || [];
                        const creditGroups = parsed?.creditGroups || [];
                        const maxGrp = Math.max(debitGroups.length, creditGroups.length);

                        for (let idx = 0; idx < maxGrp; idx++) {
                            const leftGroup = debitGroups[idx] || { label: "", total: 0, subLedgers: [] };
                            const rightGroup = creditGroups[idx] || { label: "", total: 0, subLedgers: [] };

                            pdfBody.push([
                                leftGroup.label,
                                leftGroup.total ? formatNum(leftGroup.total) : "-",
                                rightGroup.label,
                                rightGroup.total ? formatNum(rightGroup.total) : "-"
                            ]);

                            const maxSubRows = Math.max(leftGroup.subLedgers?.length || 0, rightGroup.subLedgers?.length || 0);
                            for (let sIdx = 0; sIdx < maxSubRows; sIdx++) {
                                const leftSub = leftGroup.subLedgers?.[sIdx];
                                const rightSub = rightGroup.subLedgers?.[sIdx];
                                pdfBody.push([
                                    leftSub ? `  ${sIdx + 1}. ${leftSub.name}` : "",
                                    leftSub ? formatNum(leftSub.amount) : "",
                                    rightSub ? `  ${sIdx + 1}. ${rightSub.name}` : "",
                                    rightSub ? formatNum(rightSub.amount) : ""
                                ]);
                            }
                        }

                        pdfBody.push(["", "", "Closing Balance", formatNum(parsed?.closing ?? 0)]);

                        autoTable(doc, {
                            startY: 25,
                            head: [["Particulars (Debit)", "Debit Amt", "Particulars (Credit)", "Credit amt"]],
                            body: pdfBody,
                            styles: {
                                font: activeFontName,
                                fontSize: 7.5,
                                cellPadding: 1.5,
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0],
                                textColor: [0, 0, 0],
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            },
                            theme: "grid",
                            didParseCell: (cellData) => {
                                const valL = cellData.row.cells[0]?.text?.[0] || "";
                                const valR = cellData.row.cells[2]?.text?.[0] || "";
                                const isOpeningOrClosing = valL === "Opening Balance" || valR === "Opening Balance" || valL === "Closing Balance" || valR === "Closing Balance";
                                const isGroupHeader = CASHBOX_DEBIT_GROUPS.some(g => g.label === valL) || CASHBOX_CREDIT_GROUPS.some(g => g.label === valR);

                                if (isOpeningOrClosing) {
                                    cellData.cell.styles.fillColor = [238, 238, 238];
                                    cellData.cell.styles.fontStyle = "bold";
                                    cellData.cell.styles.textColor = [180, 83, 9];
                                } else if (isGroupHeader) {
                                    cellData.cell.styles.fillColor = [226, 232, 240];
                                    cellData.cell.styles.fontStyle = "bold";
                                    cellData.cell.styles.textColor = [30, 58, 138];
                                }
                            }
                        });
                    }
                    // 5. CASH BOX EXPANDED
                    else if (data.type === "cashboxExpanded") {
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(`${fullReportTitle} - ${fromDate} to ${toDate}`, 14, 12);

                        const rows = data.detailedData?.Data || [];
                        const keys = rows.length > 0 ? Object.keys(rows[0]).slice(0, 8) : [];
                        const head = [keys.map(k => k.replace(/_/g, " ").toUpperCase())];
                        const body = rows.map((r: any) => keys.map(k => String(r[k] ?? "")));

                        autoTable(doc, {
                            startY: 25,
                            head,
                            body,
                            theme: "grid",
                            styles: {
                                font: activeFontName,
                                fontSize: 7,
                                cellPadding: 1.5,
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0],
                                textColor: [0, 0, 0],
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            }
                        });
                    }
                    // 6. OUTSTANDING MANAGEMENT / RECIEVABLE PAYABLE MGT
                    else if (data.type === "outstandingManagement") {
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(`${fullReportTitle} - As of ${dayjs(toDate).format("DD/MM/YYYY")}`, 14, 12);

                        const formatNumber = (num: number) => {
                            return new Intl.NumberFormat("en-IN", {
                                maximumFractionDigits: 2,
                                minimumFractionDigits: 0,
                            }).format(num || 0);
                        };

                        const tableBody: any[] = [];
                        tableBody.push(["Total", "", formatNumber(data.grandTotalAmount)]);

                        const flattenedRows = data.flattenedRows || [];
                        flattenedRows.forEach((row: any) => {
                            if (row.__group) {
                                tableBody.push([row.__label, "", formatNumber(row.__subtotal)]);
                            } else {
                                const accDisplay = row.invoice_count && row.invoice_count > 1
                                    ? `${row.Account_name} (${row.invoice_count})`
                                    : (row.Account_name || "-");
                                tableBody.push([
                                    "    " + accDisplay,
                                    row.days !== undefined ? String(row.days) : "-",
                                    formatNumber(row.Bal_Amount),
                                ]);
                            }
                        });

                        autoTable(doc, {
                            startY: 18,
                            head: [["Account Name", "Days", data.amountColumnTitle]],
                            body: tableBody,
                            theme: "grid",
                            styles: {
                                font: activeFontName,
                                fontSize: 8.5,
                                cellPadding: 2,
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0],
                                textColor: [0, 0, 0],
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            },
                            didParseCell: (cellData) => {
                                if (cellData.row.index === 0) {
                                    cellData.cell.styles.fillColor = [239, 246, 255];
                                    cellData.cell.styles.textColor = [30, 58, 138];
                                    cellData.cell.styles.fontStyle = "bold";
                                } else {
                                    const rowData = flattenedRows[cellData.row.index - 1];
                                    if (rowData && rowData.__group) {
                                        cellData.cell.styles.fillColor = [241, 245, 249];
                                        cellData.cell.styles.textColor = [15, 23, 42];
                                        cellData.cell.styles.fontStyle = "bold";
                                    }
                                }
                                if (cellData.column.index === 1) cellData.cell.styles.halign = "center";
                                if (cellData.column.index === 2) cellData.cell.styles.halign = "right";
                            },
                        });
                    }
                    // 7. SHET SHEET REPORT (PDF)
                    else if (data.type === "shetSheet") {
                        const activeCols: any[] = data.activeCols || [];
                        const groupedData: any[] = data.groupedData || [];

                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(fullReportTitle, 14, 12);
                        doc.setFontSize(9);
                        doc.text(`Period: ${fromDate} to ${toDate}`, 210, 12);

                        const headRow = ["S.No", ...activeCols.map((c: any) => c.label)];
                        const tableHead = [headRow];
                        const tableBody: any[] = [];
                        const groupRowIndices: number[] = [];

                        groupedData.forEach((grp: any) => {
                            groupRowIndices.push(tableBody.length);
                            tableBody.push([`${grp.groupName} (Total Bags: ${grp.totalBags || 0})`, ...activeCols.map(() => "")]);
                            grp.items.forEach((item: any, idx: number) => {
                                tableBody.push([
                                    idx + 1,
                                    ...activeCols.map((c: any) => {
                                        return formatExportCellValue(item[c.key], c.key, item, true);
                                    }),
                                ]);
                            });
                        });

                        const { pdfFontSize, cellPad, columnStyles } = getDynamicTableStyles(headRow, activeCols, 277);

                        autoTable(doc, {
                            startY: 22,
                            head: tableHead,
                            body: tableBody,
                            theme: "grid",
                            margin: { left: 10, right: 10 },
                            tableWidth: 277,
                            styles: {
                                font: activeFontName,
                                fontSize: pdfFontSize,
                                cellPadding: cellPad,
                                overflow: "linebreak",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.15,
                                textColor: [0, 0, 0]
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [255, 255, 0],
                                textColor: [0, 0, 0],
                                fontStyle: "bold",
                                lineColor: [0, 0, 0],
                                lineWidth: 0.2
                            },
                            columnStyles,
                            didParseCell: (cellData) => {
                                if (groupRowIndices.includes(cellData.row.index)) {
                                    cellData.cell.styles.fillColor = [219, 234, 254];
                                    cellData.cell.styles.textColor = [30, 58, 138];
                                    cellData.cell.styles.fontStyle = "bold";
                                }
                            },
                        });
                    }
                    // 8. STOCK DIFFERENCE REPORT (PDF)
                    else if (data.type === "stockDifference") {
                        const activeCols: any[] = data.activeCols || [];
                        const groupedData: any[] = data.groupedData || [];

                        // Header banner matching parent screen theme
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(fullReportTitle, 14, 12);
                        doc.setFontSize(9);
                        doc.text(`Period: ${fromDate} to ${toDate}`, 210, 12);

                        const headRow = ["S.No", "Stock Group", ...activeCols.map((c: any) => c.label)];
                        const tableHead = [headRow];
                        const tableBody: any[] = [];
                        const groupRowIndices: number[] = [];

                        let serial = 1;
                        groupedData.forEach((grp: any) => {
                            groupRowIndices.push(tableBody.length);

                            const groupRow: any[] = [
                                "",
                                grp.groupName,
                                ...activeCols.map((c: any) => {
                                    if (/(?:System_Qty|Actual_Qty|Difference_Qty)/i.test(c.key)) {
                                        const sum = (grp.rows || []).reduce(
                                            (s: number, r: any) => s + Number(r[c.key] || 0),
                                            0
                                        );
                                        return sum.toFixed(1);
                                    }
                                    return String((grp.rows || []).length);
                                }),
                            ];
                            tableBody.push(groupRow);

                            (grp.rows || []).forEach((row: any) => {
                                tableBody.push([
                                    serial++,
                                    "",
                                    ...activeCols.map((c: any) => {
                                        return formatExportCellValue(row[c.key], c.key, row, true);
                                    }),
                                ]);
                            });
                        });

                        // Calculate dynamic font size and column widths based on column count & printable width
                        const allColConfigs = [{ key: "Stock_Group", label: "Stock Group" }, ...activeCols];
                        const { pdfFontSize, cellPad, columnStyles } = getDynamicTableStyles(headRow, allColConfigs, 277);

                        autoTable(doc, {
                            startY: 22,
                            head: tableHead,
                            body: tableBody,
                            theme: "grid",
                            margin: { left: 10, right: 10 },
                            tableWidth: 277,
                            styles: {
                                font: activeFontName,
                                fontSize: pdfFontSize,
                                cellPadding: cellPad,
                                overflow: "linebreak",
                                lineColor: [203, 213, 225],
                                lineWidth: 0.15,
                                textColor: [0, 0, 0],
                            },
                            headStyles: {
                                font: activeFontName,
                                fillColor: [30, 58, 138], // Parent report header theme (#1E3A8A)
                                textColor: [255, 255, 255], // White text
                                fontStyle: "bold",
                                lineColor: [36, 72, 178],
                                lineWidth: 0.2,
                            },
                            columnStyles,
                            didParseCell: (cellData) => {
                                if (groupRowIndices.includes(cellData.row.index)) {
                                    cellData.cell.styles.fillColor = [241, 245, 249]; // Light slate group row (#F1F5F9)
                                    cellData.cell.styles.textColor = [15, 23, 42]; // Dark text
                                    cellData.cell.styles.fontStyle = "bold";
                                }
                            },
                        });
                    }
                    // 9. STANDARD TABLE ROWS (Supports Template Columns & Yellow Header Box Layout)
                    else {
                        doc.setFillColor(30, 58, 138);
                        doc.rect(0, 0, 297, 18, "F");
                        doc.setTextColor(255, 255, 255);
                        doc.setFontSize(13);
                        doc.text(fullReportTitle, 14, 12);
                        doc.setFontSize(9);
                        doc.text(`Period: ${fromDate} to ${toDate}`, 210, 12);

                        const rawRows = data.rows || [];
                        const templateCols = data.templateCols || [];

                        if (rawRows.length === 0) {
                            doc.setTextColor(100, 100, 100);
                            doc.text("No records found for this period.", 14, 35);
                        } else {
                            let colsToUse: Array<{ key: string; label: string }> = [];
                            if (templateCols.length > 0) {
                                colsToUse = templateCols.map((c: any) => ({
                                    key: c.key,
                                    label: c.label || c.key,
                                }));
                            } else {
                                const keys = Object.keys(rawRows[0]).slice(0, 15);
                                colsToUse = keys.map((k) => ({
                                    key: k,
                                    label: k.replace(/_/g, " ").toUpperCase(),
                                }));
                            }

                            const headRow = ["S.No", ...colsToUse.map((c) => c.label)];
                            const head = [headRow];
                            const { pdfFontSize, cellPad, columnStyles } = getDynamicTableStyles(headRow, colsToUse, 277);

                            const body = rawRows.map((r: any, idx: number) => [
                                idx + 1,
                                ...colsToUse.map((c) => {
                                    return formatExportCellValue(r[c.key], c.key, r, true);
                                }),
                            ]);

                            autoTable(doc, {
                                startY: 25,
                                head,
                                body,
                                theme: "grid",
                                margin: { left: 10, right: 10 },
                                tableWidth: 277,
                                styles: {
                                    font: activeFontName,
                                    fontSize: pdfFontSize,
                                    cellPadding: cellPad,
                                    overflow: "linebreak",
                                    lineColor: [0, 0, 0],
                                    lineWidth: 0.15,
                                    textColor: [0, 0, 0],
                                },
                                headStyles: {
                                    font: activeFontName,
                                    fillColor: [255, 255, 0], // Yellow Header (#FFFF00)
                                    textColor: [0, 0, 0], // Black text
                                    fontStyle: "bold",
                                    lineColor: [0, 0, 0],
                                    lineWidth: 0.2,
                                },
                                columnStyles,
                            });
                        }
                    }
                }
            }

            setExportStatusText("Saving PDF...");
            doc.save(`Consolidated_Report_${dayjs().format("YYYYMMDD_HHmm")}.pdf`);
            toast.success("Single PDF Document downloaded successfully! 🎉");
        } catch (err) {
            console.error("Batch PDF Export error:", err);
            toast.error("Failed to generate consolidated PDF export.");
        } finally {
            setExporting(false);
            setExportStatusText("");
        }
    };

    return (
        <Box sx={{ height: "100vh", bgcolor: "#f8fafc", display: "flex", flexDirection: "column", overflow: "hidden" }}>
            <Header />

            <Box sx={{ p: { xs: 1.5, md: 3 }, maxWidth: 1200, margin: "0 auto", width: "100%", flex: 1, display: "flex", flexDirection: "column", minHeight: 0, overflow: "hidden" }}>
                {/* Header Title Section */}
                <Box display="flex" justifyContent="space-between" alignItems="center" flexWrap="wrap" gap={2} mb={3}>
                    <Box>
                        <Typography variant="h5" fontWeight={700} color="#1E3A8A" display="flex" alignItems="center" gap={1}>
                            <LayersIcon /> Consolidated Batch Export
                        </Typography>
                        <Typography variant="body2" color="text.secondary">
                            Bundle multiple reports with customized modes, templates, and export types into a single PDF or Excel file.
                        </Typography>
                    </Box>

                    {/* Action Buttons */}
                    <Box display="flex" gap={1.5} alignItems="center">
                        {/* Save / Update Preset Button */}
                        <Button
                            variant="outlined"
                            startIcon={<SaveIcon />}
                            disabled={selectedReports.length === 0}
                            onClick={handleOpenSavePresetDialog}
                            sx={{ textTransform: "none", fontWeight: 600, borderColor: "#1E3A8A", color: "#1E3A8A" }}
                        >
                            {selectedPresetId ? "Update Preset" : "Save as Preset"}
                        </Button>

                        <Button
                            variant="contained"
                            color="success"
                            startIcon={<TableChartIcon />}
                            disabled={exporting || selectedReports.length === 0}
                            onClick={handleExportSingleExcel}
                            sx={{ textTransform: "none", fontWeight: 700 }}
                        >
                            Export Single Excel
                        </Button>

                        <Button
                            variant="contained"
                            color="error"
                            startIcon={<PictureAsPdfIcon />}
                            disabled={exporting || selectedReports.length === 0}
                            onClick={handleExportSinglePDF}
                            sx={{ textTransform: "none", fontWeight: 700 }}
                        >
                            Export Single PDF
                        </Button>
                    </Box>
                </Box>

                {/* Progress Alert */}
                {exporting && (
                    <Alert severity="info" icon={<CircularProgress size={20} />} sx={{ mb: 3 }}>
                        {exportStatusText}
                    </Alert>
                )}

                {/* Controls Bar: Preset Selector + Date Pickers + Add Report Dropdown */}
                <Paper sx={{ p: 2, mb: 3, borderRadius: 2, border: "1px solid #e2e8f0" }}>
                    <Box display="grid" gridTemplateColumns={{ xs: "1fr", sm: "1fr 1fr", md: "1.2fr 1fr 1fr 1.5fr" }} gap={2} alignItems="center">
                        {/* Saved Preset Dropdown */}
                        <FormControl size="small" fullWidth disabled={loadingPreset}>
                            <InputLabel id="preset-select-label">
                                {loadingPreset ? "Loading Preset..." : "Load Saved Preset"}
                            </InputLabel>
                            <Select
                                labelId="preset-select-label"
                                label={loadingPreset ? "Loading Preset..." : "Load Saved Preset"}
                                value={selectedPresetId}
                                onChange={(e) => handleApplyPreset(e.target.value)}
                                endAdornment={
                                    loadingPreset ? (
                                        <Box sx={{ display: "flex", alignItems: "center", mr: 2 }}>
                                            <CircularProgress size={18} thickness={5} sx={{ color: "#1E3A8A" }} />
                                        </Box>
                                    ) : null
                                }
                            >
                                <MenuItem value="">
                                    <em>None (Custom Configuration)</em>
                                </MenuItem>
                                {presetTemplates.map((p) => (
                                    <MenuItem key={p.id} value={p.id}>
                                        <Box display="flex" justifyContent="space-between" width="100%" alignItems="center">
                                            <span>{p.name}</span>
                                            <IconButton
                                                size="small"
                                                onClick={(e) => {
                                                    e.stopPropagation();
                                                    handleDeletePreset(p.id);
                                                }}
                                                sx={{ ml: 1, p: 0.2 }}
                                            >
                                                <DeleteOutlineIcon fontSize="small" color="error" />
                                            </IconButton>
                                        </Box>
                                    </MenuItem>
                                ))}
                            </Select>
                        </FormControl>

                        {/* From Date */}
                        <TextField
                            size="small"
                            type="date"
                            label="From Date"
                            value={fromDate}
                            onChange={(e) => setFromDate(e.target.value)}
                            InputLabelProps={{ shrink: true }}
                            fullWidth
                        />

                        {/* To Date */}
                        <TextField
                            size="small"
                            type="date"
                            label="To Date"
                            value={toDate}
                            onChange={(e) => setToDate(e.target.value)}
                            InputLabelProps={{ shrink: true }}
                            fullWidth
                        />

                        {/* Add Report Dropdown */}
                        <FormControl size="small" fullWidth>
                            <InputLabel id="add-report-label">+ Add Report to Batch</InputLabel>
                            <Select
                                labelId="add-report-label"
                                label="+ Add Report to Batch"
                                value=""
                                onChange={(e) => handleAddReport(e.target.value as string)}
                            >
                                {availableMenuReports.map((r) => (
                                    <MenuItem key={r.key} value={r.key}>
                                        {r.name}
                                    </MenuItem>
                                ))}
                            </Select>
                        </FormControl>
                    </Box>
                </Paper>

                {/* Selected Reports Table with Dedicated Scroll Container */}
                <Box display="flex" justifyContent="space-between" alignItems="center" mb={1.5} flexShrink={0}>
                    <Typography variant="subtitle1" fontWeight={700} color="#334155">
                        Selected Reports for Export ({selectedReports.length})
                    </Typography>
                    {selectedReports.length > 0 && (
                        <Button
                            size="small"
                            color="error"
                            onClick={() => setSelectedReports([])}
                            sx={{ textTransform: "none", fontSize: "0.8rem" }}
                        >
                            Clear All
                        </Button>
                    )}
                </Box>

                {selectedReports.length === 0 ? (
                    <Paper
                        sx={{
                            p: 6,
                            textAlign: "center",
                            borderRadius: 2,
                            border: "2px dashed #cbd5e1",
                            backgroundColor: "#f8fafc",
                            flexShrink: 0
                        }}
                    >
                        {loadingPreset ? (
                            <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", py: 4 }}>
                                <CircularProgress size={36} sx={{ color: "#1E3A8A", mb: 2 }} />
                                <Typography fontWeight={600} color="#1E3A8A">
                                    Loading preset configuration & templates...
                                </Typography>
                            </Box>
                        ) : (
                            <>
                                <TuneIcon sx={{ fontSize: 48, color: "#94a3b8", mb: 1 }} />
                                <Typography fontWeight={600} color="#64748b">
                                    No reports added yet
                                </Typography>
                                <Typography variant="body2" color="#94a3b8" mb={2}>
                                    Use the dropdown above to choose reports, configure modes, templates, and export formats.
                                </Typography>
                            </>
                        )}
                    </Paper>
                ) : (
                    <Box sx={{ position: "relative", flex: 1, minHeight: 0, display: "flex", flexDirection: "column" }}>
                        {loadingPreset && (
                            <Box
                                sx={{
                                    position: "absolute",
                                    top: 0,
                                    left: 0,
                                    right: 0,
                                    bottom: 0,
                                    bgcolor: "rgba(255, 255, 255, 0.7)",
                                    zIndex: 10,
                                    display: "flex",
                                    flexDirection: "column",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    backdropFilter: "blur(2px)",
                                    borderRadius: 2,
                                }}
                            >
                                <CircularProgress size={40} sx={{ color: "#1E3A8A", mb: 1.5 }} />
                                <Typography fontWeight={600} color="#1E3A8A" fontSize="0.9rem">
                                    Loading Preset Configuration...
                                </Typography>
                            </Box>
                        )}
                        <TableContainer
                        component={Paper}
                        variant="outlined"
                        sx={{
                            flex: 1,
                            minHeight: 0,
                            overflowY: "auto",
                            overflowX: "auto",
                            borderRadius: 2,
                            borderColor: "#cbd5e1",
                            boxShadow: "0 1px 3px 0 rgb(0 0 0 / 0.1)",
                            "&::-webkit-scrollbar": {
                                width: "8px",
                                height: "8px",
                            },
                            "&::-webkit-scrollbar-thumb": {
                                backgroundColor: "#94a3b8",
                                borderRadius: "4px",
                            },
                        }}
                    >
                        <DndContext
                            sensors={sensors}
                            collisionDetection={closestCenter}
                            onDragEnd={handleDragEnd}
                        >
                            <Table stickyHeader size="small">
                                <TableHead>
                                    <TableRow sx={{ "& th": { bgcolor: "#1E3A8A", color: "#ffffff", fontWeight: 700, fontSize: "0.82rem", py: 1.2 } }}>
                                        <TableCell width={90} align="center">S.No</TableCell>
                                        <TableCell width={240}>Report Name</TableCell>
                                        <TableCell width={240}>Report Mode(s)</TableCell>
                                        <TableCell width={220}>Saved Template</TableCell>
                                        <TableCell width={180}>Export Style</TableCell>
                                        <TableCell width={80} align="center">Action</TableCell>
                                    </TableRow>
                                </TableHead>
                                <TableBody>
                                    <SortableContext
                                        items={selectedReports.map((r) => r.id)}
                                        strategy={verticalListSortingStrategy}
                                    >
                                        {selectedReports.map((cfg, idx) => (
                                            <SortableReportRow
                                                key={cfg.id}
                                                cfg={cfg}
                                                idx={idx}
                                                onUpdateReport={handleUpdateReport}
                                                onRemoveReport={handleRemoveReport}
                                            />
                                        ))}
                                    </SortableContext>
                                </TableBody>
                            </Table>
                        </DndContext>
                    </TableContainer>
                    </Box>
                )}
            </Box>

            {/* Save / Edit Preset Dialog */}
            <Dialog open={saveDialogOpen} onClose={() => setSaveDialogOpen(false)} maxWidth="xs" fullWidth>
                <DialogTitle sx={{ fontWeight: 700, display: "flex", alignItems: "center", gap: 1 }}>
                    <BookmarkBorderIcon color="primary" /> {selectedPresetId ? "Edit / Update Batch Export Preset" : "Save Batch Export Preset"}
                </DialogTitle>
                <DialogContent>
                    <Typography variant="body2" color="text.secondary" mb={2}>
                        {selectedPresetId
                            ? `Update "${presetNameInput || "the selected preset"}" with the current set of ${selectedReports.length} report(s) and configurations.`
                            : `Save this set of ${selectedReports.length} report(s) and their options as a template to quickly reuse in the future.`}
                    </Typography>
                    <TextField
                        autoFocus
                        size="small"
                        label="Preset Template Name"
                        placeholder="e.g. Daily Management Pack"
                        fullWidth
                        value={presetNameInput}
                        onChange={(e) => setPresetNameInput(e.target.value)}
                    />
                </DialogContent>
                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button onClick={() => setSaveDialogOpen(false)} sx={{ textTransform: "none", color: "#64748b" }}>
                        Cancel
                    </Button>
                    <Button
                        variant="contained"
                        onClick={handleSavePreset}
                        sx={{ textTransform: "none", bgcolor: "#1E3A8A", "&:hover": { bgcolor: "#172554" } }}
                    >
                        {selectedPresetId ? "Update Preset" : "Save Preset"}
                    </Button>
                </DialogActions>
            </Dialog>
        </Box>
    );
};

export default BatchExportScreen;
