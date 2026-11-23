import React, { useEffect, useMemo, useState } from "react";
import {
    Box,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Paper,
    TextField,
    CircularProgress,
    Typography,
    IconButton,
    Button,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
} from "@mui/material";
import KeyboardArrowRightIcon from "@mui/icons-material/KeyboardArrowRight";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import DeleteOutlineIcon from "@mui/icons-material/DeleteOutline";
import AddCircleOutlineIcon from "@mui/icons-material/AddCircleOutline";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";

import PageHeader, { ToggleMode } from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import CommonPagination from "../../Components/CommonPagination";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import {
    OutstandingManagementReportService,
    OutstandingManagementItem,
} from "../../services/outstandingManagementReport.service";
import { SettingsService } from "../../services/reportSettings.services";

/* ================= TYPES & HELPERS ================= */

export interface AgingScale {
    id: string;
    from: number | string;
    to: number | string;
}

export interface ProcessedOutstandingRow {
    id: string;
    Acc_Id: string;
    Account_name: string;
    invoice_date: string;
    invoice_no: string;
    Bal_Amount: number;
    CR_DR: string;
    Group_Name: string;
    days: number;
    scaleLabel: string;
    invoice_count?: number;
}

const formatNumber = (num: number) => {
    return new Intl.NumberFormat("en-IN", {
        maximumFractionDigits: 2,
        minimumFractionDigits: 0,
    }).format(num || 0);
};

export const calculateDays = (invoiceDateStr: string, toDateStr: string): number => {
    if (!invoiceDateStr) return 0;
    const inv = dayjs(invoiceDateStr);
    const target = dayjs(toDateStr);
    if (!inv.isValid() || !target.isValid()) return 0;
    const diff = target.diff(inv, "day");
    return diff > 0 ? diff : 0;
};

// Compute label for a given day count against configured scales
export const getScaleLabelForDays = (days: number, scales: AgingScale[]): string => {
    for (const sc of scales) {
        const fromNum = sc.from === "" ? 0 : Number(sc.from);
        const toNum = sc.to === "" ? Infinity : Number(sc.to);
        if (days >= fromNum && days <= toNum) {
            if (sc.to === "" || !isFinite(toNum)) {
                return `Above ${fromNum} Days`;
            }
            return `${fromNum} to ${toNum} Days`;
        }
    }
    // If not matched by any interval, fallback
    return `Other Days (${days})`;
};

// Initial default scale: single range 0 to Max (user adds more as needed)
const INITIAL_DEFAULT_SCALES: AgingScale[] = [
    { id: "s1", from: 0, to: "" },
];

/* ================= COMPONENT ================= */

const OutstandingManagementReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    // Toggle: Abstract = Receivable (CR), Expanded = Payable (DR)
    const [toggleMode, setToggleMode] = useState<ToggleMode>("Abstract");
    const [rawRows, setRawRows] = useState<OutstandingManagementItem[]>([]);
    const [loading, setLoading] = useState(false);

    // Date Filters
    const [toDate, setToDate] = useState<string>(today);
    const [tempToDate, setTempToDate] = useState<string>(today);
    const [drawerOpen, setDrawerOpen] = useState<boolean>(false);

    // Aging Scales (scale like inputs in drawer) - starts with single range 0 to Max
    const [scales, setScales] = useState<AgingScale[]>(INITIAL_DEFAULT_SCALES);
    const [tempScales, setTempScales] = useState<AgingScale[]>(INITIAL_DEFAULT_SCALES);

    // Template state
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [reportName, setReportName] = useState<string>("");
    const [saveDialogOpen, setSaveDialogOpen] = useState<boolean>(false);
    const [parentReportName, setParentReportName] = useState<string>("RECIEVABLE PAYABLE");

    const [expandedKeys, setExpandedKeys] = useState<string[]>([]);

    // Pagination
    const [page, setPage] = useState<number>(1);
    const [rowsPerPage, setRowsPerPage] = useState<number>(100);

    // Account Name Header Filter
    const [accountFilterAnchor, setAccountFilterAnchor] = useState<null | HTMLElement>(null);
    const [selectedAccountNames, setSelectedAccountNames] = useState<string[] | undefined>(undefined);

    /* ================= DATA LOADING ================= */

    const loadData = async (targetDate: string) => {
        setLoading(true);
        try {
            const res = await OutstandingManagementReportService.getRecievablePayableMgt({ Todate: targetDate });
            const list: OutstandingManagementItem[] = Array.isArray(res)
                ? res
                : (res.data && Array.isArray(res.data) ? res.data : (res as any)?.data?.data || []);
            setRawRows(list);
        } catch (err) {
            console.error("Failed to load recievablepayablemgt report:", err);
            toast.error("Failed to load report data");
            setRawRows([]);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        loadData(toDate);
    }, [toDate]);

    /* ================= FILTER BY TOGGLE MODE ================= */

    // "CR_DR": "CR" show Cr fully as Recievable and DR as Payable
    const targetCrDr = toggleMode === "Abstract" ? "CR" : "DR";
    const amountColumnTitle = toggleMode === "Abstract" ? "Recievable Amt" : "Payable Amt";
    const reportTitle = toggleMode === "Abstract" ? "RECIEVABLE PAYABLE - Receivable" : "RECIEVABLE PAYABLE - Payable";

    const processedData: ProcessedOutstandingRow[] = useMemo(() => {
        const filtered = rawRows.filter((r) => String(r.CR_DR || "").toUpperCase() === targetCrDr);

        // Group by Account_name + days to consolidate duplicate invoices
        // If the Same Account name has the Same days of pending, sum amounts and show invoice count in parentheses
        const accountDaysMap = new Map<string, {
            baseItem: OutstandingManagementItem;
            days: number;
            scaleLabel: string;
            totalAmount: number;
            invoiceCount: number;
        }>();

        filtered.forEach((r) => {
            const days = calculateDays(r.invoice_date, toDate);
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

        const list: ProcessedOutstandingRow[] = [];
        let idx = 0;
        accountDaysMap.forEach((val) => {
            idx++;
            const { baseItem, days, scaleLabel, totalAmount, invoiceCount } = val;
            list.push({
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

        return list;
    }, [rawRows, targetCrDr, toDate, scales]);

    // Unique Account Names for Header Filter
    const uniqueAccountNames = useMemo(() => {
        const setNames = new Set<string>();
        processedData.forEach((row) => {
            if (row.Account_name) setNames.add(row.Account_name);
        });
        return Array.from(setNames).sort((a, b) => a.localeCompare(b));
    }, [processedData]);

    // Filter by selected Account Names
    const filteredProcessedData = useMemo(() => {
        if (!selectedAccountNames) return processedData;
        const setAllowed = new Set(selectedAccountNames);
        return processedData.filter((row) => setAllowed.has(row.Account_name));
    }, [processedData, selectedAccountNames]);

    const grandTotalAmount = useMemo(() => {
        return filteredProcessedData.reduce((acc, row) => acc + (row.Bal_Amount || 0), 0);
    }, [filteredProcessedData]);

    /* ================= GROUPING BY SCALE INTERVALS ================= */

    const groupedData = useMemo(() => {
        // Build groups ordered by defined scales
        const map = new Map<string, ProcessedOutstandingRow[]>();
        scales.forEach((sc) => {
            const fromNum = sc.from === "" ? 0 : Number(sc.from);
            const toNum = sc.to === "" ? Infinity : Number(sc.to);
            const label = sc.to === "" || !isFinite(toNum) ? `Above ${fromNum} Days` : `${fromNum} to ${toNum} Days`;
            map.set(label, []);
        });

        filteredProcessedData.forEach((row) => {
            if (!map.has(row.scaleLabel)) {
                map.set(row.scaleLabel, []);
            }
            map.get(row.scaleLabel)!.push(row);
        });

        const groups: { label: string; subtotal: number; items: ProcessedOutstandingRow[] }[] = [];
        map.forEach((items, label) => {
            const subtotal = items.reduce((sum, it) => sum + it.Bal_Amount, 0);
            groups.push({
                label,
                subtotal,
                items,
            });
        });

        return groups;
    }, [filteredProcessedData, scales]);

    // Default: only the LAST level of grouping is expanded so its Account names are listed below
    useEffect(() => {
        if (groupedData.length > 0) {
            const lastGroup = groupedData[groupedData.length - 1];
            if (lastGroup) {
                setExpandedKeys([lastGroup.label]);
            }
        }
    }, [groupedData]);

    // Toggle expand/collapse: ONLY the last level of grouping can be expanded to show Account names!
    const handleToggleExpand = (groupKey: string, level: number) => {
        const lastGroupLevel = scales.length - 1;
        // If not the last level, prevent expanding
        if (level < lastGroupLevel) {
            return;
        }

        setExpandedKeys((prev) =>
            prev.includes(groupKey) ? prev.filter((k) => k !== groupKey) : [...prev, groupKey]
        );
    };

    // Flatten rows for display with pagination
    const flattenedRows = useMemo(() => {
        const lastGroupLevel = scales.length - 1;
        const result: any[] = [];

        groupedData.forEach((group, idx) => {
            const isLastLevel = idx === lastGroupLevel;
            const isExpanded = expandedKeys.includes(group.label);

            result.push({
                __group: true,
                __key: group.label,
                __label: group.label,
                __subtotal: group.subtotal,
                __count: group.items.length,
                __level: idx,
                __isLastLevel: isLastLevel,
            });

            // ONLY show the Account names if this is the last level and it is expanded
            // Sort items ascending: lower number of days to larger number of days, and then by Account name
            if (isLastLevel && isExpanded) {
                const sortedItems = [...group.items].sort((a, b) => {
                    if (a.days !== b.days) {
                        return a.days - b.days;
                    }
                    return (a.Account_name || "").localeCompare(b.Account_name || "");
                });

                sortedItems.forEach((item) => {
                    result.push(item);
                });
            }
        });

        return result;
    }, [groupedData, expandedKeys, scales.length]);

    const paginatedRows = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        return flattenedRows.slice(start, start + rowsPerPage);
    }, [flattenedRows, page, rowsPerPage]);

    /* ================= SCALE EDITING IN DRAWER ================= */

    const handleScaleChange = (id: string, field: "from" | "to", val: string) => {
        setTempScales((prev) =>
            prev.map((sc) => (sc.id === id ? { ...sc, [field]: val === "" ? "" : Number(val) } : sc))
        );
    };

    const handleAddScale = () => {
        const newId = `s-${Date.now()}`;
        setTempScales((prev) => {
            const last = prev[prev.length - 1];
            const nextFrom = last && last.to !== "" && !isNaN(Number(last.to)) ? Number(last.to) + 1 : 0;
            return [...prev, { id: newId, from: nextFrom, to: "" }];
        });
    };

    const handleRemoveScale = (id: string) => {
        setTempScales((prev) => prev.filter((sc) => sc.id !== id));
    };

    /* ================= TEMPLATE LOADING & SAVING ================= */

    const loadTemplate = async (reportId: number) => {
        try {
            setSelectedTemplateId(reportId);
            const res = await SettingsService.getReportEditData({
                reportId,
                typeId: 1,
            });

            const templateCols = res?.data?.data?.columns || [];
            const reportInfo = res?.data?.data?.reportInfo;

            if (reportInfo?.Report_Name) {
                setReportName(reportInfo.Report_Name);
            }

            // Extract saved aging scale configuration from template columns if available
            // Format of saved scale column key: "SCALE_from_to" or stored in localStorage
            const savedScalesJson = localStorage.getItem(`template_scales_${reportId}`);
            if (savedScalesJson) {
                try {
                    const parsed = JSON.parse(savedScalesJson);
                    if (Array.isArray(parsed) && parsed.length > 0) {
                        setScales(parsed);
                        setTempScales(parsed);
                        toast.success("Template Loaded Successfully ✅");
                        return;
                    }
                } catch (e) {
                    // fallback
                }
            }

            // Or parse from template column keys
            const scaleCols = templateCols.filter((c: any) => c.key && c.key.startsWith("SCALE_"));
            if (scaleCols.length > 0) {
                const parsedScales: AgingScale[] = scaleCols.map((c: any, sIdx: number) => {
                    const parts = c.key.replace("SCALE_", "").split("_");
                    return {
                        id: `s-${sIdx + 1}`,
                        from: parts[0] === "" ? 0 : Number(parts[0]),
                        to: parts[1] === "" || parts[1] === "MAX" ? "" : Number(parts[1]),
                    };
                });
                setScales(parsedScales);
                setTempScales(parsedScales);
            }

            toast.success("Template Loaded Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to load template ❌");
        }
    };

    const handleSaveTemplate = async () => {
        try {
            if (!reportName.trim()) {
                toast.error("Enter Report Name");
                return;
            }

            const userData = JSON.parse(localStorage.getItem("user") || "{}");
            const createdBy = userData.id || 1;

            // Encode scale ranges into columns payload
            const payloadColumns = scales.map((sc, sIdx) => {
                const toStr = sc.to === "" ? "MAX" : sc.to;
                return {
                    key: `SCALE_${sc.from}_${toStr}`,
                    label: sc.to === "" ? `Above ${sc.from} Days` : `${sc.from} to ${sc.to} Days`,
                    order: sIdx + 1,
                    groupBy: sIdx + 1,
                    enabled: true,
                    isNumeric: false,
                    dataType: "nvarchar",
                };
            });

            if (selectedTemplateId) {
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 1,
                    reportName: reportName.trim(),
                    columns: payloadColumns,
                });
                localStorage.setItem(`template_scales_${selectedTemplateId}`, JSON.stringify(scales));
                toast.success("Template Updated Successfully ✅");
            } else {
                const res = await SettingsService.saveReportSettings({
                    reportName: reportName.trim(),
                    parentReport: parentReportName,
                    abstractSP: "EXTERNAL_API",
                    expandedSP: "EXTERNAL_API",
                    abstractColumns: payloadColumns,
                    expandedColumns: payloadColumns,
                    createdBy,
                });
                const newId = (res as any)?.data?.reportId;
                if (newId) {
                    localStorage.setItem(`template_scales_${newId}`, JSON.stringify(scales));
                }
                toast.success("Template Saved Successfully ✅");
            }

            setSaveDialogOpen(false);
            setTimeout(() => {
                window.location.reload();
            }, 500);
        } catch (err) {
            console.error(err);
            toast.error("Save Failed ❌");
        }
    };

    /* ================= EXPORT FUNCTIONS ================= */

    const handleExportExcel = () => {
        if (!processedData.length) {
            toast.warning("No data to export");
            return;
        }

        const workbook = XLSX.utils.book_new();
        const exportAoa: any[][] = [];

        exportAoa.push(["Account Name", "Days", amountColumnTitle]);
        exportAoa.push(["Total", "", grandTotalAmount]);

        flattenedRows.forEach((row) => {
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

        // Styling cells with professional report theme (#1E3A8A)
        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:C1");
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = worksheet[cellRef];
                if (!cell) continue;

                cell.s = {
                    font: { name: "Arial", sz: 9 },
                    border: {
                        top: { style: "thin", color: { rgb: "E2E8F0" } },
                        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
                        left: { style: "thin", color: { rgb: "E2E8F0" } },
                        right: { style: "thin", color: { rgb: "E2E8F0" } },
                    },
                };

                if (R === 0) {
                    // Table Header: Deep Blue (#1E3A8A) with White Bold text
                    cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                    cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                    cell.s.alignment = { horizontal: C === 0 ? "left" : "center", vertical: "center" };
                } else if (R === 1) {
                    // Total Row: Subtle slate-blue highlight
                    cell.s.fill = { fgColor: { rgb: "EFF6FF" } };
                    cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "1E3A8A" } };
                    cell.s.alignment = { horizontal: C === 0 ? "left" : "right" };
                } else {
                    const rowData = flattenedRows[R - 2];
                    if (rowData && rowData.__group) {
                        // Group Subtotal Row: Soft slate header (#F1F5F9)
                        cell.s.fill = { fgColor: { rgb: "F1F5F9" } };
                        cell.s.font = { name: "Arial", sz: 9.5, bold: true, color: { rgb: "0F172A" } };
                        cell.s.alignment = { horizontal: C === 0 ? "left" : "right" };
                    } else {
                        if (C === 1) cell.s.alignment = { horizontal: "center" };
                        if (C === 2) cell.s.alignment = { horizontal: "right" };
                    }
                }
            }
        }

        worksheet["!cols"] = [{ wch: 45 }, { wch: 12 }, { wch: 20 }];

        XLSX.utils.book_append_sheet(workbook, worksheet, "OutstandingManagement");
        XLSX.writeFile(workbook, `Outstanding_Management_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
        toast.success("Excel exported successfully! 📊");
    };

    const handleExportPDF = () => {
        if (!processedData.length) {
            toast.warning("No data to export");
            return;
        }

        const doc = new jsPDF("p", "mm", "a4");
        doc.setFontSize(12);
        doc.setTextColor(30, 58, 138);
        doc.text(`${reportTitle} - As of ${dayjs(toDate).format("DD/MM/YYYY")}`, 14, 12);

        const tableBody: any[] = [];
        tableBody.push(["Total", "", formatNumber(grandTotalAmount)]);

        flattenedRows.forEach((row) => {
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
            head: [["Account Name", "Days", amountColumnTitle]],
            body: tableBody,
            styles: { fontSize: 8.5, cellPadding: 2 },
            headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255], fontStyle: "bold" },
            didParseCell: (data) => {
                if (data.row.index === 0) {
                    data.cell.styles.fillColor = [239, 246, 255];
                    data.cell.styles.textColor = [30, 58, 138];
                    data.cell.styles.fontStyle = "bold";
                } else {
                    const rowData = flattenedRows[data.row.index - 1];
                    if (rowData && rowData.__group) {
                        data.cell.styles.fillColor = [241, 245, 249];
                        data.cell.styles.textColor = [15, 23, 42];
                        data.cell.styles.fontStyle = "bold";
                    }
                }
                if (data.column.index === 1) data.cell.styles.halign = "center";
                if (data.column.index === 2) data.cell.styles.halign = "right";
            },
        });

        doc.save(`Outstanding_Management_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`);
        toast.success("PDF exported successfully! 📄");
    };

    /* ================= THEME STYLES ================= */

    const themeNavy = "#1E3A8A"; // Standard system header color
    const groupRowBg = "#F1F5F9"; // Standard clean slate subtotal background

    return (
        <Box sx={{ display: "flex", flexDirection: "column", height: "100%", width: "100%", overflow: "hidden" }}>
            {/* PAGE HEADER: Save Template & Select Template enabled for Authorized User Types */}
            <PageHeader
                toggleMode={toggleMode}
                onToggleChange={(mode) => setToggleMode(mode)}
                abstractLabel="Recievable"
                expandedLabel="Payable"
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                parentReportName={parentReportName}
                onReportChange={(template) => {
                    if (!template) {
                        setSelectedTemplateId(null);
                        setReportName("");
                        setScales(INITIAL_DEFAULT_SCALES);
                        setTempScales(INITIAL_DEFAULT_SCALES);
                        return;
                    }
                    setSelectedTemplateId(template.Report_Id);
                    setReportName(template.Report_Name);
                    loadTemplate(template.Report_Id);
                }}
                onQuickSave={(parentName) => {
                    setParentReportName(parentName || "RECIEVABLE PAYABLE");
                    if (!selectedTemplateId) {
                        setReportName("");
                    }
                    setSaveDialogOpen(true);
                }}
            />

            {/* FILTER DRAWER: ToDate and Scale-Like Group By Inputs */}
            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => {
                    setDrawerOpen((prev) => {
                        if (!prev) {
                            setTempToDate(toDate);
                            setTempScales(scales.map((s) => ({ ...s })));
                        }
                        return !prev;
                    });
                }}
                onClose={() => setDrawerOpen(false)}
                hideFromDate={true}
                toDate={tempToDate}
                onToDateChange={setTempToDate}
                onApply={() => {
                    setToDate(tempToDate);
                    setScales(tempScales.map((s) => ({ ...s })));
                    setDrawerOpen(false);
                }}
            >
                <Box sx={{ mt: 2, mb: 1 }}>
                    <Typography sx={{ fontWeight: 700, color: themeNavy, fontSize: "0.88rem", mb: 0.5 }}>
                        Day Scales / Aging Ranges
                    </Typography>

                    {tempScales.map((sc, sIdx) => (
                        <Box
                            key={sc.id}
                            sx={{
                                display: "flex",
                                alignItems: "center",
                                gap: 1,
                                mb: 1.2,
                                bgcolor: "#f8fafc",
                                p: 1,
                                borderRadius: 1,
                                border: "1px solid #e2e8f0",
                            }}
                        >
                            <Typography sx={{ fontSize: "0.75rem", fontWeight: 700, color: "#475569", minWidth: 20 }}>
                                #{sIdx + 1}
                            </Typography>
                            <TextField
                                size="small"
                                type="number"
                                placeholder="From"
                                value={sc.from}
                                onChange={(e) => handleScaleChange(sc.id, "from", e.target.value)}
                                sx={{
                                    width: 80,
                                    bgcolor: "#fff",
                                    "& input": { py: 0.7, fontSize: "0.82rem", textAlign: "center" },
                                }}
                            />
                            <Typography sx={{ fontSize: "0.82rem", fontWeight: 700, color: "#64748b" }}>
                                to
                            </Typography>
                            <TextField
                                size="small"
                                type="number"
                                placeholder="Max"
                                value={sc.to}
                                onChange={(e) => handleScaleChange(sc.id, "to", e.target.value)}
                                sx={{
                                    width: 80,
                                    bgcolor: "#fff",
                                    "& input": { py: 0.7, fontSize: "0.82rem", textAlign: "center" },
                                }}
                            />
                            <Typography sx={{ fontSize: "0.75rem", color: "#64748b" }}>Days</Typography>

                            {tempScales.length > 1 && (
                                <IconButton
                                    size="small"
                                    color="error"
                                    onClick={() => handleRemoveScale(sc.id)}
                                    sx={{ ml: "auto", p: 0.5 }}
                                >
                                    <DeleteOutlineIcon fontSize="small" />
                                </IconButton>
                            )}
                        </Box>
                    ))}

                    <Button
                        size="small"
                        startIcon={<AddCircleOutlineIcon />}
                        onClick={handleAddScale}
                        sx={{
                            mt: 0.5,
                            textTransform: "none",
                            fontSize: "0.78rem",
                            fontWeight: 600,
                            color: themeNavy,
                        }}
                    >
                        Add Another Scale
                    </Button>
                </Box>
            </ReportFilterDrawer>

            {/* MAIN TABLE CONTAINER: System Reporting Theme */}
            <Box sx={{ flex: 1, minHeight: 0, p: { xs: 1, md: 2 }, display: "flex", flexDirection: "column" }}>
                <TableContainer
                    component={Paper}
                    sx={{
                        flex: 1,
                        minHeight: 0,
                        border: "1px solid #cbd5e1",
                        boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                        borderRadius: 1.5,
                        overflow: "auto",
                    }}
                >
                    <Table size="small" stickyHeader sx={{ minWidth: 650 }}>
                        {/* HEADER: Standard System Theme (#1E3A8A) */}
                        <TableHead>
                            <TableRow>
                                <TableCell
                                    sx={{
                                        bgcolor: themeNavy,
                                        color: "#ffffff",
                                        fontWeight: 700,
                                        fontSize: "0.82rem",
                                        py: 1,
                                        borderRight: "1px solid rgba(255,255,255,0.15)",
                                        width: "60%",
                                        cursor: "pointer",
                                    }}
                                    onClick={(e) => setAccountFilterAnchor(e.currentTarget)}
                                >
                                    <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                                        <span>Account Name</span>
                                        <Typography
                                            component="span"
                                            sx={{
                                                fontSize: "0.7rem",
                                                opacity: selectedAccountNames ? 1 : 0.6,
                                                bgcolor: selectedAccountNames ? "rgba(255,255,255,0.2)" : "transparent",
                                                px: 0.8,
                                                py: 0.2,
                                                borderRadius: 0.5,
                                            }}
                                        >
                                            {selectedAccountNames ? `(${selectedAccountNames.length})` : "▼"}
                                        </Typography>
                                    </Box>
                                </TableCell>
                                <TableCell
                                    align="center"
                                    sx={{
                                        bgcolor: themeNavy,
                                        color: "#ffffff",
                                        fontWeight: 700,
                                        fontSize: "0.82rem",
                                        py: 1,
                                        borderRight: "1px solid rgba(255,255,255,0.15)",
                                        width: "15%",
                                    }}
                                >
                                    Days
                                </TableCell>
                                <TableCell
                                    align="right"
                                    sx={{
                                        bgcolor: themeNavy,
                                        color: "#ffffff",
                                        fontWeight: 700,
                                        fontSize: "0.82rem",
                                        py: 1,
                                        width: "25%",
                                        pr: 2,
                                    }}
                                >
                                    {amountColumnTitle}
                                </TableCell>
                            </TableRow>
                        </TableHead>

                        <TableBody>
                            {/* TOTAL ROW (Top): Standard System Style (#EFF6FF) */}
                            <TableRow sx={{ bgcolor: "#EFF6FF", borderBottom: "2px solid #bfdbfe" }}>
                                <TableCell
                                    sx={{
                                        color: themeNavy,
                                        fontWeight: 800,
                                        fontSize: "0.85rem",
                                        py: 0.9,
                                    }}
                                >
                                    Total
                                </TableCell>
                                <TableCell sx={{ py: 0.9 }} />
                                <TableCell
                                    align="right"
                                    sx={{
                                        color: themeNavy,
                                        fontWeight: 800,
                                        fontSize: "0.85rem",
                                        py: 0.9,
                                        pr: 2,
                                    }}
                                >
                                    {formatNumber(grandTotalAmount)}
                                </TableCell>
                            </TableRow>

                            {/* LOADING STATE */}
                            {loading ? (
                                <TableRow>
                                    <TableCell colSpan={3} align="center" sx={{ py: 6 }}>
                                        <CircularProgress size={30} sx={{ color: themeNavy }} />
                                        <Typography sx={{ mt: 1, fontSize: "0.82rem", color: "#64748b" }}>
                                            Loading data...
                                        </Typography>
                                    </TableCell>
                                </TableRow>
                            ) : paginatedRows.length === 0 ? (
                                <TableRow>
                                    <TableCell colSpan={3} align="center" sx={{ py: 6 }}>
                                        <Typography sx={{ fontSize: "0.85rem", color: "#64748b" }}>
                                            No outstanding records found
                                        </Typography>
                                    </TableCell>
                                </TableRow>
                            ) : (
                                paginatedRows.map((row: any, rIdx: number) => {
                                    // GROUP ROW (Scale Header)
                                    if (row.__group) {
                                        const isExpanded = expandedKeys.includes(row.__key);
                                        const isLastLevel = Boolean(row.__isLastLevel);

                                        return (
                                            <TableRow
                                                key={`grp-${row.__key}-${rIdx}`}
                                                sx={{
                                                    bgcolor: groupRowBg,
                                                    cursor: isLastLevel ? "pointer" : "default",
                                                    borderTop: "1px solid #e2e8f0",
                                                    borderBottom: "1px solid #e2e8f0",
                                                    "&:hover": { bgcolor: isLastLevel ? "#e2e8f0" : groupRowBg },
                                                }}
                                                onClick={() => {
                                                    if (isLastLevel) {
                                                        handleToggleExpand(row.__key, row.__level);
                                                    }
                                                }}
                                            >
                                                <TableCell
                                                    sx={{
                                                        color: "#0f172a",
                                                        fontWeight: 700,
                                                        fontSize: "0.83rem",
                                                        py: 0.8,
                                                        display: "flex",
                                                        alignItems: "center",
                                                        gap: 0.5,
                                                    }}
                                                >
                                                    {isLastLevel ? (
                                                        <IconButton size="small" sx={{ p: 0.2 }}>
                                                            {isExpanded ? (
                                                                <KeyboardArrowDownIcon sx={{ fontSize: 18, color: themeNavy }} />
                                                            ) : (
                                                                <KeyboardArrowRightIcon sx={{ fontSize: 18, color: themeNavy }} />
                                                            )}
                                                        </IconButton>
                                                    ) : (
                                                        <Box sx={{ width: 22, display: "inline-block" }} />
                                                    )}
                                                    <Typography sx={{ fontWeight: 700, fontSize: "0.83rem", color: themeNavy }}>
                                                        {row.__label}
                                                    </Typography>
                                                    <Typography sx={{ fontSize: "0.72rem", color: "#64748b", ml: 0.5 }}>
                                                        ({row.__count} records)
                                                    </Typography>
                                                </TableCell>
                                                <TableCell sx={{ py: 0.8 }} />
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        color: "#0f172a",
                                                        fontWeight: 700,
                                                        fontSize: "0.83rem",
                                                        py: 0.8,
                                                        pr: 2,
                                                    }}
                                                >
                                                    {formatNumber(row.__subtotal)}
                                                </TableCell>
                                            </TableRow>
                                        );
                                    }

                                    // DETAIL ROW
                                    return (
                                        <TableRow
                                            key={`item-${row.id}-${rIdx}`}
                                            sx={{
                                                bgcolor: "#ffffff",
                                                borderBottom: "1px solid #f1f5f9",
                                                "&:hover": { bgcolor: "#f8fafc" },
                                            }}
                                        >
                                            <TableCell
                                                sx={{
                                                    color: "#334155",
                                                    fontSize: "0.8rem",
                                                    py: 0.65,
                                                    pl: 4,
                                                }}
                                            >
                                                {row.Account_name}
                                                {row.invoice_count && row.invoice_count > 1 ? ` (${row.invoice_count})` : ""}
                                            </TableCell>
                                            <TableCell
                                                align="center"
                                                sx={{
                                                    color: "#475569",
                                                    fontSize: "0.8rem",
                                                    py: 0.65,
                                                }}
                                            >
                                                {row.days}
                                            </TableCell>
                                            <TableCell
                                                align="right"
                                                sx={{
                                                    color: "#1e293b",
                                                    fontSize: "0.8rem",
                                                    fontWeight: 500,
                                                    py: 0.65,
                                                    pr: 2,
                                                }}
                                            >
                                                {formatNumber(row.Bal_Amount)}
                                            </TableCell>
                                        </TableRow>
                                    );
                                })
                            )}
                        </TableBody>
                    </Table>
                </TableContainer>

                {/* COMMON PAGINATION */}
                {flattenedRows.length > 0 && (
                    <CommonPagination
                        totalRows={flattenedRows.length}
                        page={page}
                        rowsPerPage={rowsPerPage}
                        onPageChange={setPage}
                        onRowsPerPageChange={(newRpp) => {
                            setRowsPerPage(newRpp);
                            setPage(1);
                        }}
                    />
                )}
            </Box>

            {/* ACCOUNT NAME HEADER FILTER MENU */}
            <HeaderFilterMenu
                anchorEl={accountFilterAnchor}
                open={Boolean(accountFilterAnchor)}
                onClose={() => setAccountFilterAnchor(null)}
                columnLabel="Account Name"
                options={uniqueAccountNames}
                selectedValues={selectedAccountNames}
                onFilterChange={(selected: string[] | undefined) => {
                    setSelectedAccountNames(selected);
                    setPage(1);
                }}
            />

            {/* SAVE TEMPLATE DIALOG */}
            <Dialog
                open={saveDialogOpen}
                onClose={() => setSaveDialogOpen(false)}
                fullWidth
                maxWidth="xs"
            >
                <DialogTitle sx={{ fontWeight: 700, color: themeNavy, fontSize: "1rem" }}>
                    {selectedTemplateId ? "Update Template" : "Save as Template"}
                </DialogTitle>
                <DialogContent>
                    <TextField
                        autoFocus
                        margin="dense"
                        label="Template Name"
                        fullWidth
                        size="small"
                        value={reportName}
                        onChange={(e) => setReportName(e.target.value)}
                        placeholder="Template name"
                        sx={{ mt: 1 }}
                    />
                </DialogContent>
                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button
                        onClick={() => setSaveDialogOpen(false)}
                        color="inherit"
                        size="small"
                        sx={{ textTransform: "none" }}
                    >
                        Cancel
                    </Button>
                    <Button
                        onClick={handleSaveTemplate}
                        variant="contained"
                        size="small"
                        sx={{
                            bgcolor: themeNavy,
                            "&:hover": { bgcolor: "#172554" },
                            textTransform: "none",
                            fontWeight: 600,
                        }}
                    >
                        {selectedTemplateId ? "Update" : "Save"}
                    </Button>
                </DialogActions>
            </Dialog>

            {/* SAVE TEMPLATE DIALOG */}
            <Dialog
                open={saveDialogOpen}
                onClose={() => setSaveDialogOpen(false)}
                fullWidth
                maxWidth="xs"
            >
                <DialogTitle sx={{ fontWeight: 700, color: themeNavy, fontSize: "1rem" }}>
                    {selectedTemplateId ? "Update Template" : "Save as Template"}
                </DialogTitle>
                <DialogContent>
                    <TextField
                        autoFocus
                        margin="dense"
                        label="Template Name"
                        fullWidth
                        size="small"
                        value={reportName}
                        onChange={(e) => setReportName(e.target.value)}
                        placeholder="Template name"
                        sx={{ mt: 1 }}
                    />
                </DialogContent>
                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button
                        onClick={() => setSaveDialogOpen(false)}
                        color="inherit"
                        size="small"
                        sx={{ textTransform: "none" }}
                    >
                        Cancel
                    </Button>
                    <Button
                        onClick={handleSaveTemplate}
                        variant="contained"
                        size="small"
                        sx={{
                            bgcolor: themeNavy,
                            "&:hover": { bgcolor: "#172554" },
                            textTransform: "none",
                            fontWeight: 600,
                        }}
                    >
                        {selectedTemplateId ? "Update" : "Save"}
                    </Button>
                </DialogActions>
            </Dialog>
        </Box>
    );
};

export default OutstandingManagementReport;
