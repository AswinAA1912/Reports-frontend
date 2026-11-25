import React, { useCallback, useEffect, useMemo, useState, useRef } from "react";
import {
    Box,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Paper,
    CircularProgress,
    IconButton,
    TextField,
    Autocomplete,
    Typography,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Button,
} from "@mui/material";
import dayjs from "dayjs";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import ArrowDropDownIcon from "@mui/icons-material/ArrowDropDown";
import ArrowDropUpIcon from "@mui/icons-material/ArrowDropUp";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { toast } from "react-toastify";
import { SettingsService } from "../../services/reportSettings.services";

import {
    StockDifferenceService,
    StockDifferenceItem,
    GodownMasterItem,
} from "../../services/Stock Difference.service";

/* ================= TYPES & CONSTANTS ================= */

type ColumnConfig = {
    key: string;
    label: string;
    enabled: boolean;
    order: number;
    width?: string | number;
    isNumeric?: boolean;
    align?: "left" | "right" | "center";
};

const NUMERIC_KEYS = ["System_Qty", "Actual_Qty", "Difference_Qty"];

const INITIAL_COLUMNS: ColumnConfig[] = [
    { key: "Godown_Name", label: "Godown", enabled: true, order: 0, align: "left", width: "9%" },
    { key: "Stock_Date", label: "Date", enabled: true, order: 1, align: "center", width: "8%" },
    { key: "Grade_Item_Group", label: "Grade Item Group", enabled: true, order: 2, align: "left", width: "10%" },
    { key: "Item_Name", label: "Item Name", enabled: true, order: 3, align: "left", width: "14%" },
    { key: "System_Qty", label: "System Stock", enabled: true, order: 4, isNumeric: true, align: "right", width: "10%" },
    { key: "Actual_Qty", label: "Manual Stock", enabled: true, order: 5, isNumeric: true, align: "right", width: "10%" },
    { key: "Difference_Qty", label: "Difference", enabled: true, order: 6, isNumeric: true, align: "right", width: "9%" },
    { key: "Remarks", label: "Remarks", enabled: true, order: 7, align: "left", width: "10%" },
    { key: "Created_At", label: "Created Date-Time", enabled: true, order: 8, align: "center", width: "17%" },
];

const formatOneDecimal = (val: number | null | undefined): string => {
    if (val === null || val === undefined || isNaN(Number(val))) return "0.0";
    return Number(val).toFixed(1);
};

const formatDateOnly = (val: string | null | undefined): string => {
    if (!val) return "-";
    const d = dayjs(val);
    return d.isValid() ? d.format("DD-MM-YYYY") : String(val);
};

const formatDateTime = (val: string | null | undefined): string => {
    if (!val) return "-";
    const d = dayjs(val);
    return d.isValid() ? d.format("DD-MM-YYYY hh:mm A") : String(val);
};

/* ================= MAIN COMPONENT ================= */

const StockDifferenceReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    // Drawer / filter states
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [tempFromDate, setTempFromDate] = useState(today);
    const [tempToDate, setTempToDate] = useState(today);

    // Godowns
    const [godowns, setGodowns] = useState<GodownMasterItem[]>([]);
    const [selectedGodown, setSelectedGodown] = useState<GodownMasterItem | null>(null);
    const [tempSelectedGodown, setTempSelectedGodown] = useState<GodownMasterItem | null>(null);

    // Raw Data & Loading
    const [rawRows, setRawRows] = useState<StockDifferenceItem[]>([]);
    const [loading, setLoading] = useState(false);

    // Columns config (Fixed list)
    const columns = INITIAL_COLUMNS;

    // Header filter anchor & active column
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);
    const [columnFilters, setColumnFilters] = useState<Record<string, string[] | undefined>>({});

    // Grouping: Fixed to Stock_Group
    const grouping = useMemo(() => ["Stock_Group"], []);
    const [expandedKeys, setExpandedKeys] = useState<string[]>([]);

    // String sort config
    const [sortConfig, setSortConfig] = useState<{
        key: string | null;
        direction: "asc" | "desc";
    }>({
        key: null,
        direction: "asc",
    });

    // Pagination
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(50);
    const serialRef = useRef(0);

    // Template states
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [reportName, setReportName] = useState("");
    const [parentReportName, setParentReportName] = useState("STOCK DIFFERENCE REPORT");
    const [saveDialogOpen, setSaveDialogOpen] = useState(false);

    /* ================= FETCH GODOWNS ================= */
    useEffect(() => {
        const fetchGodowns = async () => {
            try {
                const res = await StockDifferenceService.getGodowns();
                const list = res.data?.data || (Array.isArray(res.data) ? res.data : []);
                const validList: GodownMasterItem[] = Array.isArray(list)
                    ? list.filter((g) => g.Godown_Name && String(g.Godown_Name).trim() !== "" && String(g.Godown_Id) !== "0")
                    : [];

                // Prepend "All" option with Godown_Id = 0
                const allOption: GodownMasterItem = { Godown_Id: 0, Godown_Name: "All" };
                const fullList = [allOption, ...validList];
                setGodowns(fullList);

                // Default to All (Godown_Id: 0)
                setSelectedGodown(allOption);
                setTempSelectedGodown(allOption);
            } catch (err) {
                console.error("Failed to load godowns:", err);
            }
        };
        fetchGodowns();
    }, []);

    /* ================= FETCH REPORT DATA ================= */
    const loadReportData = useCallback(async (from: string, to: string, godownId?: string | number) => {
        try {
            setLoading(true);
            const res = await StockDifferenceService.getReport({
                Fromdate: from,
                Todate: to,
                Godown_Id: godownId,
            });

            const data = res.data?.data || (Array.isArray(res.data) ? res.data : []);
            const validRows: StockDifferenceItem[] = Array.isArray(data)
                ? data.map((item: any) => ({
                    ...item,
                    System_Qty: Number(item.System_Qty ?? 0),
                    Actual_Qty: Number(item.Actual_Qty ?? 0),
                    Difference_Qty: Number(item.Difference_Qty ?? 0),
                }))
                : [];

            setRawRows(validRows);
            setPage(1);

            // Auto-expand all unique Stock_Group groups
            const allGroups = Array.from(new Set(validRows.map((r) => String(r.Stock_Group ?? "Others"))));
            setExpandedKeys(allGroups.map((grp) => `Stock_Group:${grp}`));
        } catch (err) {
            console.error("Failed to load Stock Difference report:", err);
            toast.error("Failed to load Stock Difference report ❌");
            setRawRows([]);
        } finally {
            setLoading(false);
        }
    }, []);

    useEffect(() => {
        if (selectedGodown) {
            loadReportData(fromDate, toDate, selectedGodown.Godown_Id);
        }
    }, [fromDate, toDate, selectedGodown, loadReportData]);

    const handleApplyFilters = () => {
        setFromDate(tempFromDate);
        setToDate(tempToDate);
        setSelectedGodown(tempSelectedGodown);
        setDrawerOpen(false);
    };

    /* ================= TEMPLATE LOGIC ================= */
    const loadTemplate = async (templateId: number) => {
        try {
            setLoading(true);
            setSelectedTemplateId(templateId);

            const res = await SettingsService.getReportEditData({
                reportId: templateId,
                typeId: 1,
            });

            if (res.data?.success && res.data.data) {
                const info = res.data.data.reportInfo || res.data.data;
                if (info?.Report_Name) {
                    setReportName(info.Report_Name);
                }

                // Check stored Godown filter & column filters in Database Preset_Config
                let targetGodownId: string | number | null = null;
                const rawPreset = info?.Preset_Config;
                let dbConfig: any = null;
                if (rawPreset) {
                    try {
                        dbConfig = typeof rawPreset === "string" ? JSON.parse(rawPreset) : rawPreset;
                    } catch (e) {
                        console.warn("Failed to parse db Preset_Config in StockDifference", e);
                    }
                }

                if (dbConfig && typeof dbConfig === "object") {
                    if (dbConfig.godownId) targetGodownId = dbConfig.godownId;
                    if (dbConfig.columnFilters) setColumnFilters(dbConfig.columnFilters);
                    if (dbConfig.rangeFilter) setNumRangeFilter(dbConfig.rangeFilter);
                }

                if (!targetGodownId) {
                    const storedGodownId = localStorage.getItem(`template_godown_${templateId}`);
                    if (storedGodownId) {
                        targetGodownId = storedGodownId;
                    } else if (info?.Abstract_SP && !isNaN(Number(info.Abstract_SP))) {
                        targetGodownId = info.Abstract_SP;
                    }
                }

                if (targetGodownId && godowns.length > 0) {
                    const matched = godowns.find((g) => String(g.Godown_Id) === String(targetGodownId));
                    if (matched) {
                        setSelectedGodown(matched);
                        setTempSelectedGodown(matched);
                    }
                }

                toast.success("Template Loaded ✅");
            }
        } catch (err) {
            console.error("Failed to load template:", err);
            toast.error("Failed to load template ❌");
        } finally {
            setLoading(false);
        }
    };

    const handleClearTemplate = () => {
        setSelectedTemplateId(null);
        setReportName("");
        if (godowns.length > 0) {
            const allOption = godowns.find((g) => String(g.Godown_Id) === "0") || godowns[0];
            setSelectedGodown(allOption);
            setTempSelectedGodown(allOption);
        }
    };

    const handleSaveTemplate = async () => {
        if (!reportName.trim()) {
            toast.warning("Please enter a Template Name");
            return;
        }

        try {
            const userData = JSON.parse(localStorage.getItem("user") || "{}");
            const createdBy = userData?.id || 1;

            const payloadColumns = columns.map((col) => ({
                key: col.key,
                label: col.label,
                enabled: col.enabled ? 1 : 0,
                order: col.order,
                groupBy: col.key === "Stock_Group" ? 1 : 0,
            }));

            const godownIdStr = selectedGodown ? String(selectedGodown.Godown_Id) : "";
            const templateDbConfig = {
                godownId: godownIdStr,
                columnFilters,
                rangeFilter: numRangeFilter,
            };

            if (selectedTemplateId) {
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 1,
                    reportName: reportName.trim(),
                    columns: payloadColumns,
                    presetConfig: templateDbConfig,
                });
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 2,
                    reportName: reportName.trim(),
                    columns: payloadColumns,
                    presetConfig: templateDbConfig,
                });

                if (godownIdStr) {
                    localStorage.setItem(`template_godown_${selectedTemplateId}`, godownIdStr);
                }

                toast.success("Template Updated Successfully ✅");
            } else {
                const res = await SettingsService.saveReportSettings({
                    reportName: reportName.trim(),
                    parentReport: parentReportName,
                    abstractSP: godownIdStr || "EXTERNAL_API",
                    expandedSP: godownIdStr || "EXTERNAL_API",
                    abstractColumns: payloadColumns,
                    expandedColumns: payloadColumns,
                    createdBy,
                    presetConfig: templateDbConfig,
                });

                const newReportId = (res as any)?.data?.reportId || (res as any)?.data?.data?.Report_Id;
                if (newReportId && godownIdStr) {
                    localStorage.setItem(`template_godown_${newReportId}`, godownIdStr);
                }

                toast.success("Template Saved Successfully ✅");
            }

            setSaveDialogOpen(false);
            setTimeout(() => {
                window.location.reload();
            }, 500);
        } catch (err) {
            console.error("Save Template Failed:", err);
            toast.error("Save Template Failed ❌");
        }
    };

    /* ================= COLUMN FILTERING ================= */

    const filteredRows = useMemo(() => {
        return rawRows.filter((row) => {
            for (const [key, filterVals] of Object.entries(columnFilters)) {
                if (filterVals !== undefined && filterVals.length > 0) {
                    let val = String(row[key] ?? "");
                    if (key === "Stock_Date") {
                        val = formatDateOnly(row[key]);
                    } else if (key === "Created_At") {
                        val = formatDateTime(row[key]);
                    }
                    if (!filterVals.includes(val)) {
                        return false;
                    }
                }
            }
            return true;
        });
    }, [rawRows, columnFilters]);

    /* ================= NUMERICAL FILTER HOOK ================= */

    const {
        sortConfig: numSortConfig,
        rangeFilter: numRangeFilter,
        setRangeFilter: setNumRangeFilter,
        filterAnchor: numFilterAnchor,
        setFilterAnchor: setNumFilterAnchor,
        activeHeader: numActiveHeader,
        handleSort: handleNumSort,
        openFilter: openNumFilter,
        filteredAndSortedData: numFilteredAndSortedRows,
        getMinMax,
        clearRangeFilter,
    } = useNumericalFilter(filteredRows, NUMERIC_KEYS);

    /* ================= SORTING ================= */

    const sortedRows = useMemo(() => {
        if (numSortConfig) {
            return numFilteredAndSortedRows;
        }

        if (!sortConfig.key) return numFilteredAndSortedRows;

        const { key, direction } = sortConfig;

        return [...numFilteredAndSortedRows].sort((a: any, b: any) => {
            const valA = String(a[key] ?? "");
            const valB = String(b[key] ?? "");

            return direction === "asc"
                ? valA.localeCompare(valB)
                : valB.localeCompare(valA);
        });
    }, [numFilteredAndSortedRows, sortConfig, numSortConfig]);

    /* ================= ENABLED COLUMNS ================= */

    const enabledColumns = useMemo(() => {
        return [...columns].filter((c) => c.enabled).sort((a, b) => a.order - b.order);
    }, [columns]);

    /* ================= GROUPING LOGIC ================= */

    const buildGroupedData = useCallback(
        (data: any[], level: number, parentKey = ""): any[] => {
            const groupKey = grouping[level];
            if (!groupKey) return data;

            const map = new Map<string, any[]>();

            for (const row of data) {
                let val = String(row[groupKey] ?? "Others");
                if (groupKey === "Stock_Date") {
                    val = formatDateOnly(row[groupKey]);
                } else if (groupKey === "Created_At") {
                    val = formatDateTime(row[groupKey]);
                }
                if (!map.has(val)) {
                    map.set(val, []);
                }
                map.get(val)!.push(row);
            }

            return Array.from(map.entries()).map(([value, rows]) => ({
                __group: true,
                __key: `${parentKey}${groupKey}:${value}`,
                __value: value,
                __level: level,
                __rows: rows,
            }));
        },
        [grouping]
    );

    const groupedRows = useMemo(() => {
        return buildGroupedData(sortedRows, 0);
    }, [sortedRows, buildGroupedData]);

    const flattenRows = useCallback(
        (list: any[]): any[] => {
            const result: any[] = [];
            const walk = (items: any[]) => {
                for (const r of items) {
                    result.push(r);
                    if (r.__group && expandedKeys.includes(r.__key)) {
                        const children = buildGroupedData(
                            r.__rows,
                            r.__level + 1,
                            `${r.__key} > `
                        );
                        walk(children);
                    }
                }
            };
            walk(list);
            return result;
        },
        [expandedKeys, buildGroupedData]
    );

    const paginatedSourceRows = useMemo(() => {
        return flattenRows(groupedRows);
    }, [groupedRows, flattenRows]);

    const finalRows = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        const end = start + rowsPerPage;
        return paginatedSourceRows.slice(start, end);
    }, [paginatedSourceRows, page, rowsPerPage]);

    /* ================= HEADER FILTER MENU OPTIONS ================= */

    const filterOptions = useMemo(() => {
        if (!activeHeader) return [];

        const rowsForOptions = rawRows.filter((row) => {
            for (const [key, filterVals] of Object.entries(columnFilters)) {
                if (key === activeHeader) continue;
                if (filterVals !== undefined && filterVals.length > 0) {
                    let val = String(row[key] ?? "");
                    if (key === "Stock_Date") {
                        val = formatDateOnly(row[key]);
                    } else if (key === "Created_At") {
                        val = formatDateTime(row[key]);
                    }
                    if (!filterVals.includes(val)) return false;
                }
            }
            return true;
        });

        const uniqueVals = new Set<string>();
        rowsForOptions.forEach((r) => {
            let v = r[activeHeader];
            if (activeHeader === "Stock_Date") {
                v = formatDateOnly(v);
            } else if (activeHeader === "Created_At") {
                v = formatDateTime(v);
            }
            if (v !== null && v !== undefined && String(v).trim() !== "") {
                uniqueVals.add(String(v));
            }
        });

        return Array.from(uniqueVals);
    }, [activeHeader, rawRows, columnFilters]);

    /* ================= HANDLERS ================= */

    const handleHeaderClick = (e: React.MouseEvent<HTMLElement>, key: string) => {
        const col = columns.find((x) => x.key === key);
        if (col?.isNumeric) return; // handled by numerical filter hook
        setActiveHeader(key);
        setFilterAnchor(e.currentTarget);
    };

    const handleSortClick = (e: React.MouseEvent, key: string) => {
        e.stopPropagation();
        setSortConfig((prev) => {
            if (prev.key === key) {
                return { key, direction: prev.direction === "asc" ? "desc" : "asc" };
            }
            return { key, direction: "asc" };
        });
    };

    /* ================= EXPORT HANDLERS ================= */

    const handleExportExcel = () => {
        const headers = ["S.No", "Stock Group", ...enabledColumns.map((c) => c.label)];
        let serial = 1;
        const rows: any[][] = [];

        paginatedSourceRows.forEach((row: any) => {
            if (row.__group) {
                const groupRow: any[] = [
                    "",
                    row.__value,
                ];

                enabledColumns.forEach((c) => {
                    if (c.isNumeric) {
                        const sum = row.__rows.reduce(
                            (s: number, r: any) => s + Number(r[c.key] || 0),
                            0
                        );
                        groupRow.push(formatOneDecimal(sum));
                    } else {
                        groupRow.push(row.__rows.length);
                    }
                });
                rows.push(groupRow);
            } else {
                const normalRow: any[] = [serial++, ""];
                enabledColumns.forEach((c) => {
                    if (c.isNumeric) {
                        normalRow.push(formatOneDecimal(row[c.key]));
                    } else if (c.key === "Stock_Date") {
                        normalRow.push(formatDateOnly(row[c.key]));
                    } else if (c.key === "Created_At") {
                        normalRow.push(formatDateTime(row[c.key]));
                    } else {
                        normalRow.push(row[c.key] ?? "-");
                    }
                });
                rows.push(normalRow);
            }
        });

        const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Stock Difference");
        XLSX.writeFile(
            workbook,
            `Stock_Difference_Report_${fromDate}_${toDate}.xlsx`
        );
        toast.success("Excel Exported ✅");
    };

    const handleExportPDF = () => {
        const doc = new jsPDF("l", "pt", "a4");
        const headers = ["S.No", "Stock Group", ...enabledColumns.map((c) => c.label)];
        let serial = 1;

        const body: any[][] = paginatedSourceRows.map((row: any) => {
            if (row.__group) {
                return [
                    "",
                    row.__value,
                    ...enabledColumns.map((c) => {
                        if (c.isNumeric) {
                            const sum = row.__rows.reduce(
                                (s: number, r: any) => s + Number(r[c.key] || 0),
                                0
                            );
                            return formatOneDecimal(sum);
                        }
                        return String(row.__rows.length);
                    }),
                ];
            } else {
                return [
                    String(serial++),
                    "",
                    ...enabledColumns.map((c) => {
                        let val = row[c.key];
                        if (c.isNumeric) return formatOneDecimal(val);
                        if (c.key === "Stock_Date") return formatDateOnly(val);
                        if (c.key === "Created_At") return formatDateTime(val);
                        return String(val ?? "-");
                    }),
                ];
            }
        });

        const godownName = selectedGodown ? ` | Godown: ${selectedGodown.Godown_Name}` : "";
        doc.setFontSize(12);
        doc.text(
            `Stock Difference Report (${formatDateOnly(fromDate)} to ${formatDateOnly(toDate)}${godownName})`,
            40,
            25
        );

        autoTable(doc, {
            head: [headers],
            body: body,
            startY: 35,
            theme: "grid",
            styles: {
                fontSize: 8,
                cellPadding: 4,
                lineColor: [203, 213, 225],
                lineWidth: 0.5,
            },
            headStyles: {
                fillColor: [30, 58, 138],
                textColor: [255, 255, 255],
                fontStyle: "bold",
            },
        });

        doc.save(`Stock_Difference_Report_${fromDate}_${toDate}.pdf`);
        toast.success("PDF Exported ✅");
    };

    /* ================= RENDER ================= */

    return (
        <>
            <PageHeader
                parentReportName="STOCK DIFFERENCE REPORT"
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                onReportChange={(template) => {
                    if (!template || !template.Report_Id) {
                        handleClearTemplate();
                    } else {
                        setSelectedTemplateId(template.Report_Id);
                        setReportName(template.Report_Name || "");
                        loadTemplate(template.Report_Id);
                    }
                }}
                onQuickSave={(parentName) => {
                    setParentReportName(parentName || "STOCK DIFFERENCE REPORT");
                    if (!selectedTemplateId) {
                        setReportName("");
                    }
                    setSaveDialogOpen(true);
                }}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => {
                    setTempFromDate(fromDate);
                    setTempToDate(toDate);
                    setTempSelectedGodown(selectedGodown);
                    setDrawerOpen((p) => !p);
                }}
                onClose={() => setDrawerOpen(false)}
                fromDate={tempFromDate}
                toDate={tempToDate}
                onFromDateChange={(val) => setTempFromDate(val)}
                onToDateChange={(val) => setTempToDate(val)}
                onApply={handleApplyFilters}
            >
                {/* GODOWN DROPDOWN FILTER */}
                <Box mb={2}>
                    <Typography
                        variant="caption"
                        sx={{ fontWeight: 600, color: "#1E3A8A", display: "block", mb: 0.5 }}
                    >
                        Godown
                    </Typography>
                    <Autocomplete
                        options={godowns}
                        size="small"
                        getOptionLabel={(option: GodownMasterItem) => option?.Godown_Name || ""}
                        isOptionEqualToValue={(option: GodownMasterItem, val: GodownMasterItem) =>
                            val != null && String(option?.Godown_Id) === String(val?.Godown_Id)
                        }
                        filterOptions={(options, { inputValue }) => {
                            const search = inputValue
                                .toLowerCase()
                                .replace(/\s+/g, "")
                                .replace(/[^a-z0-9]/gi, "");

                            return options.filter((option: GodownMasterItem) => {
                                const label = (option?.Godown_Name || "")
                                    .toLowerCase()
                                    .replace(/\s+/g, "")
                                    .replace(/[^a-z0-9]/gi, "");

                                return label.includes(search);
                            });
                        }}
                        value={tempSelectedGodown}
                        onChange={(_, newValue: GodownMasterItem | null) => {
                            setTempSelectedGodown(newValue);
                        }}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                placeholder="Select Godown..."
                                size="small"
                            />
                        )}
                    />
                </Box>
            </ReportFilterDrawer>

            <Box sx={{ flex: 1, minHeight: 0, width: "100%", maxWidth: "100%", display: "flex", flexDirection: "column", p: 1, overflow: "hidden", boxSizing: "border-box" }}>
                {loading && (
                    <Box display="flex" justifyContent="flex-end" mb={1}>
                        <CircularProgress size={20} color="primary" />
                    </Box>
                )}

                <TableContainer
                    component={Paper}
                    elevation={1}
                    sx={{
                        borderRadius: 1,
                        border: "1px solid #cbd5e1",
                        overflowY: "auto",
                        overflowX: "hidden",
                        maxHeight: "calc(100vh - 120px)",
                        width: "100%",
                        flex: 1,
                    }}
                >
                    <Table size="small" stickyHeader sx={{ tableLayout: "fixed", width: "100%" }}>
                        <TableHead
                            sx={{
                                background: "#1E3A8A",
                                position: "sticky",
                                top: 0,
                                zIndex: 3,
                            }}
                        >
                            <TableRow>
                                <TableCell
                                    sx={{
                                        color: "#fff",
                                        fontSize: "0.72rem",
                                        fontWeight: 700,
                                        bgcolor: "#1E3A8A",
                                        borderRight: "1px solid #2448b2",
                                        width: "3%",
                                        textAlign: "center",
                                        p: "4px 2px",
                                        boxSizing: "border-box",
                                    }}
                                >
                                    SNo
                                </TableCell>
                                {enabledColumns.map((c) => (
                                    <TableCell
                                        key={c.key}
                                        align={c.align || (c.isNumeric ? "right" : "left")}
                                        sx={{
                                            color: "#fff",
                                            cursor: "pointer",
                                            fontWeight: 700,
                                            fontSize: "0.72rem",
                                            bgcolor: "#1E3A8A",
                                            borderRight: "1px solid #2448b2",
                                            width: c.width,
                                            p: "4px 4px",
                                            boxSizing: "border-box",
                                        }}
                                        onClick={(e) => {
                                            if (c.isNumeric) {
                                                openNumFilter(e, c.key);
                                            } else {
                                                handleHeaderClick(e, c.key);
                                            }
                                        }}
                                    >
                                        {c.isNumeric ? (
                                            <SortableHeaderLabel
                                                label={c.label}
                                                columnKey={c.key}
                                                sortConfig={numSortConfig}
                                                onSort={handleNumSort}
                                                onOpenFilter={(e) => openNumFilter(e, c.key)}
                                            />
                                        ) : (
                                            <Box display="flex" alignItems="center" justifyContent="space-between">
                                                <Box sx={{ display: "flex", alignItems: "center" }}>
                                                    {c.label}
                                                </Box>
                                                <IconButton
                                                    size="small"
                                                    sx={{ color: "#fff", p: 0 }}
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        handleSortClick(e, c.key);
                                                    }}
                                                >
                                                    {sortConfig.key === c.key ? (
                                                        sortConfig.direction === "asc" ? (
                                                            <ArrowDropDownIcon fontSize="small" />
                                                        ) : (
                                                            <ArrowDropUpIcon fontSize="small" />
                                                        )
                                                    ) : (
                                                        <ArrowDropDownIcon
                                                            fontSize="small"
                                                            sx={{ opacity: 0.3 }}
                                                        />
                                                    )}
                                                </IconButton>
                                            </Box>
                                        )}
                                    </TableCell>
                                ))}
                            </TableRow>
                        </TableHead>

                        <TableBody>
                            {(() => {
                                serialRef.current = (page - 1) * rowsPerPage;

                                if (finalRows.length === 0) {
                                    return (
                                        <TableRow>
                                            <TableCell
                                                colSpan={enabledColumns.length + 1}
                                                align="center"
                                                sx={{ py: 3, color: "#64748b" }}
                                            >
                                                No stock difference records found.
                                            </TableCell>
                                        </TableRow>
                                    );
                                }

                                return finalRows.map((row: any, i) => {
                                    if (row.__group) {
                                        const expanded = expandedKeys.includes(row.__key);
                                        const rowCount = row.__rows.length;

                                        return (
                                            <TableRow
                                                key={row.__key}
                                                sx={{
                                                    background: "#F1F5F9",
                                                    "&:hover": { bgcolor: "#e2e8f0" },
                                                }}
                                            >
                                                <TableCell
                                                    sx={{
                                                        borderRight: "1px solid #cbd5e1",
                                                        width: "3%",
                                                        py: 0.5,
                                                        px: 0.5,
                                                        textAlign: "center",
                                                        boxSizing: "border-box",
                                                    }}
                                                >
                                                    <IconButton
                                                        size="small"
                                                        onClick={() =>
                                                            setExpandedKeys((p) =>
                                                                p.includes(row.__key)
                                                                    ? p.filter((x) => x !== row.__key)
                                                                    : [...p, row.__key]
                                                            )
                                                        }
                                                    >
                                                        {expanded ? (
                                                            <ExpandMoreIcon fontSize="small" />
                                                        ) : (
                                                            <ChevronRightIcon fontSize="small" />
                                                        )}
                                                    </IconButton>
                                                </TableCell>

                                                {/* Group header cell spanning the first column or showing group name */}
                                                {enabledColumns.map((c, colIdx) => {
                                                    if (colIdx === 0) {
                                                        return (
                                                            <TableCell
                                                                key={c.key}
                                                                sx={{
                                                                    fontWeight: 700,
                                                                    borderRight: "1px solid #cbd5e1",
                                                                    py: 0.5,
                                                                    px: 1,
                                                                    fontSize: "0.75rem",
                                                                    overflow: "hidden",
                                                                    textOverflow: "ellipsis",
                                                                    whiteSpace: "nowrap",
                                                                    color: "#1E3A8A",
                                                                }}
                                                            >
                                                                {row.__value}
                                                            </TableCell>
                                                        );
                                                    }

                                                    if (c.isNumeric) {
                                                        const sum = row.__rows.reduce(
                                                            (s: number, r: any) => s + Number(r[c.key] || 0),
                                                            0
                                                        );
                                                        const isDiff = c.key === "Difference_Qty";
                                                        const diffColor = sum < 0 ? "#dc2626" : sum > 0 ? "#16a34a" : "inherit";

                                                        return (
                                                            <TableCell
                                                                key={c.key}
                                                                align="right"
                                                                sx={{
                                                                    fontWeight: 700,
                                                                    borderRight: "1px solid #cbd5e1",
                                                                    py: 0.5,
                                                                    px: 1,
                                                                    fontSize: "0.75rem",
                                                                    color: isDiff ? diffColor : "inherit",
                                                                }}
                                                            >
                                                                {formatOneDecimal(sum)}
                                                            </TableCell>
                                                        );
                                                    }

                                                    return (
                                                        <TableCell
                                                            key={c.key}
                                                            align={c.align || "left"}
                                                            sx={{
                                                                borderRight: "1px solid #cbd5e1",
                                                                py: 0.5,
                                                                px: 1,
                                                                fontSize: "0.75rem",
                                                                color: "#475569",
                                                            }}
                                                        >
                                                            {rowCount}
                                                        </TableCell>
                                                    );
                                                })}
                                            </TableRow>
                                        );
                                    }

                                    // Regular row
                                    return (
                                        <TableRow
                                            key={row.SD_Id ?? i}
                                            sx={{
                                                "&:hover": { bgcolor: "#f8fafc" },
                                            }}
                                        >
                                            <TableCell
                                                sx={{
                                                    borderRight: "1px solid #cbd5e1",
                                                    fontSize: "0.72rem",
                                                    py: 0.6,
                                                    px: 0.5,
                                                    textAlign: "center",
                                                    width: "3%",
                                                    boxSizing: "border-box",
                                                }}
                                            >
                                                {++serialRef.current}
                                            </TableCell>

                                            {enabledColumns.map((c) => {
                                                const align = c.align || (c.isNumeric ? "right" : "left");
                                                let displayVal: React.ReactNode = row[c.key] ?? "-";
                                                let cellColor = "inherit";

                                                // Group name should NOT be repeated on inner detail rows
                                                if (c.key === "Stock_Group") {
                                                    displayVal = "";
                                                } else if (c.isNumeric) {
                                                    const numVal = Number(row[c.key] || 0);
                                                    displayVal = formatOneDecimal(numVal);
                                                    if (c.key === "Difference_Qty") {
                                                        cellColor = numVal < 0 ? "#dc2626" : numVal > 0 ? "#16a34a" : "inherit";
                                                    }
                                                } else if (c.key === "Stock_Date") {
                                                    displayVal = formatDateOnly(row[c.key]);
                                                } else if (c.key === "Created_At") {
                                                    displayVal = formatDateTime(row[c.key]);
                                                }

                                                return (
                                                    <TableCell
                                                        key={c.key}
                                                        align={align}
                                                        sx={{
                                                            borderRight: "1px solid #cbd5e1",
                                                            fontSize: "0.75rem",
                                                            py: 0.6,
                                                            px: 1,
                                                            color: cellColor,
                                                            fontWeight: c.key === "Difference_Qty" ? 600 : 400,
                                                            wordBreak: "break-word",
                                                        }}
                                                    >
                                                        {displayVal}
                                                    </TableCell>
                                                );
                                            })}
                                        </TableRow>
                                    );
                                });
                            })()}
                        </TableBody>
                    </Table>
                </TableContainer>

                <CommonPagination
                    totalRows={paginatedSourceRows.length}
                    page={page}
                    rowsPerPage={rowsPerPage}
                    onPageChange={setPage}
                    onRowsPerPageChange={setRowsPerPage}
                />
            </Box>

            {/* HEADER FILTER MENU */}
            <HeaderFilterMenu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor)}
                onClose={() => setFilterAnchor(null)}
                columnLabel={columns.find((c) => c.key === activeHeader)?.label || ""}
                options={filterOptions}
                selectedValues={activeHeader ? columnFilters[activeHeader] : undefined}
                onFilterChange={(newSelected) => {
                    if (!activeHeader) return;
                    setColumnFilters((prev) => ({
                        ...prev,
                        [activeHeader]: newSelected && newSelected.length ? newSelected : undefined,
                    }));
                    setPage(1);
                }}
            />

            {/* NUMERICAL FILTER MENU */}
            <NumericalFilterMenu
                anchorEl={numFilterAnchor}
                open={Boolean(numFilterAnchor)}
                onClose={() => setNumFilterAnchor(null)}
                activeHeader={numActiveHeader}
                min={numActiveHeader ? getMinMax(numActiveHeader).min : 0}
                max={numActiveHeader ? getMinMax(numActiveHeader).max : 100}
                rangeFilter={numRangeFilter}
                onRangeChange={(key, range) => setNumRangeFilter(p => ({ ...p, [key]: range }))}
                onClear={clearRangeFilter}
            />

            {/* SAVE TEMPLATE DIALOG */}
            <Dialog
                open={saveDialogOpen}
                onClose={() => setSaveDialogOpen(false)}
                maxWidth="xs"
                fullWidth
            >
                <DialogTitle sx={{ color: "#1E3A8A", fontWeight: "bold" }}>
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
                        placeholder="e.g. Retail Godown Difference"
                        sx={{ mt: 1 }}
                    />
                    <Box mt={1.5} p={1} bgcolor="#f8fafc" borderRadius={1} border="1px solid #e2e8f0">
                        <Typography variant="caption" color="textSecondary" display="block">
                            Filtered Godown:
                        </Typography>
                        <Typography variant="body2" fontWeight={600} color="#1E3A8A">
                            {selectedGodown ? selectedGodown.Godown_Name : "All / Default"}
                        </Typography>
                    </Box>
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setSaveDialogOpen(false)} color="inherit">
                        Cancel
                    </Button>
                    <Button
                        onClick={handleSaveTemplate}
                        variant="contained"
                        sx={{ bgcolor: "#1E3A8A", "&:hover": { bgcolor: "#152a66" } }}
                    >
                        Save
                    </Button>
                </DialogActions>
            </Dialog>
        </>
    );
};

export default StockDifferenceReport;
