import React, { useEffect, useState, useMemo, useCallback } from "react";
import {
    Box,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Paper,
    Menu,
    MenuItem,
    TextField,
    Button,
    IconButton,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    CircularProgress,
    Slider,
    Typography,
    Tooltip,
    Switch,
} from "@mui/material";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import SyncIcon from "@mui/icons-material/Sync";
import ArrowUpwardIcon from "@mui/icons-material/ArrowUpward";
import ArrowDownwardIcon from "@mui/icons-material/ArrowDownward";
import GroupWorkIcon from "@mui/icons-material/GroupWork";
import SettingsIcon from "@mui/icons-material/Settings";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import FilterAltIcon from "@mui/icons-material/FilterAlt";

import {
    DndContext,
    closestCenter,
    PointerSensor,
    useSensor,
    useSensors,
} from "@dnd-kit/core";
import {
    SortableContext,
    verticalListSortingStrategy,
    useSortable,
    arrayMove,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";

import { toast } from "react-toastify";
import dayjs from "dayjs";
import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import { mapForExport } from "../../utils/exportMapper";
import {
    AdminUnitEconomicsReportService,
    AdminUnitEconomicsReport,
} from "../../services/unitEconomicsReport.service";
import {
    lollosColumnsService,
    LolLosColumn,
} from "../../services/lollosColumns.service";
import { SettingsService } from "../../services/reportSettings.services";

/* ================= TYPES ================= */

type ColumnConfig = {
    key: string;
    label: string;
    enabled: boolean;
    isNumeric?: boolean;
    order: number;
    groupBy?: number;
};

/* ================= CONSTANTS ================= */

const headStyle = {
    color: "#fff",
    fontWeight: 600,
    fontSize: "0.75rem",
    cursor: "pointer",
};

const bodyStyle = {
    fontSize: "0.68rem",
};

const DEFAULT_KEYS_COGS = [
    "invoice_no",
    "Ledger_Date",
    "Retailer_Name",
    "Product_Name",
    "Bill_Qty",
    "Rate",
    "Amount",
    "COGS_Rate",
    "GP_COGS",
    "TGP_COGS",
    "GP_Percentage_COGS",
];

const DEFAULT_KEYS_MIN = [
    "invoice_no",
    "Ledger_Date",
    "Retailer_Name",
    "Product_Name",
    "Bill_Qty",
    "Rate",
    "Amount",
    "Min_Rate",
    "GP_MR",
    "TGP_MR",
    "GP_Percentage_MR",
];

const NUMERIC_KEYS = [
    "Bill_Qty",
    "Rate",
    "Amount",
    "Min_Rate",
    "COGS_Rate",
    "GP_MR",
    "GP_COGS",
    "TGP_MR",
    "TGP_COGS",
    "GP_Percentage_MR",
    "GP_Percentage_COGS",
];

const CURRENCY_KEYS = [
    "Rate",
    "Amount",
    "Min_Rate",
    "COGS_Rate",
    "GP_MR",
    "GP_COGS",
    "TGP_MR",
    "TGP_COGS",
];

const PERCENTAGE_KEYS = [
    "GP_Percentage_MR",
    "GP_Percentage_COGS",
];

const REPORT_SP = "Reporting_Admin_Unit_Economics";
const PARENT_REPORT_NAME = "Unit Economics Admin Report";

/* ================= SORTABLE ROW ================= */

type SortableColumnRowProps = {
    column: ColumnConfig;
    onToggle: (key: string) => void;
    hasActiveFilter?: boolean;
};

const SortableColumnRow = ({
    column,
    onToggle,
    hasActiveFilter,
}: SortableColumnRowProps) => {
    const {
        attributes,
        listeners,
        setNodeRef,
        transform,
        transition,
    } = useSortable({ id: column.key });

    const style = {
        transform: CSS.Transform.toString(transform),
        transition,
    };

    return (
        <Box
            ref={setNodeRef}
            style={style}
            display="flex"
            alignItems="center"
            gap={1}
            mb={1}
        >
            <IconButton
                size="small"
                {...listeners}
                {...attributes}
                sx={{ cursor: "grab", p: 0.5 }}
            >
                <DragIndicatorIcon fontSize="small" />
            </IconButton>

            <Box display="flex" alignItems="center" gap={1} sx={{ flex: 1, minWidth: 0 }}>
                <Typography fontSize="0.75rem" noWrap>
                    {column.label}
                </Typography>

                {hasActiveFilter && (
                    <Tooltip title="Header filter enabled">
                        <FilterAltIcon fontSize="small" color="action" />
                    </Tooltip>
                )}
            </Box>

            <Switch
                size="small"
                checked={column.enabled}
                onChange={() => onToggle(column.key)}
            />
        </Box>
    );
};

/* ================= COMPONENT ================= */

const UnitEconomicsAdmin: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    const [data, setData] = useState<AdminUnitEconomicsReport[]>([]);
    const [rawApiData, setRawApiData] = useState<AdminUnitEconomicsReport[]>([]);
    const [lolColumns, setLolColumns] = useState<LolLosColumn[]>([]);
    const [columns, setColumns] = useState<ColumnConfig[]>([]);

    const [loading, setLoading] = useState(false);
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    /* -------- FILTERS -------- */
    const [filters, setFilters] = useState<{
        Date: { from: string; to: string };
    }>({
        Date: { from: today, to: today },
    });

    const [columnFilters, setColumnFilters] = useState<Record<string, string[] | undefined>>({});
    const [tempDate, setTempDate] = useState(filters.Date);

    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [rateType, setRateType] = useState<"cogs" | "min">("cogs");

    /* -------- HEADER FILTER -------- */
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);

    /* -------- SORTING -------- */
    const [sortConfig, setSortConfig] = useState<{
        key: string;
        direction: "asc" | "desc";
    } | null>(null);

    /* -------- RANGE FILTER -------- */
    const [rangeFilter, setRangeFilter] = useState<Record<string, [number, number]>>({});

    /* -------- COLUMN SETTINGS -------- */
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);

    /* -------- GROUP BY -------- */
    const [groupDialogOpen, setGroupDialogOpen] = useState(false);
    const [grouping, setGrouping] = useState<string[]>([]);
    const [pendingGrouping, setPendingGrouping] = useState<string[]>([]);
    const [expandedKeys, setExpandedKeys] = useState<string[]>([]);

    /* -------- TEMPLATES -------- */
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [templateLoading, setTemplateLoading] = useState(false);
    const [saveDialogOpen, setSaveDialogOpen] = useState(false);
    const [reportName, setReportName] = useState("");
    const [parentReportName, setParentReportName] = useState(PARENT_REPORT_NAME);

    /* -------- SYNC DIALOG STATE -------- */
    const [syncDialogOpen, setSyncDialogOpen] = useState(false);
    const [syncFromDate, setSyncFromDate] = useState(today);
    const [syncToDate, setSyncToDate] = useState(today);
    const [isSyncing, setIsSyncing] = useState(false);

    /* ================= DND SENSORS ================= */
    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: {
                distance: 5,
            },
        })
    );

    const isNumericColumn = (key: string) => {
        return NUMERIC_KEYS.includes(key);
    };

    const getMinMax = (key: string) => {
        const nums = rawApiData
            .map((r: any) => Number(r[key]))
            .filter((v) => !isNaN(v));
        if (nums.length === 0) return { min: 0, max: 100 };
        const minVal = Math.min(...nums);
        const maxVal = Math.max(...nums);
        return {
            min: minVal,
            max: minVal === maxVal ? minVal + 1 : maxVal,
        };
    };

    const handleSort = (columnKey: string) => {
        setSortConfig((prev) => {
            if (prev && prev.key === columnKey) {
                return {
                    key: columnKey,
                    direction: prev.direction === "asc" ? "desc" : "asc",
                };
            }
            return {
                key: columnKey,
                direction: "asc",
            };
        });
    };

    /* ================= SYNC HANDLERS ================= */
    const handleOpenSyncDialog = () => {
        setSyncFromDate(filters.Date.from);
        setSyncToDate(filters.Date.to);
        setSyncDialogOpen(true);
    };

    const handleSync = async () => {
        try {
            setIsSyncing(true);
            const res = await AdminUnitEconomicsReportService.syncReports({
                Fromdate: syncFromDate,
                Todate: syncToDate,
            });

            if (res.data?.success || res.status === 200) {
                toast.success("synced successfully");
                setSyncDialogOpen(false);

                setFilters((prev) => ({
                    ...prev,
                    Date: { from: syncFromDate, to: syncToDate }
                }));
                setFromDate(syncFromDate);
                setToDate(syncToDate);
                setTempDate({ from: syncFromDate, to: syncToDate });
            } else {
                toast.error("Sync failed ❌");
            }
        } catch (error: any) {
            console.error("Sync error", error);
            toast.error(error.response?.data?.message || "Sync failed ❌");
        } finally {
            setIsSyncing(false);
        }
    };

    /* ================= LOAD LOL COLUMNS ================= */
    useEffect(() => {
        lollosColumnsService.getLolColumns()
            .then(res => {
                const list = res.data?.data || [];
                setLolColumns(list);
            })
            .catch(err => {
                console.error("Failed to load LOL columns", err);
            });
    }, []);

    /* ================= BUILD COLUMNS HELPER ================= */
    const buildInitialColumns = useCallback((
        rows: any[],
        lolCols: LolLosColumn[],
        rateTypeMode: "cogs" | "min"
    ): ColumnConfig[] => {
        const defaultEnabled = rateTypeMode === "cogs" ? DEFAULT_KEYS_COGS : DEFAULT_KEYS_MIN;

        const baseKeyConfigs: { key: string; label: string; isNumeric: boolean }[] = [
            { key: "invoice_no", label: "Invoice No", isNumeric: false },
            { key: "Ledger_Date", label: "Date", isNumeric: false },
            { key: "Retailer_Name", label: "Retailer Name", isNumeric: false },
            { key: "Product_Name", label: "Product", isNumeric: false },
            { key: "Bill_Qty", label: "Quantity", isNumeric: true },
            { key: "Rate", label: "Rate", isNumeric: true },
            { key: "Amount", label: "Amount", isNumeric: true },
            { key: "COGS_Rate", label: "COGS Rate", isNumeric: true },
            { key: "GP_COGS", label: "GP COGS", isNumeric: true },
            { key: "TGP_COGS", label: "TGP COGS", isNumeric: true },
            { key: "GP_Percentage_COGS", label: "GP % COGS", isNumeric: true },
            { key: "Min_Rate", label: "MR Rate", isNumeric: true },
            { key: "GP_MR", label: "GP MR", isNumeric: true },
            { key: "TGP_MR", label: "TGP MR", isNumeric: true },
            { key: "GP_Percentage_MR", label: "GP % MR", isNumeric: true },
        ];

        const result: ColumnConfig[] = [];
        const seenKeys = new Set<string>();

        // 1. Add current predefined columns (enabled according to current rateType)
        baseKeyConfigs.forEach((b, idx) => {
            const isEnabled = defaultEnabled.includes(b.key);
            result.push({
                key: b.key,
                label: b.label,
                enabled: isEnabled,
                isNumeric: b.isNumeric,
                order: idx,
                groupBy: 0,
            });
            seenKeys.add(b.key.toUpperCase());
        });

        // 2. Add other columns present in API rows as disabled columns
        if (rows.length > 0) {
            Object.keys(rows[0]).forEach(key => {
                if (!seenKeys.has(key.toUpperCase())) {
                    const matchedLol = lolCols.find(
                        l => l.ColumnName && l.ColumnName.toUpperCase() === key.toUpperCase()
                    );
                    const label = (matchedLol?.Alias_Name && matchedLol.Alias_Name.trim()) || matchedLol?.ColumnName || key;

                    const isNum = NUMERIC_KEYS.includes(key) || rows.some(r => typeof r[key] === "number");

                    result.push({
                        key,
                        label,
                        enabled: false,
                        isNumeric: isNum,
                        order: result.length,
                        groupBy: 0,
                    });
                    seenKeys.add(key.toUpperCase());
                }
            });
        }

        // 3. Add remaining columns from lolfields dataset as disabled columns
        lolCols.forEach(l => {
            if (l.ColumnName && !seenKeys.has(l.ColumnName.toUpperCase())) {
                const label = (l.Alias_Name && l.Alias_Name.trim()) || l.ColumnName;

                result.push({
                    key: l.ColumnName,
                    label,
                    enabled: false,
                    isNumeric: NUMERIC_KEYS.includes(l.ColumnName),
                    order: result.length,
                    groupBy: 0,
                });
                seenKeys.add(l.ColumnName.toUpperCase());
            }
        });

        return result;
    }, []);

    /* ================= LOAD DATA ================= */
    useEffect(() => {
        const loadData = async () => {
            try {
                setLoading(true);

                let currentLol = lolColumns;
                if (currentLol.length === 0) {
                    try {
                        const lolColsRes = await lollosColumnsService.getLolColumns();
                        currentLol = lolColsRes.data?.data || [];
                        setLolColumns(currentLol);
                    } catch (e) {
                        console.error("Failed to load lol columns", e);
                    }
                }

                const res = await AdminUnitEconomicsReportService.getReports({
                    Fromdate: filters.Date.from,
                    Todate: filters.Date.to,
                });

                const responseData = res.data?.data || {};
                const rows: AdminUnitEconomicsReport[] = responseData.rows || [];
                const lolFields: any[] = responseData.lolFields || [];

                // Create map from lolFields by Ret_Id
                const lolMap = new Map<string, any>();
                lolFields.forEach((lf: any) => {
                    const id = lf.Ret_Id ?? lf.Retailer_Id ?? lf.Retailer_ID ?? lf.Id ?? lf.Auto_Id;
                    if (id !== undefined && id !== null && String(id).trim() !== "") {
                        lolMap.set(String(id).trim(), lf);
                    }
                });

                // Map Retailer_Id in rows with lolFields dataset
                const mappedRows: AdminUnitEconomicsReport[] = rows.map((row: any) => {
                    const retailerId = row.Retailer_Id ?? row.Retailer_ID ?? row.Ret_Id;
                    const lolData = (retailerId !== undefined && retailerId !== null && String(retailerId).trim() !== "")
                        ? lolMap.get(String(retailerId).trim())
                        : undefined;

                    if (!lolData) return { ...row };

                    const merged: any = {
                        ...lolData,
                        ...row,
                    };

                    // Make sure all lolColumns are populated under both ColumnName and Alias_Name
                    currentLol.forEach((lc: LolLosColumn) => {
                        if (lc.ColumnName && lolData[lc.ColumnName] !== undefined) {
                            const val = lolData[lc.ColumnName];
                            merged[lc.ColumnName] = val;
                            if (lc.Alias_Name && lc.Alias_Name.trim()) {
                                merged[lc.Alias_Name.trim()] = val;
                            }
                        }
                    });

                    return merged;
                });

                setRawApiData(mappedRows);
                setPage(1);

                setColumns(prev => {
                    if (prev.length > 0) {
                        // Already initialized, but merge any newly discovered API or lol columns
                        const existingKeys = new Set(prev.map(c => c.key.toUpperCase()));
                        const newCols = [...prev];

                        if (rows.length > 0) {
                            Object.keys(rows[0]).forEach(key => {
                                if (!existingKeys.has(key.toUpperCase())) {
                                    const matchedLol = currentLol.find(
                                        l => l.ColumnName && l.ColumnName.toUpperCase() === key.toUpperCase()
                                    );
                                    newCols.push({
                                        key,
                                        label: (matchedLol?.Alias_Name && matchedLol.Alias_Name.trim()) || matchedLol?.ColumnName || key,
                                        enabled: false,
                                        isNumeric: NUMERIC_KEYS.includes(key) || rows.some(r => typeof (r as any)[key] === "number"),
                                        order: newCols.length,
                                        groupBy: 0,
                                    });
                                    existingKeys.add(key.toUpperCase());
                                }
                            });
                        }

                        currentLol.forEach(l => {
                            if (l.ColumnName && !existingKeys.has(l.ColumnName.toUpperCase())) {
                                newCols.push({
                                    key: l.ColumnName,
                                    label: (l.Alias_Name && l.Alias_Name.trim()) || l.ColumnName,
                                    enabled: false,
                                    isNumeric: NUMERIC_KEYS.includes(l.ColumnName),
                                    order: newCols.length,
                                    groupBy: 0,
                                });
                                existingKeys.add(l.ColumnName.toUpperCase());
                            }
                        });

                        return newCols;
                    }

                    return buildInitialColumns(rows, currentLol, rateType);
                });
            } catch (err) {
                console.error("Failed to load unit economics admin report", err);
                toast.error("Failed to load unit economics report");
            } finally {
                setLoading(false);
            }
        };

        loadData();
    }, [filters.Date, buildInitialColumns, lolColumns, rateType]);

    /* ================= SYNC RATE TYPE WITH COLUMNS ================= */
    useEffect(() => {
        if (selectedTemplateId) return; // Retain user's custom template columns
        setColumns(cols => {
            if (!cols.length) return cols;
            const defaultEnabled = rateType === "cogs" ? DEFAULT_KEYS_COGS : DEFAULT_KEYS_MIN;
            const defaultDisabled = rateType === "cogs"
                ? ["Min_Rate", "GP_MR", "TGP_MR", "GP_Percentage_MR"]
                : ["COGS_Rate", "GP_COGS", "TGP_COGS", "GP_Percentage_COGS"];

            return cols.map(c => {
                if (defaultEnabled.includes(c.key)) {
                    return { ...c, enabled: true };
                }
                if (defaultDisabled.includes(c.key)) {
                    return { ...c, enabled: false };
                }
                return c;
            });
        });
    }, [rateType, selectedTemplateId]);

    /* ================= FILTER AND SORT DATA ================= */
    useEffect(() => {
        let rows = [...rawApiData];

        // Header / Categorical filters
        for (const [key, selected] of Object.entries(columnFilters)) {
            if (selected && selected.length > 0) {
                rows = rows.filter((r: any) =>
                    selected.includes(String(r[key] ?? "").trim())
                );
            }
        }

        // Numerical Range Filter
        if (rangeFilter) {
            for (const [key, range] of Object.entries(rangeFilter)) {
                rows = rows.filter((r: any) => {
                    const val = Number(r[key]);
                    return isNaN(val) || (val >= range[0] && val <= range[1]);
                });
            }
        }

        // Sorting
        if (sortConfig) {
            const { key, direction } = sortConfig;
            const col = columns.find(c => c.key === key);
            rows.sort((a: any, b: any) => {
                const valA = a[key];
                const valB = b[key];

                if (col?.isNumeric || NUMERIC_KEYS.includes(key)) {
                    const numA = Number(valA) || 0;
                    const numB = Number(valB) || 0;
                    return direction === "asc" ? numA - numB : numB - numA;
                }

                if (key === "Ledger_Date" || key === "Trans_Date") {
                    const dateA = dayjs(valA).valueOf();
                    const dateB = dayjs(valB).valueOf();
                    return direction === "asc" ? dateA - dateB : dateB - dateA;
                }

                return direction === "asc"
                    ? String(valA ?? "").localeCompare(String(valB ?? ""), undefined, { sensitivity: "base", numeric: true })
                    : String(valB ?? "").localeCompare(String(valA ?? ""), undefined, { sensitivity: "base", numeric: true });
            });
        }

        setData(rows);
        setPage(1);
    }, [rawApiData, columnFilters, rangeFilter, sortConfig, columns]);

    /* ================= ENABLED & ORDERED COLUMNS ================= */
    const sortedColumns = useMemo(() => {
        return [...columns].sort((a, b) => a.order - b.order);
    }, [columns]);

    const enabledColumns = useMemo(() => {
        return sortedColumns.filter((c) => c.enabled);
    }, [sortedColumns]);

    /* ================= GROUPING LOGIC ================= */
    const buildGroupedData = useCallback(
        (rowsList: any[], level: number, parentKey = ""): any[] => {
            const groupKey = grouping[level];
            if (!groupKey) return rowsList;

            const map = new Map<string, any[]>();

            for (const row of rowsList) {
                const rawVal = row[groupKey];
                const value = (rawVal !== undefined && rawVal !== null && rawVal !== "")
                    ? String(rawVal)
                    : "Others";
                if (!map.has(value)) {
                    map.set(value, []);
                }
                map.get(value)!.push(row);
            }

            return Array.from(map.entries()).map(([value, gRows]) => ({
                __group: true,
                __key: `${parentKey}${groupKey}:${value}`,
                __value: value,
                __level: level,
                __rows: gRows,
            }));
        },
        [grouping]
    );

    const groupedRows = useMemo(() => {
        if (!grouping.length) return data;
        return buildGroupedData(data, 0);
    }, [data, grouping, buildGroupedData]);

    const flattenRows = useCallback((items: any[]): any[] => {
        const result: any[] = [];

        const walk = (list: any[]) => {
            for (const row of list) {
                result.push(row);

                if (row.__group && expandedKeys.includes(row.__key)) {
                    walk(
                        buildGroupedData(
                            row.__rows,
                            row.__level + 1,
                            `${row.__key} > `
                        )
                    );
                }
            }
        };

        walk(items);
        return result;
    }, [expandedKeys, buildGroupedData]);

    const paginatedSourceRows = useMemo(() => {
        return grouping.length ? flattenRows(groupedRows) : data;
    }, [groupedRows, data, grouping, flattenRows]);

    const paginatedData = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        return paginatedSourceRows.slice(start, start + rowsPerPage);
    }, [paginatedSourceRows, page, rowsPerPage]);

    /* ================= FORMATTERS & TOTALS ================= */
    const formatINR = (value: number) => {
        return new Intl.NumberFormat("en-IN", {
            style: "currency",
            currency: "INR",
        }).format(value);
    };

    const getTotal = (key: string) => {
        return data.reduce((sum, r: any) => sum + (Number(r[key]) || 0), 0);
    };

    const renderCellValue = (row: any, col: ColumnConfig) => {
        let val = row[col.key];

        // Fallback: If not found under col.key, search in lolColumns for matching column name or alias
        if (val === undefined || val === null || val === "") {
            const matched = lolColumns.find(
                (l) =>
                    (l.ColumnName && l.ColumnName.toUpperCase() === col.key.toUpperCase()) ||
                    (l.Alias_Name && l.Alias_Name.trim().toUpperCase() === col.key.toUpperCase())
            );
            if (matched) {
                val = row[matched.ColumnName] ?? (matched.Alias_Name ? row[matched.Alias_Name.trim()] : undefined);
            }
        }

        if (val === undefined || val === null || val === "") return "-";

        if (col.key === "Ledger_Date" || col.key === "Trans_Date" || col.key === "Invoice_Date" || col.key === "Date_Added") {
            const d = dayjs(val);
            return d.isValid() ? d.format("DD/MM/YYYY") : String(val);
        }

        if (PERCENTAGE_KEYS.includes(col.key)) {
            return `${Number(val || 0).toFixed(2)}%`;
        }

        if (CURRENCY_KEYS.includes(col.key)) {
            return formatINR(Number(val || 0));
        }

        if (col.key === "Bill_Qty" || col.isNumeric) {
            const num = Number(val);
            if (!isNaN(num)) {
                return Number.isInteger(num) ? num.toString() : num.toFixed(2);
            }
        }

        return String(val);
    };

    const renderTotalCell = (col: ColumnConfig) => {
        if (!col.isNumeric) return "";

        if (PERCENTAGE_KEYS.includes(col.key)) {
            return "";
        }

        if (CURRENCY_KEYS.includes(col.key)) {
            return formatINR(getTotal(col.key));
        }

        if (col.key === "Bill_Qty") {
            return getTotal(col.key).toFixed(2);
        }

        const total = getTotal(col.key);
        return Number.isInteger(total) ? total.toString() : total.toFixed(2);
    };

    const renderGroupSubtotal = (gRows: any[], col: ColumnConfig) => {
        if (!col.isNumeric || PERCENTAGE_KEYS.includes(col.key)) return "";
        const sum = gRows.reduce((acc, r) => acc + (Number(r[col.key]) || 0), 0);
        if (CURRENCY_KEYS.includes(col.key)) return formatINR(sum);
        if (col.key === "Bill_Qty") return sum.toFixed(2);
        return Number.isInteger(sum) ? sum.toString() : sum.toFixed(2);
    };

    /* ================= HEADER FILTER MENU ================= */
    const openFilter = (e: React.MouseEvent<HTMLElement>, columnKey: string) => {
        setActiveHeader(columnKey);
        setFilterAnchor(e.currentTarget);
    };

    const filterOptions = useMemo(() => {
        if (!activeHeader) return [];
        const values = Array.from(
            new Set(
                rawApiData
                    .map((r: any) => r[activeHeader])
                    .filter((v: any) => v !== null && v !== undefined && v !== "")
                    .map((v: any) => String(v).trim())
            )
        );
        values.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }));
        return values;
    }, [activeHeader, rawApiData]);

    /* ================= COLUMN SETTINGS DND ================= */
    const handleDragEnd = (event: any) => {
        const { active, over } = event;
        if (!over || active.id === over.id) return;

        setColumns(cols => {
            const enabled = cols.filter(c => c.enabled);
            const disabled = cols.filter(c => !c.enabled);
            const activeList = enabled.some(c => c.key === active.id)
                ? enabled
                : disabled;

            const oldIndex = activeList.findIndex(c => c.key === active.id);
            const newIndex = activeList.findIndex(c => c.key === over.id);

            if (oldIndex === -1 || newIndex === -1) return cols;

            const reordered = arrayMove(activeList, oldIndex, newIndex);
            reordered.forEach((c, i) => (c.order = i));

            return enabled.some(c => c.key === active.id)
                ? [...reordered, ...disabled]
                : [...enabled, ...reordered];
        });
    };

    const handleResetSettings = () => {
        setSelectedTemplateId(null);
        setReportName("");
        setGrouping([]);
        setPendingGrouping([]);
        setExpandedKeys([]);
        setSortConfig(null);
        setColumnFilters({});
        setRangeFilter({});
        setColumns(cols => {
            const defaultEnabled = rateType === "cogs" ? DEFAULT_KEYS_COGS : DEFAULT_KEYS_MIN;
            return cols.map(c => ({
                ...c,
                enabled: defaultEnabled.includes(c.key),
                groupBy: 0,
            }));
        });
        setSettingsAnchor(null);
    };

    /* ================= TEMPLATES ================= */
    const applyTemplateToColumns = (
        availableCols: ColumnConfig[],
        templateCols: any[]
    ): ColumnConfig[] => {
        if (!templateCols.length) return availableCols;

        const mergedCols = availableCols.map(col => {
            const tCol = templateCols.find(tc => tc.key === col.key);
            if (tCol) {
                return {
                    ...col,
                    label: tCol.label || col.label,
                    enabled: tCol.enabled,
                    order: tCol.order,
                    groupBy: tCol.groupBy ?? 0,
                    isNumeric: tCol.isNumeric ?? col.isNumeric,
                };
            }
            return {
                ...col,
                enabled: false,
            };
        });

        const missingTemplateColumns = templateCols
            .filter(tc => !availableCols.some(ac => ac.key === tc.key))
            .map(tc => ({
                key: tc.key,
                label: tc.label,
                enabled: tc.enabled,
                order: tc.order,
                groupBy: tc.groupBy ?? 0,
                isNumeric: tc.isNumeric,
            }));

        return [...mergedCols, ...missingTemplateColumns].sort(
            (a, b) => a.order - b.order
        );
    };

    const loadTemplate = async (reportId: number) => {
        try {
            setTemplateLoading(true);
            setSelectedTemplateId(reportId);

            const res = await SettingsService.getReportEditData({
                reportId,
                typeId: 1,
            });

            const templateCols = res.data?.data?.columns || [];
            const reportInfo = res.data?.data?.reportInfo;

            setColumns(cols => applyTemplateToColumns(cols, templateCols));

            const autoGroups = templateCols
                .filter((c: any) => c.groupBy > 0 && c.enabled)
                .sort((a: any, b: any) => a.groupBy - b.groupBy)
                .map((c: any) => c.key);

            setGrouping(autoGroups);
            setPendingGrouping(autoGroups);
            setExpandedKeys([]);

            if (reportInfo?.Report_Name) {
                setReportName(reportInfo.Report_Name);
            }
        } catch (err) {
            console.error("Failed to load template", err);
            toast.error("Failed to load template ❌");
        } finally {
            setTemplateLoading(false);
        }
    };

    const handleReportChange = (template: any) => {
        if (!template || !template.Report_Id) {
            handleResetSettings();
            return;
        }
        loadTemplate(template.Report_Id);
    };

    const handleQuickSave = async () => {
        try {
            if (!reportName.trim()) {
                toast.error("Enter Report Name");
                return;
            }

            const userData = JSON.parse(localStorage.getItem("user") || "{}");
            const createdBy = userData.id || 1;

            const payload = columns.map(c => ({
                key: c.key,
                label: c.label,
                order: c.order,
                groupBy: grouping.includes(c.key) ? grouping.indexOf(c.key) + 1 : 0,
                enabled: c.enabled,
                isNumeric: c.isNumeric,
            }));

            if (selectedTemplateId) {
                await SettingsService.updateReport({
                    reportId: selectedTemplateId,
                    typeId: 1,
                    reportName: reportName.trim(),
                    columns: payload,
                });
                toast.success("Template Updated Successfully ✅");
            } else {
                await SettingsService.saveReportSettings({
                    reportName: reportName.trim(),
                    parentReport: parentReportName,
                    abstractSP: REPORT_SP,
                    expandedSP: REPORT_SP,
                    abstractColumns: payload,
                    expandedColumns: payload,
                    createdBy,
                });
                toast.success("Template Saved Successfully ✅");
            }

            setSaveDialogOpen(false);
            setTimeout(() => {
                window.location.reload();
            }, 500);
        } catch (err) {
            console.error("Save template failed", err);
            toast.error("Save Failed ❌");
        }
    };

    /* ================= EXPORT ================= */
    const exportColumns = useMemo(() => {
        return enabledColumns.map(c => ({
            key: c.key,
            label: c.label,
            type: c.isNumeric ? "number" : (c.key === "Ledger_Date" ? "date" : "string")
        }));
    }, [enabledColumns]);

    const handleExportPDF = () => {
        const { headers, data: exportData } = mapForExport(exportColumns, data);
        exportToPDF("Unit Economics Admin Report", headers, exportData);
    };

    const handleExportExcel = () => {
        const { headers, data: exportData } = mapForExport(exportColumns, data);
        exportToExcel("Unit Economics Admin Report", headers, exportData);
    };

    /* ================= RENDER ================= */
    return (
        <>
            <PageHeader
                parentReportName={parentReportName}
                onReportChange={handleReportChange}
                onQuickSave={(parentName) => {
                    setParentReportName(parentName);
                    if (!selectedTemplateId) {
                        setReportName("");
                    }
                    setSaveDialogOpen(true);
                }}
                onExportPDF={handleExportPDF}
                onExportExcel={handleExportExcel}
                infoSlot={
                    <Button
                        variant="contained"
                        size="small"
                        startIcon={<SyncIcon />}
                        onClick={handleOpenSyncDialog}
                        sx={{
                            height: 24,
                            fontSize: "0.7rem",
                            backgroundColor: "#fff",
                            color: "#1E3A8A",
                            fontWeight: 600,
                            textTransform: "none",
                            "&:hover": {
                                backgroundColor: "#f5f5f5",
                            },
                        }}
                    >
                        Sync
                    </Button>
                }
                settingsSlot={
                    <Box display="flex" gap={1}>
                        <Tooltip title="Group By">
                            <IconButton
                                size="small"
                                onClick={() => {
                                    setPendingGrouping(grouping);
                                    setGroupDialogOpen(true);
                                }}
                                sx={{
                                    height: 24,
                                    width: 24,
                                    backgroundColor: "#fff",
                                    borderRadius: 0.5,
                                }}
                            >
                                <GroupWorkIcon fontSize="small" />
                            </IconButton>
                        </Tooltip>

                        <Tooltip title="Column Settings">
                            <IconButton
                                size="small"
                                onClick={(e) => setSettingsAnchor(e.currentTarget)}
                                sx={{
                                    height: 24,
                                    width: 24,
                                    backgroundColor: "#fff",
                                    borderRadius: 0.5,
                                }}
                            >
                                <SettingsIcon fontSize="small" />
                            </IconButton>
                        </Tooltip>
                    </Box>
                }
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((p) => !p)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                toDate={toDate}
                onFromDateChange={setFromDate}
                onToDateChange={setToDate}
                showRateTypeFilter={true}
                rateTypeValue={rateType}
                onRateTypeChange={setRateType}
                onApply={() => {
                    setFilters((prev) => ({
                        ...prev,
                        Date: { from: fromDate, to: toDate }
                    }));
                }}
            />

            <AppLayout fullWidth>
                <Box sx={{ overflow: "auto", mt: 0.5, position: "relative" }}>
                    {templateLoading && (
                        <Box
                            sx={{
                                position: "absolute",
                                top: 0,
                                left: 0,
                                right: 0,
                                bottom: 0,
                                background: "rgba(255,255,255,0.7)",
                                zIndex: 10,
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                            }}
                        >
                            <CircularProgress size={30} />
                        </Box>
                    )}

                    <TableContainer
                        component={Paper}
                        sx={{
                            position: "relative",
                            maxHeight: "calc(100vh - 100px)",
                            overflow: "auto",
                        }}
                    >
                        <Table size="small">
                            {/* ===== FIXED HEADER ===== */}
                            <TableHead
                                sx={{
                                    background: "#1E3A8A",
                                    position: "sticky",
                                    top: 0,
                                    zIndex: 2,
                                }}
                            >
                                <TableRow>
                                    <TableCell sx={headStyle}>S.No</TableCell>
                                    {enabledColumns.map((col) => (
                                        <TableCell
                                            key={col.key}
                                            align={col.isNumeric ? "right" : "left"}
                                            sx={headStyle}
                                            onClick={() => handleSort(col.key)}
                                        >
                                            <Box
                                                sx={{
                                                    display: "inline-flex",
                                                    alignItems: "center",
                                                    justifyContent: col.isNumeric ? "flex-end" : "space-between",
                                                    width: "100%",
                                                }}
                                            >
                                                <span
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        openFilter(e, col.key);
                                                    }}
                                                    style={{ cursor: "pointer" }}
                                                >
                                                    {col.label}
                                                </span>
                                                <Box
                                                    component="span"
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        handleSort(col.key);
                                                    }}
                                                    sx={{
                                                        display: "inline-flex",
                                                        alignItems: "center",
                                                        marginLeft: 0.5,
                                                        width: 16,
                                                        height: 16,
                                                        cursor: "pointer",
                                                        opacity: sortConfig?.key === col.key ? 1 : 0.2,
                                                        transition: "opacity 0.2s",
                                                        "th:hover &": {
                                                            opacity: sortConfig?.key === col.key ? 1 : 0.6,
                                                        },
                                                    }}
                                                >
                                                    {sortConfig?.key === col.key ? (
                                                        sortConfig.direction === "asc" ? (
                                                            <ArrowUpwardIcon sx={{ fontSize: "0.95rem" }} />
                                                        ) : (
                                                            <ArrowDownwardIcon sx={{ fontSize: "0.95rem" }} />
                                                        )
                                                    ) : (
                                                        <ArrowUpwardIcon sx={{ fontSize: "0.95rem" }} />
                                                    )}
                                                </Box>
                                            </Box>
                                        </TableCell>
                                    ))}
                                </TableRow>

                                {/* TOTAL ROW */}
                                <TableRow
                                    sx={{
                                        background: "#f3f4f6",
                                        "& .MuiTableCell-root": {
                                            backgroundColor: "#f3f4f6",
                                            color: "#374151",
                                        },
                                    }}
                                >
                                    <TableCell sx={{ fontWeight: 700 }}>Total</TableCell>
                                    {enabledColumns.map((col) => (
                                        <TableCell
                                            key={col.key}
                                            align={col.isNumeric ? "right" : "left"}
                                            sx={{ fontWeight: 700 }}
                                        >
                                            {renderTotalCell(col)}
                                        </TableCell>
                                    ))}
                                </TableRow>
                            </TableHead>

                            {/* ===== BODY ===== */}
                            <TableBody>
                                {loading ? (
                                    <TableRow>
                                        <TableCell
                                            colSpan={enabledColumns.length + 1}
                                            align="center"
                                            sx={{ py: 3 }}
                                        >
                                            <CircularProgress size={24} />
                                        </TableCell>
                                    </TableRow>
                                ) : paginatedData.length === 0 ? (
                                    <TableRow>
                                        <TableCell
                                            colSpan={enabledColumns.length + 1}
                                            align="center"
                                            sx={{ py: 3 }}
                                        >
                                            No Data Found
                                        </TableCell>
                                    </TableRow>
                                ) : (
                                    paginatedData.map((row: any, i: number) => {
                                        if (row.__group) {
                                            const expanded = expandedKeys.includes(row.__key);
                                            return (
                                                <TableRow
                                                    key={row.__key}
                                                    sx={{ background: "#E2E8F0" }}
                                                >
                                                    <TableCell sx={bodyStyle}>
                                                        <IconButton
                                                            size="small"
                                                            onClick={() =>
                                                                setExpandedKeys(p =>
                                                                    p.includes(row.__key)
                                                                        ? p.filter(x => x !== row.__key)
                                                                        : [...p, row.__key]
                                                                )
                                                            }
                                                            sx={{ p: 0.2 }}
                                                        >
                                                            {expanded ? (
                                                                <ExpandMoreIcon fontSize="small" />
                                                            ) : (
                                                                <ChevronRightIcon fontSize="small" />
                                                            )}
                                                        </IconButton>
                                                    </TableCell>
                                                    {enabledColumns.map((c) => {
                                                        const currentGroup = grouping[row.__level];
                                                        const isCurrentGroup = c.key === currentGroup;

                                                        return (
                                                            <TableCell
                                                                key={c.key}
                                                                align={c.isNumeric ? "right" : "left"}
                                                                sx={{
                                                                    ...bodyStyle,
                                                                    fontWeight: isCurrentGroup ? 700 : 500,
                                                                }}
                                                            >
                                                                {isCurrentGroup
                                                                    ? row.__value
                                                                    : renderGroupSubtotal(row.__rows, c)}
                                                            </TableCell>
                                                        );
                                                    })}
                                                </TableRow>
                                            );
                                        }

                                        return (
                                            <TableRow key={row.invoice_no ? `${row.invoice_no}-${i}` : i} hover>
                                                <TableCell sx={bodyStyle}>
                                                    {(page - 1) * rowsPerPage + i + 1}
                                                </TableCell>
                                                {enabledColumns.map((c) => (
                                                    <TableCell
                                                        key={c.key}
                                                        align={c.isNumeric ? "right" : "left"}
                                                        sx={bodyStyle}
                                                    >
                                                        {renderCellValue(row, c)}
                                                    </TableCell>
                                                ))}
                                            </TableRow>
                                        );
                                    })
                                )}
                            </TableBody>
                        </Table>
                    </TableContainer>

                    {/* Categorical / Date Filter Menu */}
                    <HeaderFilterMenu
                        anchorEl={filterAnchor}
                        open={Boolean(filterAnchor) && Boolean(activeHeader) && !isNumericColumn(activeHeader!)}
                        onClose={() => setFilterAnchor(null)}
                        columnLabel={activeHeader ? (columns.find(c => c.key === activeHeader)?.label || activeHeader) : undefined}
                        options={filterOptions}
                        selectedValues={activeHeader ? columnFilters[activeHeader] : undefined}
                        onFilterChange={(newSelected) => {
                            if (!activeHeader) return;
                            setColumnFilters(prev => {
                                const copy = { ...prev };
                                if (newSelected === undefined) {
                                    delete copy[activeHeader];
                                } else {
                                    copy[activeHeader] = newSelected;
                                }
                                return copy;
                            });
                            setPage(1);
                        }}
                        isDateColumn={activeHeader === "Ledger_Date" || activeHeader === "Date"}
                        dateFrom={tempDate.from}
                        dateTo={tempDate.to}
                        onDateChange={(dates) => setTempDate(dates)}
                        onApplyDate={() => {
                            setFilters((p) => ({ ...p, Date: tempDate }));
                            setFilterAnchor(null);
                        }}
                    />

                    {/* Numerical Filter Menu */}
                    <Menu
                        anchorEl={filterAnchor}
                        open={Boolean(filterAnchor) && Boolean(activeHeader) && isNumericColumn(activeHeader!)}
                        onClose={() => setFilterAnchor(null)}
                    >
                        {activeHeader && isNumericColumn(activeHeader) && (() => {
                            const { min, max } = getMinMax(activeHeader);
                            const currentRange = rangeFilter[activeHeader] || [min, max];

                            const handleSliderChange = (newValue: number[]) => {
                                setRangeFilter((prev) => ({
                                    ...prev,
                                    [activeHeader]: newValue as [number, number],
                                }));
                            };

                            const handleFromChange = (value: string) => {
                                let newFrom = Number(value);
                                if (isNaN(newFrom)) return;
                                newFrom = Math.max(min, Math.min(newFrom, currentRange[1]));
                                handleSliderChange([newFrom, currentRange[1]]);
                            };

                            const handleToChange = (value: string) => {
                                let newTo = Number(value);
                                if (isNaN(newTo)) return;
                                newTo = Math.min(max, Math.max(newTo, currentRange[0]));
                                handleSliderChange([currentRange[0], newTo]);
                            };

                            return (
                                <Box p={2} sx={{ minWidth: 250 }}>
                                    <Typography variant="subtitle2" sx={{ mb: 1.5, fontWeight: 600 }}>
                                        Filter Range
                                    </Typography>
                                    <Box display="flex" alignItems="center" gap={1} mb={2}>
                                        <TextField
                                            type="number"
                                            size="small"
                                            label="Min"
                                            value={currentRange[0]}
                                            onChange={(e) => handleFromChange(e.target.value)}
                                            inputProps={{ min, max: currentRange[1] }}
                                            sx={{
                                                width: 90,
                                                "& input": {
                                                    py: 0.5,
                                                    fontSize: "0.75rem",
                                                    textAlign: "center"
                                                }
                                            }}
                                        />
                                        <Typography fontSize={12}>—</Typography>
                                        <TextField
                                            type="number"
                                            size="small"
                                            label="Max"
                                            value={currentRange[1]}
                                            onChange={(e) => handleToChange(e.target.value)}
                                            inputProps={{ min: currentRange[0], max }}
                                            sx={{
                                                width: 90,
                                                "& input": {
                                                    py: 0.5,
                                                    fontSize: "0.75rem",
                                                    textAlign: "center"
                                                }
                                            }}
                                        />
                                    </Box>
                                    <Slider
                                        value={currentRange}
                                        min={min}
                                        max={max}
                                        step={0.01}
                                        size="small"
                                        onChange={(_, newValue) => handleSliderChange(newValue as number[])}
                                        valueLabelDisplay="auto"
                                        sx={{
                                            mb: 1.5,
                                            "& .MuiSlider-thumb": {
                                                width: 12,
                                                height: 12,
                                            }
                                        }}
                                    />
                                    <Box display="flex" justifyContent="space-between">
                                        <Button
                                            size="small"
                                            onClick={() => {
                                                setRangeFilter((prev) => {
                                                    const copy = { ...prev };
                                                    delete copy[activeHeader];
                                                    return copy;
                                                });
                                                setFilterAnchor(null);
                                            }}
                                            sx={{ textTransform: "none", color: "gray" }}
                                        >
                                            Clear
                                        </Button>
                                        <Button
                                            variant="contained"
                                            size="small"
                                            onClick={() => setFilterAnchor(null)}
                                            sx={{
                                                backgroundColor: "#1E3A8A",
                                                textTransform: "none",
                                                fontWeight: 600,
                                            }}
                                        >
                                            Close
                                        </Button>
                                    </Box>
                                </Box>
                            );
                        })()}
                    </Menu>
                </Box>

                <CommonPagination
                    totalRows={paginatedSourceRows.length}
                    page={page}
                    rowsPerPage={rowsPerPage}
                    onPageChange={setPage}
                    onRowsPerPageChange={setRowsPerPage}
                />
            </AppLayout>

            {/* COLUMN SETTINGS MENU */}
            <Menu
                anchorEl={settingsAnchor}
                open={Boolean(settingsAnchor)}
                onClose={() => setSettingsAnchor(null)}
                PaperProps={{
                    sx: {
                        maxHeight: 520,
                        width: 320,
                    }
                }}
            >
                <Box p={2}>
                    <Box display="flex" justifyContent="space-between" alignItems="center" mb={1.5}>
                        <Typography fontWeight={600} fontSize="0.9rem">Column Settings</Typography>
                        <Button
                            size="small"
                            onClick={handleResetSettings}
                            sx={{ textTransform: "none", fontSize: "0.75rem" }}
                        >
                            Reset
                        </Button>
                    </Box>

                    <Box sx={{ maxHeight: 420, overflowY: "auto", pr: 0.5 }}>
                        <DndContext
                            sensors={sensors}
                            collisionDetection={closestCenter}
                            onDragEnd={handleDragEnd}
                        >
                            <Typography fontSize="0.75rem" fontWeight={700} color="#1E3A8A" mb={0.5}>
                                Enabled Columns ({sortedColumns.filter(c => c.enabled).length})
                            </Typography>

                            <SortableContext
                                items={sortedColumns.filter(c => c.enabled).map(c => c.key)}
                                strategy={verticalListSortingStrategy}
                            >
                                {sortedColumns
                                    .filter(c => c.enabled)
                                    .map(c => (
                                        <SortableColumnRow
                                            key={c.key}
                                            column={c}
                                            onToggle={key =>
                                                setColumns(cols =>
                                                    cols.map(x =>
                                                        x.key === key
                                                            ? { ...x, enabled: false }
                                                            : x
                                                    )
                                                )
                                            }
                                        />
                                    ))}
                            </SortableContext>

                            <Typography fontSize="0.75rem" fontWeight={700} color="#6B7280" mt={2} mb={0.5}>
                                Disabled Columns ({sortedColumns.filter(c => !c.enabled).length})
                            </Typography>

                            <SortableContext
                                items={sortedColumns.filter(c => !c.enabled).map(c => c.key)}
                                strategy={verticalListSortingStrategy}
                            >
                                {sortedColumns
                                    .filter(c => !c.enabled)
                                    .map(c => (
                                        <SortableColumnRow
                                            key={c.key}
                                            column={c}
                                            onToggle={key =>
                                                setColumns(cols =>
                                                    cols.map(x =>
                                                        x.key === key
                                                            ? { ...x, enabled: true }
                                                            : x
                                                    )
                                                )
                                            }
                                        />
                                    ))}
                            </SortableContext>
                        </DndContext>
                    </Box>
                </Box>
            </Menu>

            {/* GROUP BY DIALOG */}
            <Dialog
                open={groupDialogOpen}
                onClose={() => setGroupDialogOpen(false)}
                fullWidth
                maxWidth="sm"
            >
                <DialogTitle sx={{ fontWeight: 600, color: "#1E3A8A" }}>Group By Columns</DialogTitle>

                <DialogContent>
                    {[0, 1, 2].map(level => (
                        <TextField
                            key={level}
                            select
                            fullWidth
                            margin="dense"
                            size="small"
                            label={`Level ${level + 1}`}
                            value={pendingGrouping[level] || ""}
                            onChange={e => {
                                const copy = [...pendingGrouping];
                                copy[level] = e.target.value;
                                setPendingGrouping(copy.filter(Boolean));
                            }}
                            sx={{ mb: 1.5 }}
                        >
                            <MenuItem value="">No Grouping</MenuItem>
                            {enabledColumns.map(col => (
                                <MenuItem
                                    key={col.key}
                                    value={col.key}
                                    disabled={pendingGrouping.includes(col.key) && pendingGrouping[level] !== col.key}
                                >
                                    {col.label}
                                </MenuItem>
                            ))}
                        </TextField>
                    ))}
                </DialogContent>

                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button onClick={() => setGroupDialogOpen(false)} color="inherit" size="small">
                        Close
                    </Button>
                    <Button
                        variant="contained"
                        size="small"
                        onClick={() => {
                            const validGrouping = pendingGrouping.filter(Boolean);
                            setGrouping(validGrouping);
                            setColumns(prev =>
                                prev.map(col => ({
                                    ...col,
                                    enabled: validGrouping.includes(col.key)
                                        ? true
                                        : col.enabled,
                                }))
                            );
                            setGroupDialogOpen(false);
                            setPage(1);
                        }}
                        sx={{
                            backgroundColor: "#1E3A8A",
                            "&:hover": { backgroundColor: "#172E6D" }
                        }}
                    >
                        Apply
                    </Button>
                </DialogActions>
            </Dialog>

            {/* SAVE TEMPLATE DIALOG */}
            <Dialog
                open={saveDialogOpen}
                onClose={() => setSaveDialogOpen(false)}
            >
                <DialogTitle sx={{ fontWeight: 600, color: "#1E3A8A" }}>
                    {selectedTemplateId ? "Edit Template" : "Create Template"}
                </DialogTitle>

                <DialogContent sx={{ minWidth: 320, pt: 1 }}>
                    <TextField
                        fullWidth
                        size="small"
                        label="Report Name"
                        value={reportName}
                        onChange={e => setReportName(e.target.value)}
                        sx={{ mt: 1 }}
                    />
                </DialogContent>

                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button onClick={() => setSaveDialogOpen(false)} color="inherit" size="small">
                        Cancel
                    </Button>
                    <Button
                        variant="contained"
                        size="small"
                        onClick={handleQuickSave}
                        sx={{
                            backgroundColor: "#1E3A8A",
                            "&:hover": { backgroundColor: "#172E6D" }
                        }}
                    >
                        Save
                    </Button>
                </DialogActions>
            </Dialog>

            {/* SYNC DIALOG */}
            <Dialog
                open={syncDialogOpen}
                onClose={() => !isSyncing && setSyncDialogOpen(false)}
                PaperProps={{
                    sx: {
                        borderRadius: 2,
                        minWidth: 320,
                    }
                }}
            >
                <DialogTitle sx={{ fontWeight: 600, color: "#1E3A8A", pb: 1 }}>
                    Sync Unit Economics
                </DialogTitle>
                <DialogContent sx={{ display: "flex", flexDirection: "column", gap: 2.5, pt: 2 }}>
                    <TextField
                        label="From Date"
                        type="date"
                        InputLabelProps={{ shrink: true }}
                        value={syncFromDate}
                        onChange={(e) => setSyncFromDate(e.target.value)}
                        fullWidth
                        disabled={isSyncing}
                        size="small"
                    />
                    <TextField
                        label="To Date"
                        type="date"
                        InputLabelProps={{ shrink: true }}
                        value={syncToDate}
                        onChange={(e) => setSyncToDate(e.target.value)}
                        fullWidth
                        disabled={isSyncing}
                        size="small"
                    />
                </DialogContent>
                <DialogActions sx={{ px: 3, pb: 2 }}>
                    <Button
                        onClick={() => setSyncDialogOpen(false)}
                        disabled={isSyncing}
                        color="inherit"
                        size="small"
                        sx={{ textTransform: "none" }}
                    >
                        Cancel
                    </Button>
                    <Button
                        onClick={handleSync}
                        variant="contained"
                        disabled={isSyncing}
                        size="small"
                        sx={{
                            backgroundColor: "#1E3A8A",
                            textTransform: "none",
                            "&:hover": {
                                backgroundColor: "#172E6D",
                            },
                        }}
                    >
                        {isSyncing ? <CircularProgress size={20} color="inherit" /> : "Sync"}
                    </Button>
                </DialogActions>
            </Dialog>
        </>
    );
};

export default UnitEconomicsAdmin;
