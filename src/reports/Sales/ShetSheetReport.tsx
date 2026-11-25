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
    Button,
    IconButton,
    Menu,
    Switch,
    Typography,
    TextField,
    Tooltip,
    Dialog,
    DialogActions,
    DialogTitle,
    DialogContent,
    CircularProgress,
    Chip,
} from "@mui/material";
import dayjs from "dayjs";
import SettingsIcon from "@mui/icons-material/Settings";
import ArrowDropUpIcon from "@mui/icons-material/ArrowDropUp";
import ArrowDropDownIcon from "@mui/icons-material/ArrowDropDown";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import {
    DndContext,
    closestCenter,
    PointerSensor,
    useSensor,
    useSensors,
} from "@dnd-kit/core";
import {
    SortableContext,
    useSortable,
    verticalListSortingStrategy,
    arrayMove,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";

import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import CommonPagination from "../../Components/CommonPagination";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import { SettingsService } from "../../services/reportSettings.services";
import {
    ShetSheetReportService,
    ShetSheetItem,
} from "../../services/ShetSheetReport.service";

/* ================= TYPES & CONFIG ================= */

export interface ColumnConfig {
    key: string;
    label: string;
    enabled: boolean;
    isNumeric?: boolean;
    isStaffType?: boolean;
    order: number;
}

// Flat row representation for table display
export interface FlattenedShetSheetRow {
    _raw: ShetSheetItem;
    RemarksGroup: string;
    Do_Date: string;
    retailerNameGet: string;
    Bags: number;
    Godown_Name: string;
    Follow_Up: string;
    Status: string;
    Do_Inv_No: string;
    voucherTypeGet: string;
    branchNameGet: string;
    Total_Invoice_value: number;
    Created_BY_Name: string;
    Created_on: string;
    Narration: string;
    trip_no: string;
    Product_Name: string;
    UOM: string;
    Bill_Qty: number;
    Act_Qty: number;
    billedRate: number;
    itemRate: number;
    [staffTypeKey: string]: any;
}

const DEFAULT_BASE_COLUMNS: ColumnConfig[] = [
    { key: "Do_Date", label: "Date", enabled: true, isNumeric: false, order: 0 },
    { key: "retailerNameGet", label: "Party Name", enabled: true, isNumeric: false, order: 1 },
    { key: "Bags", label: "Bags", enabled: true, isNumeric: true, order: 2 },
    { key: "Godown_Name", label: "Godown", enabled: true, isNumeric: false, order: 3 },
    { key: "Follow_Up", label: "Follow Up", enabled: true, isNumeric: false, order: 4 },
    { key: "Status", label: "Status", enabled: true, isNumeric: false, order: 5 },
    { key: "Do_Inv_No", label: "DO / Inv No", enabled: false, isNumeric: false, order: 6 },
    { key: "voucherTypeGet", label: "Voucher Type", enabled: false, isNumeric: false, order: 7 },
    { key: "branchNameGet", label: "Branch", enabled: false, isNumeric: false, order: 8 },
    { key: "Total_Invoice_value", label: "Total Value", enabled: false, isNumeric: true, order: 9 },
    { key: "Product_Name", label: "Product Name", enabled: false, isNumeric: false, order: 10 },
    { key: "Bill_Qty", label: "Bill Qty", enabled: false, isNumeric: true, order: 11 },
    { key: "Act_Qty", label: "Act Qty", enabled: false, isNumeric: true, order: 12 },
    { key: "billedRate", label: "Billed Rate", enabled: false, isNumeric: true, order: 13 },
    { key: "itemRate", label: "Item Rate", enabled: false, isNumeric: true, order: 14 },
    { key: "Created_BY_Name", label: "Created By", enabled: false, isNumeric: false, order: 15 },
    { key: "Created_on", label: "Created On", enabled: false, isNumeric: false, order: 16 },
    { key: "trip_no", label: "Trip No", enabled: false, isNumeric: false, order: 17 },
    { key: "Narration", label: "Narration", enabled: false, isNumeric: false, order: 18 },
];

/* ================= DND SORTABLE ROW ================= */

const SortableColumnRow = ({
    column,
    onToggle,
}: {
    column: ColumnConfig;
    onToggle: (key: string, enabled: boolean) => void;
}) => {
    const { attributes, listeners, setNodeRef, transform, transition } = useSortable({
        id: column.key,
    });

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
            mb={0.5}
            p={0.5}
            sx={{
                bgcolor: column.isStaffType ? "#f1f5f9" : "transparent",
                borderRadius: 1,
            }}
        >
            <Box {...listeners} {...attributes} sx={{ cursor: "grab", display: "flex", alignItems: "center" }}>
                <DragIndicatorIcon fontSize="small" sx={{ color: "#94a3b8" }} />
            </Box>
            <Typography variant="body2" sx={{ flex: 1, fontSize: "0.82rem" }}>
                {column.label}
                {column.isStaffType && (
                    <Chip label="Staff Involved" size="small" sx={{ ml: 1, height: 18, fontSize: "0.65rem", bgcolor: "#e0e7ff", color: "#3730a3" }} />
                )}
            </Typography>
            <Switch
                size="small"
                checked={column.enabled}
                onChange={(e) => onToggle(column.key, e.target.checked)}
            />
        </Box>
    );
};

/* ================= MAIN COMPONENT ================= */

const ShetSheetReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const parentReportName = "SHET SHEET REPORT";

    // Date Filters
    const [fromDate, setFromDate] = useState<string>(today);
    const [toDate, setToDate] = useState<string>(today);
    const [allDetails, setAllDetails] = useState<boolean>(false);
    const [statusFilter, setStatusFilter] = useState<"pending" | "complete" | "all">("pending");
    const [drawerOpen, setDrawerOpen] = useState<boolean>(false);
    const [tempFromDate, setTempFromDate] = useState<string>(today);
    const [tempToDate, setTempToDate] = useState<string>(today);
    const [tempAllDetails, setTempAllDetails] = useState<boolean>(false);
    const [tempStatusFilter, setTempStatusFilter] = useState<"pending" | "complete" | "all">("pending");

    // Data State
    const [rawItems, setRawItems] = useState<ShetSheetItem[]>([]);
    const [loading, setLoading] = useState<boolean>(false);

    // Columns & Dynamic Staff Involved Columns
    const [columns, setColumns] = useState<ColumnConfig[]>(DEFAULT_BASE_COLUMNS);
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);

    // Template Save / Load State
    const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
    const [reportName, setReportName] = useState<string>("");
    const [saveDialogOpen, setSaveDialogOpen] = useState<boolean>(false);
    const [templateLoading, setTemplateLoading] = useState<boolean>(false);

    // Collapsed groups (Remarks grouping: Remarks || "No Remarks")
    const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({});

    // Pagination
    const [page, setPage] = useState<number>(1);
    const [rowsPerPage, setRowsPerPage] = useState<number>(100);

    // Header Filtering & Sorting
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);
    const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
    const [sortConfig, setSortConfig] = useState<{ key: string | null; direction: "asc" | "desc" }>({
        key: null,
        direction: "asc",
    });



    /* ================= FLATTENED ROWS ================= */

    // Check if Product_Name or any detailed stock column is enabled
    const isDetailedStockEnabled = useMemo(() => {
        const stockColKeys = ["Product_Name", "Bill_Qty", "Act_Qty", "billedRate", "itemRate", "UOM"];
        return columns.some((c) => c.enabled && stockColKeys.includes(c.key));
    }, [columns]);

    const flattenedRows = useMemo(() => {
        const list: FlattenedShetSheetRow[] = [];

        rawItems.forEach((item) => {
            const remarksVal = (item.Remarks && item.Remarks.trim()) || "No remarks";

            // Common staff involved mapping
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
                _raw: item,
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

            // Calculate total bags for this invoice/DO
            const totalAltBags = stocks.reduce((sum, s) => sum + Number(s.Alt_Act_Qty || 0), 0);
            const totalActBags = stocks.reduce((sum, s) => sum + Number(s.Act_Qty || 0), 0);
            const invoiceTotalBags =
                item.Pending_Bags !== null && item.Pending_Bags !== undefined
                    ? Number(item.Pending_Bags)
                    : totalAltBags > 0
                        ? totalAltBags
                        : totalActBags;

            // Unique clean Godown Name(s) without duplicates
            const uniqueGodowns = Array.from(new Set(stocks.map((s) => (s.Godown_Name || "").trim()).filter(Boolean)));
            const godownDisplay = uniqueGodowns.join(", ");

            // CASE 1: Product_Name is NOT enabled (standard summary/pivot view)
            // Show ONE row per invoice with total bags and de-duplicated godown
            if (!isDetailedStockEnabled) {
                const uniqueProducts = Array.from(new Set(stocks.map((s) => (s.Product_Name || "").trim()).filter(Boolean)));
                const sumBillQty = stocks.reduce((sum, s) => sum + Number(s.Bill_Qty || 0), 0);
                const sumActQty = stocks.reduce((sum, s) => sum + Number(s.Act_Qty || 0), 0);

                list.push({
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
                // CASE 2: Product Name or detailed stock column is enabled
                // Separate out by differing products like a pivot
                if (stocks.length === 0) {
                    list.push({
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

                        list.push({
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

        return list;
    }, [rawItems, isDetailedStockEnabled]);

    // Numerical Filter hook
    const numericKeys = useMemo(
        () => columns.filter((c) => c.isNumeric).map((c) => c.key),
        [columns]
    );

    const {
        sortConfig: numSortConfig,
        rangeFilter: numRangeFilter,
        setRangeFilter: setNumRangeFilter,
        filterAnchor: numFilterAnchor,
        setFilterAnchor: setNumFilterAnchor,
        activeHeader: numActiveHeader,
        handleSort: handleNumSort,
        openFilter: openNumFilter,
        getMinMax,
        clearRangeFilter,
    } = useNumericalFilter(flattenedRows, numericKeys);

    // DND Sensors
    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: { distance: 5 },
        })
    );

    /* ================= FETCH DATA ================= */

    const fetchData = async (
        from = fromDate,
        to = toDate,
        isAll = allDetails,
        stFilter = statusFilter
    ) => {
        try {
            setLoading(true);
            const actualFrom = isAll ? dayjs(to).subtract(60, "day").format("YYYY-MM-DD") : from;
            const res = await ShetSheetReportService.getShetSheetReport({
                Fromdate: actualFrom,
                Todate: to,
                allDetails: isAll,
                status: stFilter,
            });

            const dataList: ShetSheetItem[] = res.data?.data || (Array.isArray(res.data) ? (res.data as any) : []);
            setRawItems(dataList);

            // Extract dynamic involved staff types
            const staffTypeSet = new Set<string>();
            dataList.forEach((it) => {
                if (Array.isArray(it.involvedStaffs)) {
                    it.involvedStaffs.forEach((st) => {
                        if (st.Involved_Emp_Type && st.Involved_Emp_Type.trim()) {
                            staffTypeSet.add(st.Involved_Emp_Type.trim());
                        }
                    });
                }
            });

            // Update columns metadata with detected Staff Types
            setColumns((prevCols) => {
                const existingKeys = new Set(prevCols.map((c) => c.key));
                const newCols = [...prevCols];
                let nextOrder = prevCols.length;

                Array.from(staffTypeSet).sort().forEach((stType) => {
                    const staffColKey = `staff_${stType}`;
                    if (!existingKeys.has(staffColKey)) {
                        newCols.push({
                            key: staffColKey,
                            label: stType,
                            enabled: false, // Disabled by default; user can enable in Table Settings
                            isNumeric: false,
                            isStaffType: true,
                            order: nextOrder++,
                        });
                    }
                });

                return newCols;
            });
        } catch (err) {
            console.error("Failed to load ShetSheet report data:", err);
            toast.error("Failed to load Shet Sheet Report ❌");
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        fetchData();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    /* ================= TEMPLATE LOGIC ================= */

    const handleLoadTemplate = async (templateId: number) => {
        try {
            setTemplateLoading(true);
            setLoading(true);
            const res = await SettingsService.getReportEditData({
                reportId: templateId,
                typeId: 1,
            });

            if (res.data?.success && res.data?.data) {
                const settings = res.data.data;
                const templateCols: any[] = settings.columns || [];

                if (templateCols.length > 0) {
                    setColumns((prev) => {
                        const updated = prev.map((col, idx) => {
                            const tc = templateCols.find((x: any) => x.key === col.key);
                            return {
                                ...col,
                                enabled: tc ? Boolean(tc.enabled) : false,
                                order: tc?.order !== undefined ? Number(tc.order) : idx,
                            };
                        });
                        return updated.sort((a, b) => a.order - b.order);
                    });
                }

                setSelectedTemplateId(templateId);
                setReportName(settings.Report_Name || settings.reportInfo?.Report_Name || "");

                // Load allDetails & status setting and header columnFilters from Database Preset_Config first, then fallback to localStorage
                let isAll = false;
                let st: "pending" | "complete" | "all" = "pending";
                let savedColFilters: Record<string, string[]> = {};
                let savedRange: Record<string, [number, number]> = {};

                const dbPresetConfig = res.data.data?.reportInfo?.Preset_Config;
                let dbParsed: any = null;
                if (dbPresetConfig) {
                    try {
                        dbParsed = typeof dbPresetConfig === "string" ? JSON.parse(dbPresetConfig) : dbPresetConfig;
                    } catch (e) {
                        console.warn("Failed to parse db Preset_Config in ShetSheet", e);
                    }
                }

                if (dbParsed && typeof dbParsed === "object") {
                    if (dbParsed.allDetails !== undefined) isAll = Boolean(dbParsed.allDetails);
                    if (dbParsed.status) st = dbParsed.status;
                    if (dbParsed.columnFilters) savedColFilters = dbParsed.columnFilters;
                    if (dbParsed.rangeFilter) savedRange = dbParsed.rangeFilter;
                } else {
                    const savedAll = localStorage.getItem(`template_allDetails_${templateId}`);
                    isAll = savedAll === "true";
                    const savedStatus = localStorage.getItem(`template_status_${templateId}`) as ("pending" | "complete" | "all") | null;
                    if (savedStatus) st = savedStatus;
                    const localColFilters = localStorage.getItem(`template_colFilters_${templateId}`);
                    if (localColFilters) {
                        try { savedColFilters = JSON.parse(localColFilters); } catch (e) { /* ignore */ }
                    }
                }

                setAllDetails(isAll);
                setTempAllDetails(isAll);
                setStatusFilter(st);
                setTempStatusFilter(st);
                setColumnFilters(savedColFilters);
                if (Object.keys(savedRange).length > 0) {
                    setNumRangeFilter(savedRange);
                }

                const effectiveFrom = isAll ? dayjs(toDate).subtract(60, "day").format("YYYY-MM-DD") : fromDate;
                if (isAll) {
                    setFromDate(effectiveFrom);
                    setTempFromDate(effectiveFrom);
                }

                // Await full data load before clearing the template loading indicator so UI doesn't render partial data
                await fetchData(effectiveFrom, toDate, isAll, st);
                toast.success("Template Loaded ✅");
            }
        } catch (err) {
            console.error(err);
            toast.error("Failed to load template ❌");
        } finally {
            setLoading(false);
            setTemplateLoading(false);
        }
    };

    const handleClearTemplate = async () => {
        try {
            setTemplateLoading(true);
            setLoading(true);
            setSelectedTemplateId(null);
            setReportName("");
            setAllDetails(false);
            setTempAllDetails(false);
            setStatusFilter("pending");
            setTempStatusFilter("pending");
            setColumns((prev) =>
                prev.map((c, i) => ({
                    ...c,
                    enabled: DEFAULT_BASE_COLUMNS.some((d) => d.key === c.key && d.enabled),
                    order: i,
                }))
            );
            await fetchData(fromDate, toDate, false, "pending");
        } catch (err) {
            console.error(err);
        } finally {
            setLoading(false);
            setTemplateLoading(false);
        }
    };

    const handleQuickSave = async () => {
        if (!reportName.trim()) {
            toast.warning("Please enter a Template Name");
            return;
        }

        try {
            const createdBy = Number(localStorage.getItem("userId") || 1);
            const payloadColumns = columns.map((col) => ({
                key: col.key,
                label: col.label,
                enabled: col.enabled ? 1 : 0,
                order: col.order,
                groupBy: 0,
            }));

            const templateDbConfig = {
                allDetails,
                status: statusFilter,
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
                localStorage.setItem(`template_allDetails_${selectedTemplateId}`, String(allDetails));
                localStorage.setItem(`template_status_${selectedTemplateId}`, statusFilter);
                localStorage.setItem(`template_colFilters_${selectedTemplateId}`, JSON.stringify(columnFilters));
                localStorage.setItem(`template_rangeFilter_${selectedTemplateId}`, JSON.stringify(numRangeFilter));
                toast.success("Template Updated Successfully ✅");
                setSaveDialogOpen(false);
                setTimeout(() => {
                    window.location.reload();
                }, 500);
            } else {
                const saveRes: any = await SettingsService.saveReportSettings({
                    reportName: reportName.trim(),
                    parentReport: parentReportName,
                    abstractSP: "EXTERNAL_API",
                    expandedSP: "EXTERNAL_API",
                    abstractColumns: payloadColumns,
                    expandedColumns: payloadColumns,
                    createdBy,
                    presetConfig: templateDbConfig,
                });
                const newId = (saveRes as any)?.data?.reportId || (saveRes as any)?.data?.data?.Report_Id || (saveRes as any)?.data?.Report_Id;
                if (newId) {
                    setSelectedTemplateId(newId);
                    localStorage.setItem(`template_allDetails_${newId}`, String(allDetails));
                    localStorage.setItem(`template_status_${newId}`, statusFilter);
                    localStorage.setItem(`template_colFilters_${newId}`, JSON.stringify(columnFilters));
                    localStorage.setItem(`template_rangeFilter_${newId}`, JSON.stringify(numRangeFilter));
                }
                toast.success("Template Saved Successfully ✅");
                setSaveDialogOpen(false);
                setTimeout(() => {
                    window.location.reload();
                }, 500);
            }
        } catch (err) {
            console.error(err);
            toast.error("Save Template Failed ❌");
        }
    };



    /* ================= FILTERING & SORTING ================= */

    const filteredRows = useMemo(() => {
        let result = [...flattenedRows];

        // 1. Column Header Dropdown Filters
        Object.entries(columnFilters).forEach(([colKey, selectedValues]) => {
            if (selectedValues && selectedValues.length > 0) {
                result = result.filter((row) => {
                    const cellVal = String(row[colKey] ?? "");
                    return selectedValues.includes(cellVal);
                });
            }
        });

        // 2. Numerical Range Filters
        Object.entries(numRangeFilter).forEach(([colKey, range]) => {
            if (range && Array.isArray(range)) {
                result = result.filter((row) => {
                    const num = Number(row[colKey] || 0);
                    return num >= range[0] && num <= range[1];
                });
            }
        });

        // 3. String Header Sort
        if (sortConfig.key) {
            result.sort((a, b) => {
                const aVal = String(a[sortConfig.key!] ?? "");
                const bVal = String(b[sortConfig.key!] ?? "");
                return sortConfig.direction === "asc"
                    ? aVal.localeCompare(bVal, undefined, { numeric: true })
                    : bVal.localeCompare(aVal, undefined, { numeric: true });
            });
        }

        // 4. Numerical Sort
        if (numSortConfig && numSortConfig.key) {
            result.sort((a, b) => {
                const aVal = Number(a[numSortConfig.key!] || 0);
                const bVal = Number(b[numSortConfig.key!] || 0);
                return numSortConfig.direction === "asc" ? aVal - bVal : bVal - aVal;
            });
        }

        return result;
    }, [flattenedRows, columnFilters, numRangeFilter, sortConfig, numSortConfig]);

    // Grouping by Remarks
    const groupedData = useMemo(() => {
        const groupsMap = new Map<string, FlattenedShetSheetRow[]>();

        filteredRows.forEach((row) => {
            const grp = row.RemarksGroup;
            if (!groupsMap.has(grp)) {
                groupsMap.set(grp, []);
            }
            groupsMap.get(grp)!.push(row);
        });

        return Array.from(groupsMap.entries()).map(([groupName, items]) => ({
            groupName,
            items,
            totalBags: items.reduce((sum, r) => sum + (Number(r.Bags) || 0), 0),
        }));
    }, [filteredRows]);

    // Enabled columns sorted by order
    const enabledColumns = useMemo(() => {
        return columns.filter((c) => c.enabled).sort((a, b) => a.order - b.order);
    }, [columns]);

    // Distinct options for HeaderFilterMenu
    const filterOptions = useMemo(() => {
        if (!activeHeader) return [];
        const set = new Set<string>();
        flattenedRows.forEach((r) => {
            const val = String(r[activeHeader] ?? "");
            if (val.trim()) set.add(val);
        });
        return Array.from(set).sort();
    }, [activeHeader, flattenedRows]);

    // Min/Max for NumericalFilterMenu
    const getMinMaxFallback = (key: string) => {
        if (typeof getMinMax === "function") {
            try {
                return getMinMax(key);
            } catch (e) {
                // fallback
            }
        }
        let min = Infinity;
        let max = -Infinity;
        flattenedRows.forEach((r) => {
            const val = Number(r[key] || 0);
            if (val < min) min = val;
            if (val > max) max = val;
        });
        if (min === Infinity) min = 0;
        if (max === -Infinity) max = 100;
        return { min, max };
    };

    /* ================= TOGGLE / REORDER COLUMNS ================= */

    const handleToggleColumn = (key: string, enabled: boolean) => {
        setColumns((prev) =>
            prev.map((c) => (c.key === key ? { ...c, enabled } : c))
        );
    };

    const handleDragEnd = (event: any) => {
        const { active, over } = event;
        if (!over || active.id === over.id) return;

        setColumns((cols) => {
            const enabledList = cols.filter((c) => c.enabled);
            const disabledList = cols.filter((c) => !c.enabled);

            const isDraggingEnabled = enabledList.some((c) => c.key === active.id);
            const activeList = isDraggingEnabled ? enabledList : disabledList;

            const oldIndex = activeList.findIndex((c) => c.key === active.id);
            const newIndex = activeList.findIndex((c) => c.key === over.id);

            if (oldIndex === -1 || newIndex === -1) return cols;

            const reordered = arrayMove(activeList, oldIndex, newIndex);

            const finalCols = isDraggingEnabled
                ? [...reordered, ...disabledList]
                : [...enabledList, ...reordered];

            return finalCols.map((col, idx) => ({ ...col, order: idx }));
        });
    };

    /* ================= EXPORT LOGIC ================= */

    const handleExportExcel = () => {
        if (!groupedData.length) {
            toast.warning("No data to export");
            return;
        }

        const wb = XLSX.utils.book_new();
        const exportAoa: any[][] = [];

        // Title Row
        const titleRow = ["Shet Sheet Pending Report Screen"];
        exportAoa.push(titleRow);

        // Header Row
        const headerRow = ["S.No", ...enabledColumns.map((c) => c.label)];
        exportAoa.push(headerRow);

        groupedData.forEach((grp) => {
            // Group Row (Remarks e.g., "Vandi Varum", "No Stock", "No Shet Sheet")
            exportAoa.push([grp.groupName]);

            grp.items.forEach((item, idx) => {
                const rowData = [
                    idx + 1,
                    ...enabledColumns.map((c) => {
                        const val = item[c.key];
                        return val !== undefined && val !== null ? val : "";
                    }),
                ];
                exportAoa.push(rowData);
            });
        });

        const ws = XLSX.utils.aoa_to_sheet(exportAoa);

        // Styling cells with Excel table layout matching user screenshot
        ws["!merges"] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: enabledColumns.length } },
        ];

        const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
        let currentGroupRow = -1;

        for (let R = range.s.r; R <= range.e.r; ++R) {
            const firstCell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
            const isGroup = firstCell && groupedData.some((g) => g.groupName === firstCell.v);
            if (isGroup) {
                currentGroupRow = R;
                if (!ws["!merges"]) ws["!merges"] = [];
                ws["!merges"].push({ s: { r: R, c: 0 }, e: { r: R, c: enabledColumns.length } });
            }

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellRef];
                if (!cell) continue;

                cell.s = {
                    font: { name: "Calibri", sz: 10 },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } },
                    },
                };

                if (R === 0) {
                    // Title Header: Bright Yellow Fill, bold centered text
                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                    cell.s.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "000000" } };
                    cell.s.alignment = { horizontal: "center", vertical: "center" };
                } else if (R === 1) {
                    // Columns Header: Bright Yellow Fill, bold text
                    cell.s.fill = { fgColor: { rgb: "FFFF00" } };
                    cell.s.font = { name: "Calibri", sz: 10, bold: true, color: { rgb: "000000" } };
                    cell.s.alignment = { horizontal: C === 0 || C === 3 ? "center" : "left", vertical: "center" };
                } else if (R === currentGroupRow) {
                    // Group Header: Mild Blue Fill (#DBEAFE), bold deep blue text
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

        ws["!cols"] = [{ wch: 8 }, ...enabledColumns.map(() => ({ wch: 22 }))];

        XLSX.utils.book_append_sheet(wb, ws, "ShetSheetReport");
        XLSX.writeFile(wb, `Shet_Sheet_Report_${fromDate}_${toDate}.xlsx`);
        toast.success("Excel exported successfully! 📊");
    };

    const handleExportPDF = () => {
        if (!groupedData.length) {
            toast.warning("No data to export");
            return;
        }

        const doc = new jsPDF("l", "mm", "a4");
        doc.setFontSize(14);
        doc.setTextColor(30, 58, 138);
        doc.text("Shet Sheet Pending Report", 14, 12);
        doc.setFontSize(9);
        doc.text(`Period: ${fromDate} to ${toDate}`, 230, 12);

        const tableHeaders = [["S.No", ...enabledColumns.map((c) => c.label)]];
        const tableBody: any[] = [];
        const groupRowIndices = new Set<number>();

        groupedData.forEach((grp) => {
            groupRowIndices.add(tableBody.length);
            tableBody.push([
                {
                    content: grp.groupName,
                    colSpan: enabledColumns.length + 1,
                    styles: {
                        fillColor: [219, 234, 254], // Mild Blue (#DBEAFE)
                        textColor: [30, 58, 138], // Deep Blue (#1E3A8A)
                        fontStyle: "bold",
                        fontSize: 9.5,
                    },
                },
            ]);

            grp.items.forEach((item, idx) => {
                tableBody.push([
                    idx + 1,
                    ...enabledColumns.map((c) => {
                        const val = item[c.key];
                        return val !== undefined && val !== null ? String(val) : "";
                    }),
                ]);
            });
        });

        autoTable(doc, {
            startY: 18,
            head: tableHeaders,
            body: tableBody,
            styles: { fontSize: 8, cellPadding: 2 },
            headStyles: {
                fillColor: [30, 58, 138],
                textColor: [255, 255, 255],
                fontStyle: "bold",
            },
            theme: "grid",
            didParseCell: (cellData) => {
                if (groupRowIndices.has(cellData.row.index)) {
                    cellData.cell.styles.fillColor = [219, 234, 254]; // Mild Blue (#DBEAFE)
                    cellData.cell.styles.textColor = [30, 58, 138]; // Deep Blue (#1E3A8A)
                    cellData.cell.styles.fontStyle = "bold";
                }
            },
        });

        doc.save(`Shet_Sheet_Report_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`);
        toast.success("PDF exported successfully! 📄");
    };

    /* ================= RENDER ================= */

    return (
        <Box sx={{ display: "flex", flexDirection: "column", height: "100%", width: "100%", overflow: "hidden" }}>
            {/* PAGE HEADER */}
            <PageHeader
                parentReportName={parentReportName}
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                onReportChange={(template) => {
                    if (!template || !template.Report_Id) {
                        handleClearTemplate();
                    } else {
                        setSelectedTemplateId(template.Report_Id);
                        setReportName(template.Report_Name || "");
                        handleLoadTemplate(template.Report_Id);
                    }
                }}
                onQuickSave={() => {
                    if (!selectedTemplateId) {
                        setReportName("");
                    }
                    setSaveDialogOpen(true);
                }}
                settingsSlot={
                    <Tooltip title="Table Column Settings">
                        <IconButton
                            size="small"
                            onClick={(e) => setSettingsAnchor(e.currentTarget)}
                            sx={{
                                height: 26,
                                width: 26,
                                backgroundColor: "#fff",
                                borderRadius: 0.5,
                            }}
                        >
                            <SettingsIcon fontSize="small" />
                        </IconButton>
                    </Tooltip>
                }
            />

            {/* FILTER DRAWER */}
            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => {
                    setDrawerOpen((prev) => {
                        if (!prev) {
                            setTempFromDate(fromDate);
                            setTempToDate(toDate);
                            setTempAllDetails(allDetails);
                            setTempStatusFilter(statusFilter);
                        }
                        return !prev;
                    });
                }}
                onClose={() => setDrawerOpen(false)}
                fromDate={tempAllDetails ? dayjs(tempToDate).subtract(60, "day").format("YYYY-MM-DD") : tempFromDate}
                onFromDateChange={setTempFromDate}
                toDate={tempToDate}
                onToDateChange={(newTo) => {
                    setTempToDate(newTo);
                    if (tempAllDetails) {
                        setTempFromDate(dayjs(newTo).subtract(60, "day").format("YYYY-MM-DD"));
                    }
                }}
                disableFromDate={tempAllDetails}
                showAllDetailsFilter={true}
                allDetailsValue={tempAllDetails}
                onAllDetailsChange={(checked) => {
                    setTempAllDetails(checked);
                    if (checked) {
                        setTempFromDate(dayjs(tempToDate).subtract(60, "day").format("YYYY-MM-DD"));
                    } else {
                        setTempFromDate(today);
                    }
                }}
                showShetSheetStatusFilter={true}
                shetSheetStatusValue={tempStatusFilter}
                onShetSheetStatusChange={setTempStatusFilter}
                onApply={() => {
                    const finalFrom = tempAllDetails ? dayjs(tempToDate).subtract(60, "day").format("YYYY-MM-DD") : tempFromDate;
                    setFromDate(finalFrom);
                    setToDate(tempToDate);
                    setAllDetails(tempAllDetails);
                    setStatusFilter(tempStatusFilter);
                    setDrawerOpen(false);
                    fetchData(finalFrom, tempToDate, tempAllDetails, tempStatusFilter);
                }}
            />

            {/* MAIN CONTENT AREA */}
            <AppLayout fullWidth>
                <Box sx={{ p: 1.5, display: "flex", flexDirection: "column", flex: 1, minHeight: 0, overflow: "hidden", position: "relative" }}>
                    {(loading || templateLoading) ? (
                        <Box sx={{ display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", flex: 1, gap: 1.5 }}>
                            <CircularProgress size={42} sx={{ color: "#1E3A8A" }} />
                            {templateLoading && (
                                <Typography variant="body2" sx={{ color: "#1E3A8A", fontWeight: 600 }}>
                                    Loading template configuration & data...
                                </Typography>
                            )}
                        </Box>
                    ) : (
                        <>
                            <TableContainer
                                component={Paper}
                                sx={{
                                    flex: 1,
                                    width: "100%",
                                    overflowY: "auto",
                                    overflowX: "hidden",
                                    border: "1px solid #cbd5e1",
                                    boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                                }}
                            >
                                <Table
                                    stickyHeader
                                    size="small"
                                    sx={{
                                        tableLayout: "fixed",
                                        width: "100%",
                                    }}
                                >
                                    <TableHead>
                                        <TableRow>
                                            <TableCell
                                                sx={{
                                                    bgcolor: "#1E3A8A",
                                                    color: "#fff",
                                                    fontWeight: "bold",
                                                    fontSize: "0.78rem",
                                                    width: 50,
                                                    textAlign: "center",
                                                    borderRight: "1px solid #3b82f6",
                                                }}
                                            >
                                                S.No
                                            </TableCell>
                                            {enabledColumns.map((col) => {
                                                const isFiltered = Boolean(columnFilters[col.key] || numRangeFilter[col.key]);

                                                return (
                                                    <TableCell
                                                        key={col.key}
                                                        sx={{
                                                            bgcolor: "#1E3A8A",
                                                            color: "#fff",
                                                            fontWeight: "bold",
                                                            fontSize: "0.78rem",
                                                            borderRight: "1px solid #3b82f6",
                                                            cursor: "pointer",
                                                            p: "6px 8px",
                                                            overflow: "hidden",
                                                            textOverflow: "ellipsis",
                                                        }}
                                                        onClick={(e) => {
                                                            if (col.isNumeric) {
                                                                openNumFilter(e, col.key);
                                                            } else {
                                                                setFilterAnchor(e.currentTarget);
                                                                setActiveHeader(col.key);
                                                            }
                                                        }}
                                                    >
                                                        {col.isNumeric ? (
                                                            <SortableHeaderLabel
                                                                label={col.label}
                                                                columnKey={col.key}
                                                                sortConfig={numSortConfig}
                                                                onSort={handleNumSort}
                                                                onOpenFilter={(e) => openNumFilter(e, col.key)}
                                                            />
                                                        ) : (
                                                            <Box display="flex" alignItems="center" justifyContent="space-between">
                                                                <Typography variant="caption" sx={{ fontWeight: "bold", fontSize: "0.78rem", color: "#fff" }}>
                                                                    {col.label}
                                                                    {isFiltered && (
                                                                        <Chip label="•" size="small" sx={{ ml: 0.5, height: 14, width: 14, p: 0, bgcolor: "#f59e0b", color: "#fff" }} />
                                                                    )}
                                                                </Typography>
                                                                <IconButton
                                                                    size="small"
                                                                    sx={{ color: "#fff", p: 0, ml: 0.5 }}
                                                                    onClick={(e) => {
                                                                        e.stopPropagation();
                                                                        setSortConfig((prev) => ({
                                                                            key: col.key,
                                                                            direction: prev.key === col.key && prev.direction === "asc" ? "desc" : "asc",
                                                                        }));
                                                                    }}
                                                                >
                                                                    {sortConfig.key === col.key ? (
                                                                        sortConfig.direction === "asc" ? (
                                                                            <ArrowDropDownIcon fontSize="small" />
                                                                        ) : (
                                                                            <ArrowDropUpIcon fontSize="small" />
                                                                        )
                                                                    ) : (
                                                                        <ArrowDropDownIcon fontSize="small" sx={{ opacity: 0.3 }} />
                                                                    )}
                                                                </IconButton>
                                                            </Box>
                                                        )}
                                                    </TableCell>
                                                );
                                            })}
                                        </TableRow>
                                    </TableHead>

                                    <TableBody>
                                        {groupedData.length === 0 ? (
                                            <TableRow>
                                                <TableCell colSpan={enabledColumns.length + 1} align="center" sx={{ py: 6, color: "#64748b" }}>
                                                    No records found for the selected period.
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            groupedData.map((group) => {
                                                const isCollapsed = Boolean(collapsedGroups[group.groupName]);

                                                return (
                                                    <React.Fragment key={group.groupName}>
                                                        {/* GROUP HEADER ROW */}
                                                        <TableRow
                                                            sx={{
                                                                bgcolor: "#dbeafe", // Mild Blue
                                                                cursor: "pointer",
                                                                "&:hover": { bgcolor: "#bfdbfe" },
                                                                borderTop: "1px solid #93c5fd",
                                                                borderBottom: "1px solid #93c5fd",
                                                            }}
                                                            onClick={() => {
                                                                setCollapsedGroups((prev) => ({
                                                                    ...prev,
                                                                    [group.groupName]: !prev[group.groupName],
                                                                }));
                                                            }}
                                                        >
                                                            <TableCell
                                                                colSpan={enabledColumns.length + 1}
                                                                sx={{
                                                                    color: "#1e3a8a",
                                                                    fontWeight: "bold",
                                                                    fontSize: "0.85rem",
                                                                    py: 0.75,
                                                                }}
                                                            >
                                                                <Box display="flex" alignItems="center" justifyContent="space-between">
                                                                    <Box display="flex" alignItems="center" gap={1}>
                                                                        {isCollapsed ? <ChevronRightIcon fontSize="small" sx={{ color: "#1e3a8a" }} /> : <ExpandMoreIcon fontSize="small" sx={{ color: "#1e3a8a" }} />}
                                                                        <Typography sx={{ fontWeight: 700, fontSize: "0.85rem", color: "#1e3a8a" }}>
                                                                            {group.groupName} ({group.items.length})
                                                                        </Typography>
                                                                    </Box>
                                                                    <Typography sx={{ fontWeight: 600, fontSize: "0.8rem", color: "#1e40af" }}>
                                                                        Total Bags: {group.totalBags}
                                                                    </Typography>
                                                                </Box>
                                                            </TableCell>
                                                        </TableRow>

                                                        {/* GROUP ITEMS */}
                                                        {!isCollapsed &&
                                                            group.items.map((row, rowIdx) => (
                                                                <TableRow
                                                                    key={`${group.groupName}-${row._raw.Do_Id || rowIdx}-${rowIdx}`}
                                                                    hover
                                                                    sx={{
                                                                        "&:nth-of-type(even)": { bgcolor: "#f8fafc" },
                                                                        "&:hover": { bgcolor: "#f1f5f9" },
                                                                    }}
                                                                >
                                                                    <TableCell align="center" sx={{ fontSize: "0.78rem", py: 0.5, borderRight: "1px solid #e2e8f0" }}>
                                                                        {rowIdx + 1}
                                                                    </TableCell>
                                                                    {enabledColumns.map((col) => (
                                                                        <TableCell
                                                                            key={col.key}
                                                                            align={col.isNumeric ? "right" : "left"}
                                                                            sx={{
                                                                                fontSize: "0.78rem",
                                                                                py: 0.5,
                                                                                px: 1,
                                                                                borderRight: "1px solid #e2e8f0",
                                                                                overflow: "hidden",
                                                                                textOverflow: "ellipsis",
                                                                                wordBreak: "break-word",
                                                                            }}
                                                                        >
                                                                            {row[col.key] !== undefined && row[col.key] !== null
                                                                                ? String(row[col.key])
                                                                                : "-"}
                                                                        </TableCell>
                                                                    ))}
                                                                </TableRow>
                                                            ))}
                                                    </React.Fragment>
                                                );
                                            })
                                        )}
                                    </TableBody>
                                </Table>
                            </TableContainer>

                            <CommonPagination
                                totalRows={filteredRows.length}
                                page={page}
                                rowsPerPage={rowsPerPage}
                                onPageChange={setPage}
                                onRowsPerPageChange={setRowsPerPage}
                            />
                        </>
                    )}
                </Box>
            </AppLayout>

            {/* COLUMN SETTINGS MENU */}
            <Menu
                anchorEl={settingsAnchor}
                open={Boolean(settingsAnchor)}
                onClose={() => setSettingsAnchor(null)}
                PaperProps={{
                    sx: {
                        width: 320,
                        maxHeight: 520,
                        p: 1.5,
                    },
                }}
            >
                <Box display="flex" justifyContent="space-between" alignItems="center" mb={1}>
                    <Typography variant="subtitle2" sx={{ fontWeight: "bold", color: "#1E3A8A" }}>
                        Configure Columns
                    </Typography>
                    <Button
                        size="small"
                        color="info"
                        sx={{ fontSize: "0.75rem", textTransform: "none", p: 0.5 }}
                        onClick={handleClearTemplate}
                    >
                        Reset
                    </Button>
                </Box>

                <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
                    <Box sx={{ maxHeight: 420, overflowY: "auto", pr: 0.5 }}>
                        {/* ENABLED COLUMNS */}
                        <Typography fontSize="0.75rem" fontWeight={700} color="#1e3a8a" mb={0.5}>
                            Enabled Columns ({columns.filter((c) => c.enabled).length})
                        </Typography>

                        <SortableContext
                            items={columns.filter((c) => c.enabled).map((c) => c.key)}
                            strategy={verticalListSortingStrategy}
                        >
                            {columns
                                .filter((c) => c.enabled)
                                .map((col) => (
                                    <SortableColumnRow
                                        key={col.key}
                                        column={col}
                                        onToggle={(key) => handleToggleColumn(key, false)}
                                    />
                                ))}
                        </SortableContext>

                        {/* DISABLED COLUMNS */}
                        <Typography fontSize="0.75rem" fontWeight={700} color="#64748b" mt={2} mb={0.5}>
                            Disabled Columns ({columns.filter((c) => !c.enabled).length})
                        </Typography>

                        <SortableContext
                            items={columns.filter((c) => !c.enabled).map((c) => c.key)}
                            strategy={verticalListSortingStrategy}
                        >
                            {columns
                                .filter((c) => !c.enabled)
                                .map((col) => (
                                    <SortableColumnRow
                                        key={col.key}
                                        column={col}
                                        onToggle={(key) => handleToggleColumn(key, true)}
                                    />
                                ))}
                        </SortableContext>
                    </Box>
                </DndContext>
            </Menu>

            {/* HEADER FILTER MENU */}
            <HeaderFilterMenu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor)}
                onClose={() => {
                    setFilterAnchor(null);
                    setActiveHeader(null);
                }}
                columnLabel={columns.find((c) => c.key === activeHeader)?.label || activeHeader || undefined}
                options={filterOptions}
                selectedValues={activeHeader ? columnFilters[activeHeader] : undefined}
                onFilterChange={(selected: string[] | undefined) => {
                    if (activeHeader) {
                        setColumnFilters((prev) => {
                            const copy = { ...prev };
                            if (selected === undefined) {
                                delete copy[activeHeader];
                            } else {
                                copy[activeHeader] = selected;
                            }
                            return copy;
                        });
                    }
                }}
            />

            {/* NUMERICAL RANGE FILTER MENU */}
            <NumericalFilterMenu
                anchorEl={numFilterAnchor}
                open={Boolean(numFilterAnchor)}
                onClose={() => setNumFilterAnchor(null)}
                activeHeader={numActiveHeader}
                min={numActiveHeader ? getMinMaxFallback(numActiveHeader).min : 0}
                max={numActiveHeader ? getMinMaxFallback(numActiveHeader).max : 100}
                rangeFilter={numRangeFilter}
                onRangeChange={(key, range) => setNumRangeFilter((p: any) => ({ ...p, [key]: range }))}
                onClear={clearRangeFilter}
            />

            {/* SAVE TEMPLATE DIALOG */}
            <Dialog open={saveDialogOpen} onClose={() => setSaveDialogOpen(false)} maxWidth="xs" fullWidth>
                <DialogTitle sx={{ fontWeight: "bold", color: "#1E3A8A" }}>
                    {selectedTemplateId ? "Update Template" : "Save as New Template"}
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
                        sx={{ mt: 1 }}
                    />
                </DialogContent>
                <DialogActions sx={{ p: 2 }}>
                    <Button onClick={() => setSaveDialogOpen(false)} color="inherit">
                        Cancel
                    </Button>
                    <Button onClick={handleQuickSave} variant="contained" sx={{ bgcolor: "#1E3A8A" }}>
                        Save
                    </Button>
                </DialogActions>
            </Dialog>
        </Box>
    );
};

export default ShetSheetReport;
