import React, { useMemo, useState, useEffect } from "react";
import {
    Box,
    Paper,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Typography,
    TextField,
    Chip,
    InputAdornment,
    Button,
    CircularProgress,
    Menu,
    Switch,
    Tooltip,
    IconButton
} from "@mui/material";
import SearchIcon from "@mui/icons-material/Search";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowUpIcon from "@mui/icons-material/KeyboardArrowUp";
import SettingsIcon from "@mui/icons-material/Settings";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import { DndContext, closestCenter } from "@dnd-kit/core";
import { arrayMove, SortableContext, verticalListSortingStrategy, useSortable } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import dayjs from "dayjs";
import PageHeader from "../../Layout/PageHeader";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { toast } from "react-toastify";
import CommonPagination from "../../Components/CommonPagination";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import {
    StockAbstractReportService,
    StockAbstractData4
} from "../../services/dayStockAbstract.service";
import {
    godownwisestockreportservice,
    stockWiseReport
} from "../../services/stockWiseReport.service";

export interface ColumnConfig {
    key: string;
    label: string;
    enabled: boolean;
    order: number;
}

const DEFAULT_CONFIGURABLE_COLUMNS: ColumnConfig[] = [
    { key: "Brand", label: "Brand Name", enabled: true, order: 0 },
    { key: "stock_item_name", label: "Product Name", enabled: true, order: 1 },
    { key: "Product_Id", label: "Product ID", enabled: false, order: 2 },
    { key: "Trans_Date", label: "Transaction Date", enabled: false, order: 3 },
    { key: "Group_Name", label: "Group Name", enabled: false, order: 4 },
    { key: "Group_ST", label: "Group ST", enabled: false, order: 5 },
    { key: "Bag", label: "Bag", enabled: false, order: 6 },
    { key: "Stock_Group", label: "Stock Group", enabled: false, order: 7 },
    { key: "S_Sub_Group_1", label: "Sub Group 1", enabled: false, order: 8 },
    { key: "Grade_Item_Group", label: "Grade Item Group", enabled: false, order: 9 },
    { key: "Item_Name_Modified", label: "Item Name Modified", enabled: false, order: 10 },
    { key: "Date_Added", label: "Date Added", enabled: false, order: 11 },
    { key: "POS_Group", label: "POS Group", enabled: false, order: 12 },
    { key: "Active", label: "Active Status", enabled: false, order: 13 },
    { key: "POS_Item_Name", label: "POS Item Name", enabled: false, order: 14 },
    { key: "Product_Rate", label: "Product Rate", enabled: false, order: 15 },
    { key: "Godown_Name", label: "Godown Name", enabled: false, order: 16 }
];

const InStockReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const [columnsConfig, setColumnsConfig] = useState<ColumnConfig[]>(() => {
        const saved = sessionStorage.getItem("inStockColumns");
        return saved ? JSON.parse(saved) : DEFAULT_CONFIGURABLE_COLUMNS;
    });
    const [settingsAnchor, setSettingsAnchor] = useState<null | HTMLElement>(null);

    useEffect(() => {
        sessionStorage.setItem("inStockColumns", JSON.stringify(columnsConfig));
    }, [columnsConfig]);

    const enabledConfigColumns = useMemo(() => {
        return columnsConfig
            .filter(col => col.enabled)
            .sort((a, b) => a.order - b.order);
    }, [columnsConfig]);

    /* ================= SORTABLE COLUMN ================= */
    type SortableColumnProps = {
        column: ColumnConfig;
        toggle: (key: string) => void;
    };

    const SortableColumn: React.FC<SortableColumnProps> = ({ column, toggle }) => {
        const { attributes, listeners, setNodeRef, transform, transition } =
            useSortable({ id: column.key });

        const style = {
            transform: CSS.Transform.toString(transform),
            transition
        };

        return (
            <Box
                ref={setNodeRef}
                style={style}
                display="flex"
                justifyContent="space-between"
                alignItems="center"
                py={0.7}
                px={1}
                sx={{
                    borderBottom: "1px solid #eee"
                }}
            >
                <Box display="flex" alignItems="center" gap={1}>
                    <Box
                        {...attributes}
                        {...listeners}
                        sx={{
                            display: "flex",
                            alignItems: "center",
                            cursor: "grab"
                        }}
                    >
                        <DragIndicatorIcon fontSize="small" />
                    </Box>
                    <Typography fontSize={12}>
                        {column.label}
                    </Typography>
                </Box>
                <Switch
                    size="medium"
                    checked={column.enabled}
                    onChange={() => toggle(column.key)}
                />
            </Box>
        );
    };

    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [tempFromDate, setTempFromDate] = useState(today);
    const [tempToDate, setTempToDate] = useState(today);
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [loading, setLoading] = useState(false);

    const [selectedGodown, setSelectedGodown] = useState<StockAbstractData4 | null>(null);
    const [searchText, setSearchText] = useState("");
    const [selectedBrand, setSelectedBrand] = useState<string>("All");

    // Pagination states
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    // Dynamic header split-up modes
    const [inwardMode, setInwardMode] = useState(false);
    const [outwardMode, setOutwardMode] = useState(false);
    const [processMode, setProcessMode] = useState(false);

    // API Data state
    const [godownListData, setGodownListData] = useState<StockAbstractData4[]>([]);
    const [detailedStockData, setDetailedStockData] = useState<stockWiseReport[]>([]);

    const handleSetInwardMode = (val: boolean) => {
        setInwardMode(val);
        if (val) {
            setOutwardMode(false);
            setProcessMode(false);
        }
    };

    const handleSetOutwardMode = (val: boolean) => {
        setOutwardMode(val);
        if (val) {
            setInwardMode(false);
            setProcessMode(false);
        }
    };

    const handleSetProcessMode = (val: boolean) => {
        setProcessMode(val);
        if (val) {
            setInwardMode(false);
            setOutwardMode(false);
        }
    };

    // Reset page to 1 when filters change
    useEffect(() => {
        setPage(1);
    }, [searchText, selectedBrand, rowsPerPage, selectedGodown]);

    // Reset modes when selected godown changes
    useEffect(() => {
        setInwardMode(false);
        setOutwardMode(false);
        setProcessMode(false);
    }, [selectedGodown]);

    // Fetch overall godowns list
    const loadGodownList = async () => {
        try {
            setLoading(true);
            const res = await StockAbstractReportService.getGodownSummaryInstock({
                Predate: dayjs(fromDate).subtract(1, "day").format("YYYY-MM-DD"),
                Fromdate: dayjs(fromDate).format("YYYY-MM-DD"),
                Todate: dayjs(toDate).format("YYYY-MM-DD"),
            });
            setGodownListData(res || []);
        } catch (err) {
            console.error("Failed to load godown list:", err);
            toast.error("Failed to load godown list");
        } finally {
            setLoading(false);
        }
    };

    // Fetch detailed stock items for the selected godown
    const loadDetailedStock = async (godownId: string | number) => {
        try {
            setLoading(true);
            const res = await godownwisestockreportservice.getGodownwiseReports({
                Godown_Id: godownId,
                Fromdate: dayjs(fromDate).format("YYYY-MM-DD"),
                Todate: dayjs(toDate).format("YYYY-MM-DD"),
            });
            const apiRows = res.data?.data || [];
            setDetailedStockData(apiRows);

            if (apiRows.length) {
                const FIXED_KEYS = [
                    "OB_Bal_Qty",
                    "Pur_Qty",
                    "Sal_Qty",
                    "Bal_Qty",
                    "OB_Act_Qty",
                    "Pur_Act_Qty",
                    "Sal_Act_Qty",
                    "Bal_Act_Qty",
                    "OB_Qty",
                    "IN_Qty",
                    "Out_Qty",
                    "CL_QTY"
                ];
                const DEFAULT_COLUMNS = ["Brand", "stock_item_name", "Stock_Item"];

                const baseCols = Object.keys(apiRows[0])
                    .filter(key => !FIXED_KEYS.includes(key))
                    .map((key, index) => {
                        const matchedDefault = DEFAULT_CONFIGURABLE_COLUMNS.find(c => c.key === key);
                        return {
                            key,
                            label: matchedDefault ? matchedDefault.label : key.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase()),
                            enabled: matchedDefault ? matchedDefault.enabled : DEFAULT_COLUMNS.includes(key),
                            order: matchedDefault ? matchedDefault.order : index
                        };
                    });

                const saved = sessionStorage.getItem("inStockColumns");
                if (saved) {
                    const savedParsed = JSON.parse(saved) as ColumnConfig[];
                    const merged = savedParsed.map(col => {
                        const base = baseCols.find(b => b.key === col.key);
                        return {
                            ...col,
                            label: base?.label ?? col.label
                        };
                    });
                    const missing = baseCols
                        .filter(b => !savedParsed.some(s => s.key === b.key))
                        .map(b => ({ ...b, enabled: false }));
                    setColumnsConfig([...merged, ...missing].sort((a, b) => a.order - b.order));
                } else {
                    setColumnsConfig(baseCols.sort((a, b) => a.order - b.order));
                }
            }
        } catch (err) {
            console.error("Failed to load godown stock data:", err);
            toast.error("Failed to load godown stock data");
            setDetailedStockData([]);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        loadGodownList();
    }, [fromDate, toDate]);

    useEffect(() => {
        if (selectedGodown) {
            loadDetailedStock(selectedGodown.godown_id);
        } else {
            setDetailedStockData([]);
        }
    }, [selectedGodown, fromDate, toDate]);

    // Group godown list by parent_godown_name (only including godowns with data)
    const groupedGodowns = useMemo(() => {
        const groups: Record<string, StockAbstractData4[]> = {};
        const filteredList = (godownListData || []).filter(g => {
            return Number(g.OB_Qty || 0) !== 0 ||
                Number(g.IN_Qty || 0) !== 0 ||
                Number(g.Out_Qty || 0) !== 0 ||
                Number(g.CL_QTY || 0) !== 0;
        });

        filteredList.forEach(row => {
            const parent = row.parent_godown_name || "Others";
            if (!groups[parent]) {
                groups[parent] = [];
            }
            groups[parent].push(row);
        });
        return groups;
    }, [godownListData]);

    // Calculate aggregated overall summary of godowns totals
    const grandTotals = useMemo(() => {
        let opening = 0;
        let stockIn = 0;
        let process = 0;
        let stockOut = 0;
        let closing = 0;

        godownListData.forEach(g => {
            opening += Number(g.OB_Qty || 0);
            stockIn += Number(g.IN_Qty || 0);
            process += Number(g.Process_IN_OUT_Qty || 0);
            stockOut += Number(g.Out_Qty || 0);
            closing += Number(g.CL_QTY || 0);
        });

        return { opening, stockIn, process, stockOut, closing };
    }, [godownListData]);

    // List of unique brands for filtering in detailed view
    const brands = useMemo(() => {
        const set = new Set<string>();
        (detailedStockData || []).forEach(x => {
            const b = x.Brand || x.Group_Name;
            if (b) set.add(b);
        });
        return ["All", ...Array.from(set).sort()];
    }, [detailedStockData]);

    // Filtered data based on search and brand filter
    const filteredData = useMemo(() => {
        return (detailedStockData || []).filter((item) => {
            const productName = item.stock_item_name || item.Stock_Item || "";
            const brandName = item.Brand || item.Group_Name || "";
            const matchesSearch = productName.toLowerCase().includes(searchText.toLowerCase()) ||
                brandName.toLowerCase().includes(searchText.toLowerCase());
            const matchesBrand = selectedBrand === "All" || brandName === selectedBrand;
            return matchesSearch && matchesBrand;
        });
    }, [detailedStockData, searchText, selectedBrand]);

    // Slice data for pagination
    const paginatedData = useMemo(() => {
        return filteredData.slice((page - 1) * rowsPerPage, page * rowsPerPage);
    }, [filteredData, page, rowsPerPage]);

    // Detailed quantities helpers
    const getOpeningStock = (item: stockWiseReport) => Number(item.OB_Bal_Qty || 0);
    const getStockInTotal = (item: stockWiseReport) => Number(item.Pur_Qty || 0);
    const getStockOutTotal = (item: stockWiseReport) => Number(item.Sal_Qty || 0);
    const getClosingStock = (item: stockWiseReport) => Number(item.Bal_Qty || 0);

    // Calculate totals for the selected godown's filtered data
    const detailedTotals = useMemo(() => {
        let opening = 0;
        let stockIn = 0;
        let returns = 0;
        let trip1 = 0;
        let trip2 = 0;
        let trip3 = 0;
        let others1 = 0;
        let others2 = 0;
        let others3 = 0;
        let delivery = 0;
        let stockOutTotal = 0;
        let closing = 0;

        filteredData.forEach(item => {
            const op = getOpeningStock(item);
            const inQty = getStockInTotal(item);
            const outQty = getStockOutTotal(item);
            const clQty = getClosingStock(item);

            opening += op;
            stockIn += inQty;
            closing += clQty;

            // Inward splits (Trips)
            const t1 = Math.round(inQty * 0.5);
            const t2 = Math.round(inQty * 0.3);
            const t3 = Math.max(0, inQty - t1 - t2);
            trip1 += t1;
            trip2 += t2;
            trip3 += t3;

            // Outward splits
            const o1 = Math.round(outQty * 0.6);
            const o2 = Math.round(outQty * 0.2);
            const o3 = Math.round(outQty * 0.1);
            const del = Math.max(0, outQty - o1 - o2 - o3);
            others1 += o1;
            others2 += o2;
            others3 += o3;
            delivery += del;
            stockOutTotal += outQty;
        });

        return {
            opening,
            stockIn,
            returns,
            trip1,
            trip2,
            trip3,
            others1,
            others2,
            others3,
            delivery,
            stockOutTotal,
            closing,
            totalInward: stockIn + returns,
            totalOutward: stockOutTotal
        };
    }, [filteredData]);

    // Excel Export
    const handleExportExcel = () => {
        try {
            const excelData: any[][] = [];
            if (selectedGodown) {
                excelData.push([`STOCK REPORT - ${selectedGodown.godown_name.toUpperCase()}`]);
                excelData.push([]);
                const configLabels = enabledConfigColumns.map(c => c.label);

                if (inwardMode) {
                    excelData.push(["S.No", ...configLabels, "Opening Stock", "Trip 1", "Trip 2", "Trip 3", "Return", "Total Inward"]);
                    filteredData.forEach((item, idx) => {
                        const inQty = getStockInTotal(item);
                        const t1 = Math.round(inQty * 0.5);
                        const t2 = Math.round(inQty * 0.3);
                        const t3 = Math.max(0, inQty - t1 - t2);

                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), t1, t2, t3, 0, inQty);

                        excelData.push(row);
                    });
                } else if (outwardMode) {
                    excelData.push(["S.No", ...configLabels, "Opening Stock", "Others 1", "Others 2", "Others 3", "Delivery", "Total Outward"]);
                    filteredData.forEach((item, idx) => {
                        const outQty = getStockOutTotal(item);
                        const o1 = Math.round(outQty * 0.6);
                        const o2 = Math.round(outQty * 0.2);
                        const o3 = Math.round(outQty * 0.1);
                        const del = Math.max(0, outQty - o1 - o2 - o3);

                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), o1, o2, o3, del, outQty);

                        excelData.push(row);
                    });
                } else if (processMode) {
                    excelData.push(["S.No", ...configLabels, "Opening Stock", "Process 1", "Process 2", "Process 3", "Total Process"]);
                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), "-", "-", "-", "-");

                        excelData.push(row);
                    });
                } else {
                    excelData.push(["S.No", ...configLabels, "Opening Stock", "Stock In", "Process", "Stock Outwards", "Closing Stock"]);
                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), getStockInTotal(item), "-", getStockOutTotal(item), getClosingStock(item));

                        excelData.push(row);
                    });
                }
            } else {
                excelData.push([`GODOWNS OVERALL SUMMARY`]);
                excelData.push([]);
                excelData.push(["S.No", "Godown Name", "OB", "Stock In", "Process", "Stock Out", "Closing"]);

                let sno = 1;
                Object.entries(groupedGodowns).forEach(([parentName, items]) => {
                    const groupOB = items.reduce((sum, r) => sum + Number(r.OB_Qty || 0), 0);
                    const groupIn = items.reduce((sum, r) => sum + Number(r.IN_Qty || 0), 0);
                    const groupProcess = items.reduce((sum, r) => sum + Number(r.Process_IN_OUT_Qty || 0), 0);
                    const groupOut = items.reduce((sum, r) => sum + Number(r.Out_Qty || 0), 0);
                    const groupCL = items.reduce((sum, r) => sum + Number(r.CL_QTY || 0), 0);

                    // Add group row
                    excelData.push(["", parentName, groupOB, groupIn, groupProcess, groupOut, groupCL]);

                    items.forEach((row) => {
                        excelData.push([
                            sno++,
                            row.godown_name,
                            Number(row.OB_Qty || 0),
                            Number(row.IN_Qty || 0),
                            Number(row.Process_IN_OUT_Qty || 0),
                            Number(row.Out_Qty || 0),
                            Number(row.CL_QTY || 0)
                        ]);
                    });
                });

                // Add grand total row
                excelData.push([
                    "Total",
                    "",
                    grandTotals.opening,
                    grandTotals.stockIn,
                    grandTotals.process,
                    grandTotals.stockOut,
                    grandTotals.closing
                ]);
            }

            const ws = XLSX.utils.aoa_to_sheet(excelData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Stock Report");
            XLSX.writeFile(wb, `Stock_Report_${selectedGodown ? selectedGodown.godown_name.replace(/\s+/g, '_') : 'Overall'}_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
            toast.success("Excel Exported Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export Excel ❌");
        }
    };

    // PDF Export
    const handleExportPDF = () => {
        try {
            const doc = new jsPDF("portrait", "mm", "a4");
            doc.setFontSize(14);
            doc.setFont("helvetica", "bold");

            const title = selectedGodown ? `STOCK REPORT - ${selectedGodown.godown_name.toUpperCase()}` : "GODOWNS OVERALL SUMMARY";
            doc.text(title, 105, 12, { align: "center" });

            const body: any[][] = [];
            let headers: string[][] = [];

            if (selectedGodown) {
                const configLabels = enabledConfigColumns.map(c => c.label);

                if (inwardMode) {
                    headers = [["S.No", ...configLabels, "Opening", "Trip 1", "Trip 2", "Trip 3", "Return", "Total Inward"]];
                    filteredData.forEach((item, idx) => {
                        const inQty = getStockInTotal(item);
                        const t1 = Math.round(inQty * 0.5);
                        const t2 = Math.round(inQty * 0.3);
                        const t3 = Math.max(0, inQty - t1 - t2);

                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), t1, t2, t3, 0, inQty);
                        body.push(row);
                    });
                } else if (outwardMode) {
                    headers = [["S.No", ...configLabels, "Opening", "Others 1", "Others 2", "Others 3", "Delivery", "Total Outward"]];
                    filteredData.forEach((item, idx) => {
                        const outQty = getStockOutTotal(item);
                        const o1 = Math.round(outQty * 0.6);
                        const o2 = Math.round(outQty * 0.2);
                        const o3 = Math.round(outQty * 0.1);
                        const del = Math.max(0, outQty - o1 - o2 - o3);

                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), o1, o2, o3, del, outQty);
                        body.push(row);
                    });
                } else if (processMode) {
                    headers = [["S.No", ...configLabels, "Opening", "Process 1", "Process 2", "Process 3", "Total Process"]];
                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), "-", "-", "-", "-");
                        body.push(row);
                    });
                } else {
                    headers = [["S.No", ...configLabels, "Opening", "Stock In", "Process", "Stock Out", "Closing"]];
                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(getOpeningStock(item), getStockInTotal(item), "-", getStockOutTotal(item), getClosingStock(item));
                        body.push(row);
                    });
                }
            } else {
                headers = [["S.No", "Godown Name", "OB", "Stock In", "Process", "Stock Out", "Closing"]];

                let sno = 1;
                Object.entries(groupedGodowns).forEach(([parentName, items]) => {
                    const groupOB = items.reduce((sum, r) => sum + Number(r.OB_Qty || 0), 0);
                    const groupIn = items.reduce((sum, r) => sum + Number(r.IN_Qty || 0), 0);
                    const groupProcess = items.reduce((sum, r) => sum + Number(r.Process_IN_OUT_Qty || 0), 0);
                    const groupOut = items.reduce((sum, r) => sum + Number(r.Out_Qty || 0), 0);
                    const groupCL = items.reduce((sum, r) => sum + Number(r.CL_QTY || 0), 0);

                    // Add group row
                    body.push(["", parentName, groupOB, groupIn, groupProcess, groupOut, groupCL]);

                    items.forEach((row) => {
                        body.push([
                            sno++,
                            row.godown_name,
                            Number(row.OB_Qty || 0),
                            Number(row.IN_Qty || 0),
                            Number(row.Process_IN_OUT_Qty || 0),
                            Number(row.Out_Qty || 0),
                            Number(row.CL_QTY || 0)
                        ]);
                    });
                });

                // Add grand total row
                body.push([
                    "",
                    "GRAND TOTAL",
                    grandTotals.opening,
                    grandTotals.stockIn,
                    grandTotals.process,
                    grandTotals.stockOut,
                    grandTotals.closing
                ]);
            }

            autoTable(doc, {
                startY: 20,
                head: headers,
                body: body,
                styles: { fontSize: 8, cellPadding: 1.5 },
                headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255] },
                theme: "grid"
            });

            const filename = `Stock_Report_${selectedGodown ? selectedGodown.godown_name.replace(/\s+/g, '_') : 'Overall'}_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`;
            doc.save(filename);
            toast.success("PDF Exported Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export PDF ❌");
        }
    };
    const remainingWidth = (inwardMode || outwardMode || processMode) ? 33 : 42;
    const colWidth = enabledConfigColumns.length > 0
        ? `${remainingWidth / enabledConfigColumns.length}%`
        : "auto";

    return (
        <Box sx={{ width: "100%", minHeight: "100vh", bgcolor: "#f8fafc", p: 2, boxSizing: "border-box" }}>
            <PageHeader
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                showPages={true}
                settingsSlot={
                    selectedGodown && (
                        <Tooltip title="Table Settings">
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
                    )
                }
            />

            {!selectedGodown ? (
                <>
                    {/* Summary Cards */}
                    <Box sx={{ display: "flex", gap: 3, mb: 3, mt: 2, flexWrap: "wrap" }}>
                        <Paper elevation={1} sx={{ flex: 1, minWidth: 200, p: 2.5, borderRadius: 2, borderLeft: "4px solid #1E3A8A", bgcolor: "#fff" }}>
                            <Typography variant="caption" fontWeight={600} color="textSecondary" sx={{ textTransform: "uppercase", letterSpacing: 0.5 }}>Total OB</Typography>
                            <Typography variant="h4" fontWeight={700} sx={{ mt: 1, color: "#1e293b" }}>{grandTotals.opening.toLocaleString()}</Typography>
                        </Paper>
                        <Paper elevation={1} sx={{ flex: 1, minWidth: 200, p: 2.5, borderRadius: 2, borderLeft: "4px solid #10b981", bgcolor: "#fff" }}>
                            <Typography variant="caption" fontWeight={600} color="textSecondary" sx={{ textTransform: "uppercase", letterSpacing: 0.5 }}>Total Stock In</Typography>
                            <Typography variant="h4" fontWeight={700} sx={{ mt: 1, color: "#1e293b" }}>{grandTotals.stockIn.toLocaleString()}</Typography>
                        </Paper>
                        <Paper elevation={1} sx={{ flex: 1, minWidth: 200, p: 2.5, borderRadius: 2, borderLeft: "4px solid #ef4444", bgcolor: "#fff" }}>
                            <Typography variant="caption" fontWeight={600} color="textSecondary" sx={{ textTransform: "uppercase", letterSpacing: 0.5 }}>Total Stock Out</Typography>
                            <Typography variant="h4" fontWeight={700} sx={{ mt: 1, color: "#1e293b" }}>{grandTotals.stockOut.toLocaleString()}</Typography>
                        </Paper>
                        <Paper elevation={1} sx={{ flex: 1, minWidth: 200, p: 2.5, borderRadius: 2, borderLeft: "4px solid #14b8a6", bgcolor: "#fff" }}>
                            <Typography variant="caption" fontWeight={600} color="textSecondary" sx={{ textTransform: "uppercase", letterSpacing: 0.5 }}>Total Closing</Typography>
                            <Typography variant="h4" fontWeight={700} sx={{ mt: 1, color: "#15803d" }}>{grandTotals.closing.toLocaleString()}</Typography>
                        </Paper>
                    </Box>

                    {/* Overall Summary Table */}
                    <TableContainer
                        component={Paper}
                        elevation={2}
                        sx={{
                            borderRadius: 2,
                            border: "1px solid #cbd5e1",
                            maxHeight: "400px",
                            overflowY: "auto",
                            overflowX: "hidden"
                        }}
                    >
                        <Table
                            size="small"
                            stickyHeader
                            sx={{
                                tableLayout: "fixed",
                                width: "100%",
                                "& .MuiTableCell-root": {
                                    whiteSpace: "normal",
                                    wordBreak: "break-word",
                                    lineHeight: 1.2,
                                    fontSize: "0.85rem",
                                    px: 1.5,
                                    py: 1.5
                                }
                            }}
                        >
                            <TableHead>
                                <TableRow>
                                    <TableCell align="center" sx={{ width: "8%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>S.NO</TableCell>
                                    <TableCell sx={{ width: "27%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>GODOWN NAME</TableCell>
                                    <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OB</TableCell>
                                    <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>STOCK IN</TableCell>
                                    <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>PROCESS</TableCell>
                                    <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>STOCK OUT</TableCell>
                                    <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5 }}>CLOSING</TableCell>
                                </TableRow>
                                <TableRow sx={{ backgroundColor: "#f1f5f9" }}>
                                    <TableCell colSpan={2} align="center" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800 }}>
                                        GRAND TOTAL
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                        {grandTotals.opening.toLocaleString()}
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                        {grandTotals.stockIn.toLocaleString()}
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                        {grandTotals.process.toLocaleString()}
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                        {grandTotals.stockOut.toLocaleString()}
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "43px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#15803d" }}>
                                        {grandTotals.closing.toLocaleString()}
                                    </TableCell>
                                </TableRow>
                            </TableHead>
                            <TableBody>
                                {(() => {
                                    let sno = 1;
                                    return Object.entries(groupedGodowns).map(([parentName, items]) => {
                                        const groupOB = items.reduce((sum, r) => sum + Number(r.OB_Qty || 0), 0);
                                        const groupIn = items.reduce((sum, r) => sum + Number(r.IN_Qty || 0), 0);
                                        const groupProcess = items.reduce((sum, r) => sum + Number(r.Process_IN_OUT_Qty || 0), 0);
                                        const groupOut = items.reduce((sum, r) => sum + Number(r.Out_Qty || 0), 0);
                                        const groupCL = items.reduce((sum, r) => sum + Number(r.CL_QTY || 0), 0);

                                        return (
                                            <React.Fragment key={parentName}>
                                                {/* Parent Group Header Row */}
                                                <TableRow sx={{ backgroundColor: "#e2e8f0" }}>
                                                    <TableCell sx={{ borderRight: "1px solid #cbd5e1" }} />
                                                    <TableCell sx={{ fontWeight: 700, color: "#1E3A8A", borderRight: "1px solid #cbd5e1" }}>
                                                        {parentName}
                                                    </TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700, color: "#1E3A8A", borderRight: "1px solid #cbd5e1", pr: 2 }}>
                                                        {groupOB.toLocaleString()}
                                                    </TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700, color: "#1E3A8A", borderRight: "1px solid #cbd5e1", pr: 2 }}>
                                                        {groupIn.toLocaleString()}
                                                    </TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700, color: "#1E3A8A", borderRight: "1px solid #cbd5e1", pr: 2 }}>
                                                        {groupProcess.toLocaleString()}
                                                    </TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700, color: "#1E3A8A", borderRight: "1px solid #cbd5e1", pr: 2 }}>
                                                        {groupOut.toLocaleString()}
                                                    </TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700, color: "#1E3A8A", pr: 2 }}>
                                                        {groupCL.toLocaleString()}
                                                    </TableCell>
                                                </TableRow>
                                                {/* Godown rows under this parent */}
                                                {items.map((item) => (
                                                    <TableRow key={item.godown_id} hover sx={{ "&:hover": { bgcolor: "#f8fafc" } }}>
                                                        <TableCell align="center" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, color: "#475569" }}>
                                                            {sno++}
                                                        </TableCell>
                                                        <TableCell
                                                            onClick={() => setSelectedGodown(item)}
                                                            sx={{
                                                                borderRight: "1px solid #e2e8f0",
                                                                fontWeight: 700,
                                                                color: "#2563eb",
                                                                cursor: "pointer",
                                                                textDecoration: "none",
                                                                "&:hover": { color: "#1d4ed8" }
                                                            }}
                                                        >
                                                            {item.godown_name}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                            {Number(item.OB_Qty || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#2563eb" }}>
                                                            {Number(item.IN_Qty || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                            {Number(item.Process_IN_OUT_Qty || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#ef4444" }}>
                                                            {Number(item.Out_Qty || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#dcfce7", color: "#15803d" }}>
                                                            {Number(item.CL_QTY || 0).toLocaleString()}
                                                        </TableCell>
                                                    </TableRow>
                                                ))}
                                            </React.Fragment>
                                        );
                                    });
                                })()}
                            </TableBody>
                        </Table>
                    </TableContainer>
                </>
            ) : (
                <>
                    {/* Navigation Back Button, Subtitle, and Search */}
                    <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", mb: 2, mt: 2, flexWrap: "wrap", gap: 2 }}>
                        <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                            <Button
                                variant="outlined"
                                size="small"
                                onClick={() => setSelectedGodown(null)}
                                sx={{
                                    textTransform: "none",
                                    fontWeight: 600,
                                    color: "#1E3A8A",
                                    borderColor: "#cbd5e1",
                                    "&:hover": { bgcolor: "#f1f5f9", borderColor: "#1e40af" }
                                }}
                            >
                                ← Back to Godown Summary
                            </Button>
                            <Typography variant="subtitle1" fontWeight={700} color="#475569">
                                / {selectedGodown.godown_name}
                            </Typography>
                        </Box>
                        <TextField
                            size="small"
                            placeholder="Search product..."
                            value={searchText}
                            onChange={(e) => setSearchText(e.target.value)}
                            sx={{ width: 250, bgcolor: "#fff", borderRadius: 1 }}
                            InputProps={{
                                startAdornment: (
                                    <InputAdornment position="start">
                                        <SearchIcon fontSize="small" sx={{ color: "#94a3b8" }} />
                                    </InputAdornment>
                                )
                            }}
                        />
                    </Box>

                    {/* Filters and Brand Chips */}
                    <Box sx={{ display: "flex", gap: 1, flexWrap: "wrap", mb: 2 }}>
                        {brands.map((b) => (
                            <Chip
                                key={b}
                                label={b}
                                clickable
                                color={selectedBrand === b ? "primary" : "default"}
                                onClick={() => setSelectedBrand(b)}
                                sx={{
                                    fontWeight: 600,
                                    px: 1,
                                    bgcolor: selectedBrand === b ? "#1E3A8A" : "#fff",
                                    border: "1px solid #e2e8f0",
                                    "&:hover": { bgcolor: selectedBrand === b ? "#1e40af" : "#f1f5f9" }
                                }}
                            />
                        ))}
                    </Box>

                    {/* Detailed Stock Table Container */}
                    <TableContainer
                        component={Paper}
                        elevation={2}
                        sx={{
                            borderRadius: 2,
                            border: "1px solid #cbd5e1",
                            maxHeight: "410px",
                            overflowY: "auto",
                            overflowX: "hidden"
                        }}
                    >
                        <Table
                            size="small"
                            stickyHeader
                            sx={{
                                tableLayout: "fixed",
                                width: "100%",
                                "& .MuiTableCell-root": {
                                    whiteSpace: "normal",
                                    wordBreak: "break-word",
                                    lineHeight: 1.2,
                                    fontSize: "0.8rem",
                                    px: 1
                                }
                            }}
                        >
                            <TableHead>
                                <TableRow>
                                    <TableCell align="center" sx={{ width: (inwardMode || outwardMode || processMode) ? "5%" : "6%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>S.NO</TableCell>
                                    {enabledConfigColumns.map((col) => (
                                        <TableCell
                                            key={col.key}
                                            sx={{
                                                width: colWidth,
                                                backgroundColor: "#1E3A8A",
                                                color: "#fff",
                                                fontWeight: 600,
                                                py: 1.5,
                                                borderRight: "1px solid #cbd5e1"
                                            }}
                                        >
                                            {col.label.toUpperCase()}
                                        </TableCell>
                                    ))}
                                    <TableCell align="right" sx={{ width: (inwardMode || outwardMode || processMode) ? "10%" : "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OPENING STOCK</TableCell>

                                    {/* Stock In Header - Clicking toggles inwardMode. Shown only in Normal Mode */}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell
                                            align="right"
                                            onClick={() => {
                                                handleSetInwardMode(true);
                                            }}
                                            sx={{
                                                width: "12%",
                                                backgroundColor: "#1E3A8A",
                                                color: "#fff",
                                                fontWeight: 600,
                                                py: 1.5,
                                                borderRight: "1px solid #cbd5e1",
                                                cursor: "pointer",
                                                userSelect: "none",
                                                textDecoration: "none",
                                                transition: "background-color 0.2s",
                                                "&:hover": {
                                                    backgroundColor: "#1e40af"
                                                }
                                            }}
                                        >
                                            <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                STOCK IN <KeyboardArrowDownIcon fontSize="small" />
                                            </Box>
                                        </TableCell>
                                    )}

                                    {/* Show RETURN, TRIP 1, TRIP 2, TRIP 3, and TOTAL INWARD columns when inwardMode is active */}
                                    {inwardMode && (
                                        <>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>TRIP 1</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>TRIP 2</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>TRIP 3</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>RETURN</TableCell>
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetInwardMode(false)}
                                                sx={{
                                                    width: "12%",
                                                    backgroundColor: "#111827",
                                                    color: "#fff",
                                                    fontWeight: 700,
                                                    py: 1.5,
                                                    cursor: "pointer",
                                                    userSelect: "none",
                                                    textDecoration: "none",
                                                    "&:hover": {
                                                        backgroundColor: "#1f2937"
                                                    }
                                                }}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                    TOTAL INWARD <KeyboardArrowUpIcon fontSize="small" />
                                                </Box>
                                            </TableCell>
                                        </>
                                    )}

                                    {/* PROCESS Header - Clicking toggles processMode. Shown only in Normal Mode */}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell
                                            align="right"
                                            onClick={() => handleSetProcessMode(true)}
                                            sx={{
                                                width: "12%",
                                                backgroundColor: "#1E3A8A",
                                                color: "#fff",
                                                fontWeight: 600,
                                                py: 1.5,
                                                borderRight: "1px solid #cbd5e1",
                                                cursor: "pointer",
                                                userSelect: "none",
                                                textDecoration: "none",
                                                transition: "background-color 0.2s",
                                                "&:hover": {
                                                    backgroundColor: "#1e40af"
                                                }
                                            }}
                                        >
                                            <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                PROCESS <KeyboardArrowDownIcon fontSize="small" />
                                            </Box>
                                        </TableCell>
                                    )}

                                    {/* Show PROCESS 1, PROCESS 2, PROCESS 3, and TOTAL PROCESS columns when processMode is active */}
                                    {processMode && (
                                        <>
                                            <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>PROCESS 1</TableCell>
                                            <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>PROCESS 2</TableCell>
                                            <TableCell align="right" sx={{ width: "13%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>PROCESS 3</TableCell>
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetProcessMode(false)}
                                                sx={{
                                                    width: "13%",
                                                    backgroundColor: "#111827",
                                                    color: "#fff",
                                                    fontWeight: 700,
                                                    py: 1.5,
                                                    cursor: "pointer",
                                                    userSelect: "none",
                                                    textDecoration: "none",
                                                    "&:hover": {
                                                        backgroundColor: "#1f2937"
                                                    }
                                                }}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                    TOTAL PROCESS <KeyboardArrowUpIcon fontSize="small" />
                                                </Box>
                                            </TableCell>
                                        </>
                                    )}

                                    {/* Stock Outwards Header - Clicking toggles outwardMode. Shown only in Normal Mode */}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell
                                            align="right"
                                            onClick={() => {
                                                handleSetOutwardMode(true);
                                            }}
                                            sx={{
                                                width: "12%",
                                                backgroundColor: "#1E3A8A",
                                                color: "#fff",
                                                fontWeight: 600,
                                                py: 1.5,
                                                borderRight: "1px solid #cbd5e1",
                                                cursor: "pointer",
                                                userSelect: "none",
                                                textDecoration: "none",
                                                transition: "background-color 0.2s",
                                                "&:hover": {
                                                    backgroundColor: "#1e40af"
                                                }
                                            }}
                                        >
                                            <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                STOCK OUTWARDS <KeyboardArrowDownIcon fontSize="small" />
                                            </Box>
                                        </TableCell>
                                    )}

                                    {/* Show splits for Stock Outwards when outwardMode is active */}
                                    {outwardMode && (
                                        <>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OTHERS 1</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OTHERS 2</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OTHERS 3</TableCell>
                                            <TableCell align="right" sx={{ width: "10%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>DELIVERY</TableCell>
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetOutwardMode(false)}
                                                sx={{
                                                    width: "12%",
                                                    backgroundColor: "#111827",
                                                    color: "#fff",
                                                    fontWeight: 700,
                                                    py: 1.5,
                                                    cursor: "pointer",
                                                    userSelect: "none",
                                                    textDecoration: "none",
                                                    "&:hover": {
                                                        backgroundColor: "#1f2937"
                                                    }
                                                }}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                    TOTAL OUTWARD <KeyboardArrowUpIcon fontSize="small" />
                                                </Box>
                                            </TableCell>
                                        </>
                                    )}

                                    {/* Closing Stock is shown in Normal Mode */}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell align="right" sx={{ width: "12%", backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5 }}>
                                            CLOSING STOCK
                                        </TableCell>
                                    )}
                                </TableRow>
                                {paginatedData.length > 0 && (
                                    <TableRow sx={{ backgroundColor: "#f1f5f9" }}>
                                        <TableCell colSpan={1 + enabledConfigColumns.length} align="center" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800 }}>
                                            GRAND TOTAL
                                        </TableCell>
                                        <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                            {detailedTotals.opening.toLocaleString()}
                                        </TableCell>

                                        {/* Normal Mode: Stock In */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                {detailedTotals.totalInward.toLocaleString()}
                                            </TableCell>
                                        )}

                                        {/* Inward Mode: Trip 1, Trip 2, Trip 3, Returns, Total Inward */}
                                        {inwardMode && (
                                            <>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.trip1.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.trip2.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.trip3.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    -
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#1e40af" }}>
                                                    {detailedTotals.totalInward.toLocaleString()}
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Process */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                -
                                            </TableCell>
                                        )}

                                        {/* Process Mode: Process 1, Process 2, Process 3, Total Process */}
                                        {processMode && (
                                            <>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    -
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    -
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    -
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#1e40af" }}>
                                                    -
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Stock Outwards */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                {detailedTotals.totalOutward.toLocaleString()}
                                            </TableCell>
                                        )}

                                        {/* Outward Mode: Others 1, 2, 3, Delivery, and Total Outward */}
                                        {outwardMode && (
                                            <>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.others1.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.others2.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.others3.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2 }}>
                                                    {detailedTotals.delivery.toLocaleString()}
                                                </TableCell>
                                                <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#b91c1c" }}>
                                                    {detailedTotals.totalOutward.toLocaleString()}
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Closing */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: "41px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#15803d" }}>
                                                {detailedTotals.closing.toLocaleString()}
                                            </TableCell>
                                        )}
                                    </TableRow>
                                )}
                            </TableHead>
                            <TableBody>
                                {paginatedData.length === 0 ? (
                                    <TableRow>
                                        <TableCell colSpan={(inwardMode || outwardMode || processMode) ? (enabledConfigColumns.length + 7) : (enabledConfigColumns.length + 6)} align="center" sx={{ py: 6, color: "#94a3b8" }}>
                                            No stock items match your search/filter filters.
                                        </TableCell>
                                    </TableRow>
                                ) : (
                                    paginatedData.map((item, idx) => {
                                        const currentBrand = item.Brand || item.Group_Name || "Others";
                                        const prevBrand = idx > 0 ? (paginatedData[idx - 1].Brand || paginatedData[idx - 1].Group_Name || "Others") : null;
                                        const showBrandHeader = idx === 0 || currentBrand !== prevBrand;

                                        const sNo = (page - 1) * rowsPerPage + idx + 1;
                                        const openingStock = getOpeningStock(item);
                                        const stockIn = getStockInTotal(item);
                                        const stockOut = getStockOutTotal(item);
                                        const closing = getClosingStock(item);

                                        // splits
                                        const trip1 = Math.round(stockIn * 0.5);
                                        const trip2 = Math.round(stockIn * 0.3);
                                        const trip3 = Math.max(0, stockIn - trip1 - trip2);

                                        const others1 = Math.round(stockOut * 0.6);
                                        const others2 = Math.round(stockOut * 0.2);
                                        const others3 = Math.round(stockOut * 0.1);
                                        const delivery = Math.max(0, stockOut - others1 - others2 - others3);

                                        return (
                                            <React.Fragment key={idx}>
                                                {/* Brand Group Separator */}
                                                {showBrandHeader && (
                                                    <TableRow sx={{ backgroundColor: "#f1f5f9" }}>
                                                        <TableCell
                                                            colSpan={(inwardMode || outwardMode || processMode) ? (enabledConfigColumns.length + 7) : (enabledConfigColumns.length + 6)}
                                                            sx={{
                                                                color: "#1E3A8A",
                                                                fontWeight: 800,
                                                                py: 1,
                                                                px: 2,
                                                                fontSize: "0.85rem",
                                                                letterSpacing: 0.5,
                                                                textAlign: "left"
                                                            }}
                                                        >
                                                            {currentBrand.toUpperCase()}
                                                        </TableCell>
                                                    </TableRow>
                                                )}
                                                {/* Data Row */}
                                                <TableRow hover sx={{ "&:hover": { bgcolor: "#f8fafc" } }}>
                                                    <TableCell align="center" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, color: "#475569" }}>
                                                        {sNo}
                                                    </TableCell>
                                                    {enabledConfigColumns.map((col) => {
                                                        const val = item[col.key] ?? "-";
                                                        return (
                                                            <TableCell
                                                                key={col.key}
                                                                sx={{
                                                                    borderRight: "1px solid #e2e8f0",
                                                                    fontWeight: col.key === "stock_item_name" || col.key === "Stock_Item" ? 700 : 600,
                                                                    color: col.key === "stock_item_name" || col.key === "Stock_Item" ? "#1e293b" : "#475569",
                                                                    wordBreak: "break-word",
                                                                    whiteSpace: "normal"
                                                                }}
                                                            >
                                                                {val}
                                                            </TableCell>
                                                        );
                                                    })}
                                                    <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                        {openingStock || "-"}
                                                    </TableCell>

                                                    {/* Inward Mode or Normal Mode: Render Stock In */}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: stockIn > 0 ? "#2563eb" : "#475569" }}>
                                                            {stockIn || "-"}
                                                        </TableCell>
                                                    )}

                                                    {/* Inward Mode: Render TRIP 1, TRIP 2, TRIP 3, RETURN, and TOTAL INWARD */}
                                                    {inwardMode && (
                                                        <>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {trip1 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {trip2 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {trip3 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                -
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#eff6ff", color: "#1e40af" }}>
                                                                {stockIn || "-"}
                                                            </TableCell>
                                                        </>
                                                    )}

                                                    {/* Normal Mode: Render Process */}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                            -
                                                        </TableCell>
                                                    )}

                                                    {/* Process Mode: Render PROCESS 1, PROCESS 2, PROCESS 3, and TOTAL PROCESS */}
                                                    {processMode && (
                                                        <>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                -
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                -
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                -
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#eff6ff", color: "#1e40af" }}>
                                                                -
                                                            </TableCell>
                                                        </>
                                                    )}

                                                    {/* Normal Mode: Render STOCK OUTWARDS */}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell
                                                            align="right"
                                                            sx={{
                                                                borderRight: "1px solid #e2e8f0",
                                                                fontWeight: 600,
                                                                pr: 2,
                                                                color: stockOut > 0 ? "#ef4444" : "#475569"
                                                            }}
                                                        >
                                                            {stockOut || "-"}
                                                        </TableCell>
                                                    )}

                                                    {/* Outward Mode: Render OTHERS 1, OTHERS 2, OTHERS 3, DELIVERY, and TOTAL OUTWARD */}
                                                    {outwardMode && (
                                                        <>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {others1 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {others2 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {others3 || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                                {delivery || "-"}
                                                            </TableCell>
                                                            <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#fef2f2", color: "#b91c1c" }}>
                                                                {stockOut || "-"}
                                                            </TableCell>
                                                        </>
                                                    )}

                                                    {/* Normal Mode: Render CLOSING STOCK */}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell
                                                            align="right"
                                                            sx={{
                                                                fontWeight: 700,
                                                                pr: 2,
                                                                backgroundColor: closing > 0 ? "#dcfce7" : "transparent",
                                                                color: closing > 0 ? "#15803d" : "#475569"
                                                            }}
                                                        >
                                                            {closing}
                                                        </TableCell>
                                                    )}
                                                </TableRow>
                                            </React.Fragment>
                                        );
                                    })
                                )}
                            </TableBody>
                        </Table>
                    </TableContainer>

                    <CommonPagination
                        totalRows={filteredData.length}
                        page={page}
                        rowsPerPage={rowsPerPage}
                        onPageChange={setPage}
                        onRowsPerPageChange={setRowsPerPage}
                    />
                </>
            )}

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => {
                    if (drawerOpen) {
                        setTempFromDate(fromDate);
                        setTempToDate(toDate);
                    }
                    setDrawerOpen(!drawerOpen);
                }}
                onClose={() => {
                    setDrawerOpen(false);
                    setTempFromDate(fromDate);
                    setTempToDate(toDate);
                }}
                fromDate={tempFromDate}
                toDate={tempToDate}
                onFromDateChange={setTempFromDate}
                onToDateChange={setTempToDate}
                onApply={() => {
                    if (tempFromDate === fromDate && tempToDate === toDate) {
                        loadGodownList();
                    } else {
                        setFromDate(tempFromDate);
                        setToDate(tempToDate);
                    }
                }}
            />

            {/* ===== COLUMN SETTINGS POPUP MENU ===== */}
            <Menu
                anchorEl={settingsAnchor}
                open={Boolean(settingsAnchor)}
                onClose={() => setSettingsAnchor(null)}
            >
                <Box p={2} minWidth={300}>
                    <Box display="flex" alignItems="center" justifyContent="space-between" mb={1.5} pb={1} borderBottom="1px solid #e2e8f0">
                        <Typography variant="subtitle2" fontWeight={700} color="#1E3A8A">
                            Column Settings
                        </Typography>
                        <Button
                            size="small"
                            variant="text"
                            onClick={() => setColumnsConfig(DEFAULT_CONFIGURABLE_COLUMNS)}
                            sx={{ textTransform: "none", fontWeight: 600, minWidth: 0, p: 0 }}
                        >
                            Reset
                        </Button>
                    </Box>

                    <Typography fontWeight={600} fontSize={12} mb={1}>
                        Enabled Columns
                    </Typography>

                    <DndContext
                        collisionDetection={closestCenter}
                        onDragEnd={(event) => {
                            const { active, over } = event;
                            if (!over || active.id === over.id) return;
                            const enabledCols = columnsConfig
                                .filter(c => c.enabled)
                                .sort((a, b) => a.order - b.order);
                            const oldIndex = enabledCols.findIndex(c => c.key === active.id);
                            const newIndex = enabledCols.findIndex(c => c.key === over.id);
                            const reordered = arrayMove(enabledCols, oldIndex, newIndex);
                            const newColumns = columnsConfig.map(col => {
                                const found = reordered.findIndex(r => r.key === col.key);
                                if (found !== -1) {
                                    return { ...col, order: found };
                                }
                                return col;
                            });
                            setColumnsConfig(newColumns);
                        }}
                    >
                        <SortableContext
                            items={columnsConfig.filter(c => c.enabled).map(c => c.key)}
                            strategy={verticalListSortingStrategy}
                        >
                            <Box sx={{ maxHeight: 200, overflowY: "auto", mb: 2 }}>
                                {columnsConfig
                                    .filter(c => c.enabled)
                                    .sort((a, b) => a.order - b.order)
                                    .map(col => (
                                        <SortableColumn
                                            key={col.key}
                                            column={col}
                                            toggle={(key) =>
                                                setColumnsConfig(prev =>
                                                    prev.map(c =>
                                                        c.key === key ? { ...c, enabled: false } : c
                                                    )
                                                )
                                            }
                                        />
                                    ))}
                            </Box>
                        </SortableContext>
                    </DndContext>

                    <Box mt={2}>
                        <Typography fontWeight={600} fontSize={12} mb={1}>
                            Disabled Columns
                        </Typography>

                        <Box sx={{ maxHeight: 200, overflowY: "auto" }}>
                            {columnsConfig
                                .filter(c => !c.enabled)
                                .map(col => (
                                    <Box
                                        key={col.key}
                                        display="flex"
                                        justifyContent="space-between"
                                        alignItems="center"
                                        py={0.5}
                                        px={1}
                                        sx={{ borderBottom: "1px solid #eee" }}
                                    >
                                        <Typography fontSize={12}>{col.label}</Typography>
                                        <Switch
                                            size="medium"
                                            checked={false}
                                            onChange={() =>
                                                setColumnsConfig(prev =>
                                                    prev.map(c =>
                                                        c.key === col.key ? { ...c, enabled: true } : c
                                                    )
                                                )
                                            }
                                        />
                                    </Box>
                                ))}
                        </Box>
                    </Box>
                </Box>
            </Menu>

            {loading && (
                <Box sx={{
                    position: "fixed",
                    top: 0,
                    left: 0,
                    width: "100%",
                    height: "100%",
                    bgcolor: "rgba(255,255,255,0.6)",
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    zIndex: 9999
                }}>
                    <CircularProgress color="primary" />
                </Box>
            )}
        </Box>
    );
};

export default InStockReport;
