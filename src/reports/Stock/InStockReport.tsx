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
import CloseIcon from "@mui/icons-material/Close";
import AddIcon from "@mui/icons-material/Add";
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
    stockWiseReport,
    stockInOutProcessService
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

    const isPivotMode = useMemo(() => {
        return !enabledConfigColumns.some(col => col.key === "stock_item_name" || col.key === "Stock_Item");
    }, [enabledConfigColumns]);

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
    const [qtyMode, setQtyMode] = useState<"qty" | "actQty">(() => {
        const saved = sessionStorage.getItem("inStockQtyMode");
        return (saved as "qty" | "actQty") || "qty";
    });

    useEffect(() => {
        sessionStorage.setItem("inStockQtyMode", qtyMode);
    }, [qtyMode]);

    const formatQtyVal = (val: any): string => {
        if (val === null || val === undefined || val === "" || Number(val) === 0) return "-";
        const num = Number(val);
        if (isNaN(num)) return String(val);
        return Number(num.toFixed(2)).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    };

    const qtyKeys = useMemo(() => {
        if (qtyMode === "qty") {
            return {
                opening: "OB_Bal_Qty" as keyof stockWiseReport,
                in: "Pur_Qty" as keyof stockWiseReport,
                procIn: "Proc_IN_Qty" as keyof stockWiseReport,
                procOut: "Proc_OUT_Qty" as keyof stockWiseReport,
                out: "Sal_Qty" as keyof stockWiseReport,
                closing: "Bal_Qty" as keyof stockWiseReport
            };
        } else {
            return {
                opening: "OB_Act_Qty" as keyof stockWiseReport,
                in: "Pur_Act_Qty" as keyof stockWiseReport,
                procIn: "Proc_IN_Act_Qty" as keyof stockWiseReport,
                procOut: "Proc_OUT_Act_Qty" as keyof stockWiseReport,
                out: "Sal_Act_Qty" as keyof stockWiseReport,
                closing: "Bal_Act_Qty" as keyof stockWiseReport
            };
        }
    }, [qtyMode]);

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

    // Hidden dynamic columns
    const [hiddenInwardColumns, setHiddenInwardColumns] = useState<string[]>([]);
    const [hiddenOutwardColumns, setHiddenOutwardColumns] = useState<string[]>([]);
    const [hiddenProcessColumns, setHiddenProcessColumns] = useState<string[]>([]);

    const handleHideInwardColumn = (col: string) => {
        setHiddenInwardColumns(prev => [...prev, col]);
    };
    const handleShowInwardColumn = (col: string) => {
        setHiddenInwardColumns(prev => prev.filter(c => c !== col));
    };

    const handleHideOutwardColumn = (col: string) => {
        setHiddenOutwardColumns(prev => [...prev, col]);
    };
    const handleShowOutwardColumn = (col: string) => {
        setHiddenOutwardColumns(prev => prev.filter(c => c !== col));
    };

    const handleHideProcessColumn = (col: string) => {
        setHiddenProcessColumns(prev => [...prev, col]);
    };
    const handleShowProcessColumn = (col: string) => {
        setHiddenProcessColumns(prev => prev.filter(c => c !== col));
    };

    // Dynamic sticky header height calculation
    const headerRowRef = React.useRef<HTMLTableRowElement>(null);
    const [headerHeight, setHeaderHeight] = useState(41);

    // API Data state
    const [godownListData, setGodownListData] = useState<StockAbstractData4[]>([]);
    const [detailedStockData, setDetailedStockData] = useState<stockWiseReport[]>([]);
    const [processApiData, setProcessApiData] = useState<any[]>([]);

    const handleSetInwardMode = (val: boolean) => {
        setInwardMode(val);
        if (val) {
            setOutwardMode(false);
            setProcessMode(false);
        } else {
            setHiddenInwardColumns([]);
        }
    };

    const handleSetOutwardMode = (val: boolean) => {
        setOutwardMode(val);
        if (val) {
            setInwardMode(false);
            setProcessMode(false);
        } else {
            setHiddenOutwardColumns([]);
        }
    };

    const handleSetProcessMode = (val: boolean) => {
        setProcessMode(val);
        if (val) {
            setInwardMode(false);
            setOutwardMode(false);
        } else {
            setHiddenProcessColumns([]);
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

            try {
                const processRes = await stockInOutProcessService.getStockInOutProcess({
                    Todate: dayjs(toDate).format("YYYY-MM-DD"),
                    Fromdate: dayjs(fromDate).format("YYYY-MM-DD"),
                });
                setProcessApiData(processRes.data?.data || []);
            } catch (err) {
                console.error("Failed to load stock in out process details:", err);
                setProcessApiData([]);
            }

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
            setProcessApiData([]);
        }
    }, [selectedGodown, fromDate, toDate]);

    // Group godown list by parent_godown_name (only including godowns with data)
    const groupedGodowns = useMemo(() => {
        const groups: Record<string, StockAbstractData4[]> = {};
        const filteredList = (godownListData || []).filter(g => {
            if (qtyMode === "actQty") {
                return Number(g.ACt_OB_Qty || 0) !== 0 ||
                    Number(g.ACt_In_Qty || 0) !== 0 ||
                    Number(g.ACt_Out_Qty || 0) !== 0 ||
                    Number(g.CL_ACt_QTY || 0) !== 0;
            }
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
    }, [godownListData, qtyMode]);

    // Calculate aggregated overall summary of godowns totals
    const grandTotals = useMemo(() => {
        let opening = 0;
        let stockIn = 0;
        let process = 0;
        let stockOut = 0;
        let closing = 0;

        godownListData.forEach(g => {
            if (qtyMode === "actQty") {
                opening += Number(g.ACt_OB_Qty || 0);
                stockIn += Number(g.ACt_In_Qty || 0);
                process += Number(g.Process_Act_IN_OUT_Qty || 0);
                stockOut += Number(g.ACt_Out_Qty || 0);
                closing += Number(g.CL_ACt_QTY || 0);
            } else {
                opening += Number(g.OB_Qty || 0);
                stockIn += Number(g.IN_Qty || 0);
                process += Number(g.Process_IN_OUT_Qty || 0);
                stockOut += Number(g.Out_Qty || 0);
                closing += Number(g.CL_QTY || 0);
            }
        });

        return { opening, stockIn, process, stockOut, closing };
    }, [godownListData, qtyMode]);

    // Map stockinoutprocess API data by item_id and godown_name
    const mappedProcessData = useMemo(() => {
        const map: Record<string, any[]> = {};
        processApiData.forEach((record) => {
            const key = `${record.item_id}_${String(record.godown_name || "").toLowerCase().trim()}`;
            if (!map[key]) {
                map[key] = [];
            }
            map[key].push(record);
        });
        return map;
    }, [processApiData]);

    // Helper to get mapped details for a product
    const getProductDetails = useMemo(() => {
        return (item: stockWiseReport) => {
            const godownName = String(item.Godown_Name || selectedGodown?.godown_name || "").toLowerCase().trim();
            
            let records: any[] = [];
            const productIds = (item as any).Product_Ids;
            if (productIds && Array.isArray(productIds)) {
                productIds.forEach((pId) => {
                    const key = `${pId}_${godownName}`;
                    const recs = mappedProcessData[key] || [];
                    records = [...records, ...recs];
                });
            } else {
                const key = `${item.Product_Id}_${godownName}`;
                records = mappedProcessData[key] || [];
            }

            const isTrip = (r: any) => {
                return (
                    r.trip_voucher_number !== null ||
                    r.trip_id !== null ||
                    (r.Trip_No !== null && r.Trip_No !== undefined && String(r.Trip_No).trim() !== "") ||
                    (r.trip_no !== null && r.trip_no !== undefined && String(r.trip_no).trim() !== "")
                );
            };

            const trips = records.filter(
                (r) =>
                    r.stock_direction?.toUpperCase() === "IN" &&
                    isTrip(r)
            );

            const returns = records.filter(
                (r) =>
                    r.stock_direction?.toUpperCase() === "IN" &&
                    !isTrip(r)
            );
            const returnQty = returns.reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            const stockInQty = records
                .filter((r) => r.stock_direction?.toUpperCase() === "IN")
                .reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            const procInQty = records
                .filter((r) => r.stock_direction?.toUpperCase() === "PROCESS IN")
                .reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            const procOutQty = records
                .filter((r) => r.stock_direction?.toUpperCase() === "PROCESS OUT")
                .reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            const outwardQty = records
                .filter((r) => r.stock_direction?.toUpperCase() === "OUT")
                .reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            const outTrips = records.filter(
                (r) =>
                    r.stock_direction?.toUpperCase() === "OUT" &&
                    isTrip(r)
            );
            const deliveries = records.filter(
                (r) =>
                    r.stock_direction?.toUpperCase() === "OUT" &&
                    !isTrip(r)
            );
            const deliveryQty = deliveries.reduce((sum, r) => sum + Number(r.quantity || 0), 0);

            return {
                trips,
                returnQty,
                stockInQty,
                procInQty,
                procOutQty,
                outwardQty,
                outTrips,
                deliveryQty
            };
        };
    }, [mappedProcessData, selectedGodown]);



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
    const filteredDetailedData = useMemo(() => {
        return (detailedStockData || []).filter((item) => {
            const productName = item.stock_item_name || item.Stock_Item || "";
            const brandName = item.Brand || item.Group_Name || "";
            const matchesSearch = productName.toLowerCase().includes(searchText.toLowerCase()) ||
                brandName.toLowerCase().includes(searchText.toLowerCase());
            const matchesBrand = selectedBrand === "All" || brandName === selectedBrand;

            if (!matchesSearch || !matchesBrand) return false;

            if (inwardMode) {
                const { stockInQty } = getProductDetails(item);
                const totalStockIn = processApiData.length > 0 ? stockInQty : Number(item[qtyKeys.in] || 0);
                if (totalStockIn <= 0) return false;
            }
            if (processMode) {
                const { procInQty, procOutQty } = getProductDetails(item);
                const totalProcIn = processApiData.length > 0 ? procInQty : Number(item[qtyKeys.procIn] || 0);
                const totalProcOut = processApiData.length > 0 ? procOutQty : Number(item[qtyKeys.procOut] || 0);
                if (totalProcIn <= 0 && totalProcOut <= 0) return false;
            }
            if (outwardMode) {
                const { outwardQty } = getProductDetails(item);
                const totalStockOut = processApiData.length > 0 ? outwardQty : Number(item[qtyKeys.out] || 0);
                if (totalStockOut <= 0) return false;
            }

            return true;
        });
    }, [detailedStockData, searchText, selectedBrand, inwardMode, processMode, outwardMode, getProductDetails, qtyKeys, processApiData]);

    const filteredData = useMemo(() => {
        if (!isPivotMode) return filteredDetailedData;

        const groups: Record<string, stockWiseReport> = {};

        filteredDetailedData.forEach(item => {
            const groupKeyParts: string[] = [];
            
            // Always group by Brand
            const brandVal = item.Brand || item.Group_Name || "Others";
            groupKeyParts.push(brandVal);
            
            // Add other enabled columns that are NOT product-specific identifiers
            const excludedKeys = ["stock_item_name", "Stock_Item", "Product_Id", "POS_Item_Name", "Item_Name_Modified"];
            enabledConfigColumns.forEach(col => {
                if (!excludedKeys.includes(col.key) && col.key !== "Brand" && col.key !== "Group_Name") {
                    groupKeyParts.push(String(item[col.key] ?? ""));
                }
            });

            const key = groupKeyParts.join(" | ");

            if (!groups[key]) {
                const groupedItem: any = {
                    ...item,
                    Product_Ids: [],
                };

                // Initialize numeric fields to 0
                const numericKeys = [
                    "OB_Bal_Qty", "Pur_Qty", "Sal_Qty", "Bal_Qty",
                    "OB_Act_Qty", "Pur_Act_Qty", "Sal_Act_Qty", "Bal_Act_Qty",
                    "OB_Qty", "IN_Qty", "Out_Qty", "CL_QTY",
                    "ACt_OB_Qty", "ACt_In_Qty", "Process_Act_IN_OUT_Qty", "ACt_Out_Qty", "CL_ACt_QTY",
                    "Process_IN_OUT_Qty"
                ];
                numericKeys.forEach(k => {
                    groupedItem[k] = 0;
                });

                groups[key] = groupedItem;
            }

            const g = groups[key];
            if (item.Product_Id) {
                (g as any).Product_Ids.push(item.Product_Id);
            }

            // Sum numeric fields
            const numericKeys = [
                "OB_Bal_Qty", "Pur_Qty", "Sal_Qty", "Bal_Qty",
                "OB_Act_Qty", "Pur_Act_Qty", "Sal_Act_Qty", "Bal_Act_Qty",
                "OB_Qty", "IN_Qty", "Out_Qty", "CL_QTY",
                "ACt_OB_Qty", "ACt_In_Qty", "Process_Act_IN_OUT_Qty", "ACt_Out_Qty", "CL_ACt_QTY",
                "Process_IN_OUT_Qty"
            ];
            numericKeys.forEach(k => {
                if (item[k] !== undefined && item[k] !== null) {
                    g[k] = (Number(g[k]) || 0) + (Number(item[k]) || 0);
                }
            });
        });

        return Object.values(groups);
    }, [filteredDetailedData, enabledConfigColumns, isPivotMode]);

    const getTripLabel = React.useCallback((t: any): string => {
        const rawVal = t.Trip_No || t.trip_no || t.trip_voucher_number || t.trip_id;
        if (!rawVal) return "N/A";
        const str = String(rawVal).trim();
        if (str.toLowerCase().startsWith("trip")) {
            return str;
        }
        return `Trip - ${str}`;
    }, []);

    const getQtyForTrip = React.useCallback((tripsList: any[], label: string): number => {
        return tripsList
            .filter(t => getTripLabel(t) === label)
            .reduce((sum, t) => sum + Number(t.quantity || 0), 0);
    }, [getTripLabel]);

    // Unique inward trip headers in filteredData
    const inwardTripHeaders = useMemo(() => {
        if (!inwardMode) return [];
        const set = new Set<string>();
        filteredData.forEach((item) => {
            const { trips } = getProductDetails(item);
            trips.forEach((t) => {
                set.add(getTripLabel(t));
            });
        });
        return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    }, [filteredData, inwardMode, getProductDetails, getTripLabel]);

    // Unique outward trip headers in filteredData
    const outwardTripHeaders = useMemo(() => {
        if (!outwardMode) return [];
        const set = new Set<string>();
        filteredData.forEach((item) => {
            const { outTrips } = getProductDetails(item);
            outTrips.forEach((t) => {
                set.add(getTripLabel(t));
            });
        });
        return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    }, [filteredData, outwardMode, getProductDetails, getTripLabel]);

    // Calculate total quantity for each trip in inwardMode
    const inwardTripTotals = useMemo(() => {
        const totals: Record<string, number> = {};
        inwardTripHeaders.forEach(label => {
            let sum = 0;
            filteredData.forEach(item => {
                const { trips } = getProductDetails(item);
                sum += getQtyForTrip(trips, label);
            });
            totals[label] = sum;
        });
        return totals;
    }, [filteredData, inwardTripHeaders, getProductDetails, getQtyForTrip]);

    // Calculate total quantity for each trip in outwardMode
    const outwardTripTotals = useMemo(() => {
        const totals: Record<string, number> = {};
        outwardTripHeaders.forEach(label => {
            let sum = 0;
            filteredData.forEach(item => {
                const { outTrips } = getProductDetails(item);
                sum += getQtyForTrip(outTrips, label);
            });
            totals[label] = sum;
        });
        return totals;
    }, [filteredData, outwardTripHeaders, getProductDetails, getQtyForTrip]);

    useEffect(() => {
        const handleResize = () => {
            if (headerRowRef.current) {
                setHeaderHeight(headerRowRef.current.offsetHeight);
            }
        };
        window.addEventListener("resize", handleResize);
        
        handleResize();
        const t = setTimeout(handleResize, 100);

        return () => {
            window.removeEventListener("resize", handleResize);
            clearTimeout(t);
        };
    }, [filteredData, inwardMode, outwardMode, processMode, columnsConfig, searchText, page, rowsPerPage]);

    // Slice data for pagination
    const paginatedData = useMemo(() => {
        return filteredData.slice((page - 1) * rowsPerPage, page * rowsPerPage);
    }, [filteredData, page, rowsPerPage]);

    // Detailed quantities helpers
    const getOpeningStock = (item: stockWiseReport) => Number(item[qtyKeys.opening] || 0);
    const getStockInTotal = (item: stockWiseReport) => Number(item[qtyKeys.in] || 0);
    const getStockOutTotal = (item: stockWiseReport) => Number(item[qtyKeys.out] || 0);
    const getClosingStock = (item: stockWiseReport) => Number(item[qtyKeys.closing] || 0);
    const getProcIn = (item: stockWiseReport) => Number(item[qtyKeys.procIn] || 0);
    const getProcOut = (item: stockWiseReport) => Number(item[qtyKeys.procOut] || 0);

    // Calculate totals for the selected godown's filtered data
    const detailedTotals = useMemo(() => {
        let opening = 0;
        let stockIn = 0;
        let procIn = 0;
        let procOut = 0;
        let tripQtyTotal = 0;
        let returnQtyTotal = 0;
        let outTripQtyTotal = 0;
        let deliveryQtyTotal = 0;
        let stockOutTotal = 0;
        let closing = 0;

        filteredData.forEach(item => {
            const op = getOpeningStock(item);
            const inQty = getStockInTotal(item);
            const outQty = getStockOutTotal(item);
            const clQty = getClosingStock(item);

            opening += op;
            closing += clQty;

            const { trips, returnQty, stockInQty, procInQty, procOutQty, outwardQty, outTrips, deliveryQty } = getProductDetails(item);

            if (processApiData.length > 0) {
                const itemTripQty = trips.reduce((sum, r) => sum + Number(r.quantity || 0), 0);
                const itemReturnQty = returnQty;
                const itemStockIn = stockInQty;
                const itemProcIn = procInQty;
                const itemProcOut = procOutQty;
                const itemOutTripQty = outTrips.reduce((sum, r) => sum + Number(r.quantity || 0), 0);
                const itemDeliveryQty = deliveryQty;
                const itemStockOut = outwardQty;

                tripQtyTotal += itemTripQty;
                returnQtyTotal += itemReturnQty;
                outTripQtyTotal += itemOutTripQty;
                deliveryQtyTotal += itemDeliveryQty;
                procIn += itemProcIn;
                procOut += itemProcOut;
                stockIn += itemStockIn;
                stockOutTotal += itemStockOut;
            } else {
                const pIn = getProcIn(item);
                const pOut = getProcOut(item);
                procIn += pIn;
                procOut += pOut;
                stockIn += inQty;
                stockOutTotal += outQty;

                const t1 = Math.round(inQty * 0.5);
                const t2 = Math.round(inQty * 0.3);
                const t3 = Math.max(0, inQty - t1 - t2);
                tripQtyTotal += t1 + t2 + t3;

                const o1 = Math.round(outQty * 0.6);
                const o2 = Math.round(outQty * 0.2);
                const o3 = Math.round(outQty * 0.1);
                const del = Math.max(0, outQty - o1 - o2 - o3);
                outTripQtyTotal += o1 + o2 + o3;
                deliveryQtyTotal += del;
            }
        });

        return {
            opening,
            stockIn,
            procIn,
            procOut,
            tripQtyTotal,
            returnQtyTotal,
            outTripQtyTotal,
            deliveryQtyTotal,
            stockOutTotal,
            closing,
            totalInward: stockIn,
            totalOutward: stockOutTotal,
            totalProcess: procIn - procOut
        };
    }, [filteredData, qtyKeys, getProductDetails, processApiData]);

    // Recalculated row-level Stock In Quantity summing only visible columns
    const getRecalculatedStockInQty = React.useCallback((item: stockWiseReport) => {
        const { trips, returnQty } = getProductDetails(item);
        let sum = 0;
        inwardTripHeaders.forEach((tripLabel) => {
            if (!hiddenInwardColumns.includes(tripLabel)) {
                sum += getQtyForTrip(trips, tripLabel);
            }
        });
        if (!hiddenInwardColumns.includes("RETURN")) {
            sum += returnQty;
        }
        return sum;
    }, [getProductDetails, inwardTripHeaders, hiddenInwardColumns, getQtyForTrip]);

    // Recalculated row-level Stock Out Quantity summing only visible columns
    const getRecalculatedStockOutQty = React.useCallback((item: stockWiseReport) => {
        const { outTrips, deliveryQty } = getProductDetails(item);
        let sum = 0;
        outwardTripHeaders.forEach((tripLabel) => {
            if (!hiddenOutwardColumns.includes(tripLabel)) {
                sum += getQtyForTrip(outTrips, tripLabel);
            }
        });
        if (!hiddenOutwardColumns.includes("PENDING DELIVERY")) {
            sum += deliveryQty;
        }
        return sum;
    }, [getProductDetails, outwardTripHeaders, hiddenOutwardColumns, getQtyForTrip]);

    // Recalculated row-level Process Quantity summing only visible columns
    const getRecalculatedProcessQty = React.useCallback((item: stockWiseReport) => {
        let pIn = 0;
        let pOut = 0;
        if (processApiData.length > 0) {
            const { procInQty, procOutQty } = getProductDetails(item);
            pIn = procInQty;
            pOut = procOutQty;
        } else {
            pIn = getProcIn(item);
            pOut = getProcOut(item);
        }
        
        let sum = 0;
        if (!hiddenProcessColumns.includes("PROCESS IN")) {
            sum += pIn;
        }
        if (!hiddenProcessColumns.includes("PROCESS OUT")) {
            sum -= pOut;
        }
        return sum;
    }, [getProductDetails, hiddenProcessColumns, processApiData, qtyKeys]);

    // Memoized recalculated grand totals
    const recalculatedTotals = useMemo(() => {
        // 1. Inward total
        let inwardTotal = 0;
        if (processApiData.length > 0) {
            inwardTripHeaders.forEach(tripLabel => {
                if (!hiddenInwardColumns.includes(tripLabel)) {
                    inwardTotal += inwardTripTotals[tripLabel] || 0;
                }
            });
            if (!hiddenInwardColumns.includes("RETURN")) {
                inwardTotal += detailedTotals.returnQtyTotal;
            }
        } else {
            inwardTotal = detailedTotals.stockIn;
        }

        // 2. Outward total
        let outwardTotal = 0;
        if (processApiData.length > 0) {
            outwardTripHeaders.forEach(tripLabel => {
                if (!hiddenOutwardColumns.includes(tripLabel)) {
                    outwardTotal += outwardTripTotals[tripLabel] || 0;
                }
            });
            if (!hiddenOutwardColumns.includes("PENDING DELIVERY")) {
                outwardTotal += detailedTotals.deliveryQtyTotal;
            }
        } else {
            outwardTotal = detailedTotals.stockOutTotal;
        }

        // 3. Process total
        let visibleProcIn = 0;
        let visibleProcOut = 0;
        visibleProcIn = !hiddenProcessColumns.includes("PROCESS IN") ? detailedTotals.procIn : 0;
        visibleProcOut = !hiddenProcessColumns.includes("PROCESS OUT") ? detailedTotals.procOut : 0;
        const processTotal = visibleProcIn - visibleProcOut;

        return {
            inwardTotal,
            outwardTotal,
            processTotal
        };
    }, [
        processApiData,
        inwardTripHeaders,
        inwardTripTotals,
        hiddenInwardColumns,
        detailedTotals,
        outwardTripHeaders,
        outwardTripTotals,
        hiddenOutwardColumns,
        hiddenProcessColumns
    ]);

    // Excel Export
    const handleExportExcel = () => {
        try {
            const excelData: any[][] = [];
            const fmt = (val: any) => {
                if (val === null || val === undefined || val === "" || val === "-") return "-";
                const num = Number(val);
                return isNaN(num) ? val : Number(num.toFixed(2));
            };
            if (selectedGodown) {
                excelData.push([`STOCK REPORT - ${selectedGodown.godown_name.toUpperCase()}`]);
                excelData.push([]);
                const configLabels = enabledConfigColumns.map(c => c.label);

                if (inwardMode) {
                    const visibleTrips = inwardTripHeaders.filter(t => !hiddenInwardColumns.includes(t));
                    const isReturnVisible = !hiddenInwardColumns.includes("RETURN");
                    
                    const inwardHeaderRow = ["S.No", ...configLabels, ...visibleTrips];
                    if (isReturnVisible) inwardHeaderRow.push("Return");
                    inwardHeaderRow.push("Total Stock In");
                    excelData.push(inwardHeaderRow);

                    filteredData.forEach((item, idx) => {
                        const { trips, returnQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        
                        visibleTrips.forEach(tripLabel => {
                            const qty = getQtyForTrip(trips, tripLabel);
                            row.push(qty === 0 ? "-" : fmt(qty));
                        });

                        if (isReturnVisible) {
                            row.push(fmt(returnQty));
                        }
                        row.push(fmt(getRecalculatedStockInQty(item)));
                        excelData.push(row);
                    });
                } else if (outwardMode) {
                    const visibleTrips = outwardTripHeaders.filter(t => !hiddenOutwardColumns.includes(t));
                    const isPendingVisible = !hiddenOutwardColumns.includes("PENDING DELIVERY");
                    
                    const outwardHeaderRow = ["S.No", ...configLabels, ...visibleTrips];
                    if (isPendingVisible) outwardHeaderRow.push("Pending Delivery");
                    outwardHeaderRow.push("Total Outward");
                    excelData.push(outwardHeaderRow);

                    filteredData.forEach((item, idx) => {
                        const { outTrips, deliveryQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        
                        visibleTrips.forEach(tripLabel => {
                            const qty = getQtyForTrip(outTrips, tripLabel);
                            row.push(qty === 0 ? "-" : fmt(qty));
                        });

                        if (isPendingVisible) {
                            row.push(fmt(deliveryQty));
                        }
                        row.push(fmt(getRecalculatedStockOutQty(item)));
                        excelData.push(row);
                    });
                } else if (processMode) {
                    const isProcessInVisible = !hiddenProcessColumns.includes("PROCESS IN");
                    const isProcessOutVisible = !hiddenProcessColumns.includes("PROCESS OUT");

                    const processHeaderRow = ["S.No", ...configLabels];
                    if (isProcessInVisible) processHeaderRow.push("Process In");
                    if (isProcessOutVisible) processHeaderRow.push("Process Out");
                    processHeaderRow.push("Total Process");
                    excelData.push(processHeaderRow);

                    filteredData.forEach((item, idx) => {
                        const { procInQty, procOutQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        
                        if (isProcessInVisible) row.push(fmt(procInQty));
                        if (isProcessOutVisible) row.push(fmt(procOutQty));
                        row.push(fmt(getRecalculatedProcessQty(item)));
                        excelData.push(row);
                    });
                } else {
                    excelData.push(["S.No", ...configLabels, "Opening Stock", "Stock In", "Process", "Stock Outwards", "Closing Stock"]);
                    filteredData.forEach((item, idx) => {
                        const { stockInQty, procInQty, procOutQty, outwardQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        row.push(
                            fmt(getOpeningStock(item)),
                            fmt((processApiData.length > 0 ? stockInQty : getStockInTotal(item))),
                            fmt((processApiData.length > 0 ? (procInQty - procOutQty) : (getProcIn(item) - getProcOut(item)))),
                            fmt((processApiData.length > 0 ? outwardQty : getStockOutTotal(item))),
                            fmt(getClosingStock(item))
                        );
                        excelData.push(row);
                    });
                }
            } else {
                excelData.push([`GODOWNS OVERALL SUMMARY`]);
                excelData.push([]);
                excelData.push(["S.No", "Godown Name", "OB", "Stock In", "Process", "Stock Out", "Closing"]);

                let sno = 1;
                Object.entries(groupedGodowns).forEach(([parentName, items]) => {
                    const groupOB = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_OB_Qty : r.OB_Qty) || 0), 0);
                    const groupIn = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_In_Qty : r.IN_Qty) || 0), 0);
                    const groupProcess = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.Process_Act_IN_OUT_Qty : r.Process_IN_OUT_Qty) || 0), 0);
                    const groupOut = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_Out_Qty : r.Out_Qty) || 0), 0);
                    const groupCL = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.CL_ACt_QTY : r.CL_QTY) || 0), 0);

                    // Add group row
                    excelData.push(["", parentName, fmt(groupOB), fmt(groupIn), fmt(groupProcess), fmt(groupOut), fmt(groupCL)]);

                    items.forEach((row) => {
                        excelData.push([
                            sno++,
                            row.godown_name,
                            fmt(Number((qtyMode === "actQty" ? row.ACt_OB_Qty : row.OB_Qty) || 0)),
                            fmt(Number((qtyMode === "actQty" ? row.ACt_In_Qty : row.IN_Qty) || 0)),
                            fmt(Number((qtyMode === "actQty" ? row.Process_Act_IN_OUT_Qty : row.Process_IN_OUT_Qty) || 0)),
                            fmt(Number((qtyMode === "actQty" ? row.ACt_Out_Qty : row.Out_Qty) || 0)),
                            fmt(Number((qtyMode === "actQty" ? row.CL_ACt_QTY : row.CL_QTY) || 0))
                        ]);
                    });
                });

                // Add grand total row
                excelData.push([
                    "Total",
                    "",
                    fmt(grandTotals.opening),
                    fmt(grandTotals.stockIn),
                    fmt(grandTotals.process),
                    fmt(grandTotals.stockOut),
                    fmt(grandTotals.closing)
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

            const fmtStr = (val: any) => {
                if (val === null || val === undefined || val === "" || val === "-") return "-";
                const num = Number(val);
                if (isNaN(num)) return String(val);
                if (num === 0) return "-";
                return Number(num.toFixed(2)).toLocaleString();
            };

            if (selectedGodown) {
                const configLabels = enabledConfigColumns.map(c => c.label);

                if (inwardMode) {
                    const visibleTrips = inwardTripHeaders.filter(t => !hiddenInwardColumns.includes(t));
                    const isReturnVisible = !hiddenInwardColumns.includes("RETURN");
                    
                    const inwardHeaderRow = ["S.No", ...configLabels, ...visibleTrips];
                    if (isReturnVisible) inwardHeaderRow.push("Return");
                    inwardHeaderRow.push("Total Stock In");
                    headers = [inwardHeaderRow];

                    filteredData.forEach((item, idx) => {
                        const { trips, returnQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));

                        visibleTrips.forEach(tripLabel => {
                            const qty = getQtyForTrip(trips, tripLabel);
                            row.push(fmtStr(qty));
                        });

                        if (isReturnVisible) {
                            row.push(fmtStr(returnQty));
                        }
                        row.push(fmtStr(getRecalculatedStockInQty(item)));
                        body.push(row);
                    });
                } else if (outwardMode) {
                    const visibleTrips = outwardTripHeaders.filter(t => !hiddenOutwardColumns.includes(t));
                    const isPendingVisible = !hiddenOutwardColumns.includes("PENDING DELIVERY");

                    const outwardHeaderRow = ["S.No", ...configLabels, ...visibleTrips];
                    if (isPendingVisible) outwardHeaderRow.push("Pending Delivery");
                    outwardHeaderRow.push("Total Outward");
                    headers = [outwardHeaderRow];

                    filteredData.forEach((item, idx) => {
                        const { outTrips, deliveryQty } = getProductDetails(item);
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));

                        visibleTrips.forEach(tripLabel => {
                            const qty = getQtyForTrip(outTrips, tripLabel);
                            row.push(fmtStr(qty));
                        });

                        if (isPendingVisible) {
                            row.push(fmtStr(deliveryQty));
                        }
                        row.push(fmtStr(getRecalculatedStockOutQty(item)));
                        body.push(row);
                    });
                } else if (processMode) {
                    const isProcessInVisible = !hiddenProcessColumns.includes("PROCESS IN");
                    const isProcessOutVisible = !hiddenProcessColumns.includes("PROCESS OUT");

                    const processHeaderRow = ["S.No", ...configLabels];
                    if (isProcessInVisible) processHeaderRow.push("Process In");
                    if (isProcessOutVisible) processHeaderRow.push("Process Out");
                    processHeaderRow.push("Total Process");
                    headers = [processHeaderRow];

                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        const { procInQty, procOutQty } = getProductDetails(item);

                        if (isProcessInVisible) row.push(fmtStr(procInQty));
                        if (isProcessOutVisible) row.push(fmtStr(procOutQty));
                        row.push(fmtStr(getRecalculatedProcessQty(item)));
                        body.push(row);
                    });
                } else {
                    headers = [["S.No", ...configLabels, "Opening", "Stock In", "Process", "Stock Out", "Closing"]];
                    filteredData.forEach((item, idx) => {
                        const row: any[] = [idx + 1];
                        enabledConfigColumns.forEach(c => row.push(item[c.key] ?? "-"));
                        const { stockInQty, procInQty, procOutQty, outwardQty } = getProductDetails(item);
                        row.push(
                            fmtStr(getOpeningStock(item)),
                            fmtStr((processApiData.length > 0 ? stockInQty : getStockInTotal(item))),
                            fmtStr((processApiData.length > 0 ? (procInQty - procOutQty) : (getProcIn(item) - getProcOut(item)))),
                            fmtStr((processApiData.length > 0 ? outwardQty : getStockOutTotal(item))),
                            fmtStr(getClosingStock(item))
                        );
                        body.push(row);
                    });
                }
            } else {
                headers = [["S.No", "Godown Name", "OB", "Stock In", "Process", "Stock Out", "Closing"]];

                let sno = 1;
                Object.entries(groupedGodowns).forEach(([parentName, items]) => {
                    const groupOB = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_OB_Qty : r.OB_Qty) || 0), 0);
                    const groupIn = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_In_Qty : r.IN_Qty) || 0), 0);
                    const groupProcess = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.Process_Act_IN_OUT_Qty : r.Process_IN_OUT_Qty) || 0), 0);
                    const groupOut = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_Out_Qty : r.Out_Qty) || 0), 0);
                    const groupCL = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.CL_ACt_QTY : r.CL_QTY) || 0), 0);

                    // Add group row
                    body.push(["", parentName, fmtStr(groupOB), fmtStr(groupIn), fmtStr(groupProcess), fmtStr(groupOut), fmtStr(groupCL)]);

                    items.forEach((row) => {
                        body.push([
                            sno++,
                            row.godown_name,
                            fmtStr(Number((qtyMode === "actQty" ? row.ACt_OB_Qty : row.OB_Qty) || 0)),
                            fmtStr(Number((qtyMode === "actQty" ? row.ACt_In_Qty : row.IN_Qty) || 0)),
                            fmtStr(Number((qtyMode === "actQty" ? row.Process_Act_IN_OUT_Qty : row.Process_IN_OUT_Qty) || 0)),
                            fmtStr(Number((qtyMode === "actQty" ? row.ACt_Out_Qty : row.Out_Qty) || 0)),
                            fmtStr(Number((qtyMode === "actQty" ? row.CL_ACt_QTY : row.CL_QTY) || 0))
                        ]);
                    });
                });

                // Add grand total row
                body.push([
                    "",
                    "Total",
                    fmtStr(grandTotals.opening),
                    fmtStr(grandTotals.stockIn),
                    fmtStr(grandTotals.process),
                    fmtStr(grandTotals.stockOut),
                    fmtStr(grandTotals.closing)
                ]);
            }

            autoTable(doc, {
                head: headers,
                body: body,
                startY: 18,
                theme: "grid",
                styles: { fontSize: 8 },
                headStyles: { fillColor: [30, 58, 138] },
            });

            const filename = `Stock_Report_${selectedGodown ? selectedGodown.godown_name.replace(/\s+/g, '_') : 'Overall'}_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`;
            doc.save(filename);
            toast.success("PDF Exported Successfully ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export PDF ❌");
        }
    };


    const getTotalColumns = () => {
        const L = enabledConfigColumns.length;
        if (inwardMode) {
            const visibleTripsCount = inwardTripHeaders.filter(t => !hiddenInwardColumns.includes(t)).length;
            const isReturnVisible = !hiddenInwardColumns.includes("RETURN") ? 1 : 0;
            return L + visibleTripsCount + isReturnVisible + 2; // S.No + Config + Trips + Return + Total Stock In
        }
        if (outwardMode) {
            const visibleTripsCount = outwardTripHeaders.filter(t => !hiddenOutwardColumns.includes(t)).length;
            const isPendingVisible = !hiddenOutwardColumns.includes("PENDING DELIVERY") ? 1 : 0;
            return L + visibleTripsCount + isPendingVisible + 2; // S.No + Config + Trips + Pending Delivery + Total Outward
        }
        if (processMode) {
            const visibleIn = !hiddenProcessColumns.includes("PROCESS IN") ? 1 : 0;
            const visibleOut = !hiddenProcessColumns.includes("PROCESS OUT") ? 1 : 0;
            return L + visibleIn + visibleOut + 2; // S.No + Config + Proc In + Proc Out + Total Process
        }
        return L + 6;
    };

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
                            maxHeight: "calc(100vh - 260px)",
                            overflowY: "auto",
                            overflowX: "auto"
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
                                        const groupOB = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_OB_Qty : r.OB_Qty) || 0), 0);
                                        const groupIn = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_In_Qty : r.IN_Qty) || 0), 0);
                                        const groupProcess = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.Process_Act_IN_OUT_Qty : r.Process_IN_OUT_Qty) || 0), 0);
                                        const groupOut = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.ACt_Out_Qty : r.Out_Qty) || 0), 0);
                                        const groupCL = items.reduce((sum, r) => sum + Number((qtyMode === "actQty" ? r.CL_ACt_QTY : r.CL_QTY) || 0), 0);

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
                                                            {Number((qtyMode === "actQty" ? item.ACt_OB_Qty : item.OB_Qty) || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#2563eb" }}>
                                                            {Number((qtyMode === "actQty" ? item.ACt_In_Qty : item.IN_Qty) || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569" }}>
                                                            {Number((qtyMode === "actQty" ? item.Process_Act_IN_OUT_Qty : item.Process_IN_OUT_Qty) || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#ef4444" }}>
                                                            {Number((qtyMode === "actQty" ? item.ACt_Out_Qty : item.Out_Qty) || 0).toLocaleString()}
                                                        </TableCell>
                                                        <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#dcfce7", color: "#15803d" }}>
                                                            {Number((qtyMode === "actQty" ? item.CL_ACt_QTY : item.CL_QTY) || 0).toLocaleString()}
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

                    {/* Hidden Columns Bar */}
                    {inwardMode && hiddenInwardColumns.length > 0 && (
                        <Box sx={{ display: "flex", gap: 1, alignItems: "center", mb: 2, flexWrap: "wrap", p: 1.2, bgcolor: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 1.5 }}>
                            <Typography variant="body2" sx={{ fontWeight: 700, color: "#1e40af", fontSize: "0.72rem", textTransform: "uppercase", letterSpacing: 0.5 }}>
                                Hidden Columns (Click to Restore):
                            </Typography>
                            {hiddenInwardColumns.map((col) => (
                                <Chip
                                    key={col}
                                    label={col}
                                    size="small"
                                    onClick={() => handleShowInwardColumn(col)}
                                    color="primary"
                                    icon={<AddIcon fontSize="small" />}
                                    sx={{ fontWeight: 600, bgcolor: "#1e40af", color: "#fff", "& .MuiChip-icon": { color: "#fff" } }}
                                />
                            ))}
                        </Box>
                    )}

                    {outwardMode && hiddenOutwardColumns.length > 0 && (
                        <Box sx={{ display: "flex", gap: 1, alignItems: "center", mb: 2, flexWrap: "wrap", p: 1.2, bgcolor: "#fef2f2", border: "1px solid #fecaca", borderRadius: 1.5 }}>
                            <Typography variant="body2" sx={{ fontWeight: 700, color: "#991b1b", fontSize: "0.72rem", textTransform: "uppercase", letterSpacing: 0.5 }}>
                                Hidden Columns (Click to Restore):
                            </Typography>
                            {hiddenOutwardColumns.map((col) => (
                                <Chip
                                    key={col}
                                    label={col}
                                    size="small"
                                    onClick={() => handleShowOutwardColumn(col)}
                                    color="error"
                                    icon={<AddIcon fontSize="small" />}
                                    sx={{ fontWeight: 600, bgcolor: "#b91c1c", color: "#fff", "& .MuiChip-icon": { color: "#fff" } }}
                                />
                            ))}
                        </Box>
                    )}

                    {processMode && hiddenProcessColumns.length > 0 && (
                        <Box sx={{ display: "flex", gap: 1, alignItems: "center", mb: 2, flexWrap: "wrap", p: 1.2, bgcolor: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 1.5 }}>
                            <Typography variant="body2" sx={{ fontWeight: 700, color: "#166534", fontSize: "0.72rem", textTransform: "uppercase", letterSpacing: 0.5 }}>
                                Hidden Columns (Click to Restore):
                            </Typography>
                            {hiddenProcessColumns.map((col) => (
                                <Chip
                                    key={col}
                                    label={col}
                                    size="small"
                                    onClick={() => handleShowProcessColumn(col)}
                                    color="success"
                                    icon={<AddIcon fontSize="small" />}
                                    sx={{ fontWeight: 600, bgcolor: "#15803d", color: "#fff", "& .MuiChip-icon": { color: "#fff" } }}
                                />
                            ))}
                        </Box>
                    )}

                    {/* Detailed Stock Table Container */}
                    <TableContainer
                        component={Paper}
                        elevation={2}
                        sx={{
                            borderRadius: 2,
                            border: "1px solid #cbd5e1",
                            maxHeight: "calc(100vh - 250px)",
                            overflowY: "auto",
                            overflowX: "auto"
                        }}
                    >
                        <Table
                            size="small"
                            stickyHeader
                            sx={{
                                tableLayout: (inwardMode || outwardMode) ? "auto" : "fixed",
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
                                <TableRow ref={headerRowRef}>
                                    <TableCell align="center" sx={{ width: 60, minWidth: 60, backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>S.NO</TableCell>
                                    {enabledConfigColumns.map((col) => {
                                        const isProduct = col.key === "stock_item_name" || col.key === "Stock_Item";
                                        const w = isProduct ? 280 : 150;
                                        return (
                                            <TableCell
                                                key={col.key}
                                                sx={{
                                                    width: w,
                                                    minWidth: w,
                                                    backgroundColor: "#1E3A8A",
                                                    color: "#fff",
                                                    fontWeight: 600,
                                                    py: 1.5,
                                                    borderRight: "1px solid #cbd5e1"
                                                }}
                                            >
                                                {col.label.toUpperCase()}
                                            </TableCell>
                                        );
                                    })}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell align="right" sx={{ width: 120, minWidth: 120, backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5, borderRight: "1px solid #cbd5e1" }}>OPENING STOCK</TableCell>
                                    )}

                                    {/* Stock In Header - Clicking toggles inwardMode. Shown only in Normal Mode */}
                                    {!inwardMode && !outwardMode && !processMode && (
                                        <TableCell
                                            align="right"
                                            onClick={() => {
                                                handleSetInwardMode(true);
                                            }}
                                            sx={{
                                                width: 120,
                                                minWidth: 120,
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

                                    {/* Show TRIP DETAILS, RETURN, and TOTAL STOCK IN columns when inwardMode is active */}
                                    {inwardMode && (
                                        <>
                                            {inwardTripHeaders.map((tripLabel) => {
                                                if (hiddenInwardColumns.includes(tripLabel)) return null;
                                                return (
                                                    <TableCell
                                                        key={tripLabel}
                                                        align="right"
                                                        sx={{
                                                            width: 120,
                                                            minWidth: 120,
                                                            backgroundColor: "#1E3A8A",
                                                            color: "#fff",
                                                            fontWeight: 600,
                                                            py: 0.5,
                                                            borderRight: "1px solid #cbd5e1"
                                                        }}
                                                    >
                                                        <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                            {tripLabel.toUpperCase()}
                                                            <IconButton
                                                                size="small"
                                                                onClick={() => handleHideInwardColumn(tripLabel)}
                                                                sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                            >
                                                                <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                            </IconButton>
                                                        </Box>
                                                    </TableCell>
                                                );
                                            })}
                                            {!hiddenInwardColumns.includes("RETURN") && (
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        width: 120,
                                                        minWidth: 120,
                                                        backgroundColor: "#1E3A8A",
                                                        color: "#fff",
                                                        fontWeight: 600,
                                                        py: 0.5,
                                                        borderRight: "1px solid #cbd5e1"
                                                    }}
                                                >
                                                    <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                        RETURN
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => handleHideInwardColumn("RETURN")}
                                                            sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                        >
                                                            <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                        </IconButton>
                                                    </Box>
                                                </TableCell>
                                            )}
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetInwardMode(false)}
                                                sx={{
                                                    width: 140,
                                                    minWidth: 140,
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
                                                    TOTAL STOCK IN <KeyboardArrowUpIcon fontSize="small" />
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
                                                width: 120,
                                                minWidth: 120,
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

                                    {/* Show PROCESS IN, PROCESS OUT, and TOTAL PROCESS columns when processMode is active */}
                                    {processMode && (
                                        <>
                                            {!hiddenProcessColumns.includes("PROCESS IN") && (
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        width: 120,
                                                        minWidth: 120,
                                                        backgroundColor: "#1E3A8A",
                                                        color: "#fff",
                                                        fontWeight: 600,
                                                        py: 0.5,
                                                        borderRight: "1px solid #cbd5e1"
                                                    }}
                                                >
                                                    <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                        PROCESS IN
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => handleHideProcessColumn("PROCESS IN")}
                                                            sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                        >
                                                            <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                        </IconButton>
                                                    </Box>
                                                </TableCell>
                                            )}
                                            {!hiddenProcessColumns.includes("PROCESS OUT") && (
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        width: 120,
                                                        minWidth: 120,
                                                        backgroundColor: "#1E3A8A",
                                                        color: "#fff",
                                                        fontWeight: 600,
                                                        py: 0.5,
                                                        borderRight: "1px solid #cbd5e1"
                                                    }}
                                                >
                                                    <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                        PROCESS OUT
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => handleHideProcessColumn("PROCESS OUT")}
                                                            sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                        >
                                                            <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                        </IconButton>
                                                    </Box>
                                                </TableCell>
                                            )}
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetProcessMode(false)}
                                                sx={{
                                                    width: 140,
                                                    minWidth: 140,
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
                                                width: 120,
                                                minWidth: 120,
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
                                            {outwardTripHeaders.map((tripLabel) => {
                                                if (hiddenOutwardColumns.includes(tripLabel)) return null;
                                                return (
                                                    <TableCell
                                                        key={tripLabel}
                                                        align="right"
                                                        sx={{
                                                            width: 120,
                                                            minWidth: 120,
                                                            backgroundColor: "#1E3A8A",
                                                            color: "#fff",
                                                            fontWeight: 600,
                                                            py: 0.5,
                                                            borderRight: "1px solid #cbd5e1"
                                                        }}
                                                    >
                                                        <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                            {tripLabel.toUpperCase()}
                                                            <IconButton
                                                                size="small"
                                                                onClick={() => handleHideOutwardColumn(tripLabel)}
                                                                sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                            >
                                                                <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                            </IconButton>
                                                        </Box>
                                                    </TableCell>
                                                );
                                            })}
                                            {!hiddenOutwardColumns.includes("PENDING DELIVERY") && (
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        width: 130,
                                                        minWidth: 130,
                                                        backgroundColor: "#1E3A8A",
                                                        color: "#fff",
                                                        fontWeight: 600,
                                                        py: 0.5,
                                                        borderRight: "1px solid #cbd5e1"
                                                    }}
                                                >
                                                    <Box display="flex" alignItems="center" justifyContent="flex-end" gap={0.5}>
                                                        PENDING DELIVERY
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => handleHideOutwardColumn("PENDING DELIVERY")}
                                                            sx={{ color: "rgba(255,255,255,0.7)", p: 0.2, "&:hover": { color: "#fff" } }}
                                                        >
                                                            <CloseIcon fontSize="inherit" sx={{ fontSize: 14 }} />
                                                        </IconButton>
                                                    </Box>
                                                </TableCell>
                                            )}
                                            <TableCell
                                                align="right"
                                                onClick={() => handleSetOutwardMode(false)}
                                                sx={{
                                                    width: 140,
                                                    minWidth: 140,
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
                                        <TableCell align="right" sx={{ width: 120, minWidth: 120, backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 600, py: 1.5 }}>
                                            CLOSING STOCK
                                        </TableCell>
                                    )}
                                </TableRow>
                                {paginatedData.length > 0 && (
                                    <TableRow sx={{ backgroundColor: "#f1f5f9" }}>
                                        <TableCell colSpan={1 + enabledConfigColumns.length} align="center" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800 }}>
                                            GRAND TOTAL
                                        </TableCell>
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                {formatQtyVal(detailedTotals.opening)}
                                            </TableCell>
                                        )}

                                        {/* Normal Mode: Stock In */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                {formatQtyVal(detailedTotals.totalInward)}
                                            </TableCell>
                                        )}

                                        {/* Inward Mode: Trip Details & Returns & Total Stock In */}
                                        {inwardMode && (
                                            <>
                                                {inwardTripHeaders.map((tripLabel) => {
                                                    if (hiddenInwardColumns.includes(tripLabel)) return null;
                                                    return (
                                                        <TableCell key={tripLabel} align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 110, minWidth: 110 }}>
                                                            {formatQtyVal(inwardTripTotals[tripLabel])}
                                                        </TableCell>
                                                    );
                                                })}
                                                {!hiddenInwardColumns.includes("RETURN") && (
                                                    <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                        {formatQtyVal(detailedTotals.returnQtyTotal)}
                                                    </TableCell>
                                                )}
                                                <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#1e40af", width: 140, minWidth: 140 }}>
                                                    {formatQtyVal(recalculatedTotals.inwardTotal)}
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Process */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                {formatQtyVal(detailedTotals.totalProcess)}
                                            </TableCell>
                                        )}

                                        {/* Process Mode: PROCESS IN, PROCESS OUT, TOTAL PROCESS */}
                                        {processMode && (
                                            <>
                                                {!hiddenProcessColumns.includes("PROCESS IN") && (
                                                    <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                        {formatQtyVal(detailedTotals.procIn)}
                                                    </TableCell>
                                                )}
                                                {!hiddenProcessColumns.includes("PROCESS OUT") && (
                                                    <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                        {formatQtyVal(detailedTotals.procOut)}
                                                    </TableCell>
                                                )}
                                                <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#1e40af", width: 140, minWidth: 140 }}>
                                                    {formatQtyVal(recalculatedTotals.processTotal)}
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Stock Outwards */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 120, minWidth: 120 }}>
                                                {formatQtyVal(detailedTotals.totalOutward)}
                                            </TableCell>
                                        )}

                                        {/* Outward Mode: Out Details & Delivery & Total Outward */}
                                        {outwardMode && (
                                            <>
                                                {outwardTripHeaders.map((tripLabel) => {
                                                    if (hiddenOutwardColumns.includes(tripLabel)) return null;
                                                    return (
                                                        <TableCell key={tripLabel} align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 110, minWidth: 110 }}>
                                                            {formatQtyVal(outwardTripTotals[tripLabel])}
                                                        </TableCell>
                                                    );
                                                })}
                                                {!hiddenOutwardColumns.includes("PENDING DELIVERY") && (
                                                    <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, pr: 2, width: 130, minWidth: 130 }}>
                                                        {formatQtyVal(detailedTotals.deliveryQtyTotal)}
                                                    </TableCell>
                                                )}
                                                <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#b91c1c", width: 140, minWidth: 140 }}>
                                                    {formatQtyVal(recalculatedTotals.outwardTotal)}
                                                </TableCell>
                                            </>
                                        )}

                                        {/* Normal Mode: Closing */}
                                        {!inwardMode && !outwardMode && !processMode && (
                                            <TableCell align="right" sx={{ position: "sticky", top: headerHeight, zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, pr: 2, color: "#15803d", width: 120, minWidth: 120 }}>
                                                {formatQtyVal(detailedTotals.closing)}
                                            </TableCell>
                                        )}
                                    </TableRow>
                                )}
                            </TableHead>
                            <TableBody>
                                {paginatedData.length === 0 ? (
                                    <TableRow>
                                        <TableCell colSpan={getTotalColumns()} align="center" sx={{ py: 6, color: "#94a3b8" }}>
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

                                        const { trips, returnQty, stockInQty, procInQty, procOutQty, outwardQty, outTrips, deliveryQty } = getProductDetails(item);

                                        return (
                                            <React.Fragment key={idx}>
                                                {/* Brand Group Separator */}
                                                {showBrandHeader && (
                                                    <TableRow sx={{ backgroundColor: "#f1f5f9" }}>
                                                        <TableCell
                                                            colSpan={getTotalColumns()}
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
                                                    <TableCell align="center" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, color: "#475569", width: 60, minWidth: 60 }}>
                                                        {sNo}
                                                    </TableCell>
                                                    {enabledConfigColumns.map((col) => {
                                                        const val = item[col.key] ?? "-";
                                                        const isProduct = col.key === "stock_item_name" || col.key === "Stock_Item";
                                                        const w = isProduct ? 280 : 150;
                                                        return (
                                                            <TableCell
                                                                key={col.key}
                                                                sx={{
                                                                    borderRight: "1px solid #e2e8f0",
                                                                    fontWeight: isProduct ? 700 : 600,
                                                                    color: isProduct ? "#1e293b" : "#475569",
                                                                    wordBreak: "break-word",
                                                                    whiteSpace: "normal",
                                                                    width: w,
                                                                    minWidth: w
                                                                }}
                                                            >
                                                                {val}
                                                            </TableCell>
                                                        );
                                                    })}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569", width: 120, minWidth: 120 }}>
                                                            {formatQtyVal(openingStock)}
                                                        </TableCell>
                                                    )}

                                                    {/* Inward Mode or Normal Mode: Render Stock In */}
                                                    {!inwardMode && !outwardMode && !processMode && (
                                                        <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: (processApiData.length > 0 ? stockInQty : stockIn) > 0 ? "#2563eb" : "#475569", width: 120, minWidth: 120 }}>
                                                            {formatQtyVal(processApiData.length > 0 ? stockInQty : stockIn)}
                                                        </TableCell>
                                                    )}

                                                    {/* Inward Mode: Render dynamic TRIP columns, RETURN and TOTAL STOCK IN */}
                                                    {inwardMode && (
                                                        <>
                                                            {inwardTripHeaders.map((tripLabel) => {
                                                                if (hiddenInwardColumns.includes(tripLabel)) return null;
                                                                const qty = getQtyForTrip(trips, tripLabel);
                                                                return (
                                                                    <TableCell key={tripLabel} align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: qty > 0 ? "#1e293b" : "#94a3b8", width: 110, minWidth: 110 }}>
                                                                        {formatQtyVal(qty)}
                                                                    </TableCell>
                                                                );
                                                            })}
                                                            {!hiddenInwardColumns.includes("RETURN") && (
                                                                <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: returnQty > 0 ? "#475569" : "#475569", width: 120, minWidth: 120 }}>
                                                                    {formatQtyVal(returnQty)}
                                                                </TableCell>
                                                            )}
                                                            <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#eff6ff", color: "#1e40af", width: 140, minWidth: 140 }}>
                                                                {formatQtyVal(getRecalculatedStockInQty(item))}
                                                             </TableCell>
                                                         </>
                                                     )}

                                                     {/* Normal Mode: Render Process */}
                                                     {!inwardMode && !outwardMode && !processMode && (
                                                         <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569", width: 120, minWidth: 120 }}>
                                                             {formatQtyVal(processApiData.length > 0 ? (procInQty - procOutQty) : (getProcIn(item) - getProcOut(item)))}
                                                         </TableCell>
                                                     )}

                                                     {/* Process Mode: Render PROCESS IN, PROCESS OUT, and TOTAL PROCESS */}
                                                     {processMode && (
                                                         <>
                                                             {!hiddenProcessColumns.includes("PROCESS IN") && (
                                                                 <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569", width: 120, minWidth: 120 }}>
                                                                     {formatQtyVal(procInQty)}
                                                                 </TableCell>
                                                             )}
                                                             {!hiddenProcessColumns.includes("PROCESS OUT") && (
                                                                 <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569", width: 120, minWidth: 120 }}>
                                                                     {formatQtyVal(procOutQty)}
                                                                 </TableCell>
                                                             )}
                                                             <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#eff6ff", color: "#1e40af", width: 140, minWidth: 140 }}>
                                                                 {formatQtyVal(getRecalculatedProcessQty(item))}
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
                                                                 color: (processApiData.length > 0 ? outwardQty : stockOut) > 0 ? "#ef4444" : "#475569",
                                                                 width: 120,
                                                                 minWidth: 120
                                                             }}
                                                         >
                                                             {formatQtyVal(processApiData.length > 0 ? outwardQty : stockOut)}
                                                         </TableCell>
                                                     )}

                                                     {/* Outward Mode: Render dynamic TRIP columns, PENDING DELIVERY, and TOTAL OUTWARD */}
                                                     {outwardMode && (
                                                         <>
                                                             {outwardTripHeaders.map((tripLabel) => {
                                                                 if (hiddenOutwardColumns.includes(tripLabel)) return null;
                                                                 const qty = getQtyForTrip(outTrips, tripLabel);
                                                                 return (
                                                                     <TableCell key={tripLabel} align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: qty > 0 ? "#1e293b" : "#94a3b8", width: 110, minWidth: 110 }}>
                                                                         {formatQtyVal(qty)}
                                                                     </TableCell>
                                                                 );
                                                             })}
                                                             {!hiddenOutwardColumns.includes("PENDING DELIVERY") && (
                                                                 <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600, pr: 2, color: "#475569", width: 130, minWidth: 130 }}>
                                                                     {formatQtyVal(deliveryQty)}
                                                                 </TableCell>
                                                             )}
                                                             <TableCell align="right" sx={{ fontWeight: 700, pr: 2, backgroundColor: "#fef2f2", color: "#b91c1c", width: 140, minWidth: 140 }}>
                                                                 {formatQtyVal(getRecalculatedStockOutQty(item))}
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
                                                                color: closing > 0 ? "#15803d" : "#475569",
                                                                width: 120,
                                                                minWidth: 120
                                                            }}
                                                        >
                                                            {formatQtyVal(closing)}
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
                showQtyModeFilter={true}
                qtyModeValue={qtyMode}
                onQtyModeChange={setQtyMode}
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
