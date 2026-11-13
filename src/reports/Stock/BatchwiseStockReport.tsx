import React, { useCallback, useEffect, useMemo, useState } from "react";
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
    FormControl,
    FormLabel,
    RadioGroup,
    FormControlLabel,
    Radio,
    Typography,
    Dialog,
    DialogTitle,
    DialogContent,
    IconButton,
    Chip,
    Tooltip,
    Popover,
    TextField,
    Button,
    Autocomplete,
} from "@mui/material";
import FilterAltIcon from "@mui/icons-material/FilterAlt";
import CalendarMonthIcon from "@mui/icons-material/CalendarMonth";
import CloseIcon from "@mui/icons-material/Close";
import dayjs from "dayjs";
import AppLayout, { useToggleMode } from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import HeaderFilterMenu from "../../Components/HeaderFilterMenu";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import {
    BatchStockReportService,
    BatchStockItem,
    GodownMasterItem,
} from "../../services/batchStockReport.service";
import {
    transactionBatchStockService,
    TransactionBatchStockItem,
} from "../../services/stockWiseReport.service";

const formatNumber = (val: number): string => {
    return Number(val || 0).toLocaleString("en-IN", {
        minimumFractionDigits: 0,
        maximumFractionDigits: 3,
    });
};

const parseNum = (val: any): number => {
    if (typeof val === "number") return isNaN(val) ? 0 : val;
    if (!val) return 0;
    const clean = String(val).replace(/,/g, "").trim();
    const parsed = parseFloat(clean);
    return isNaN(parsed) ? 0 : parsed;
};

const isValidBatch = (b: any): boolean => {
    if (b === null || b === undefined) return false;
    const s = String(b).trim();
    if (!s || s === "—" || s === "-" || s.toLowerCase() === "null") return false;
    return true;
};

const BatchwiseStockReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const { toggleMode, setToggleMode } = useToggleMode();
    const isExpanded = toggleMode === "Expanded";

    // Data states
    const [rawRows, setRawRows] = useState<BatchStockItem[]>([]);
    const [loading, setLoading] = useState(false);

    // Godown master states (Expanded side)
    const [godowns, setGodowns] = useState<GodownMasterItem[]>([]);
    const [selectedGodownId, setSelectedGodownId] = useState<string>("");
    const [tempSelectedGodownId, setTempSelectedGodownId] = useState<string>("");

    // Filter drawer states
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [tempFromDate, setTempFromDate] = useState(today);
    const [tempToDate, setTempToDate] = useState(today);

    // Qty Mode Filter: "qty" (Qty) or "actQty" (Actual Qty)
    const [qtyMode, setQtyMode] = useState<"qty" | "actQty">("qty");
    const [tempQtyMode, setTempQtyMode] = useState<"qty" | "actQty">("qty");

    // Data Scope Filter: "batch" (Batch datas, default) or "all" (All datas)
    const [dataScope, setDataScope] = useState<"all" | "batch">("batch");
    const [tempDataScope, setTempDataScope] = useState<"all" | "batch">("batch");

    // Header Filter states for Item Name and Batch
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeFilterColumn, setActiveFilterColumn] = useState<"itemName" | "batch" | null>(null);
    const [selectedItemNames, setSelectedItemNames] = useState<string[] | undefined>(undefined);
    const [selectedBatches, setSelectedBatches] = useState<string[] | undefined>(undefined);

    // Pagination
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    /* ===== BATCH TRANSACTION POPUP MODAL STATES ===== */
    const [batchModalOpen, setBatchModalOpen] = useState(false);
    const [selectedBatchItem, setSelectedBatchItem] = useState<{
        productId: string | number;
        productName: string;
        batchNo: string;
        apiBatch: string;
        godownId?: string | number;
        godownName?: string;
        isGodownWise?: boolean;
    } | null>(null);
    const [batchTransactions, setBatchTransactions] = useState<TransactionBatchStockItem[]>([]);
    const [batchModalLoading, setBatchModalLoading] = useState(false);
    const [batchModalError, setBatchModalError] = useState<string | null>(null);

    /* ===== BATCH-SPECIFIC DATE FILTER STATES (Inside Pop-up) ===== */
    const [batchFromDate, setBatchFromDate] = useState<string>(today);
    const [batchToDate, setBatchToDate] = useState<string>(today);
    const [tempBatchFromDate, setTempBatchFromDate] = useState<string>(today);
    const [tempBatchToDate, setTempBatchToDate] = useState<string>(today);
    const [isBatchDateCustom, setIsBatchDateCustom] = useState<boolean>(false);
    const [batchDateAnchorEl, setBatchDateAnchorEl] = useState<HTMLElement | null>(null);

    /* ================= FETCH GODOWN MASTER ================= */
    useEffect(() => {
        const fetchGodowns = async () => {
            try {
                const res = await BatchStockReportService.getGodownMaster();
                const list = res.data?.data || (Array.isArray(res.data) ? res.data : []);
                const validList = Array.isArray(list)
                    ? list.filter((g) => g.Godown_Name && String(g.Godown_Name).trim() !== "" && String(g.Godown_Id) !== "0")
                    : [];
                setGodowns(validList);
            } catch (err) {
                console.error("Error fetching godown master:", err);
            }
        };
        fetchGodowns();
    }, []);

    /* ================= FETCH MAIN DATA ================= */
    const loadData = async (from: string, to: string, godownId: string, expanded: boolean) => {
        try {
            if (expanded && !godownId) {
                setRawRows([]);
                setLoading(false);
                return;
            }

            setLoading(true);
            let res;
            if (expanded) {
                res = await BatchStockReportService.getBatchStockReportGodown({
                    Fromdate: from,
                    Todate: to,
                    Godown_Id: godownId,
                });
            } else {
                res = await BatchStockReportService.getBatchStockReport({
                    Fromdate: from,
                    Todate: to,
                });
            }

            const data = res.data?.data || (Array.isArray(res.data) ? res.data : []);
            setRawRows(Array.isArray(data) ? data : []);
            setPage(1);
        } catch (err) {
            console.error("Error fetching batch stock report:", err);
            setRawRows([]);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        loadData(fromDate, toDate, selectedGodownId, isExpanded);
    }, [fromDate, toDate, selectedGodownId, isExpanded]);

    /* ================= APPLY FILTER DRAWER ================= */
    const handleApplyFilters = () => {
        setFromDate(tempFromDate);
        setToDate(tempToDate);
        setQtyMode(tempQtyMode);
        setDataScope(tempDataScope);
        if (isExpanded) {
            setSelectedGodownId(tempSelectedGodownId);
        }
        setDrawerOpen(false);
        setPage(1);
    };

    /* ================= PROCESSED GRID ROWS ================= */
    const processedRows = useMemo(() => {
        const isAct = qtyMode === "actQty";

        const scopedRows = dataScope === "batch"
            ? rawRows.filter((item) => isValidBatch(item.Batch_No))
            : rawRows;

        return scopedRows.map((item, idx) => {
            const ob = isAct
                ? parseNum(item.OB_Act_Qty)
                : parseNum(item.OB_Bal_Qty);

            const inQty = isAct
                ? parseNum(item.IN_Act_Qty)
                : parseNum(item.IN_Qty);

            const outQty = isAct
                ? parseNum(item.OUT_Act_Qty)
                : parseNum(item.OUT_Qty);

            // Frontend formula for Closing: OB + IN - OUT
            const cls = ob + inQty - outQty;

            const batchLabel = isValidBatch(item.Batch_No)
                ? String(item.Batch_No).trim()
                : "—";

            return {
                id: `${item.Product_Id}-${item.Batch_No ?? "null"}-${idx}`,
                productId: item.Product_Id,
                rawBatch: item.Batch_No,
                godownId: item.Godown_Id,
                godownName: item.Godown_Name,
                itemName: item.stock_item_name || "—",
                batch: batchLabel,
                ob,
                inQty,
                outQty,
                cls,
                rawItem: item,
            };
        });
    }, [rawRows, qtyMode, dataScope]);

    /* ================= FILTERED ROWS (ITEM NAME & BATCH HEADER FILTERS) ================= */
    const filteredRows = useMemo(() => {
        return processedRows.filter((row) => {
            if (selectedItemNames && selectedItemNames.length > 0) {
                if (!selectedItemNames.includes(row.itemName)) return false;
            }
            if (selectedBatches && selectedBatches.length > 0) {
                if (!selectedBatches.includes(row.batch)) return false;
            }
            return true;
        });
    }, [processedRows, selectedItemNames, selectedBatches]);

    /* ================= OPTIONS FOR HEADER FILTERS ================= */
    const itemNameOptions = useMemo(() => {
        const unique = Array.from(new Set(processedRows.map((r) => r.itemName))).filter(Boolean);
        unique.sort((a, b) => a.localeCompare(b));
        return unique;
    }, [processedRows]);

    const batchOptions = useMemo(() => {
        const unique = Array.from(new Set(processedRows.map((r) => r.batch))).filter(Boolean);
        unique.sort((a, b) => a.localeCompare(b));
        return unique;
    }, [processedRows]);

    /* ================= TOTALS ================= */
    const totals = useMemo(() => {
        let totalOB = 0;
        let totalIN = 0;
        let totalOUT = 0;
        let totalCLS = 0;

        filteredRows.forEach((r) => {
            totalOB += r.ob;
            totalIN += r.inQty;
            totalOUT += r.outQty;
            totalCLS += r.cls;
        });

        return {
            totalOB,
            totalIN,
            totalOUT,
            totalCLS,
        };
    }, [filteredRows]);

    /* ================= PAGINATION ================= */
    const paginatedRows = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        return filteredRows.slice(start, start + rowsPerPage);
    }, [filteredRows, page, rowsPerPage]);

    /* ================= POPUP BATCH DETAILS FETCH ================= */
    const loadBatchTransactions = useCallback(
        async (
            selected: NonNullable<typeof selectedBatchItem>,
            targetFromDate = batchFromDate,
            targetToDate = batchToDate
        ) => {
            const productId = selected.productId || "";
            if (!productId) {
                setBatchModalError("Product ID is missing.");
                setBatchTransactions([]);
                return;
            }

            setBatchModalLoading(true);
            setBatchModalError(null);

            try {
                let res;
                const batchParam = selected.apiBatch;

                if (selected.isGodownWise && selected.godownId) {
                    res = await transactionBatchStockService.getTransactionBatchStockByGodown({
                        fromDate: targetFromDate,
                        toDate: targetToDate,
                        Product_Id: productId,
                        Godown_Id: selected.godownId,
                        Batch: batchParam,
                    });
                } else {
                    res = await transactionBatchStockService.getTransactionBatchStock({
                        fromDate: targetFromDate,
                        toDate: targetToDate,
                        Product_Id: productId,
                        Batch: batchParam,
                    });
                }

                const rawData = res.data?.data || [];
                setBatchTransactions(Array.isArray(rawData) ? rawData : []);
            } catch (err: any) {
                console.error("Error fetching batch transactions:", err);
                setBatchModalError(
                    err.response?.data?.message || "Failed to load batch transaction data."
                );
                setBatchTransactions([]);
            } finally {
                setBatchModalLoading(false);
            }
        },
        [batchFromDate, batchToDate]
    );

    useEffect(() => {
        if (!batchModalOpen || !selectedBatchItem) return;
        loadBatchTransactions(selectedBatchItem, batchFromDate, batchToDate);
    }, [batchModalOpen, selectedBatchItem, batchFromDate, batchToDate, loadBatchTransactions]);

    const handleOpenItemBatchModal = (row: typeof processedRows[0]) => {
        const rawBatch = row.rawBatch;
        let displayBatch = "Unassigned";
        let apiBatch = "";

        if (rawBatch !== null && rawBatch !== undefined && String(rawBatch).trim().toLowerCase() !== "null") {
            const trimmed = String(rawBatch).trim();
            if (trimmed !== "" && trimmed !== "—" && trimmed !== "-") {
                displayBatch = trimmed;
                apiBatch = trimmed;
            }
        }

        const isGodown = Boolean(isExpanded || row.godownId);
        const godownId = isGodown ? (row.godownId || selectedGodownId || "10") : undefined;
        const matchedGodown = godowns.find((g) => String(g.Godown_Id) === String(godownId));
        const godownName = row.godownName || (matchedGodown ? matchedGodown.Godown_Name : (isGodown ? "Main Location" : undefined));

        setSelectedBatchItem({
            productId: row.productId,
            productName: row.itemName,
            batchNo: displayBatch,
            apiBatch,
            godownId,
            godownName,
            isGodownWise: Boolean(isGodown && godownId),
        });

        setBatchFromDate(fromDate);
        setBatchToDate(toDate);
        setTempBatchFromDate(fromDate);
        setTempBatchToDate(toDate);
        setIsBatchDateCustom(false);
        setBatchDateAnchorEl(null);
        setBatchModalOpen(true);
    };

    /* ===== PROCESSED BATCH LEDGER (SORTED BY DATE) ===== */
    const processedBatchRows = useMemo(() => {
        const isAct = qtyMode === "actQty";

        const sorted = [...batchTransactions].sort((a, b) => {
            const isA_OB =
                a.invoice_no === "OB" ||
                a.ord === 0 ||
                String(a.Particulars || "").toLowerCase().includes("opening balance");
            const isB_OB =
                b.invoice_no === "OB" ||
                b.ord === 0 ||
                String(b.Particulars || "").toLowerCase().includes("opening balance");

            if (isA_OB && !isB_OB) return -1;
            if (!isA_OB && isB_OB) return 1;

            const dateA = a.Ledger_Date ? new Date(a.Ledger_Date).getTime() : 0;
            const dateB = b.Ledger_Date ? new Date(b.Ledger_Date).getTime() : 0;
            if (dateA !== dateB) return dateA - dateB;

            const ordA = Number(a.ord || 0);
            const ordB = Number(b.ord || 0);
            if (ordA !== ordB) return ordA - ordB;

            return Number(a.Trans_Id || 0) - Number(b.Trans_Id || 0);
        });

        let running = 0;
        return sorted.map((row, index) => {
            const inQtyVal = isAct
                ? Number(row.In_Act_Qty !== undefined && row.In_Act_Qty !== null ? row.In_Act_Qty : (row.In_Qty || 0))
                : Number(row.In_Qty || 0);
            const outQtyVal = isAct
                ? Number(row.Out_Act_Qty !== undefined && row.Out_Act_Qty !== null ? row.Out_Act_Qty : (row.Out_Qty || 0))
                : Number(row.Out_Qty || 0);

            running += inQtyVal - outQtyVal;

            return {
                ...row,
                sNo: index + 1,
                inQtyVal,
                outQtyVal,
                clsQtyVal: running,
            };
        });
    }, [batchTransactions, qtyMode]);

    const batchTotals = useMemo(() => {
        let totalIn = 0;
        let totalOut = 0;

        processedBatchRows.forEach((r) => {
            totalIn += r.inQtyVal;
            totalOut += r.outQtyVal;
        });

        const finalCls =
            processedBatchRows.length > 0
                ? processedBatchRows[processedBatchRows.length - 1].clsQtyVal
                : 0;

        return { totalIn, totalOut, finalCls };
    }, [processedBatchRows]);

    /* ================= EXPORT HANDLERS (GRID ONLY) ================= */
    const currentGodownName = useMemo(() => {
        if (!selectedGodownId) return "";
        const found = godowns.find((g) => String(g.Godown_Id) === String(selectedGodownId));
        return found ? found.Godown_Name : "";
    }, [selectedGodownId, godowns]);

    const handleExportExcel = () => {
        const isExp = isExpanded;
        const headers = isExp
            ? ["S.No", "Godown Name", "Item Name", "Batch", "OB", "IN", "OUT", "CLS"]
            : ["S.No", "Item Name", "Batch", "OB", "IN", "OUT", "CLS"];

        const exportData = filteredRows.map((row, index) => {
            if (isExp) {
                return [
                    index + 1,
                    currentGodownName || row.godownName || "—",
                    row.itemName,
                    row.batch,
                    row.ob,
                    row.inQty,
                    row.outQty,
                    row.cls,
                ];
            }
            return [
                index + 1,
                row.itemName,
                row.batch,
                row.ob,
                row.inQty,
                row.outQty,
                row.cls,
            ];
        });

        if (isExp) {
            exportData.push([
                "",
                "",
                "TOTAL",
                "",
                totals.totalOB,
                totals.totalIN,
                totals.totalOUT,
                totals.totalCLS,
            ]);
        } else {
            exportData.push([
                "",
                "TOTAL",
                "",
                totals.totalOB,
                totals.totalIN,
                totals.totalOUT,
                totals.totalCLS,
            ]);
        }

        const fileName = isExp
            ? `Batchwise_Stock_Report_Expanded_${currentGodownName ? currentGodownName.replace(/[^a-zA-Z0-9_-]/g, "_") + "_" : ""}${fromDate}_to_${toDate}`
            : `Batchwise_Stock_Report_Abstract_${fromDate}_to_${toDate}`;

        exportToExcel(
            fileName,
            headers,
            exportData
        );
    };

    const handleExportPDF = () => {
        const isExp = isExpanded;
        const headers = isExp
            ? ["S.No", "Godown Name", "Item Name", "Batch", "OB", "IN", "OUT", "CLS"]
            : ["S.No", "Item Name", "Batch", "OB", "IN", "OUT", "CLS"];

        const exportData = filteredRows.map((row, index) => {
            if (isExp) {
                return [
                    index + 1,
                    currentGodownName || row.godownName || "—",
                    row.itemName,
                    row.batch,
                    formatNumber(row.ob),
                    formatNumber(row.inQty),
                    formatNumber(row.outQty),
                    formatNumber(row.cls),
                ];
            }
            return [
                index + 1,
                row.itemName,
                row.batch,
                formatNumber(row.ob),
                formatNumber(row.inQty),
                formatNumber(row.outQty),
                formatNumber(row.cls),
            ];
        });

        if (isExp) {
            exportData.push([
                "",
                "",
                "TOTAL",
                "",
                formatNumber(totals.totalOB),
                formatNumber(totals.totalIN),
                formatNumber(totals.totalOUT),
                formatNumber(totals.totalCLS),
            ]);
        } else {
            exportData.push([
                "",
                "TOTAL",
                "",
                formatNumber(totals.totalOB),
                formatNumber(totals.totalIN),
                formatNumber(totals.totalOUT),
                formatNumber(totals.totalCLS),
            ]);
        }

        const title = isExp
            ? `Batchwise Stock Report (Expanded - ${currentGodownName || "Godown Wise"}) (${fromDate} to ${toDate})`
            : `Batchwise Stock Report (Abstract) (${fromDate} to ${toDate})`;

        exportToPDF(
            title,
            headers,
            exportData
        );
    };

    return (
        <>
            <PageHeader
                parentReportName="BATCHWISE STOCK REPORT"
                toggleMode={toggleMode}
                onToggleChange={setToggleMode}
                abstractLabel="Abstract"
                expandedLabel="Expanded"
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => {
                    setTempFromDate(fromDate);
                    setTempToDate(toDate);
                    setTempQtyMode(qtyMode);
                    setTempDataScope(dataScope);
                    setTempSelectedGodownId(selectedGodownId);
                    setDrawerOpen((p) => !p);
                }}
                onClose={() => setDrawerOpen(false)}
                fromDate={tempFromDate}
                toDate={tempToDate}
                onFromDateChange={(val) => setTempFromDate(val)}
                onToDateChange={(val) => setTempToDate(val)}
                onApply={handleApplyFilters}
            >
                {/* GODOWN DROPDOWN FILTER (EXPANDED SIDE ONLY - MATCHING STOCK IN HAND REPORT) */}
                {isExpanded && (
                    <Box mb={2}>
                        <Autocomplete
                            options={godowns}
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
                            value={
                                godowns.find((g) => String(g.Godown_Id) === String(tempSelectedGodownId)) || null
                            }
                            onChange={(_, newValue: GodownMasterItem | null) => {
                                setTempSelectedGodownId(newValue != null ? String(newValue.Godown_Id) : "");
                            }}
                            renderInput={(params) => (
                                <TextField
                                    {...params}
                                    label="Godown"
                                    placeholder="Search..."
                                />
                            )}
                            sx={{ mb: 2 }}
                        />
                    </Box>
                )}

                {/* QUANTITY TYPE FILTER */}
                <FormControl sx={{ mb: 2.5, display: "block" }}>
                    <FormLabel
                        sx={{
                            fontWeight: 600,
                            color: "#1E3A8A",
                            fontSize: "0.875rem",
                            display: "block",
                            mb: 0.5,
                        }}
                    >
                        Quantity Type
                    </FormLabel>
                    <RadioGroup
                        value={tempQtyMode}
                        onChange={(e) => {
                            const next = e.target.value as "qty" | "actQty";
                            setTempQtyMode(next);
                            setQtyMode(next);
                        }}
                    >
                        <FormControlLabel
                            value="qty"
                            control={
                                <Radio
                                    size="small"
                                    sx={{
                                        color: "#1E3A8A",
                                        "&.Mui-checked": { color: "#1E3A8A" },
                                    }}
                                />
                            }
                            label={
                                <Typography sx={{ fontSize: "0.825rem" }}>
                                    Qty
                                </Typography>
                            }
                        />
                        <FormControlLabel
                            value="actQty"
                            control={
                                <Radio
                                    size="small"
                                    sx={{
                                        color: "#1E3A8A",
                                        "&.Mui-checked": { color: "#1E3A8A" },
                                    }}
                                />
                            }
                            label={
                                <Typography sx={{ fontSize: "0.825rem" }}>
                                    Actual Qty
                                </Typography>
                            }
                        />
                    </RadioGroup>
                </FormControl>

                {/* DATA SCOPE FILTER: ALL DATAS vs BATCH DATAS */}
                <FormControl sx={{ mb: 2, display: "block" }}>
                    <FormLabel
                        sx={{
                            fontWeight: 600,
                            color: "#1E3A8A",
                            fontSize: "0.875rem",
                            display: "block",
                            mb: 0.5,
                        }}
                    >
                        Data View
                    </FormLabel>
                    <RadioGroup
                        value={tempDataScope}
                        onChange={(e) => {
                            const next = e.target.value as "all" | "batch";
                            setTempDataScope(next);
                            setDataScope(next);
                        }}
                    >
                        <FormControlLabel
                            value="all"
                            control={
                                <Radio
                                    size="small"
                                    sx={{
                                        color: "#1E3A8A",
                                        "&.Mui-checked": { color: "#1E3A8A" },
                                    }}
                                />
                            }
                            label={
                                <Typography sx={{ fontSize: "0.825rem" }}>
                                    All datas
                                </Typography>
                            }
                        />
                        <FormControlLabel
                            value="batch"
                            control={
                                <Radio
                                    size="small"
                                    sx={{
                                        color: "#1E3A8A",
                                        "&.Mui-checked": { color: "#1E3A8A" },
                                    }}
                                />
                            }
                            label={
                                <Typography sx={{ fontSize: "0.825rem" }}>
                                    Batch datas
                                </Typography>
                            }
                        />
                    </RadioGroup>
                </FormControl>
            </ReportFilterDrawer>

            <AppLayout fullWidth>
                <Box px={2} pb={1} pt={1}>
                    {loading ? (
                        <Box
                            display="flex"
                            justifyContent="center"
                            alignItems="center"
                            minHeight="300px"
                        >
                            <CircularProgress color="primary" />
                        </Box>
                    ) : (
                        <>
                            <TableContainer
                                component={Paper}
                                elevation={1}
                                sx={{
                                    borderRadius: 2,
                                    border: "1px solid #cbd5e1",
                                    overflow: "auto",
                                    maxHeight: "calc(100vh - 120px)",
                                }}
                            >
                                <Table size="small" stickyHeader>
                                    <TableHead
                                        sx={{
                                            background: "#1E3A8A",
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 3,
                                            "& .MuiTableCell-root": {
                                                bgcolor: "#1E3A8A !important",
                                                color: "#fff !important",
                                                fontWeight: 700,
                                                fontSize: "0.75rem",
                                                borderRight: "1px solid #2448b2",
                                                whiteSpace: "nowrap",
                                                py: 1,
                                            },
                                        }}
                                    >
                                        <TableRow>
                                            <TableCell
                                                sx={{ width: 60, textAlign: "center" }}
                                            >
                                                S.No
                                            </TableCell>
                                            <TableCell
                                                sx={{ minWidth: 260, cursor: "pointer" }}
                                                onClick={(e) => {
                                                    setFilterAnchor(e.currentTarget);
                                                    setActiveFilterColumn("itemName");
                                                }}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="space-between">
                                                    <span>Item Name</span>
                                                    {selectedItemNames !== undefined && (
                                                        <FilterAltIcon
                                                            fontSize="small"
                                                            sx={{ color: "#FDE047", ml: 0.5, fontSize: 16 }}
                                                        />
                                                    )}
                                                </Box>
                                            </TableCell>
                                            <TableCell
                                                sx={{ minWidth: 140, cursor: "pointer" }}
                                                onClick={(e) => {
                                                    setFilterAnchor(e.currentTarget);
                                                    setActiveFilterColumn("batch");
                                                }}
                                            >
                                                <Box display="flex" alignItems="center" justifyContent="space-between">
                                                    <span>Batch</span>
                                                    {selectedBatches !== undefined && (
                                                        <FilterAltIcon
                                                            fontSize="small"
                                                            sx={{ color: "#FDE047", ml: 0.5, fontSize: 16 }}
                                                        />
                                                    )}
                                                </Box>
                                            </TableCell>
                                            <TableCell
                                                align="right"
                                                sx={{ minWidth: 100 }}
                                            >
                                                OB
                                            </TableCell>
                                            <TableCell
                                                align="right"
                                                sx={{ minWidth: 100 }}
                                            >
                                                IN
                                            </TableCell>
                                            <TableCell
                                                align="right"
                                                sx={{ minWidth: 100 }}
                                            >
                                                OUT
                                            </TableCell>
                                            <TableCell
                                                align="right"
                                                sx={{ minWidth: 110 }}
                                            >
                                                CLS
                                            </TableCell>
                                        </TableRow>

                                        {/* TOTAL ROW */}
                                        <TableRow
                                            sx={{
                                                background: "#f3f4f6",
                                                position: "sticky",
                                                top: 36,
                                                zIndex: 2,
                                                "& .MuiTableCell-root": {
                                                    bgcolor: "#f3f4f6 !important",
                                                    color: "#0F172A !important",
                                                    fontWeight: 700,
                                                    fontSize: "0.75rem",
                                                    py: 0.8,
                                                },
                                            }}
                                        >
                                            <TableCell />
                                            <TableCell sx={{ fontWeight: 700 }}>
                                                TOTAL
                                            </TableCell>
                                            <TableCell />
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatNumber(totals.totalOB)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatNumber(totals.totalIN)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatNumber(totals.totalOUT)}
                                            </TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatNumber(totals.totalCLS)}
                                            </TableCell>
                                        </TableRow>
                                    </TableHead>

                                    <TableBody>
                                        {paginatedRows.length === 0 ? (
                                            <TableRow>
                                                <TableCell
                                                    colSpan={7}
                                                    align="center"
                                                    sx={{ py: 4, color: "#64748B" }}
                                                >
                                                    {isExpanded && !selectedGodownId
                                                        ? "Please select a Godown from Filter to get data"
                                                        : "No Records Found"}
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            paginatedRows.map((row, idx) => {
                                                const sNo =
                                                    (page - 1) * rowsPerPage + idx + 1;
                                                return (
                                                    <TableRow
                                                        key={row.id}
                                                        hover
                                                        sx={{
                                                            "&:hover": { bgcolor: "#f8fafc" },
                                                            "& .MuiTableCell-root": {
                                                                fontSize: "0.75rem",
                                                                py: 0.8,
                                                                borderRight: "1px solid #e2e8f0",
                                                            },
                                                        }}
                                                    >
                                                        <TableCell
                                                            align="center"
                                                            sx={{ color: "#64748B" }}
                                                        >
                                                            {sNo}
                                                        </TableCell>
                                                        <TableCell
                                                            onClick={() => handleOpenItemBatchModal(row)}
                                                            sx={{
                                                                fontWeight: 500,
                                                                color: "#1E3A8A",
                                                                cursor: "pointer",
                                                                "&:hover": { color: "#2563EB" },
                                                            }}
                                                        >
                                                            {row.itemName}
                                                        </TableCell>
                                                        <TableCell
                                                            onClick={() => handleOpenItemBatchModal(row)}
                                                            sx={{
                                                                color: "#334155",
                                                                cursor: "pointer",
                                                                "&:hover": { color: "#1E3A8A", fontWeight: 600 },
                                                            }}
                                                        >
                                                            {row.batch}
                                                        </TableCell>
                                                        <TableCell align="right">
                                                            {formatNumber(row.ob)}
                                                        </TableCell>
                                                        <TableCell
                                                            align="right"
                                                            sx={{
                                                                color:
                                                                    row.inQty > 0
                                                                        ? "#16A34A"
                                                                        : "inherit",
                                                                fontWeight:
                                                                    row.inQty > 0 ? 600 : 400,
                                                            }}
                                                        >
                                                            {formatNumber(row.inQty)}
                                                        </TableCell>
                                                        <TableCell
                                                            align="right"
                                                            sx={{
                                                                color:
                                                                    row.outQty > 0
                                                                        ? "#DC2626"
                                                                        : "inherit",
                                                                fontWeight:
                                                                    row.outQty > 0 ? 600 : 400,
                                                            }}
                                                        >
                                                            {formatNumber(row.outQty)}
                                                        </TableCell>
                                                        <TableCell
                                                            align="right"
                                                            sx={{
                                                                color: "#1E3A8A",
                                                                fontWeight: 700,
                                                            }}
                                                        >
                                                            {formatNumber(row.cls)}
                                                        </TableCell>
                                                    </TableRow>
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

            {/* Header Filter Menu for Item Name */}
            <HeaderFilterMenu
                anchorEl={activeFilterColumn === "itemName" ? filterAnchor : null}
                open={Boolean(filterAnchor) && activeFilterColumn === "itemName"}
                onClose={() => {
                    setFilterAnchor(null);
                    setActiveFilterColumn(null);
                }}
                columnLabel="Item Name"
                options={itemNameOptions}
                selectedValues={selectedItemNames}
                onFilterChange={(newSelected) => {
                    setSelectedItemNames(newSelected);
                    setPage(1);
                }}
            />

            {/* Header Filter Menu for Batch */}
            <HeaderFilterMenu
                anchorEl={activeFilterColumn === "batch" ? filterAnchor : null}
                open={Boolean(filterAnchor) && activeFilterColumn === "batch"}
                onClose={() => {
                    setFilterAnchor(null);
                    setActiveFilterColumn(null);
                }}
                columnLabel="Batch"
                options={batchOptions}
                selectedValues={selectedBatches}
                onFilterChange={(newSelected) => {
                    setSelectedBatches(newSelected);
                    setPage(1);
                }}
            />

            {/* ================= BATCH TRANSACTION DETAILS POP-UP MODAL ================= */}
            <Dialog
                open={batchModalOpen}
                onClose={() => {
                    setBatchModalOpen(false);
                    setBatchDateAnchorEl(null);
                }}
                maxWidth="md"
                fullWidth
                PaperProps={{
                    sx: {
                        borderRadius: 3,
                        boxShadow: "0 20px 45px rgba(0,0,0,0.3)",
                        overflow: "hidden",
                    },
                }}
            >
                {/* Header with Item Name, Batch No, and Date Filter */}
                <DialogTitle
                    sx={{
                        m: 0,
                        p: 2.2,
                        bgcolor: "#1E3A8A",
                        color: "#FFFFFF",
                        display: "flex",
                        alignItems: "flex-start",
                        justifyContent: "space-between",
                        borderBottom: "1px solid rgba(255, 255, 255, 0.12)",
                    }}
                >
                    <Box sx={{ pr: 2 }}>
                        <Box sx={{ display: "flex", alignItems: "center", gap: 1.2, flexWrap: "wrap" }}>
                            <Typography
                                variant="h6"
                                component="span"
                                sx={{
                                    fontWeight: 700,
                                    fontSize: "1.15rem",
                                    color: "#FFFFFF",
                                    letterSpacing: "0.2px",
                                }}
                            >
                                {selectedBatchItem?.productName || "Item Details"}
                            </Typography>

                            <Chip
                                label={`Batch: ${selectedBatchItem?.batchNo ?? "Unassigned"}`}
                                size="small"
                                sx={{
                                    bgcolor: "rgba(255, 255, 255, 0.2)",
                                    color: "#FFFFFF",
                                    fontWeight: 700,
                                    fontSize: "0.78rem",
                                    border: "1px solid rgba(255, 255, 255, 0.35)",
                                    px: 0.5,
                                }}
                            />

                            {selectedBatchItem?.godownName && (
                                <Chip
                                    label={`Godown: ${selectedBatchItem.godownName}`}
                                    size="small"
                                    sx={{
                                        bgcolor: "rgba(255, 255, 255, 0.2)",
                                        color: "#FFFFFF",
                                        fontWeight: 600,
                                        fontSize: "0.72rem",
                                        border: "1px solid rgba(255, 255, 255, 0.25)",
                                    }}
                                />
                            )}

                            <Chip
                                label={qtyMode === "actQty" ? "Mode: Act_qty" : "Mode: Qty"}
                                size="small"
                                sx={{
                                    bgcolor: qtyMode === "actQty" ? "#7C3AED" : "rgba(255, 255, 255, 0.15)",
                                    color: "#FFFFFF",
                                    fontWeight: 600,
                                    fontSize: "0.72rem",
                                    border: "1px solid rgba(255, 255, 255, 0.2)",
                                }}
                            />
                        </Box>

                        <Box sx={{ display: "flex", alignItems: "center", gap: 1.5, mt: 0.8, flexWrap: "wrap" }}>
                            {/* Interactive Date Filter Badge */}
                            <Tooltip title="Click to change From & To Date for this Batch" arrow>
                                <Box
                                    component="button"
                                    type="button"
                                    onClick={(e) => {
                                        setTempBatchFromDate(batchFromDate);
                                        setTempBatchToDate(batchToDate);
                                        setBatchDateAnchorEl(e.currentTarget);
                                    }}
                                    sx={{
                                        display: "inline-flex",
                                        alignItems: "center",
                                        gap: 0.8,
                                        px: 1.2,
                                        py: 0.45,
                                        borderRadius: "6px",
                                        bgcolor: isBatchDateCustom ? "rgba(245, 158, 11, 0.28)" : "rgba(255, 255, 255, 0.16)",
                                        border: isBatchDateCustom ? "1.5px solid #F59E0B" : "1px solid rgba(255, 255, 255, 0.35)",
                                        color: "#FFFFFF",
                                        cursor: "pointer",
                                        outline: "none",
                                        boxShadow: isBatchDateCustom ? "0 0 12px rgba(245, 158, 11, 0.35)" : "none",
                                        transition: "all 0.2s ease-in-out",
                                        "&:hover": {
                                            bgcolor: isBatchDateCustom ? "rgba(245, 158, 11, 0.4)" : "rgba(255, 255, 255, 0.28)",
                                            borderColor: "#FFFFFF",
                                            transform: "translateY(-1px)",
                                            boxShadow: "0 3px 8px rgba(0,0,0,0.25)",
                                        },
                                    }}
                                >
                                    <CalendarMonthIcon sx={{ fontSize: 16, color: isBatchDateCustom ? "#FDE047" : "#93C5FD" }} />
                                    <Typography variant="caption" sx={{ color: "#FFFFFF", fontSize: "0.76rem", fontWeight: 500 }}>
                                        From: <strong>{dayjs(batchFromDate).format("DD/MM/YYYY")}</strong> | To:{" "}
                                        <strong>{dayjs(batchToDate).format("DD/MM/YYYY")}</strong>
                                    </Typography>
                                    {isBatchDateCustom && (
                                        <Chip
                                            label="Custom Date"
                                            size="small"
                                            sx={{
                                                height: 18,
                                                fontSize: "0.65rem",
                                                fontWeight: 700,
                                                bgcolor: "#F59E0B",
                                                color: "#1F2937",
                                                lineHeight: 1,
                                                px: 0.3,
                                            }}
                                        />
                                    )}
                                    <Chip
                                        label="Change Date"
                                        size="small"
                                        sx={{
                                            height: 18,
                                            fontSize: "0.65rem",
                                            fontWeight: 600,
                                            bgcolor: "rgba(255, 255, 255, 0.22)",
                                            color: "#FFFFFF",
                                            lineHeight: 1,
                                            px: 0.4,
                                        }}
                                    />
                                </Box>
                            </Tooltip>

                            <Typography variant="caption" sx={{ color: "rgba(255, 255, 255, 0.85)", fontSize: "0.75rem" }}>
                                Total Entries: <strong>{processedBatchRows.length}</strong>
                            </Typography>
                        </Box>
                    </Box>

                    <IconButton
                        aria-label="close"
                        onClick={() => {
                            setBatchModalOpen(false);
                            setBatchDateAnchorEl(null);
                        }}
                        sx={{
                            color: "#FFFFFF",
                            bgcolor: "rgba(255, 255, 255, 0.08)",
                            "&:hover": { bgcolor: "rgba(255, 255, 255, 0.2)" },
                            p: 0.8,
                        }}
                    >
                        <CloseIcon sx={{ fontSize: 20 }} />
                    </IconButton>
                </DialogTitle>

                {/* Batch-specific Date Filter Popover */}
                <Popover
                    open={Boolean(batchDateAnchorEl)}
                    anchorEl={batchDateAnchorEl}
                    onClose={() => setBatchDateAnchorEl(null)}
                    anchorOrigin={{ vertical: "bottom", horizontal: "left" }}
                    transformOrigin={{ vertical: "top", horizontal: "left" }}
                    PaperProps={{
                        sx: {
                            mt: 1,
                            p: 2.2,
                            width: 340,
                            borderRadius: 2.5,
                            boxShadow: "0 12px 30px rgba(0, 0, 0, 0.25)",
                            border: "1px solid #E2E8F0",
                        },
                    }}
                    sx={{ zIndex: 1400 }}
                >
                    <Box sx={{ display: "flex", flexDirection: "column", gap: 1.8 }}>
                        <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                                <CalendarMonthIcon sx={{ color: "#1E3A8A", fontSize: 20 }} />
                                <Typography sx={{ fontSize: "0.95rem", fontWeight: 700, color: "#1E293B" }}>
                                    Batch Date Filter
                                </Typography>
                            </Box>
                            {isBatchDateCustom && (
                                <Chip
                                    label="Custom Active"
                                    size="small"
                                    sx={{
                                        height: 20,
                                        fontSize: "0.68rem",
                                        fontWeight: 700,
                                        bgcolor: "#FEF3C7",
                                        color: "#B45309",
                                    }}
                                />
                            )}
                        </Box>

                        <Box sx={{ display: "flex", gap: 1.5 }}>
                            <TextField
                                label="From Date"
                                type="date"
                                size="small"
                                fullWidth
                                value={tempBatchFromDate}
                                onChange={(e) => setTempBatchFromDate(e.target.value)}
                                InputLabelProps={{ shrink: true }}
                                inputProps={{ max: tempBatchToDate || undefined }}
                            />
                            <TextField
                                label="To Date"
                                type="date"
                                size="small"
                                fullWidth
                                value={tempBatchToDate}
                                onChange={(e) => setTempBatchToDate(e.target.value)}
                                InputLabelProps={{ shrink: true }}
                                inputProps={{ min: tempBatchFromDate || undefined }}
                            />
                        </Box>

                        {dayjs(tempBatchFromDate).isAfter(dayjs(tempBatchToDate)) && (
                            <Typography sx={{ color: "#DC2626", fontSize: "0.75rem", fontWeight: 500 }}>
                                * From Date cannot be after To Date.
                            </Typography>
                        )}

                        <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", pt: 1, borderTop: "1px solid #E2E8F0" }}>
                            <Button
                                size="small"
                                color="inherit"
                                onClick={() => {
                                    setBatchFromDate(fromDate);
                                    setBatchToDate(toDate);
                                    setTempBatchFromDate(fromDate);
                                    setTempBatchToDate(toDate);
                                    setIsBatchDateCustom(false);
                                    setBatchDateAnchorEl(null);
                                }}
                                sx={{ textTransform: "none", fontSize: "0.75rem", color: "#64748B" }}
                            >
                                Reset to Report
                            </Button>

                            <Box sx={{ display: "flex", gap: 1 }}>
                                <Button
                                    size="small"
                                    variant="outlined"
                                    onClick={() => setBatchDateAnchorEl(null)}
                                    sx={{ textTransform: "none", fontSize: "0.75rem" }}
                                >
                                    Cancel
                                </Button>
                                <Button
                                    size="small"
                                    variant="contained"
                                    disabled={
                                        !tempBatchFromDate ||
                                        !tempBatchToDate ||
                                        dayjs(tempBatchFromDate).isAfter(dayjs(tempBatchToDate))
                                    }
                                    onClick={() => {
                                        setBatchFromDate(tempBatchFromDate);
                                        setBatchToDate(tempBatchToDate);
                                        setIsBatchDateCustom(tempBatchFromDate !== fromDate || tempBatchToDate !== toDate);
                                        setBatchDateAnchorEl(null);
                                    }}
                                    sx={{
                                        textTransform: "none",
                                        fontSize: "0.75rem",
                                        bgcolor: "#1E3A8A",
                                        "&:hover": { bgcolor: "#1E40AF" },
                                    }}
                                >
                                    Apply Filter
                                </Button>
                            </Box>
                        </Box>
                    </Box>
                </Popover>

                {/* Dialog Content: Table with Vocuher No, Date, Particulars, etc. matching screenshot */}
                <DialogContent sx={{ p: 0, bgcolor: "#FFFFFF", overflowX: "hidden" }}>
                    {batchModalLoading ? (
                        <Box
                            sx={{
                                display: "flex",
                                flexDirection: "column",
                                justifyContent: "center",
                                alignItems: "center",
                                py: 8,
                                gap: 1.5,
                            }}
                        >
                            <CircularProgress size={36} sx={{ color: "#1E3A8A" }} />
                            <Typography sx={{ color: "#64748B", fontSize: "0.85rem", fontWeight: 500 }}>
                                Loading batch transactions...
                            </Typography>
                        </Box>
                    ) : batchModalError ? (
                        <Box sx={{ textAlign: "center", py: 5, px: 2 }}>
                            <Typography sx={{ color: "#DC2626", fontWeight: 600, mb: 1 }}>{batchModalError}</Typography>
                            <Button
                                variant="outlined"
                                size="small"
                                onClick={() => {
                                    if (selectedBatchItem) {
                                        loadBatchTransactions(selectedBatchItem, batchFromDate, batchToDate);
                                    }
                                }}
                            >
                                Retry
                            </Button>
                        </Box>
                    ) : processedBatchRows.length === 0 ? (
                        <Box sx={{ textAlign: "center", py: 8 }}>
                            <Typography sx={{ color: "#64748B", fontSize: "0.9rem", fontWeight: 500 }}>
                                No transaction records found for Batch &ldquo;{selectedBatchItem?.batchNo}&rdquo; within the selected date range.
                            </Typography>
                        </Box>
                    ) : (
                        <TableContainer
                            sx={{
                                maxHeight: "65vh",
                                overflowY: "auto",
                                overflowX: "hidden",
                            }}
                        >
                            <Table size="small" stickyHeader sx={{ width: "100%", tableLayout: "auto" }}>
                                <TableHead>
                                    {/* ROW 1: COLUMN HEADERS */}
                                    <TableRow sx={{ height: 32 }}>
                                        <TableCell
                                            align="center"
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                width: 48,
                                                py: 0.4,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            S.NO
                                        </TableCell>

                                        <TableCell
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                width: 95,
                                                py: 0.4,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            Date
                                        </TableCell>

                                        <TableCell
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                width: 130,
                                                py: 0.4,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            Vocuher No
                                        </TableCell>

                                        <TableCell
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                whiteSpace: "nowrap",
                                                zIndex: 3,
                                            }}
                                        >
                                            Particulars
                                        </TableCell>

                                        <TableCell
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                width: 130,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            Voucher Type
                                        </TableCell>

                                        <TableCell
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                width: 160,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            Retailer
                                        </TableCell>

                                        <TableCell
                                            align="right"
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                width: 120,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            IN
                                        </TableCell>

                                        <TableCell
                                            align="right"
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                width: 120,
                                                whiteSpace: "nowrap",
                                                borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                                                zIndex: 3,
                                            }}
                                        >
                                            OUT
                                        </TableCell>

                                        <TableCell
                                            align="right"
                                            sx={{
                                                bgcolor: "#1E3A8A !important",
                                                color: "#FFFFFF !important",
                                                fontWeight: 700,
                                                fontSize: "0.76rem",
                                                py: 0.4,
                                                width: 125,
                                                whiteSpace: "nowrap",
                                                zIndex: 3,
                                            }}
                                        >
                                            CLS
                                        </TableCell>
                                    </TableRow>

                                    {/* ROW 2: FIXED TOTAL ROW */}
                                    <TableRow
                                        sx={{
                                            height: 32,
                                            "& .MuiTableCell-root": {
                                                bgcolor: "#F1F5F9 !important",
                                                position: "sticky",
                                                top: "32px !important",
                                                zIndex: 2,
                                                py: 0.35,
                                                whiteSpace: "nowrap",
                                                borderBottom: "2px solid #CBD5E1",
                                            },
                                        }}
                                    >
                                        <TableCell align="center" />
                                        <TableCell />
                                        <TableCell />
                                        <TableCell sx={{ fontWeight: 700, color: "#0F172A", fontSize: "0.76rem", whiteSpace: "nowrap" }}>
                                            TOTAL
                                        </TableCell>
                                        <TableCell />
                                        <TableCell />
                                        <TableCell
                                            align="right"
                                            sx={{
                                                fontWeight: 700,
                                                color: "#16A34A",
                                                fontSize: "0.76rem",
                                                whiteSpace: "nowrap",
                                            }}
                                        >
                                            {formatNumber(batchTotals.totalIn)}
                                        </TableCell>
                                        <TableCell
                                            align="right"
                                            sx={{
                                                fontWeight: 700,
                                                color: "#DC2626",
                                                fontSize: "0.76rem",
                                                whiteSpace: "nowrap",
                                            }}
                                        >
                                            {formatNumber(batchTotals.totalOut)}
                                        </TableCell>
                                        <TableCell
                                            align="right"
                                            sx={{
                                                fontWeight: 700,
                                                color: "#1D4ED8",
                                                fontSize: "0.78rem",
                                                whiteSpace: "nowrap",
                                            }}
                                        >
                                            {formatNumber(batchTotals.finalCls)}
                                        </TableCell>
                                    </TableRow>
                                </TableHead>

                                <TableBody>
                                    {processedBatchRows.map((row) => {
                                        const isOB =
                                            row.invoice_no === "OB" ||
                                            String(row.Particulars || "").toLowerCase().includes("opening balance");

                                        return (
                                            <TableRow
                                                key={`${row.Trans_Id || row.sNo}-${row.sNo}`}
                                                hover
                                                sx={{
                                                    height: 32,
                                                    bgcolor: isOB ? "#F8FAFC" : row.sNo % 2 === 0 ? "#FAFAFA" : "#FFFFFF",
                                                    "&:hover": { bgcolor: "#EFF6FF" },
                                                    borderBottom: "1px solid #F1F5F9",
                                                }}
                                            >
                                                <TableCell
                                                    align="center"
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        color: "#64748B",
                                                        fontWeight: 500,
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {row.sNo}
                                                </TableCell>

                                                <TableCell
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        color: "#475569",
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {row.Ledger_Date ? dayjs(row.Ledger_Date).format("DD/MM/YYYY") : "—"}
                                                </TableCell>

                                                <TableCell
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        color: "#595a5cff",
                                                        fontWeight: row.invoice_no === "OB" ? 500 : 600,
                                                        whiteSpace: "nowrap",
                                                        overflow: "hidden",
                                                        textOverflow: "ellipsis",
                                                        maxWidth: 140,
                                                    }}
                                                    title={row.invoice_no || ""}
                                                >
                                                    {row.invoice_no || "—"}
                                                </TableCell>

                                                <TableCell
                                                    sx={{
                                                        py: 0.35,
                                                        whiteSpace: "nowrap",
                                                        overflow: "hidden",
                                                        textOverflow: "ellipsis",
                                                        maxWidth: 240,
                                                    }}
                                                    title={row.Particulars || (isOB ? "Opening Balance" : "—")}
                                                >
                                                    <Typography
                                                        component="span"
                                                        sx={{
                                                            fontSize: "0.75rem",
                                                            fontWeight: isOB ? 700 : 500,
                                                            color: isOB ? "#0F172A" : "#1E293B",
                                                            whiteSpace: "nowrap",
                                                        }}
                                                    >
                                                        {row.Particulars || (isOB ? "Opening Balance" : "—")}
                                                    </Typography>
                                                </TableCell>

                                                <TableCell
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        color: "#334155",
                                                        whiteSpace: "nowrap",
                                                        overflow: "hidden",
                                                        textOverflow: "ellipsis",
                                                        maxWidth: 140,
                                                    }}
                                                    title={row.voucher_name || ""}
                                                >
                                                    {row.voucher_name || "—"}
                                                </TableCell>

                                                <TableCell
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        color: "#334155",
                                                        whiteSpace: "nowrap",
                                                        overflow: "hidden",
                                                        textOverflow: "ellipsis",
                                                        maxWidth: 170,
                                                    }}
                                                    title={row.Retailer_Name || ""}
                                                >
                                                    {row.Retailer_Name || "—"}
                                                </TableCell>

                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        fontWeight: row.inQtyVal > 0 ? 600 : 400,
                                                        color: row.inQtyVal > 0 ? "#16A34A" : "#94A3B8",
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {formatNumber(row.inQtyVal)}
                                                </TableCell>

                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        fontWeight: row.outQtyVal > 0 ? 600 : 400,
                                                        color: row.outQtyVal > 0 ? "#DC2626" : "#94A3B8",
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {formatNumber(row.outQtyVal)}
                                                </TableCell>

                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        py: 0.35,
                                                        fontSize: "0.75rem",
                                                        fontWeight: 700,
                                                        color: "#1D4ED8",
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {formatNumber(row.clsQtyVal)}
                                                </TableCell>
                                            </TableRow>
                                        );
                                    })}
                                </TableBody>
                            </Table>
                        </TableContainer>
                    )}
                </DialogContent>
            </Dialog>
        </>
    );
};

export default BatchwiseStockReport;
