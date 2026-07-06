import React, { useEffect, useState, useMemo } from "react";
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
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    CircularProgress,
    Checkbox,
} from "@mui/material";
import SyncIcon from "@mui/icons-material/Sync";
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


/* ================= STYLES ================= */
const headStyle = {
    color: "#fff",
    fontWeight: 600,
    fontSize: "0.75rem",
    cursor: "pointer",
};

const bodyStyle = {
    fontSize: "0.68rem",
};

const UnitEconomicsAdmin: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    const [data, setData] = useState<AdminUnitEconomicsReport[]>([]);
    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    /* -------- FILTERS -------- */
    const [filters, setFilters] = useState({
        Date: { from: today, to: today },
        Product: [] as string[],
        invoice_no: [] as string[],
        Retailer_Name: [] as string[],
    });

    const [tempDate, setTempDate] = useState(filters.Date);

    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [rateType, setRateType] = useState<"cogs" | "min">("cogs");

    /* -------- HEADER FILTER -------- */
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);
    const [searchText, setSearchText] = useState("");

    /* -------- SUMMARY -------- */
    const [summaryColumn, setSummaryColumn] = useState<keyof AdminUnitEconomicsReport | null>(null);

    /* -------- SYNC DIALOG STATE -------- */
    const [syncDialogOpen, setSyncDialogOpen] = useState(false);
    const [syncFromDate, setSyncFromDate] = useState(today);
    const [syncToDate, setSyncToDate] = useState(today);
    const [isSyncing, setIsSyncing] = useState(false);

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

                // Set the page filters to the synced dates, triggering a reload
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

    const exportColumns = useMemo(() => {
        const baseCols = [
            { label: "S.No", key: "sno" },
            { label: "Invoice No", key: "invoice_no" },
            { label: "Date", key: "Ledger_Date", type: "date" },
            { label: "Retailer Name", key: "Retailer_Name" },
            { label: "Product Name", key: "Product_Name" },
            { label: "Quantity", key: "Bill_Qty", type: "number" },
            { label: "Rate", key: "Rate", type: "number" },
            { label: "Amount", key: "Amount", type: "number" },
        ];

        if (rateType === "cogs") {
            return [
                ...baseCols,
                { label: "COGS Rate", key: "COGS_Rate", type: "number" },
                { label: "GP COGS", key: "GP_COGS", type: "number" },
                { label: "TGP COGS", key: "TGP_COGS", type: "number" },
                { label: "GP % COGS", key: "GP_Percentage_COGS", type: "number" },
            ];
        } else {
            return [
                ...baseCols,
                { label: "MR Rate", key: "Min_Rate", type: "number" },
                { label: "GP MR", key: "GP_MR", type: "number" },
                { label: "TGP MR", key: "TGP_MR", type: "number" },
                { label: "GP % MR", key: "GP_Percentage_MR", type: "number" },
            ];
        }
    }, [rateType]);

    /* ================= LOAD DATA ================= */
    const [rawApiData, setRawApiData] = useState<AdminUnitEconomicsReport[]>([]);

    useEffect(() => {
        const loadData = async () => {
            const res = await AdminUnitEconomicsReportService.getReports({
                Fromdate: filters.Date.from,
                Todate: filters.Date.to,
            });

            const responseData = res.data.data;
            let rows: AdminUnitEconomicsReport[] = responseData.rows || [];
            setRawApiData(rows);
            setPage(1);
            setSummaryColumn(null);
        };

        loadData();
    }, [filters.Date]);

    useEffect(() => {
        let rows = rawApiData;
        if (filters.Product && filters.Product.length > 0) {
            rows = rows.filter((r) => filters.Product.includes(r.Product_Name));
        }
        if (filters.invoice_no && filters.invoice_no.length > 0) {
            rows = rows.filter((r) => filters.invoice_no.includes(r.invoice_no));
        }
        if (filters.Retailer_Name && filters.Retailer_Name.length > 0) {
            rows = rows.filter((r) => filters.Retailer_Name.includes(r.Retailer_Name));
        }
        setData(rows);
        setPage(1);
    }, [rawApiData, filters.Product, filters.invoice_no, filters.Retailer_Name]);

    /* ================= PAGINATION ================= */
    const paginatedData = data.slice(
        (page - 1) * rowsPerPage,
        page * rowsPerPage
    );

    /* ================= DROPDOWNS ================= */
    const products = useMemo(
        () =>
            [...new Set(rawApiData.map((d) => d.Product_Name))]
                .filter(Boolean)
                .filter((p) =>
                    p.toLowerCase().includes(searchText.toLowerCase())
                ),
        [rawApiData, searchText]
    );

    const invoiceNos = useMemo(
        () =>
            [...new Set(rawApiData.map((d) => d.invoice_no))]
                .filter(Boolean)
                .filter((inv) =>
                    inv.toLowerCase().includes(searchText.toLowerCase())
                ),
        [rawApiData, searchText]
    );

    const retailerNames = useMemo(
        () =>
            [...new Set(rawApiData.map((d) => d.Retailer_Name))]
                .filter(Boolean)
                .filter((ret) =>
                    ret.toLowerCase().includes(searchText.toLowerCase())
                ),
        [rawApiData, searchText]
    );

    /* ================= SUMMARY ================= */
    const getTotal = (key: keyof AdminUnitEconomicsReport) => {
        return data.reduce((sum, r) => sum + (Number(r[key]) || 0), 0);
    };

    /* ================= HEADER CLICK ================= */
    const openFilter = (e: React.MouseEvent<HTMLElement>, column: string) => {
        setActiveHeader(column);
        setFilterAnchor(e.currentTarget);
        setSearchText("");

        if (["Bill_Qty", "Rate", "Amount", "Min_Rate", "COGS_Rate", "GP_MR", "GP_COGS", "TGP_MR", "TGP_COGS", "GP_Percentage_MR", "GP_Percentage_COGS"].includes(column)) {
            if (summaryColumn === column) {
                setSummaryColumn(null);
            } else {
                setSummaryColumn(column as keyof AdminUnitEconomicsReport);
            }
        }
    };

    const handleToggleItemDirect = (column: "Product" | "invoice_no" | "Retailer_Name", item: string) => {
        setFilters((prev) => {
            const list = prev[column] || [];
            const newList = list.includes(item)
                ? list.filter((x) => x !== item)
                : [...list, item];
            return {
                ...prev,
                [column]: newList,
            };
        });
    };

    const handleToggleSelectAllDirect = (column: "Product" | "invoice_no" | "Retailer_Name", visibleItems: string[]) => {
        setFilters((prev) => {
            const list = prev[column] || [];
            const allSelected = visibleItems.every((x) => list.includes(x));
            let newList: string[];
            if (allSelected) {
                newList = list.filter((x) => !visibleItems.includes(x));
            } else {
                newList = [...new Set([...list, ...visibleItems])];
            }
            return {
                ...prev,
                [column]: newList,
            };
        });
    };

    useEffect(() => {
        setSummaryColumn(null);
    }, [filters.Date, filters.Product, filters.invoice_no, filters.Retailer_Name]);



    /* ================= EXPORT ================= */
    const handleExportPDF = () => {
        const { headers, data: exportData } = mapForExport(exportColumns, data);
        exportToPDF("Unit Economics Report", headers, exportData);
    };

    const handleExportExcel = () => {
        const { headers, data: exportData } = mapForExport(exportColumns, data);
        exportToExcel("Unit Economics Report", headers, exportData);
    };

    const formatINR = (value: number) => {
        return new Intl.NumberFormat("en-IN", {
            style: "currency",
            currency: "INR",
        }).format(value);
    };


    /* ================= RENDER ================= */
    return (
        <>
            <PageHeader
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

            <AppLayout fullWidth >
                <Box sx={{ overflow: "auto", mt: 0.5 }}>
                    <TableContainer
                        component={Paper}
                        sx={{
                            position: 'relative',
                            maxHeight: "calc(100vh - 100px)",
                            overflow: "auto"
                        }}
                    >
                        <Table size="small">
                            {/* ===== FIXED HEADER ===== */}
                            <TableHead sx={{
                                background: "#1E3A8A",
                                position: "sticky",
                                top: 0,
                                zIndex: 2
                            }}>
                                <TableRow>
                                    <TableCell sx={headStyle}>S.No</TableCell>
                                    <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Invoice No")}>Invoice No</TableCell>
                                    <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Date")}>Date</TableCell>
                                    <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Retailer Name")}>Retailer Name</TableCell>
                                    <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Product")}>Product</TableCell>
                                    <TableCell align="right" sx={headStyle} > Quantity </TableCell>
                                    <TableCell align="right" sx={headStyle} >Rate</TableCell>
                                    <TableCell align="right" sx={headStyle} >Amount</TableCell>
                                    {rateType === "cogs" ? (
                                        <>
                                            <TableCell align="right" sx={headStyle} >COGS Rate</TableCell>
                                            <TableCell align="right" sx={headStyle} >GP COGS</TableCell>
                                            <TableCell align="right" sx={headStyle} >TGP COGS</TableCell>
                                            <TableCell align="right" sx={headStyle} >GP % COGS</TableCell>
                                        </>
                                    ) : (
                                        <>
                                            <TableCell align="right" sx={headStyle} >MR Rate</TableCell>
                                            <TableCell align="right" sx={headStyle} >GP MR</TableCell>
                                            <TableCell align="right" sx={headStyle} >TGP MR</TableCell>
                                            <TableCell align="right" sx={headStyle} >GP % MR</TableCell>
                                        </>
                                    )}
                                </TableRow>
                            </TableHead>

                            {/* ===== FIXED SUMMARY ROW ABOVE BODY ===== */}
                            <TableBody>
                                <TableRow
                                    sx={{
                                        background: "#f3f4f6",
                                        position: "sticky",
                                        top: 37,
                                        zIndex: 1,
                                    }}
                                >
                                    {/* S.No */}
                                    <TableCell sx={{ fontWeight: 700 }}>Total</TableCell>

                                    {/* Invoice No */}
                                    <TableCell />

                                    {/* Date */}
                                    <TableCell />

                                    {/* Retailer Name */}
                                    <TableCell />

                                    {/* Product */}
                                    <TableCell />

                                    {/* Quantity */}
                                    <TableCell align="right" sx={{ fontWeight: 700 }}>
                                        {getTotal("Bill_Qty").toFixed(2)}
                                    </TableCell>

                                    {/* Rate (empty) */}
                                    <TableCell />

                                    {/* Amount */}
                                    <TableCell align="right" sx={{ fontWeight: 700 }}>
                                        {formatINR(getTotal("Amount"))}
                                    </TableCell>

                                    {rateType === "cogs" ? (
                                        <>
                                            {/* COGS Rate (empty) */}
                                            <TableCell />
                                            {/* GP COGS (empty) */}
                                            <TableCell />
                                            {/* TGP COGS */}
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatINR(getTotal("TGP_COGS"))}
                                            </TableCell>
                                            {/* GP % COGS (empty) */}
                                            <TableCell />
                                        </>
                                    ) : (
                                        <>
                                            {/* MR Rate (empty) */}
                                            <TableCell />
                                            {/* GP MR (empty) */}
                                            <TableCell />
                                            {/* TGP MR */}
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                {formatINR(getTotal("TGP_MR"))}
                                            </TableCell>
                                            {/* GP % MR (empty) */}
                                            <TableCell />
                                        </>
                                    )}
                                </TableRow>
                            </TableBody>

                            {/* ===== BODY ===== */}
                            <TableBody>
                                {paginatedData.map((row, i) => (
                                    <TableRow key={`${row.invoice_no}-${i}`}>
                                        <TableCell sx={bodyStyle}>
                                            {(page - 1) * rowsPerPage + i + 1}
                                        </TableCell>
                                        <TableCell sx={bodyStyle}>{row.invoice_no}</TableCell>
                                        <TableCell sx={bodyStyle}>
                                            {dayjs(row.Ledger_Date).format("DD/MM/YYYY")}
                                        </TableCell>
                                        <TableCell sx={bodyStyle}>{row.Retailer_Name}</TableCell>
                                        <TableCell sx={bodyStyle}>{row.Product_Name}</TableCell>
                                        <TableCell align="right" sx={bodyStyle}>{Number(row.Bill_Qty).toFixed(2)}</TableCell>
                                        <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.Rate))}</TableCell>
                                        <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.Amount))}</TableCell>
                                        {rateType === "cogs" ? (
                                            <>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.COGS_Rate))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.GP_COGS))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.TGP_COGS))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{Number(row.GP_Percentage_COGS).toFixed(2)}%</TableCell>
                                            </>
                                        ) : (
                                            <>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.Min_Rate))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.GP_MR))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{formatINR(Number(row.TGP_MR))}</TableCell>
                                                <TableCell align="right" sx={bodyStyle}>{Number(row.GP_Percentage_MR).toFixed(2)}%</TableCell>
                                            </>
                                        )}
                                    </TableRow>
                                ))}
                            </TableBody>
                        </Table>
                    </TableContainer>

                    {/* ===== FILTER MENU (SAME PATTERN AS REFERENCE) ===== */}
                    <Menu
                        anchorEl={filterAnchor}
                        open={Boolean(filterAnchor)}
                        onClose={() => setFilterAnchor(null)}
                    >
                        {activeHeader === "Product" && (
                            <Box p={2} sx={{ minWidth: 240 }}>
                                <TextField
                                    size="small"
                                    fullWidth
                                    placeholder="Search Product"
                                    value={searchText}
                                    onChange={(e) => setSearchText(e.target.value)}
                                    sx={{ mb: 1 }}
                                />

                                <Box sx={{ maxHeight: 200, overflowY: 'auto' }}>
                                    <MenuItem
                                        sx={{ fontWeight: 600 }}
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            handleToggleSelectAllDirect("Product", products);
                                        }}
                                    >
                                        <Checkbox
                                            size="small"
                                            checked={products.length > 0 && products.every((x) => filters.Product.includes(x))}
                                            indeterminate={products.some((x) => filters.Product.includes(x)) && !products.every((x) => filters.Product.includes(x))}
                                        />
                                        Select All
                                    </MenuItem>

                                    {products.map((p) => (
                                        <MenuItem
                                            key={p}
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                handleToggleItemDirect("Product", p);
                                            }}
                                        >
                                            <Checkbox
                                                size="small"
                                                checked={filters.Product.includes(p)}
                                            />
                                            {p}
                                        </MenuItem>
                                    ))}
                                </Box>
                            </Box>
                        )}

                        {activeHeader === "Invoice No" && (
                            <Box p={2} sx={{ minWidth: 240 }}>
                                <TextField
                                    size="small"
                                    fullWidth
                                    placeholder="Search Invoice No"
                                    value={searchText}
                                    onChange={(e) => setSearchText(e.target.value)}
                                    sx={{ mb: 1 }}
                                />

                                <Box sx={{ maxHeight: 200, overflowY: 'auto' }}>
                                    <MenuItem
                                        sx={{ fontWeight: 600 }}
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            handleToggleSelectAllDirect("invoice_no", invoiceNos);
                                        }}
                                    >
                                        <Checkbox
                                            size="small"
                                            checked={invoiceNos.length > 0 && invoiceNos.every((x) => filters.invoice_no.includes(x))}
                                            indeterminate={invoiceNos.some((x) => filters.invoice_no.includes(x)) && !invoiceNos.every((x) => filters.invoice_no.includes(x))}
                                        />
                                        Select All
                                    </MenuItem>

                                    {invoiceNos.map((inv) => (
                                        <MenuItem
                                            key={inv}
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                handleToggleItemDirect("invoice_no", inv);
                                            }}
                                        >
                                            <Checkbox
                                                size="small"
                                                checked={filters.invoice_no.includes(inv)}
                                            />
                                            {inv}
                                        </MenuItem>
                                    ))}
                                </Box>
                            </Box>
                        )}

                        {activeHeader === "Retailer Name" && (
                            <Box p={2} sx={{ minWidth: 240 }}>
                                <TextField
                                    size="small"
                                    fullWidth
                                    placeholder="Search Retailer"
                                    value={searchText}
                                    onChange={(e) => setSearchText(e.target.value)}
                                    sx={{ mb: 1 }}
                                />

                                <Box sx={{ maxHeight: 200, overflowY: 'auto' }}>
                                    <MenuItem
                                        sx={{ fontWeight: 600 }}
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            handleToggleSelectAllDirect("Retailer_Name", retailerNames);
                                        }}
                                    >
                                        <Checkbox
                                            size="small"
                                            checked={retailerNames.length > 0 && retailerNames.every((x) => filters.Retailer_Name.includes(x))}
                                            indeterminate={retailerNames.some((x) => filters.Retailer_Name.includes(x)) && !retailerNames.every((x) => filters.Retailer_Name.includes(x))}
                                        />
                                        Select All
                                    </MenuItem>

                                    {retailerNames.map((ret) => (
                                        <MenuItem
                                            key={ret}
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                handleToggleItemDirect("Retailer_Name", ret);
                                            }}
                                        >
                                            <Checkbox
                                                size="small"
                                                checked={filters.Retailer_Name.includes(ret)}
                                            />
                                            {ret}
                                        </MenuItem>
                                    ))}
                                </Box>
                            </Box>
                        )}

                        {activeHeader === "Date" && (
                            <Box p={2} display="flex" flexDirection="column" gap={1}>
                                <TextField
                                    type="date"
                                    value={tempDate.from}
                                    onChange={(e) =>
                                        setTempDate((p) => ({ ...p, from: e.target.value }))
                                    }
                                />
                                <TextField
                                    type="date"
                                    value={tempDate.to}
                                    onChange={(e) =>
                                        setTempDate((p) => ({ ...p, to: e.target.value }))
                                    }
                                />
                                <Button
                                    variant="contained"
                                    onClick={() => {
                                        setFilters((p) => ({ ...p, Date: tempDate }));
                                        setFilterAnchor(null);
                                    }}
                                    sx={{
                                        backgroundColor: "#1E3A8A",
                                        fontWeight: 600,
                                    }}
                                >
                                    Apply
                                </Button>
                            </Box>
                        )}

                    </Menu>
                </Box>
                <CommonPagination
                    totalRows={data.length}
                    page={page}
                    rowsPerPage={rowsPerPage}
                    onPageChange={setPage}
                    onRowsPerPageChange={setRowsPerPage}
                />
            </AppLayout>

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
