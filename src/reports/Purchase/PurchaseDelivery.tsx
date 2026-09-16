/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
import React, { useState, useMemo } from "react";
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
    IconButton,
    Collapse,
    Chip,
    Button,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Grid,
    TextField,
    Tooltip,
    InputAdornment,
    MenuItem,
} from "@mui/material";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import LocalShippingIcon from "@mui/icons-material/LocalShipping";
import CloseIcon from "@mui/icons-material/Close";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
import AccessTimeIcon from "@mui/icons-material/AccessTime";
import ReceiptLongIcon from "@mui/icons-material/ReceiptLong";
import SearchIcon from "@mui/icons-material/Search";
import ClearIcon from "@mui/icons-material/Clear";
import VisibilityIcon from "@mui/icons-material/Visibility";
import ErrorOutlineIcon from "@mui/icons-material/ErrorOutline";
import Inventory2OutlinedIcon from "@mui/icons-material/Inventory2Outlined";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import CommonPagination from "../../Components/CommonPagination";
import {
    PurchaseDeliveryFunnelItem,
    FunnelTripItem,
    FunnelInvoiceItem,
    DUMMY_PURCHASE_DELIVERY_DATA,
    getFunnelMetrics,
    getOrderItems,
} from "../../services/purchaseDelivery.service";

/* ================= HELPER FORMATTERS ================= */

const formatCurrency = (val: number): string => {
    return "₹ " + Number(val || 0).toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
};

/* ================= COMPONENT ================= */

const PurchaseDelivery: React.FC = () => {
    // State
    const [data] = useState<PurchaseDeliveryFunnelItem[]>(DUMMY_PURCHASE_DELIVERY_DATA);
    const [searchQuery, setSearchQuery] = useState<string>("");
    const [activeStageFilter, setActiveStageFilter] = useState<string>("ALL");
    const [expandedRows, setExpandedRows] = useState<Record<string, boolean>>({
        "pd-1": false, // Default first row expanded for nice preview
    });

    // Filter Drawer State
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState<string>("");
    const [toDate, setToDate] = useState<string>("");
    const [filterStatus, setFilterStatus] = useState<string>("ALL");
    const [filterTripCount, setFilterTripCount] = useState<string>("ALL");

    // Pagination State
    const [page, setPage] = useState<number>(1);
    const [rowsPerPage, setRowsPerPage] = useState<number>(100);

    // Modal State for viewing trip details
    const [selectedTrip, setSelectedTrip] = useState<FunnelTripItem | null>(null);
    const [selectedInvoice, setSelectedInvoice] = useState<FunnelInvoiceItem | null>(null);
    const [selectedOrderDetails, setSelectedOrderDetails] = useState<PurchaseDeliveryFunnelItem | null>(null);

    // Toggle row expand
    const toggleRow = (id: string) => {
        setExpandedRows((prev) => ({
            ...prev,
            [id]: !prev[id],
        }));
    };

    // Filtered dataset
    const filteredData = useMemo(() => {
        return data.filter((item) => {
            // 1. Search Query (Matches Order ID, Retailer, City, Batch, Trip No, Invoice No)
            if (searchQuery.trim() !== "") {
                const q = searchQuery.toLowerCase().trim();
                const matchOrderId = item.orderId.toLowerCase().includes(q);
                const matchRetailer = item.retailerName.toLowerCase().includes(q);
                const matchCity = item.city.toLowerCase().includes(q);
                const matchBatch = item.batch.toLowerCase().includes(q);
                const matchTrips = item.trips.some(
                    (t) =>
                        t.tripNo.toLowerCase().includes(q) ||
                        t.vehicleNo.toLowerCase().includes(q) ||
                        t.driverName.toLowerCase().includes(q)
                );
                const matchInvoices = item.invoices.some((inv) =>
                    inv.invoiceNo.toLowerCase().includes(q)
                );

                if (
                    !matchOrderId &&
                    !matchRetailer &&
                    !matchCity &&
                    !matchBatch &&
                    !matchTrips &&
                    !matchInvoices
                ) {
                    return false;
                }
            }

            // 2. Stage Quick Filters
            if (activeStageFilter === "COMPLETED") {
                if (item.status !== "Completed") return false;
            } else if (activeStageFilter === "NOT_COMPLETED") {
                if (item.status !== "Not Completed") return false;
            } else if (activeStageFilter === "NO_TRIPS") {
                if (item.trips.length > 0) return false;
            } else if (activeStageFilter === "NO_INVOICE") {
                if (item.invoices.length > 0) return false;
            } else if (activeStageFilter === "BALANCE_DUE") {
                if (item.balanceToSettle <= 0) return false;
            } else if (activeStageFilter === "MULTI_TRIP") {
                if (item.trips.length < 2) return false;
            }

            // 3. Status filter from drawer
            if (filterStatus === "COMPLETED" && item.status !== "Completed") return false;
            if (filterStatus === "NOT_COMPLETED" && item.status !== "Not Completed") return false;

            // 4. Trip count filter from drawer
            if (filterTripCount === "NONE" && item.trips.length > 0) return false;
            if (filterTripCount === "SINGLE" && item.trips.length !== 1) return false;
            if (filterTripCount === "MULTIPLE" && item.trips.length <= 1) return false;

            // 5. Date filter (Order Date DD/MM/YYYY)
            if (fromDate || toDate) {
                const parts = item.orderDate.split("/");
                if (parts.length === 3) {
                    const itemDate = dayjs(`${parts[2]}-${parts[1]}-${parts[0]}`);
                    if (itemDate.isValid()) {
                        if (fromDate && itemDate.isBefore(dayjs(fromDate), "day")) return false;
                        if (toDate && itemDate.isAfter(dayjs(toDate), "day")) return false;
                    }
                }
            }

            return true;
        });
    }, [data, searchQuery, activeStageFilter, filterStatus, filterTripCount, fromDate, toDate]);

    // Overall Funnel Metrics across filtered data
    const metrics = useMemo(() => {
        return getFunnelMetrics(filteredData);
    }, [filteredData]);

    // Grand Totals for main table summary row
    const tableGrandTotals = useMemo(() => {
        let totalTrips = 0;
        let totalInvoices = 0;
        let totalBalance = 0;
        let grandTotalAmount = 0;

        filteredData.forEach((item) => {
            totalTrips += item.trips?.length || 0;
            totalInvoices += item.invoices?.length || 0;
            totalBalance += item.balanceToSettle;
            grandTotalAmount += item.totalAmount;
        });

        return {
            totalTrips,
            totalInvoices,
            totalBalance,
            grandTotalAmount,
        };
    }, [filteredData]);

    // Paginated Data
    const paginatedData = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        return filteredData.slice(start, start + rowsPerPage);
    }, [filteredData, page, rowsPerPage]);

    /* ================= EXPORT HANDLERS ================= */

    const handleExportExcel = () => {
        try {
            if (!filteredData.length) {
                toast.warning("No data to export");
                return;
            }

            const excelRows: any[][] = [];

            // Header Banner
            excelRows.push(["PURCHASE DELIVERY - FUNNEL TRACKING REPORT"]);
            excelRows.push([
                `Generated on: ${dayjs().format("DD-MM-YYYY HH:mm")} | Total Orders: ${filteredData.length
                } | Completed: ${metrics.completedOrders} | Not Completed: ${metrics.notCompletedOrders
                }`,
            ]);
            excelRows.push([]);

            // Table Header row (Matches Funnel columns)
            excelRows.push([
                "S.No",
                "Order Id",
                "Trip Count",
                "Invoice Count",
                "Payment (Balance to Settle)",
                "Total Amount",
                "Status",
                "Party / Retailer",
                "City",
                "Order Date",
            ]);

            // Data Rows
            filteredData.forEach((row, idx) => {
                excelRows.push([
                    idx + 1,
                    row.orderId,
                    row.trips.length,
                    row.invoices.length,
                    row.balanceToSettle,
                    row.totalAmount,
                    row.status,
                    row.retailerName,
                    row.city,
                    row.orderDate,
                ]);
            });

            // Totals summary row
            excelRows.push([
                "TOTAL",
                `${filteredData.length} Orders`,
                tableGrandTotals.totalTrips,
                tableGrandTotals.totalInvoices,
                tableGrandTotals.totalBalance,
                tableGrandTotals.grandTotalAmount,
                `${metrics.completionRate}% Completed`,
                "-",
                "-",
                "-",
            ]);

            const ws = XLSX.utils.aoa_to_sheet(excelRows);

            // Apply Excel styles
            if (ws && ws["!ref"]) {
                const range = XLSX.utils.decode_range(ws["!ref"]);
                const borderStyle = {
                    top: { style: "thin", color: { rgb: "CFCFCF" } },
                    bottom: { style: "thin", color: { rgb: "CFCFCF" } },
                    left: { style: "thin", color: { rgb: "CFCFCF" } },
                    right: { style: "thin", color: { rgb: "CFCFCF" } },
                };

                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
                        if (!cell) continue;

                        cell.s = {
                            font: { name: "Arial", sz: 10 },
                            border: borderStyle,
                        };

                        // Main Title
                        if (R === 0) {
                            cell.s.font = { name: "Arial", sz: 13, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                            cell.s.alignment = { horizontal: "center" };
                        }
                        // Subtitle
                        else if (R === 1) {
                            cell.s.font = { name: "Arial", sz: 9, italic: true, color: { rgb: "334155" } };
                            cell.s.fill = { fgColor: { rgb: "F1F5F9" } };
                        }
                        // Column Headers
                        else if (R === 3) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                            cell.s.alignment = { horizontal: "center", vertical: "center" };
                        }
                        // Summary Totals Row (last row)
                        else if (R === range.e.r) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "0F172A" } };
                            cell.s.fill = { fgColor: { rgb: "E2E8F0" } };
                        }
                        // Regular data rows
                        else if (R > 3) {
                            // Status column color styling (index 6)
                            if (C === 6) {
                                const isCompleted = cell.v === "Completed";
                                cell.s.font = {
                                    name: "Arial",
                                    sz: 10,
                                    bold: true,
                                    color: { rgb: isCompleted ? "15803D" : "B91C1C" },
                                };
                                cell.s.fill = { fgColor: { rgb: isCompleted ? "DCFCE7" : "FEE2E2" } };
                                cell.s.alignment = { horizontal: "center" };
                            }
                            // Payment balance column (index 4)
                            else if (C === 4) {
                                const bal = Number(cell.v || 0);
                                cell.s.font = {
                                    name: "Arial",
                                    sz: 10,
                                    bold: true,
                                    color: { rgb: bal > 0 ? "B45309" : "15803D" },
                                };
                            }
                        }
                    }
                }
            }

            // Column widths
            ws["!cols"] = [
                { wch: 6 },  // S.No
                { wch: 15 }, // Order Id
                { wch: 12 }, // Trip Count
                { wch: 14 }, // Invoice Count
                { wch: 22 }, // Balance to Settle
                { wch: 18 }, // Total Amount
                { wch: 16 }, // Status
                { wch: 26 }, // Party / Retailer
                { wch: 16 }, // City
                { wch: 14 }, // Order Date
            ];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "PurchaseFunnel");
            XLSX.writeFile(wb, `PurchaseDelivery_Funnel_Report_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
            toast.success("Excel report downloaded successfully! ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export Excel report ❌");
        }
    };

    const handleExportPDF = () => {
        try {
            if (!filteredData.length) {
                toast.warning("No data to export");
                return;
            }

            const doc = new jsPDF("landscape", "mm", "a4");

            // Title & Metadata
            doc.setFontSize(14);
            doc.setTextColor(30, 58, 138);
            doc.text("PURCHASE DELIVERY - FUNNEL TRACKING REPORT", 14, 15);

            doc.setFontSize(8.5);
            doc.setTextColor(100, 116, 139);
            doc.text(
                `Generated: ${dayjs().format("DD-MM-YYYY HH:mm")} | Total Orders: ${filteredData.length
                } | Completed: ${metrics.completedOrders} (${metrics.completionRate}%) | Balance Due: ${formatCurrency(
                    tableGrandTotals.totalBalance
                )}`,
                14,
                21
            );

            const tableHeaders = [
                [
                    "S.No",
                    "Order ID",
                    "Trip",
                    "Invoice",
                    "Payment (Balance)",
                    "Total Amount",
                    "Status",
                    "Party / Retailer",
                ],
            ];

            const tableBody = filteredData.map((row, idx) => [
                idx + 1,
                row.orderId,
                row.trips.length,
                row.invoices.length,
                formatCurrency(row.balanceToSettle),
                formatCurrency(row.totalAmount),
                row.status,
                `${row.retailerName} (${row.city})`,
            ]);

            // Summary totals row in PDF
            tableBody.push([
                "TOTAL",
                `${filteredData.length} Orders`,
                tableGrandTotals.totalTrips,
                tableGrandTotals.totalInvoices,
                formatCurrency(tableGrandTotals.totalBalance),
                formatCurrency(tableGrandTotals.grandTotalAmount),
                `${metrics.completionRate}% Done`,
                "-",
            ]);

            autoTable(doc, {
                head: tableHeaders,
                body: tableBody,
                startY: 26,
                theme: "grid",
                styles: { fontSize: 8, cellPadding: 2 },
                headStyles: {
                    fillColor: [30, 58, 138],
                    textColor: [255, 255, 255],
                    fontStyle: "bold",
                },
                alternateRowStyles: { fillColor: [248, 250, 252] },
                didParseCell: (cellData) => {
                    // Highlight Status column
                    if (cellData.section === "body" && cellData.column.index === 8) {
                        const val = String(cellData.cell.raw);
                        if (val === "Completed") {
                            cellData.cell.styles.textColor = [21, 128, 61]; // Green
                            cellData.cell.styles.fontStyle = "bold";
                        } else if (val === "Not Completed") {
                            cellData.cell.styles.textColor = [185, 28, 28]; // Red
                            cellData.cell.styles.fontStyle = "bold";
                        }
                    }
                    // Highlight totals row
                    if (
                        cellData.section === "body" &&
                        cellData.row.index === tableBody.length - 1
                    ) {
                        cellData.cell.styles.fillColor = [226, 232, 240];
                        cellData.cell.styles.fontStyle = "bold";
                    }
                },
            });

            doc.save(
                `PurchaseDelivery_Funnel_Report_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`
            );
            toast.success("PDF report downloaded successfully! ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export PDF report ❌");
        }
    };

    /* ================= RESET FILTERS ================= */

    const handleResetFilters = () => {
        setSearchQuery("");
        setActiveStageFilter("ALL");
        setFromDate("");
        setToDate("");
        setFilterStatus("ALL");
        setFilterTripCount("ALL");
        setPage(1);
        toast.info("Filters reset to default");
    };

    return (
        <Box
            sx={{
                width: "100%",
                height: "100%",
                maxHeight: "100%",
                display: "flex",
                flexDirection: "column",
                overflow: "hidden", // 🚫 Prevent whole screen scroll
                bgcolor: "#f8fafc",
            }}
        >
            {/* 1. TOP PAGE HEADER */}
            <PageHeader
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                showPages={true}
            />

            {/* 2. FILTER DRAWER */}
            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((prev) => !prev)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                onFromDateChange={setFromDate}
                toDate={toDate}
                onToDateChange={setToDate}
                onApply={() => {
                    setPage(1);
                    setDrawerOpen(false);
                    toast.success("Filters applied successfully!");
                }}
            >
                <Box sx={{ display: "flex", flexDirection: "column", gap: 2, mt: 1 }}>
                    <TextField
                        select
                        size="small"
                        label="Overall Status"
                        value={filterStatus}
                        onChange={(e) => setFilterStatus(e.target.value)}
                        fullWidth
                    >
                        <MenuItem value="ALL">All Statuses</MenuItem>
                        <MenuItem value="COMPLETED">Completed Only</MenuItem>
                        <MenuItem value="NOT_COMPLETED">Not Completed Only</MenuItem>
                    </TextField>

                    <TextField
                        select
                        size="small"
                        label="Trip Assignment"
                        value={filterTripCount}
                        onChange={(e) => setFilterTripCount(e.target.value)}
                        fullWidth
                    >
                        <MenuItem value="ALL">All Orders</MenuItem>
                        <MenuItem value="NONE">Unassigned (0 Trips)</MenuItem>
                        <MenuItem value="SINGLE">Single Trip (1 Trip)</MenuItem>
                        <MenuItem value="MULTIPLE">Multi-Trip (2+ Trips)</MenuItem>
                    </TextField>
                </Box>
            </ReportFilterDrawer>

            {/* 3. MAIN CONTENT CONTAINER (CONTAINS EVERYTHING INSIDE SCREEN) */}
            <Box
                sx={{
                    flex: 1,
                    minHeight: 0,
                    display: "flex",
                    flexDirection: "column",
                    overflow: "hidden", // 🚫 Keeps table inside screen with internal scroll
                    px: { xs: 1.5, md: 2.5 },
                    pt: 1,
                    pb: 1,
                    gap: 1,
                }}
            >
                {/* TITLE & QUICK SEARCH BAR */}
                <Box
                    sx={{
                        display: "flex",
                        flexDirection: { xs: "column", md: "row" },
                        justifyContent: "space-between",
                        alignItems: { xs: "flex-start", md: "center" },
                        gap: 1.2,
                        flexShrink: 0,
                    }}
                >
                    {/* Title & Badge */}
                    <Box sx={{ display: "flex", alignItems: "center", gap: 1.2, flexWrap: "wrap" }}>
                        <Typography
                            variant="h6"
                            sx={{
                                fontWeight: 800,
                                color: "#1e3a8a",
                                fontSize: { xs: "1.05rem", md: "1.15rem" },
                                letterSpacing: "-0.01em",
                            }}
                        >
                            Purchase Delivery Funnel Tracking
                        </Typography>

                        <Chip
                            size="small"
                            label={`${filteredData.length} Orders`}
                            sx={{
                                bgcolor: "#1e3a8a",
                                color: "#ffffff",
                                fontWeight: 700,
                                fontSize: "0.72rem",
                                height: 22,
                            }}
                        />

                        <Chip
                            size="small"
                            label={`${metrics.completionRate}% Completed`}
                            sx={{
                                bgcolor: metrics.completionRate > 50 ? "#dcfce7" : "#fef3c7",
                                color: metrics.completionRate > 50 ? "#15803d" : "#b45309",
                                fontWeight: 800,
                                fontSize: "0.70rem",
                                height: 22,
                                border: "1px solid",
                                borderColor: metrics.completionRate > 50 ? "#86efac" : "#fde68a",
                            }}
                        />
                    </Box>

                    {/* Search Input & Action Buttons */}
                    <Box
                        sx={{
                            display: "flex",
                            alignItems: "center",
                            gap: 1,
                            width: { xs: "100%", md: "auto" },
                            flexWrap: "wrap",
                        }}
                    >
                        <TextField
                            size="small"
                            placeholder="Search Order, Trip, Invoice, Retailer..."
                            value={searchQuery}
                            onChange={(e) => {
                                setSearchQuery(e.target.value);
                                setPage(1);
                            }}
                            InputProps={{
                                startAdornment: (
                                    <InputAdornment position="start">
                                        <SearchIcon sx={{ color: "#64748b", fontSize: 18 }} />
                                    </InputAdornment>
                                ),
                                endAdornment: searchQuery ? (
                                    <InputAdornment position="end">
                                        <IconButton size="small" onClick={() => setSearchQuery("")}>
                                            <ClearIcon sx={{ fontSize: 15 }} />
                                        </IconButton>
                                    </InputAdornment>
                                ) : null,
                            }}
                            sx={{
                                width: { xs: "100%", sm: 260 },
                                "& .MuiOutlinedInput-root": {
                                    bgcolor: "#ffffff",
                                    borderRadius: 1.5,
                                    fontSize: "0.80rem",
                                    height: 34,
                                },
                            }}
                        />
                    </Box>
                </Box>

                {/* 5. MAIN FUNNEL TABLE CARD (SCROLL ENABLED INSIDE TABLE, ENTIRE SCREEN FITTED) */}
                <Paper
                    elevation={0}
                    sx={{
                        flex: 1,
                        minHeight: 0,
                        display: "flex",
                        flexDirection: "column",
                        borderRadius: 2,
                        border: "1px solid #e2e8f0",
                        bgcolor: "#ffffff",
                        boxShadow: "0 1px 3px rgba(0,0,0,0.04)",
                        overflow: "hidden",
                    }}
                >
                    <TableContainer
                        sx={{
                            flex: 1,
                            minHeight: 0,
                            overflow: "auto",
                        }}
                    >
                        <Table
                            size="small"
                            stickyHeader
                            aria-label="purchase delivery funnel table"
                            sx={{ minWidth: 800, width: "100%" }}
                        >
                            {/* TABLE HEAD */}
                            <TableHead>
                                {/* PRIMARY COLUMNS MATCHING EXCEL FUNNEL TRACKING */}
                                <TableRow
                                    sx={{
                                        "& th": {
                                            bgcolor: "#1E3A8A",
                                            fontWeight: 700,
                                            color: "#ffffff",
                                            fontSize: "0.80rem",
                                            borderRight: "1px solid rgba(255, 255, 255, 0.15)",
                                            py: 1.1,
                                            px: 1,
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 5,
                                        },
                                    }}
                                >
                                    <TableCell align="center" sx={{ width: 44, px: 0.5 }}>
                                        #
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: 60 }}>
                                        S.No
                                    </TableCell>
                                    <TableCell sx={{ minWidth: 160 }}>
                                        Order Id
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: 95 }}>
                                        Trip
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: 95 }}>
                                        Invoice
                                    </TableCell>
                                    <TableCell align="right" sx={{ width: 170 }}>
                                        Payment
                                    </TableCell>
                                    <TableCell align="right" sx={{ width: 150 }}>
                                        Total
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: 140, borderRight: "none" }}>
                                        Status
                                    </TableCell>
                                </TableRow>

                                {/* STICKY GRAND TOTALS ROW DIRECTLY UNDER TABLE HEADER */}
                                <TableRow
                                    sx={{
                                        bgcolor: "#f1f5f9",
                                        borderBottom: "2px solid #cbd5e1",
                                        "& th, & td": {
                                            fontWeight: 800,
                                            fontSize: "0.78rem",
                                            py: 0.8,
                                            px: 1,
                                            color: "#0f172a",
                                            borderRight: "1px solid #e2e8f0",
                                            position: "sticky",
                                            top: "41px",
                                            zIndex: 4,
                                            bgcolor: "#f1f5f9",
                                        },
                                    }}
                                >
                                    <TableCell align="center">-</TableCell>
                                    <TableCell align="center">TOTAL</TableCell>
                                    <TableCell sx={{ color: "#1e3a8a", fontWeight: 800 }}>
                                        {filteredData.length} Orders
                                    </TableCell>
                                    {/* Trip Total Count */}
                                    <TableCell align="center" sx={{ color: "#0369a1", fontWeight: 800, fontSize: "0.82rem" }}>
                                        {tableGrandTotals.totalTrips}
                                    </TableCell>
                                    {/* Invoice Total Count */}
                                    <TableCell align="center" sx={{ color: "#c2410c", fontWeight: 800, fontSize: "0.82rem" }}>
                                        {tableGrandTotals.totalInvoices}
                                    </TableCell>
                                    {/* Balance Total */}
                                    <TableCell align="right" sx={{ color: "#b45309", fontWeight: 800 }}>
                                        {formatCurrency(tableGrandTotals.totalBalance)}
                                    </TableCell>
                                    {/* Total Amount */}
                                    <TableCell align="right" sx={{ color: "#15803d", fontWeight: 800 }}>
                                        {formatCurrency(tableGrandTotals.grandTotalAmount)}
                                    </TableCell>
                                    {/* Overall Status summary */}
                                    <TableCell align="center" sx={{ borderRight: "none" }}>
                                        <Chip
                                            size="small"
                                            label={`${metrics.completionRate}% Done`}
                                            sx={{
                                                bgcolor: "#dcfce7",
                                                color: "#15803d",
                                                fontWeight: 800,
                                                fontSize: "0.70rem",
                                                height: 20,
                                            }}
                                        />
                                    </TableCell>
                                </TableRow>
                            </TableHead>

                            {/* TABLE BODY */}
                            <TableBody>
                                {paginatedData.length > 0 ? (
                                    paginatedData.map((row, idx) => {
                                        const isExpanded = Boolean(expandedRows[row.id]);
                                        const tripsCount = row.trips?.length || 0;
                                        const invoicesCount = row.invoices?.length || 0;
                                        const isCompleted = row.status === "Completed";
                                        const orderItems = getOrderItems(row);

                                        return (
                                            <React.Fragment key={row.id}>
                                                {/* MAIN DATA ROW */}
                                                <TableRow
                                                    hover
                                                    sx={{
                                                        cursor: "pointer",
                                                        bgcolor: isExpanded
                                                            ? "#f8fafc"
                                                            : idx % 2 === 1
                                                                ? "#fafbfc"
                                                                : "#ffffff",
                                                        transition: "background-color 0.2s ease",
                                                        borderBottom: isExpanded ? "none" : "1px solid #e2e8f0",
                                                        "&:hover": {
                                                            bgcolor: "#f1f5f9 !important",
                                                        },
                                                    }}
                                                    onClick={() => toggleRow(row.id)}
                                                >
                                                    {/* # CHEVRON */}
                                                    <TableCell
                                                        align="center"
                                                        sx={{
                                                            width: 44,
                                                            py: 1,
                                                            px: 0.5,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        <IconButton
                                                            size="small"
                                                            onClick={(e) => {
                                                                e.stopPropagation();
                                                                toggleRow(row.id);
                                                            }}
                                                            sx={{
                                                                p: 0.3,
                                                                color: isExpanded ? "#1e3a8a" : "#64748b",
                                                                transform: isExpanded ? "rotate(180deg)" : "rotate(0deg)",
                                                                transition: "transform 0.2s ease, color 0.2s ease",
                                                            }}
                                                        >
                                                            <KeyboardArrowDownIcon sx={{ fontSize: 18 }} />
                                                        </IconButton>
                                                    </TableCell>

                                                    {/* S.NO */}
                                                    <TableCell
                                                        align="center"
                                                        sx={{
                                                            fontWeight: 700,
                                                            color: "#475569",
                                                            fontSize: "0.80rem",
                                                            width: 60,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        {(page - 1) * rowsPerPage + idx + 1}
                                                    </TableCell>

                                                    {/* ORDER ID */}
                                                    <TableCell
                                                        sx={{
                                                            minWidth: 160,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        <Box sx={{ display: "flex", flexDirection: "column", gap: 0.2 }}>
                                                            <Typography
                                                                sx={{
                                                                    fontWeight: 800,
                                                                    color: "#1e3a8a",
                                                                    fontSize: "0.82rem",
                                                                    letterSpacing: 0.3,
                                                                }}
                                                            >
                                                                {row.orderId}
                                                            </Typography>
                                                            <Typography
                                                                sx={{
                                                                    fontSize: "0.70rem",
                                                                    color: "#64748b",
                                                                    fontWeight: 500,
                                                                }}
                                                            >
                                                                {row.orderDate}
                                                            </Typography>
                                                        </Box>
                                                    </TableCell>

                                                    {/* TRIP (COUNT ALONE) */}
                                                    <TableCell
                                                        align="center"
                                                        sx={{
                                                            width: 95,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        <Typography
                                                            sx={{
                                                                fontWeight: 800,
                                                                fontSize: "0.82rem",
                                                                color: tripsCount > 0 ? "#0f172a" : "#94a3b8",
                                                            }}
                                                        >
                                                            {tripsCount}
                                                        </Typography>
                                                    </TableCell>

                                                    {/* INVOICE (COUNT ALONE) */}
                                                    <TableCell
                                                        align="center"
                                                        sx={{
                                                            width: 95,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        <Typography
                                                            sx={{
                                                                fontWeight: 800,
                                                                fontSize: "0.82rem",
                                                                color: invoicesCount > 0 ? "#0f172a" : "#94a3b8",
                                                            }}
                                                        >
                                                            {invoicesCount}
                                                        </Typography>
                                                    </TableCell>

                                                    {/* PAYMENT (BALANCE TO SETTLE) */}
                                                    <TableCell
                                                        align="right"
                                                        sx={{
                                                            width: 170,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        <Box sx={{ display: "flex", flexDirection: "column", alignItems: "flex-end" }}>
                                                            <Typography
                                                                sx={{
                                                                    fontWeight: 800,
                                                                    color: row.balanceToSettle > 0 ? "#b45309" : "#15803d",
                                                                    fontSize: "0.82rem",
                                                                }}
                                                            >
                                                                {formatCurrency(row.balanceToSettle)}
                                                            </Typography>
                                                            <Typography
                                                                sx={{
                                                                    fontSize: "0.68rem",
                                                                    color: "#64748b",
                                                                    fontWeight: 600,
                                                                }}
                                                            >
                                                                Paid: {formatCurrency(row.paidAmount)}
                                                            </Typography>
                                                        </Box>
                                                    </TableCell>

                                                    {/* TOTAL (TOTAL AMOUNT) */}
                                                    <TableCell
                                                        align="right"
                                                        sx={{
                                                            fontWeight: 800,
                                                            color: "#0f172a",
                                                            fontSize: "0.84rem",
                                                            width: 150,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "1px solid #eef2f6",
                                                        }}
                                                    >
                                                        {formatCurrency(row.totalAmount)}
                                                    </TableCell>

                                                    {/* STATUS */}
                                                    <TableCell
                                                        align="center"
                                                        sx={{
                                                            width: 140,
                                                            py: 1,
                                                            px: 1,
                                                            borderRight: "none",
                                                        }}
                                                    >
                                                        <Tooltip title={row.statusReason || row.status} arrow>
                                                            <Box sx={{ display: "inline-flex", flexDirection: "column", alignItems: "center" }}>
                                                                <Chip
                                                                    size="small"
                                                                    icon={
                                                                        isCompleted ? (
                                                                            <CheckCircleOutlineIcon sx={{ fontSize: 14 }} />
                                                                        ) : (
                                                                            <AccessTimeIcon sx={{ fontSize: 14 }} />
                                                                        )
                                                                    }
                                                                    label={row.status}
                                                                    sx={{
                                                                        fontWeight: 800,
                                                                        fontSize: "0.70rem",
                                                                        height: 22,
                                                                        bgcolor: isCompleted ? "#dcfce7" : "#fee2e2",
                                                                        color: isCompleted ? "#15803d" : "#b91c1c",
                                                                        border: "1px solid",
                                                                        borderColor: isCompleted ? "#86efac" : "#fca5a5",
                                                                        "& .MuiChip-icon": {
                                                                            color: isCompleted ? "#16a34a" : "#dc2626",
                                                                        },
                                                                    }}
                                                                />
                                                                {!isCompleted && row.statusReason && (
                                                                    <Typography
                                                                        sx={{
                                                                            fontSize: "0.65rem",
                                                                            color: "#b91c1c",
                                                                            fontWeight: 600,
                                                                            mt: 0.3,
                                                                            maxWidth: 130,
                                                                            whiteSpace: "nowrap",
                                                                            overflow: "hidden",
                                                                            textOverflow: "ellipsis",
                                                                        }}
                                                                    >
                                                                        {row.statusReason}
                                                                    </Typography>
                                                                )}
                                                            </Box>
                                                        </Tooltip>
                                                    </TableCell>
                                                </TableRow>

                                                {/* EXPANDED ROW: ORDER FUNNEL DETAILS */}
                                                <TableRow sx={{ bgcolor: "#f8fafc" }}>
                                                    <TableCell
                                                        colSpan={8}
                                                        sx={{
                                                            py: 0,
                                                            px: { xs: 1, md: 2.5 },
                                                            bgcolor: "#f8fafc",
                                                            borderBottom: isExpanded ? "2px solid #cbd5e1" : "none",
                                                        }}
                                                    >
                                                        <Collapse in={isExpanded} timeout="auto" unmountOnExit>
                                                            <Box
                                                                sx={{
                                                                    my: 1.5,
                                                                    p: 2,
                                                                    bgcolor: "#ffffff",
                                                                    borderRadius: 2,
                                                                    border: "1px solid #cbd5e1",
                                                                    borderLeft: "4px solid #1e3a8a",
                                                                    boxShadow: "0 2px 6px rgba(0,0,0,0.04)",
                                                                }}
                                                            >
                                                                {/* Header Strip inside expansion listing full Party/Retailer & Order Details */}
                                                                <Box
                                                                    sx={{
                                                                        display: "flex",
                                                                        justifyContent: "space-between",
                                                                        alignItems: "center",
                                                                        pb: 1.5,
                                                                        mb: 1.5,
                                                                        borderBottom: "1px solid #e2e8f0",
                                                                        flexWrap: "wrap",
                                                                        gap: 1,
                                                                    }}
                                                                >
                                                                    <Box sx={{ display: "flex", alignItems: "center", gap: 1.2, flexWrap: "wrap" }}>
                                                                        <ReceiptLongIcon sx={{ fontSize: 24, color: "#1e3a8a" }} />
                                                                        <Box>
                                                                            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                                                                                <Typography sx={{ fontWeight: 800, fontSize: "0.95rem", color: "#0f172a" }}>
                                                                                    {row.retailerName}
                                                                                </Typography>
                                                                                <Chip
                                                                                    size="small"
                                                                                    label={`${row.city}, ${row.state}`}
                                                                                    sx={{
                                                                                        fontWeight: 600,
                                                                                        fontSize: "0.70rem",
                                                                                        bgcolor: "#f1f5f9",
                                                                                        color: "#475569",
                                                                                        height: 20,
                                                                                    }}
                                                                                />
                                                                            </Box>
                                                                            <Typography sx={{ fontSize: "0.72rem", color: "#64748b", fontWeight: 500, mt: 0.2 }}>
                                                                                Order ID: <strong>{row.orderId}</strong> • Date: {row.orderDate} • Total Order Value: <strong>{formatCurrency(row.totalAmount)}</strong>
                                                                            </Typography>
                                                                        </Box>
                                                                    </Box>

                                                                    <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                                                                        <Chip
                                                                            size="small"
                                                                            label={isCompleted ? "Fully Completed" : "Action Required"}
                                                                            sx={{
                                                                                height: 22,
                                                                                fontSize: "0.70rem",
                                                                                fontWeight: 800,
                                                                                bgcolor: isCompleted ? "#dcfce7" : "#fee2e2",
                                                                                color: isCompleted ? "#15803d" : "#b91c1c",
                                                                            }}
                                                                        />
                                                                        <Button
                                                                            size="small"
                                                                            variant="outlined"
                                                                            startIcon={<VisibilityIcon sx={{ fontSize: 15 }} />}
                                                                            onClick={(e) => {
                                                                                e.stopPropagation();
                                                                                setSelectedOrderDetails(row);
                                                                            }}
                                                                            sx={{
                                                                                textTransform: "none",
                                                                                fontSize: "0.72rem",
                                                                                fontWeight: 700,
                                                                                height: 26,
                                                                                py: 0,
                                                                                px: 1.2,
                                                                                borderColor: "#cbd5e1",
                                                                                color: "#1e3a8a",
                                                                                "&:hover": {
                                                                                    borderColor: "#1e3a8a",
                                                                                    bgcolor: "#eff6ff",
                                                                                },
                                                                            }}
                                                                        >
                                                                            Full Journey
                                                                        </Button>
                                                                    </Box>
                                                                </Box>

                                                                {/* 1. ORDER ITEM DETAILS & INVOICE / TRIP MAPPING */}
                                                                <Box>
                                                                    <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", mb: 1, flexWrap: "wrap", gap: 1 }}>
                                                                        <Box sx={{ display: "flex", alignItems: "center", gap: 0.8 }}>
                                                                            <Inventory2OutlinedIcon sx={{ fontSize: 18, color: "#1e3a8a" }} />
                                                                            <Typography sx={{ fontSize: "0.80rem", fontWeight: 800, color: "#1e3a8a" }}>
                                                                                Order Items & Dispatch Fulfillment ({orderItems.length} Items)
                                                                            </Typography>
                                                                        </Box>
                                                                        <Typography sx={{ fontSize: "0.70rem", color: "#64748b" }}>
                                                                            Item-wise amounts mapped according to assigned trips & invoices
                                                                        </Typography>
                                                                    </Box>

                                                                    <TableContainer
                                                                        component={Paper}
                                                                        variant="outlined"
                                                                        sx={{
                                                                            borderRadius: 1.5,
                                                                            borderColor: "#e2e8f0",
                                                                            overflow: "hidden",
                                                                            bgcolor: "#ffffff",
                                                                        }}
                                                                    >
                                                                        <Table size="small" aria-label="order items details">
                                                                            <TableHead>
                                                                                <TableRow sx={{ bgcolor: "#f8fafc", borderBottom: "1px solid #e2e8f0" }}>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8, width: 40 }} align="center">
                                                                                        #
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8, minWidth: 200 }}>
                                                                                        Item Description / Commodity
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8 }} align="right">
                                                                                        Quantity
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8 }} align="right">
                                                                                        Rate (₹)
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8 }} align="right">
                                                                                        Item Amount (₹)
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8, minWidth: 150 }}>
                                                                                        Assigned Trip
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8, minWidth: 160 }}>
                                                                                        Invoice Details
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", color: "#475569", py: 0.8 }} align="center">
                                                                                        Delivery Status
                                                                                    </TableCell>
                                                                                </TableRow>
                                                                            </TableHead>
                                                                            <TableBody>
                                                                                {orderItems.map((item, itIdx) => (
                                                                                    <TableRow
                                                                                        key={item.itemId || itIdx}
                                                                                        sx={{
                                                                                            bgcolor: itIdx % 2 === 1 ? "#fafbfc" : "#ffffff",
                                                                                            "&:hover": { bgcolor: "#f1f5f9" },
                                                                                            borderBottom: "1px solid #f1f5f9",
                                                                                        }}
                                                                                    >
                                                                                        <TableCell align="center" sx={{ fontSize: "0.72rem", color: "#64748b", py: 0.8 }}>
                                                                                            {itIdx + 1}
                                                                                        </TableCell>
                                                                                        <TableCell sx={{ py: 0.8 }}>
                                                                                            <Typography sx={{ fontSize: "0.75rem", fontWeight: 700, color: "#0f172a" }}>
                                                                                                {item.itemName}
                                                                                            </Typography>
                                                                                        </TableCell>
                                                                                        <TableCell align="right" sx={{ fontSize: "0.72rem", color: "#334155", fontWeight: 600, py: 0.8 }}>
                                                                                            {item.quantity} {item.unit || "Bags"}
                                                                                        </TableCell>
                                                                                        <TableCell align="right" sx={{ fontSize: "0.72rem", color: "#334155", py: 0.8 }}>
                                                                                            {formatCurrency(item.rate)}
                                                                                        </TableCell>
                                                                                        <TableCell align="right" sx={{ fontSize: "0.75rem", fontWeight: 800, color: "#0f172a", py: 0.8 }}>
                                                                                            {formatCurrency(item.amount)}
                                                                                        </TableCell>
                                                                                        <TableCell sx={{ py: 0.8 }}>
                                                                                            {item.tripNo && item.tripNo !== "Unassigned" ? (
                                                                                                <Box sx={{ display: "flex", alignItems: "center", gap: 0.6, flexWrap: "wrap" }}>
                                                                                                    <LocalShippingIcon sx={{ fontSize: 14, color: "#0284c7" }} />
                                                                                                    <Typography sx={{ fontSize: "0.74rem", fontWeight: 700, color: "#0369a1" }}>
                                                                                                        Trip: {item.tripNo}
                                                                                                    </Typography>
                                                                                                    {item.vehicleNo && (
                                                                                                        <Typography sx={{ fontSize: "0.68rem", color: "#475569", fontWeight: 500 }}>
                                                                                                            ({item.vehicleNo})
                                                                                                        </Typography>
                                                                                                    )}
                                                                                                </Box>
                                                                                            ) : (
                                                                                                <Typography sx={{ fontSize: "0.70rem", color: "#b91c1c", fontWeight: 600 }}>
                                                                                                    Trip Pending
                                                                                                </Typography>
                                                                                            )}
                                                                                        </TableCell>
                                                                                        <TableCell sx={{ py: 0.8 }}>
                                                                                            {item.invoiceNo && item.invoiceNo !== "Unassigned" ? (
                                                                                                <Box sx={{ display: "flex", alignItems: "center", gap: 0.6, flexWrap: "wrap" }}>
                                                                                                    <ReceiptLongIcon sx={{ fontSize: 14, color: "#c2410c" }} />
                                                                                                    <Typography sx={{ fontSize: "0.74rem", fontWeight: 700, color: "#c2410c" }}>
                                                                                                        Inv: {item.invoiceNo}
                                                                                                    </Typography>
                                                                                                    {item.invoiceDate && (
                                                                                                        <Typography sx={{ fontSize: "0.68rem", color: "#64748b" }}>
                                                                                                            ({item.invoiceDate})
                                                                                                        </Typography>
                                                                                                    )}
                                                                                                </Box>
                                                                                            ) : (
                                                                                                <Typography sx={{ fontSize: "0.70rem", color: "#b91c1c", fontWeight: 600 }}>
                                                                                                    Invoice Pending
                                                                                                </Typography>
                                                                                            )}
                                                                                        </TableCell>
                                                                                        <TableCell align="center" sx={{ py: 0.8 }}>
                                                                                            <Chip
                                                                                                size="small"
                                                                                                label={item.status || "Pending"}
                                                                                                sx={{
                                                                                                    height: 20,
                                                                                                    fontSize: "0.66rem",
                                                                                                    fontWeight: 700,
                                                                                                    bgcolor:
                                                                                                        item.status === "Delivered"
                                                                                                            ? "#dcfce7"
                                                                                                            : item.status === "In Transit"
                                                                                                                ? "#fef3c7"
                                                                                                                : "#fee2e2",
                                                                                                    color:
                                                                                                        item.status === "Delivered"
                                                                                                            ? "#15803d"
                                                                                                            : item.status === "In Transit"
                                                                                                                ? "#b45309"
                                                                                                                : "#b91c1c",
                                                                                                }}
                                                                                            />
                                                                                        </TableCell>
                                                                                    </TableRow>
                                                                                ))}
                                                                                {/* Order Items Subtotal */}
                                                                                <TableRow sx={{ bgcolor: "#f8fafc", borderTop: "2px solid #e2e8f0" }}>
                                                                                    <TableCell align="center" sx={{ fontWeight: 800, fontSize: "0.72rem", color: "#1e3a8a" }}>
                                                                                        ∑
                                                                                    </TableCell>
                                                                                    <TableCell sx={{ fontWeight: 800, fontSize: "0.74rem", color: "#1e3a8a" }}>
                                                                                        Total ({orderItems.length} Items)
                                                                                    </TableCell>
                                                                                    <TableCell align="right" sx={{ fontWeight: 800, fontSize: "0.74rem", color: "#0f172a" }}>
                                                                                        {orderItems.reduce((acc, curr) => acc + (curr.quantity || 0), 0)} Bags
                                                                                    </TableCell>
                                                                                    <TableCell align="right" sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                                                                                        -
                                                                                    </TableCell>
                                                                                    <TableCell align="right" sx={{ fontWeight: 800, fontSize: "0.78rem", color: "#15803d" }}>
                                                                                        {formatCurrency(row.totalAmount)}
                                                                                    </TableCell>
                                                                                    <TableCell colSpan={3} sx={{ fontSize: "0.70rem", color: "#64748b", fontStyle: "italic" }}>
                                                                                        {row.trips.length} trip(s) assigned • {row.invoices.length} invoice(s) generated
                                                                                    </TableCell>
                                                                                </TableRow>
                                                                            </TableBody>
                                                                        </Table>
                                                                    </TableContainer>
                                                                </Box>
                                                            </Box>
                                                        </Collapse>
                                                </TableCell>
                                            </TableRow>
                                                </React.Fragment>
                            );
                                        })
                            ) : (
                            <TableRow>
                                <TableCell colSpan={8} align="center" sx={{ py: 6 }}>
                                    <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 1 }}>
                                        <ErrorOutlineIcon sx={{ fontSize: 36, color: "#94a3b8" }} />
                                        <Typography sx={{ color: "#64748b", fontWeight: 600, fontSize: "0.9rem" }}>
                                            No matching purchase delivery records found
                                        </Typography>
                                        <Button
                                            size="small"
                                            variant="outlined"
                                            onClick={handleResetFilters}
                                            sx={{ textTransform: "none", mt: 0.5 }}
                                        >
                                            Clear All Filters
                                        </Button>
                                    </Box>
                                </TableCell>
                            </TableRow>
                                    )}
                        </TableBody>
                    </Table>
                </TableContainer>
            </Paper>

            {/* 6. PINNED PAGINATION */}
            <Box sx={{ flexShrink: 0 }}>
                <CommonPagination
                    totalRows={filteredData.length}
                    page={page}
                    rowsPerPage={rowsPerPage}
                    onPageChange={setPage}
                    onRowsPerPageChange={(rows) => {
                        setRowsPerPage(rows);
                        setPage(1);
                    }}
                />
            </Box>
        </Box>

            {/* 7. TRIP DETAILS MODAL */ }
    <Dialog
        open={Boolean(selectedTrip)}
        onClose={() => setSelectedTrip(null)}
        maxWidth="sm"
        fullWidth
    >
        <DialogTitle
            sx={{
                bgcolor: "#1e3a8a",
                color: "#ffffff",
                fontWeight: 700,
                fontSize: "0.95rem",
                py: 1.2,
                px: 2,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
            }}
        >
            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                <LocalShippingIcon sx={{ fontSize: 20 }} />
                Trip #{selectedTrip?.tripNo} Details
            </Box>
            <IconButton
                size="small"
                onClick={() => setSelectedTrip(null)}
                sx={{ color: "#ffffff" }}
            >
                <CloseIcon sx={{ fontSize: 18 }} />
            </IconButton>
        </DialogTitle>
        <DialogContent sx={{ p: 2.5 }}>
            {selectedTrip && (
                <Grid container spacing={2}>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Trip Number
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.tripNo}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Vehicle Number
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.vehicleNo}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Driver Name
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.driverName}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Driver Contact
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.driverPhone || "N/A"}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Transporter
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.transporter}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            LR Number
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.lrNo}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Dispatch Date
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedTrip.dispatchDate}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Delivery Status
                        </Typography>
                        <Chip
                            size="small"
                            label={selectedTrip.status}
                            sx={{
                                fontWeight: 700,
                                bgcolor:
                                    selectedTrip.status === "Delivered"
                                        ? "#dcfce7"
                                        : selectedTrip.status === "In Transit"
                                            ? "#fef3c7"
                                            : "#e0f2fe",
                                color:
                                    selectedTrip.status === "Delivered"
                                        ? "#15803d"
                                        : selectedTrip.status === "In Transit"
                                            ? "#b45309"
                                            : "#0369a1",
                            }}
                        />
                    </Grid>
                </Grid>
            )}
        </DialogContent>
        <DialogActions sx={{ px: 2, pb: 2 }}>
            <Button
                size="small"
                variant="contained"
                onClick={() => setSelectedTrip(null)}
                sx={{ bgcolor: "#1e3a8a", textTransform: "none" }}
            >
                Close
            </Button>
        </DialogActions>
    </Dialog>

    {/* 8. INVOICE DETAILS MODAL */ }
    <Dialog
        open={Boolean(selectedInvoice)}
        onClose={() => setSelectedInvoice(null)}
        maxWidth="xs"
        fullWidth
    >
        <DialogTitle
            sx={{
                bgcolor: "#1e3a8a",
                color: "#ffffff",
                fontWeight: 700,
                fontSize: "0.95rem",
                py: 1.2,
                px: 2,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
            }}
        >
            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                <ReceiptLongIcon sx={{ fontSize: 20 }} />
                Invoice #{selectedInvoice?.invoiceNo}
            </Box>
            <IconButton
                size="small"
                onClick={() => setSelectedInvoice(null)}
                sx={{ color: "#ffffff" }}
            >
                <CloseIcon sx={{ fontSize: 18 }} />
            </IconButton>
        </DialogTitle>
        <DialogContent sx={{ p: 2.5 }}>
            {selectedInvoice && (
                <Grid container spacing={2}>
                    <Grid item xs={12}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Invoice Number
                        </Typography>
                        <Typography sx={{ fontSize: "0.9rem", fontWeight: 800, color: "#1e3a8a" }}>
                            {selectedInvoice.invoiceNo}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Invoice Date
                        </Typography>
                        <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                            {selectedInvoice.invoiceDate}
                        </Typography>
                    </Grid>
                    <Grid item xs={6}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Invoice Amount
                        </Typography>
                        <Typography sx={{ fontSize: "0.95rem", fontWeight: 800, color: "#15803d" }}>
                            {formatCurrency(selectedInvoice.amount)}
                        </Typography>
                    </Grid>
                    <Grid item xs={12}>
                        <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>
                            Invoice Status
                        </Typography>
                        <Chip
                            size="small"
                            label={selectedInvoice.status}
                            sx={{
                                mt: 0.5,
                                fontWeight: 700,
                                bgcolor: "#ffedd5",
                                color: "#c2410c",
                            }}
                        />
                    </Grid>
                </Grid>
            )}
        </DialogContent>
        <DialogActions sx={{ px: 2, pb: 2 }}>
            <Button
                size="small"
                variant="contained"
                onClick={() => setSelectedInvoice(null)}
                sx={{ bgcolor: "#1e3a8a", textTransform: "none" }}
            >
                Close
            </Button>
        </DialogActions>
    </Dialog>

    {/* 9. FULL ORDER JOURNEY MODAL */ }
    <Dialog
        open={Boolean(selectedOrderDetails)}
        onClose={() => setSelectedOrderDetails(null)}
        maxWidth="md"
        fullWidth
    >
        <DialogTitle
            sx={{
                bgcolor: "#1e3a8a",
                color: "#ffffff",
                fontWeight: 700,
                fontSize: "1rem",
                py: 1.5,
                px: 2.5,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
            }}
        >
            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                <ReceiptLongIcon sx={{ fontSize: 22 }} />
                Funnel Lifecycle: Order {selectedOrderDetails?.orderId}
            </Box>
            <IconButton
                size="small"
                onClick={() => setSelectedOrderDetails(null)}
                sx={{ color: "#ffffff" }}
            >
                <CloseIcon sx={{ fontSize: 18 }} />
            </IconButton>
        </DialogTitle>
        <DialogContent sx={{ p: 3 }}>
            {selectedOrderDetails && (
                <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
                    {/* Top Overview */}
                    <Box
                        sx={{
                            p: 2,
                            bgcolor: "#f8fafc",
                            borderRadius: 2,
                            border: "1px solid #e2e8f0",
                            display: "flex",
                            justifyContent: "space-between",
                            flexWrap: "wrap",
                            gap: 2,
                        }}
                    >
                        <Box>
                            <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>Retailer / Supplier</Typography>
                            <Typography sx={{ fontSize: "0.95rem", fontWeight: 800, color: "#0f172a" }}>
                                {selectedOrderDetails.retailerName}
                            </Typography>
                            <Typography sx={{ fontSize: "0.75rem", color: "#475569" }}>
                                {selectedOrderDetails.city}, {selectedOrderDetails.state}
                            </Typography>
                        </Box>
                        <Box>
                            <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>Order Placed</Typography>
                            <Typography sx={{ fontSize: "0.85rem", fontWeight: 700, color: "#1e293b" }}>
                                {selectedOrderDetails.orderDate}
                            </Typography>
                        </Box>
                        <Box>
                            <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>Order Items</Typography>
                            <Chip
                                size="small"
                                label={`${(selectedOrderDetails.items || getOrderItems(selectedOrderDetails)).length} Items`}
                                sx={{ fontWeight: 700, bgcolor: "#e2e8f0", color: "#1e293b" }}
                            />
                        </Box>
                        <Box>
                            <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>Overall Status</Typography>
                            <Chip
                                size="small"
                                label={selectedOrderDetails.status}
                                sx={{
                                    fontWeight: 800,
                                    bgcolor: selectedOrderDetails.status === "Completed" ? "#dcfce7" : "#fee2e2",
                                    color: selectedOrderDetails.status === "Completed" ? "#15803d" : "#b91c1c",
                                }}
                            />
                        </Box>
                    </Box>

                    {/* Financial Reconcile Bar */}
                    <Grid container spacing={2}>
                        <Grid item xs={4}>
                            <Box sx={{ p: 1.5, bgcolor: "#f1f5f9", borderRadius: 1.5, textAlign: "center" }}>
                                <Typography sx={{ fontSize: "0.72rem", color: "#64748b" }}>Total Amount</Typography>
                                <Typography sx={{ fontSize: "1.05rem", fontWeight: 800, color: "#0f172a" }}>
                                    {formatCurrency(selectedOrderDetails.totalAmount)}
                                </Typography>
                            </Box>
                        </Grid>
                        <Grid item xs={4}>
                            <Box sx={{ p: 1.5, bgcolor: "#ecfdf5", borderRadius: 1.5, textAlign: "center" }}>
                                <Typography sx={{ fontSize: "0.72rem", color: "#166534" }}>Paid Amount</Typography>
                                <Typography sx={{ fontSize: "1.05rem", fontWeight: 800, color: "#15803d" }}>
                                    {formatCurrency(selectedOrderDetails.paidAmount)}
                                </Typography>
                            </Box>
                        </Grid>
                        <Grid item xs={4}>
                            <Box sx={{ p: 1.5, bgcolor: "#fff7ed", borderRadius: 1.5, textAlign: "center" }}>
                                <Typography sx={{ fontSize: "0.72rem", color: "#9a3412" }}>Balance to Settle</Typography>
                                <Typography
                                    sx={{
                                        fontSize: "1.05rem",
                                        fontWeight: 800,
                                        color: selectedOrderDetails.balanceToSettle > 0 ? "#c2410c" : "#15803d",
                                    }}
                                >
                                    {formatCurrency(selectedOrderDetails.balanceToSettle)}
                                </Typography>
                            </Box>
                        </Grid>
                    </Grid>

                    {/* Order Items Table in Modal */}
                    <Box>
                        <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", mb: 1, flexWrap: "wrap", gap: 1 }}>
                            <Typography sx={{ fontSize: "0.82rem", fontWeight: 800, color: "#1e3a8a" }}>
                                Order Items Breakdown ({(selectedOrderDetails.items || getOrderItems(selectedOrderDetails)).length} Items)
                            </Typography>
                            <Typography sx={{ fontSize: "0.70rem", color: "#64748b" }}>
                                Item-wise amounts mapped according to assigned trips & invoices
                            </Typography>
                        </Box>
                        <TableContainer component={Paper} variant="outlined">
                            <Table size="small">
                                <TableHead>
                                    <TableRow sx={{ bgcolor: "#f1f5f9" }}>
                                        <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem", width: 40 }} align="center">#</TableCell>
                                        <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Item Description</TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Qty</TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Rate</TableCell>
                                        <TableCell align="right" sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Item Amount</TableCell>
                                        <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Trip Assigned</TableCell>
                                        <TableCell sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Invoice Ref</TableCell>
                                        <TableCell align="center" sx={{ fontWeight: 700, fontSize: "0.72rem" }}>Status</TableCell>
                                    </TableRow>
                                </TableHead>
                                <TableBody>
                                    {(selectedOrderDetails.items || getOrderItems(selectedOrderDetails)).map((item, mIdx) => (
                                        <TableRow key={item.itemId || mIdx} sx={{ bgcolor: mIdx % 2 === 1 ? "#fafbfc" : "#ffffff" }}>
                                            <TableCell align="center" sx={{ fontSize: "0.72rem" }}>{mIdx + 1}</TableCell>
                                            <TableCell sx={{ fontSize: "0.74rem", fontWeight: 700, color: "#0f172a" }}>{item.itemName}</TableCell>
                                            <TableCell align="right" sx={{ fontSize: "0.72rem" }}>{item.quantity} {item.unit || "Bags"}</TableCell>
                                            <TableCell align="right" sx={{ fontSize: "0.72rem" }}>{formatCurrency(item.rate)}</TableCell>
                                            <TableCell align="right" sx={{ fontSize: "0.74rem", fontWeight: 800, color: "#0f172a" }}>{formatCurrency(item.amount)}</TableCell>
                                            <TableCell sx={{ fontSize: "0.72rem" }}>
                                                {item.tripNo && item.tripNo !== "Unassigned" ? (
                                                    <Chip size="small" label={`Trip ${item.tripNo}`} sx={{ height: 18, fontSize: "0.62rem", bgcolor: "#e0f2fe", color: "#0369a1", fontWeight: 700 }} />
                                                ) : (
                                                    <Typography sx={{ fontSize: "0.68rem", color: "#b91c1c" }}>Pending</Typography>
                                                )}
                                            </TableCell>
                                            <TableCell sx={{ fontSize: "0.72rem" }}>
                                                {item.invoiceNo && item.invoiceNo !== "Unassigned" ? (
                                                    <Chip size="small" label={`Inv ${item.invoiceNo}`} sx={{ height: 18, fontSize: "0.62rem", bgcolor: "#ffedd5", color: "#c2410c", fontWeight: 700 }} />
                                                ) : (
                                                    <Typography sx={{ fontSize: "0.68rem", color: "#b91c1c" }}>Pending</Typography>
                                                )}
                                            </TableCell>
                                            <TableCell align="center">
                                                <Chip
                                                    size="small"
                                                    label={item.status || "Pending"}
                                                    sx={{
                                                        height: 18,
                                                        fontSize: "0.62rem",
                                                        fontWeight: 700,
                                                        bgcolor: item.status === "Delivered" ? "#dcfce7" : item.status === "In Transit" ? "#fef3c7" : "#fee2e2",
                                                        color: item.status === "Delivered" ? "#15803d" : item.status === "In Transit" ? "#b45309" : "#b91c1c",
                                                    }}
                                                />
                                            </TableCell>
                                        </TableRow>
                                    ))}
                                </TableBody>
                            </Table>
                        </TableContainer>
                    </Box>

                    {/* Trips List */}
                    <Box>
                        <Typography sx={{ fontSize: "0.82rem", fontWeight: 800, color: "#1e3a8a", mb: 1 }}>
                            Assigned Logistics ({selectedOrderDetails.trips.length} Trips)
                        </Typography>
                        {selectedOrderDetails.trips.length > 0 ? (
                            <TableContainer component={Paper} variant="outlined">
                                <Table size="small">
                                    <TableHead>
                                        <TableRow sx={{ bgcolor: "#f1f5f9" }}>
                                            <TableCell sx={{ fontWeight: 700 }}>Trip No</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Vehicle No</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Driver</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Transporter</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Dispatch</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Status</TableCell>
                                        </TableRow>
                                    </TableHead>
                                    <TableBody>
                                        {selectedOrderDetails.trips.map((t) => (
                                            <TableRow key={t.tripId}>
                                                <TableCell sx={{ fontWeight: 700 }}>{t.tripNo}</TableCell>
                                                <TableCell>{t.vehicleNo}</TableCell>
                                                <TableCell>{t.driverName}</TableCell>
                                                <TableCell>{t.transporter}</TableCell>
                                                <TableCell>{t.dispatchDate}</TableCell>
                                                <TableCell>
                                                    <Chip size="small" label={t.status} sx={{ height: 20, fontSize: "0.68rem" }} />
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </TableContainer>
                        ) : (
                            <Typography sx={{ fontSize: "0.75rem", color: "#ef4444", fontStyle: "italic" }}>
                                No trips assigned. Order is waiting for vehicle allocation.
                            </Typography>
                        )}
                    </Box>

                    {/* Invoices List */}
                    <Box>
                        <Typography sx={{ fontSize: "0.82rem", fontWeight: 800, color: "#1e3a8a", mb: 1 }}>
                            Invoices Raised ({selectedOrderDetails.invoices.length} Invoices)
                        </Typography>
                        {selectedOrderDetails.invoices.length > 0 ? (
                            <TableContainer component={Paper} variant="outlined">
                                <Table size="small">
                                    <TableHead>
                                        <TableRow sx={{ bgcolor: "#f1f5f9" }}>
                                            <TableCell sx={{ fontWeight: 700 }}>Invoice No</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Date</TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 700 }}>Amount</TableCell>
                                            <TableCell sx={{ fontWeight: 700 }}>Status</TableCell>
                                        </TableRow>
                                    </TableHead>
                                    <TableBody>
                                        {selectedOrderDetails.invoices.map((inv) => (
                                            <TableRow key={inv.invoiceId}>
                                                <TableCell sx={{ fontWeight: 700 }}>{inv.invoiceNo}</TableCell>
                                                <TableCell>{inv.invoiceDate}</TableCell>
                                                <TableCell align="right" sx={{ fontWeight: 700, color: "#15803d" }}>
                                                    {formatCurrency(inv.amount)}
                                                </TableCell>
                                                <TableCell>
                                                    <Chip size="small" label={inv.status} sx={{ height: 20, fontSize: "0.68rem" }} />
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </TableContainer>
                        ) : (
                            <Typography sx={{ fontSize: "0.75rem", color: "#c2410c", fontStyle: "italic" }}>
                                No invoices generated yet.
                            </Typography>
                        )}
                    </Box>
                </Box>
            )}
        </DialogContent>
        <DialogActions sx={{ px: 3, pb: 2.5 }}>
            <Button
                size="small"
                variant="contained"
                onClick={() => setSelectedOrderDetails(null)}
                sx={{ bgcolor: "#1e3a8a", textTransform: "none" }}
            >
                Close
            </Button>
        </DialogActions>
    </Dialog>
        </Box >
    );
};

export default PurchaseDelivery;
