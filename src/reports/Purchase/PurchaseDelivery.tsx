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
    Chip,
    Button,
    TextField,
    InputAdornment,
    IconButton,
} from "@mui/material";
import SearchIcon from "@mui/icons-material/Search";
import ClearIcon from "@mui/icons-material/Clear";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
import AccessTimeIcon from "@mui/icons-material/AccessTime";
import ErrorOutlineIcon from "@mui/icons-material/ErrorOutline";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import CommonPagination from "../../Components/CommonPagination";
import {
    PurchaseDeliveryItem,
    PURCHASE_DELIVERY_DATA,
} from "../../services/purchaseDelivery.service";

/* ================= HELPER FORMATTERS ================= */

const formatCurrency = (val: number | string): string => {
    if (!val && val !== 0) return "";
    const num = typeof val === "string" ? parseFloat(val.replace(/,/g, "")) : val;
    if (isNaN(num)) return String(val);
    return "₹\u00A0" + num.toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
};

/* ================= COMPONENT ================= */

const PurchaseDelivery: React.FC = () => {
    // State
    const [data] = useState<PurchaseDeliveryItem[]>(PURCHASE_DELIVERY_DATA);
    const [searchQuery, setSearchQuery] = useState<string>("");
    const [activeStatusFilter, setActiveStatusFilter] = useState<string>("ALL");

    const today = dayjs().format("YYYY-MM-DD");

    // Filter Drawer State
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [fromDate, setFromDate] = useState<string>(today);
    const [toDate, setToDate] = useState<string>(today);
    const [tempFromDate, setTempFromDate] = useState<string>(today);
    const [tempToDate, setTempToDate] = useState<string>(today);

    // Pagination State
    const [page, setPage] = useState<number>(1);
    const [rowsPerPage, setRowsPerPage] = useState<number>(100);

    // Filtered dataset
    const filteredData = useMemo(() => {
        return data.filter((item) => {
            // 1. Search Query
            if (searchQuery.trim() !== "") {
                const q = searchQuery.toLowerCase().trim();
                const matchStockGroup = item.stockGroup?.toLowerCase().includes(q);
                const matchBatch = item.inwardBatchWithItemName?.toLowerCase().includes(q);
                const matchPurOrder = item.purOrderNo?.toLowerCase().includes(q);
                const matchInwardJou = String(item.inwardJouNo).toLowerCase().includes(q);
                const matchPurInv = item.purInvNo?.toLowerCase().includes(q);
                const matchPaymentNo = item.paymentNo?.toLowerCase().includes(q);
                const matchStatus = item.status?.toLowerCase().includes(q);

                if (
                    !matchStockGroup &&
                    !matchBatch &&
                    !matchPurOrder &&
                    !matchInwardJou &&
                    !matchPurInv &&
                    !matchPaymentNo &&
                    !matchStatus
                ) {
                    return false;
                }
            }

            // 2. Status Quick Filter
            if (activeStatusFilter === "COMPLETED") {
                if (item.status !== "COMPLETED") return false;
            } else if (activeStatusFilter === "NOT_COMPLETED") {
                if (item.status === "COMPLETED") return false;
            }

            // 3. Date filter (Order Date)
            if (fromDate || toDate) {
                if (item.orderDate) {
                    const itemDate = dayjs(item.orderDate);
                    if (itemDate.isValid()) {
                        if (fromDate && itemDate.isBefore(dayjs(fromDate), "day")) return false;
                        if (toDate && itemDate.isAfter(dayjs(toDate), "day")) return false;
                    }
                }
            }

            return true;
        });
    }, [data, searchQuery, activeStatusFilter, fromDate, toDate]);

    // Metrics summary
    const metrics = useMemo(() => {
        const total = filteredData.length;
        const completed = filteredData.filter((i) => i.status === "COMPLETED").length;
        const notCompleted = total - completed;
        const completionRate = total > 0 ? Math.round((completed / total) * 100) : 0;

        let totalPaymentAmt = 0;
        filteredData.forEach((item) => {
            if (item.paymentAmt) {
                const num =
                    typeof item.paymentAmt === "string"
                        ? parseFloat(item.paymentAmt.replace(/,/g, ""))
                        : item.paymentAmt;
                if (!isNaN(num)) totalPaymentAmt += num;
            }
        });

        return {
            total,
            completed,
            notCompleted,
            completionRate,
            totalPaymentAmt,
        };
    }, [filteredData]);

    // Paginated Data
    const paginatedData = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        return filteredData.slice(start, start + rowsPerPage);
    }, [filteredData, page, rowsPerPage]);

    // Reset filters
    const handleResetFilters = () => {
        setSearchQuery("");
        setActiveStatusFilter("ALL");
        setFromDate(today);
        setToDate(today);
        setTempFromDate(today);
        setTempToDate(today);
        setPage(1);
    };

    /* ================= EXPORT HANDLERS ================= */

    const handleExportExcel = () => {
        try {
            if (!filteredData.length) {
                toast.warning("No data to export");
                return;
            }

            const excelRows: any[][] = [];

            // Header Banner
            excelRows.push(["PURCHASE DELIVERY REPORT"]);
            excelRows.push([
                `Generated on: ${dayjs().format("DD-MM-YYYY HH:mm")} | Total Orders: ${filteredData.length} | Completed: ${metrics.completed} | Pending: ${metrics.notCompleted}`,
            ]);
            excelRows.push([]);

            // Table Header row (matching user's image exactly)
            excelRows.push([
                "S.no",
                "Stock group",
                "Inward Batch with item name",
                "Pur .order no",
                "Inward Jou no",
                "Pur Inv no",
                "Payment no",
                "Payment amt",
                "status",
            ]);

            // Data Rows
            filteredData.forEach((row, idx) => {
                excelRows.push([
                    idx + 1,
                    row.stockGroup,
                    row.inwardBatchWithItemName,
                    row.purOrderNo,
                    row.inwardJouNo,
                    row.purInvNo,
                    row.paymentNo,
                    row.paymentAmt || "",
                    row.status,
                ]);
            });

            // Totals summary row
            excelRows.push([
                "TOTAL",
                "-",
                `${filteredData.length} Records`,
                "-",
                "-",
                "-",
                "-",
                metrics.totalPaymentAmt > 0 ? metrics.totalPaymentAmt : "-",
                `${metrics.completionRate}% Done`,
            ]);

            const ws = XLSX.utils.aoa_to_sheet(excelRows);

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

                        if (R === 0) {
                            cell.s.font = { name: "Arial", sz: 13, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                            cell.s.alignment = { horizontal: "center" };
                        } else if (R === 1) {
                            cell.s.font = { name: "Arial", sz: 9, italic: true, color: { rgb: "334155" } };
                            cell.s.fill = { fgColor: { rgb: "F1F5F9" } };
                        } else if (R === 3) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                            cell.s.alignment = { horizontal: "center", vertical: "center" };
                        } else if (R === range.e.r) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "0F172A" } };
                            cell.s.fill = { fgColor: { rgb: "E2E8F0" } };
                        } else if (R > 3) {
                            if (C === 8) {
                                const isCompleted = cell.v === "COMPLETED";
                                cell.s.font = {
                                    name: "Arial",
                                    sz: 10,
                                    bold: true,
                                    color: { rgb: isCompleted ? "15803D" : "B91C1C" },
                                };
                                cell.s.fill = { fgColor: { rgb: isCompleted ? "DCFCE7" : "FEE2E2" } };
                                cell.s.alignment = { horizontal: "center" };
                            }
                        }
                    }
                }
            }

            // Column widths
            ws["!cols"] = [
                { wch: 8 },  // S.no
                { wch: 16 }, // Stock group
                { wch: 34 }, // Inward Batch with item name
                { wch: 18 }, // Pur .order no
                { wch: 15 }, // Inward Jou no
                { wch: 18 }, // Pur Inv no
                { wch: 18 }, // Payment no
                { wch: 16 }, // Payment amt
                { wch: 16 }, // status
            ];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "PurchaseDelivery");
            XLSX.writeFile(wb, `PurchaseDelivery_Report_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
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
            doc.text("PURCHASE DELIVERY REPORT", 14, 15);

            doc.setFontSize(8.5);
            doc.setTextColor(100, 116, 139);
            doc.text(
                `Generated: ${dayjs().format("DD-MM-YYYY HH:mm")} | Total: ${filteredData.length} | Completed: ${metrics.completed} (${metrics.completionRate}%)`,
                14,
                21
            );

            const tableHeaders = [
                [
                    "S.no",
                    "Stock group",
                    "Inward Batch with item name",
                    "Pur .order no",
                    "Inward Jou no",
                    "Pur Inv no",
                    "Payment no",
                    "Payment amt",
                    "status",
                ],
            ];

            const tableBody = filteredData.map((row, idx) => [
                idx + 1,
                row.stockGroup,
                row.inwardBatchWithItemName,
                row.purOrderNo,
                row.inwardJouNo,
                row.purInvNo,
                row.paymentNo,
                row.paymentAmt ? formatCurrency(row.paymentAmt) : "",
                row.status,
            ]);

            tableBody.push([
                "TOTAL",
                "-",
                `${filteredData.length} Records`,
                "-",
                "-",
                "-",
                "-",
                metrics.totalPaymentAmt > 0 ? formatCurrency(metrics.totalPaymentAmt) : "-",
                `${metrics.completionRate}% Done`,
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
                    if (cellData.section === "body" && cellData.column.index === 8) {
                        const val = String(cellData.cell.raw);
                        if (val === "COMPLETED") {
                            cellData.cell.styles.textColor = [21, 128, 61];
                            cellData.cell.styles.fontStyle = "bold";
                        } else if (val === "NOT COMPLETED") {
                            cellData.cell.styles.textColor = [185, 28, 28];
                            cellData.cell.styles.fontStyle = "bold";
                        }
                    }
                    if (
                        cellData.section === "body" &&
                        cellData.row.index === tableBody.length - 1
                    ) {
                        cellData.cell.styles.fontStyle = "bold";
                        cellData.cell.styles.fillColor = [241, 245, 249];
                        cellData.cell.styles.textColor = [15, 23, 42];
                    }
                },
            });

            doc.save(`PurchaseDelivery_Report_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`);
            toast.success("PDF report downloaded successfully! ✅");
        } catch (err) {
            console.error(err);
            toast.error("Failed to export PDF report ❌");
        }
    };

    return (
        <Box
            sx={{
                width: "100%",
                height: "100%",
                maxHeight: "100%",
                display: "flex",
                flexDirection: "column",
                overflow: "hidden",
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
                onClose={() => {
                    setTempFromDate(fromDate);
                    setTempToDate(toDate);
                    setDrawerOpen(false);
                }}
                fromDate={tempFromDate}
                onFromDateChange={setTempFromDate}
                toDate={tempToDate}
                onToDateChange={setTempToDate}
                fromDateLabel="From Date"
                toDateLabel="To Date"
                onApply={() => {
                    setFromDate(tempFromDate);
                    setToDate(tempToDate);
                    setPage(1);
                    setDrawerOpen(false);
                    toast.success("Filters applied successfully!");
                }}
            />

            {/* 3. MAIN CONTENT CONTAINER */}
            <Box
                sx={{
                    flex: 1,
                    minHeight: 0,
                    display: "flex",
                    flexDirection: "column",
                    overflow: "hidden",
                    px: { xs: 1.5, md: 2.5 },
                    pt: 1,
                    pb: 1,
                    gap: 1,
                }}
            >
                {/* TITLE & CONTROLS BAR */}
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
                    {/* Title & Badges */}
                    <Box sx={{ display: "flex", alignItems: "center", gap: 1, flexWrap: "wrap" }}>
                        <Typography
                            variant="h6"
                            sx={{
                                fontWeight: 800,
                                color: "#1e3a8a",
                                fontSize: { xs: "1.05rem", md: "1.15rem" },
                                letterSpacing: "-0.01em",
                            }}
                        >
                            Purchase Delivery Report
                        </Typography>

                        <Chip
                            size="small"
                            label={`${filteredData.length} Records`}
                            sx={{
                                bgcolor: "#1e3a8a",
                                color: "#ffffff",
                                fontWeight: 700,
                                fontSize: "0.72rem",
                                height: 22,
                            }}
                        />

                        {/* Quick Status Chips */}
                        <Chip
                            size="small"
                            label={`All (${data.length})`}
                            onClick={() => {
                                setActiveStatusFilter("ALL");
                                setPage(1);
                            }}
                            variant={activeStatusFilter === "ALL" ? "filled" : "outlined"}
                            sx={{
                                height: 24,
                                fontSize: "0.72rem",
                                fontWeight: 700,
                                cursor: "pointer",
                                bgcolor: activeStatusFilter === "ALL" ? "#e0f2fe" : "transparent",
                                color: activeStatusFilter === "ALL" ? "#0369a1" : "#64748b",
                                borderColor: activeStatusFilter === "ALL" ? "#0284c7" : "#cbd5e1",
                            }}
                        />
                        <Chip
                            size="small"
                            label={`Completed (${data.filter((i) => i.status === "COMPLETED").length})`}
                            onClick={() => {
                                setActiveStatusFilter("COMPLETED");
                                setPage(1);
                            }}
                            variant={activeStatusFilter === "COMPLETED" ? "filled" : "outlined"}
                            sx={{
                                height: 24,
                                fontSize: "0.72rem",
                                fontWeight: 700,
                                cursor: "pointer",
                                bgcolor: activeStatusFilter === "COMPLETED" ? "#dcfce7" : "transparent",
                                color: activeStatusFilter === "COMPLETED" ? "#15803d" : "#64748b",
                                borderColor: activeStatusFilter === "COMPLETED" ? "#86efac" : "#cbd5e1",
                            }}
                        />
                        <Chip
                            size="small"
                            label={`Not Completed (${data.filter((i) => i.status !== "COMPLETED").length})`}
                            onClick={() => {
                                setActiveStatusFilter("NOT_COMPLETED");
                                setPage(1);
                            }}
                            variant={activeStatusFilter === "NOT_COMPLETED" ? "filled" : "outlined"}
                            sx={{
                                height: 24,
                                fontSize: "0.72rem",
                                fontWeight: 700,
                                cursor: "pointer",
                                bgcolor: activeStatusFilter === "NOT_COMPLETED" ? "#fee2e2" : "transparent",
                                color: activeStatusFilter === "NOT_COMPLETED" ? "#b91c1c" : "#64748b",
                                borderColor: activeStatusFilter === "NOT_COMPLETED" ? "#fca5a5" : "#cbd5e1",
                            }}
                        />
                    </Box>

                    {/* Search Input */}
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
                            placeholder="Search Order, Batch, Group, Invoice..."
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
                                width: { xs: "100%", sm: 280 },
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

                {/* 4. MAIN TABLE (FLAT TABLE WITHOUT EXPANSION MATCHING USER IMAGE) */}
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
                            overflowY: "auto",
                            overflowX: "hidden",
                        }}
                    >
                        <Table
                            size="small"
                            stickyHeader
                            aria-label="purchase delivery table"
                            sx={{ width: "100%", tableLayout: "fixed" }}
                        >
                            {/* TABLE HEAD */}
                            <TableHead>
                                {/* HEADER MATCHING USER IMAGE */}
                                <TableRow
                                    sx={{
                                        "& th": {
                                            bgcolor: "#1E3A8A",
                                            fontWeight: 700,
                                            color: "#ffffff",
                                            fontSize: "0.78rem",
                                            borderRight: "1px solid rgba(255, 255, 255, 0.15)",
                                            py: 0.9,
                                            px: 0.8,
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 5,
                                            overflow: "hidden",
                                            textOverflow: "ellipsis",
                                            whiteSpace: "nowrap",
                                        },
                                    }}
                                >
                                    <TableCell align="center" sx={{ width: "4%" }}>
                                        S.no
                                    </TableCell>
                                    <TableCell sx={{ width: "9.5%" }}>
                                        Stock group
                                    </TableCell>
                                    <TableCell sx={{ width: "21.5%" }}>
                                        Inward Batch with item name
                                    </TableCell>
                                    <TableCell sx={{ width: "12%" }}>
                                        Pur .order no
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: "10%" }}>
                                        Inward Jou no
                                    </TableCell>
                                    <TableCell sx={{ width: "12%" }}>
                                        Pur Inv no
                                    </TableCell>
                                    <TableCell sx={{ width: "12%" }}>
                                        Payment no
                                    </TableCell>
                                    <TableCell align="right" sx={{ width: "11%" }}>
                                        Payment amt
                                    </TableCell>
                                    <TableCell align="center" sx={{ width: "11%", borderRight: "none" }}>
                                        status
                                    </TableCell>
                                </TableRow>

                                {/* STICKY GRAND TOTALS ROW DIRECTLY UNDER TABLE HEADER */}
                                <TableRow
                                    sx={{
                                        bgcolor: "#f1f5f9",
                                        borderBottom: "2px solid #cbd5e1",
                                        "& th, & td": {
                                            fontWeight: 800,
                                            fontSize: "0.76rem",
                                            py: 0.8,
                                            px: 0.8,
                                            color: "#0f172a",
                                            borderRight: "1px solid #e2e8f0",
                                            position: "sticky",
                                            top: "37px",
                                            zIndex: 4,
                                            bgcolor: "#f1f5f9",
                                            overflow: "hidden",
                                            textOverflow: "ellipsis",
                                        },
                                    }}
                                >
                                    <TableCell align="center">TOTAL</TableCell>
                                    <TableCell> </TableCell>
                                    <TableCell sx={{ color: "#1e3a8a", fontWeight: 800 }}>
                                        {filteredData.length} Records
                                    </TableCell>
                                    <TableCell> </TableCell>
                                    <TableCell align="center"> </TableCell>
                                    <TableCell> </TableCell>
                                    <TableCell> </TableCell>
                                    <TableCell
                                        align="right"
                                        sx={{
                                            color: "#15803d",
                                            fontWeight: 800,
                                            fontSize: "0.78rem",
                                            whiteSpace: "nowrap",
                                        }}
                                    >
                                        {metrics.totalPaymentAmt > 0 ? formatCurrency(metrics.totalPaymentAmt) : "-"}
                                    </TableCell>
                                    <TableCell align="center" sx={{ borderRight: "none" }}>
                                        <Chip
                                            size="small"
                                            label={`${metrics.completionRate}% Done`}
                                            sx={{
                                                bgcolor: "#dcfce7",
                                                color: "#15803d",
                                                fontWeight: 800,
                                                fontSize: "0.68rem",
                                                height: 20,
                                            }}
                                        />
                                    </TableCell>
                                </TableRow>
                            </TableHead>

                            {/* TABLE BODY (NO EXPANSION - FLAT ROWS) */}
                            <TableBody>
                                {paginatedData.length > 0 ? (
                                    paginatedData.map((row, idx) => {
                                        const isCompleted = row.status === "COMPLETED";

                                        return (
                                            <TableRow
                                                key={row.id}
                                                hover
                                                sx={{
                                                    bgcolor: idx % 2 === 1 ? "#fafbfc" : "#ffffff",
                                                    borderBottom: "1px solid #e2e8f0",
                                                    "&:hover": {
                                                        bgcolor: "#f1f5f9 !important",
                                                    },
                                                }}
                                            >
                                                {/* S.NO */}
                                                <TableCell
                                                    align="center"
                                                    sx={{
                                                        fontWeight: 700,
                                                        color: "#475569",
                                                        fontSize: "0.78rem",
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                    }}
                                                >
                                                    {(page - 1) * rowsPerPage + idx + 1}
                                                </TableCell>

                                                {/* STOCK GROUP */}
                                                <TableCell
                                                    sx={{
                                                        fontSize: "0.78rem",
                                                        color: "#334155",
                                                        fontWeight: 600,
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                        wordBreak: "break-word",
                                                    }}
                                                >
                                                    {row.stockGroup}
                                                </TableCell>

                                                {/* INWARD BATCH WITH ITEM NAME */}
                                                <TableCell
                                                    sx={{
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                    }}
                                                >
                                                    <Typography
                                                        sx={{
                                                            fontSize: "0.80rem",
                                                            fontWeight: 700,
                                                            color: "#0f172a",
                                                            wordBreak: "break-word",
                                                            lineHeight: 1.25,
                                                        }}
                                                    >
                                                        {row.inwardBatchWithItemName}
                                                    </Typography>
                                                </TableCell>

                                                {/* PUR .ORDER NO */}
                                                <TableCell
                                                    sx={{
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                    }}
                                                >
                                                    <Typography
                                                        sx={{
                                                            fontWeight: 700,
                                                            color: "#1e3a8a",
                                                            fontSize: "0.78rem",
                                                            wordBreak: "break-word",
                                                        }}
                                                    >
                                                        {row.purOrderNo}
                                                    </Typography>
                                                </TableCell>

                                                {/* INWARD JOU NO */}
                                                <TableCell
                                                    align="center"
                                                    sx={{
                                                        fontWeight: 600,
                                                        color: "#475569",
                                                        fontSize: "0.78rem",
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                    }}
                                                >
                                                    {row.inwardJouNo}
                                                </TableCell>

                                                {/* PUR INV NO */}
                                                <TableCell
                                                    sx={{
                                                        fontWeight: 600,
                                                        color: "#475569",
                                                        fontSize: "0.78rem",
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                        wordBreak: "break-word",
                                                    }}
                                                >
                                                    {row.purInvNo}
                                                </TableCell>

                                                {/* PAYMENT NO */}
                                                <TableCell
                                                    sx={{
                                                        fontWeight: 600,
                                                        color: "#475569",
                                                        fontSize: "0.78rem",
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                        wordBreak: "break-word",
                                                    }}
                                                >
                                                    {row.paymentNo}
                                                </TableCell>

                                                {/* PAYMENT AMT */}
                                                <TableCell
                                                    align="right"
                                                    sx={{
                                                        fontWeight: 700,
                                                        color: row.paymentAmt ? "#0f172a" : "#94a3b8",
                                                        fontSize: "0.78rem",
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "1px solid #eef2f6",
                                                        whiteSpace: "nowrap",
                                                    }}
                                                >
                                                    {row.paymentAmt ? (
                                                        typeof row.paymentAmt === "number" ? (
                                                            formatCurrency(row.paymentAmt)
                                                        ) : (
                                                            row.paymentAmt
                                                        )
                                                    ) : (
                                                        ""
                                                    )}
                                                </TableCell>

                                                {/* STATUS */}
                                                <TableCell
                                                    align="center"
                                                    sx={{
                                                        py: 0.8,
                                                        px: 0.8,
                                                        borderRight: "none",
                                                    }}
                                                >
                                                    <Chip
                                                        size="small"
                                                        icon={
                                                            isCompleted ? (
                                                                <CheckCircleOutlineIcon sx={{ fontSize: 13 }} />
                                                            ) : (
                                                                <AccessTimeIcon sx={{ fontSize: 13 }} />
                                                            )
                                                        }
                                                        label={row.status}
                                                        sx={{
                                                            fontWeight: 800,
                                                            fontSize: "0.66rem",
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
                                                </TableCell>
                                            </TableRow>
                                        );
                                    })
                                ) : (
                                    <TableRow>
                                        <TableCell colSpan={9} align="center" sx={{ py: 6 }}>
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

                {/* 5. PINNED PAGINATION */}
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
        </Box>
    );
};

export default PurchaseDelivery;
