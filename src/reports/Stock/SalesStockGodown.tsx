import React, { useMemo, useState } from "react";
import {
    Box,
    Table,
    TableBody,
    TableCell,
    TableContainer,
    TableHead,
    TableRow,
    Paper,
    Typography,
    Card,
    CardContent,
    Grid
} from "@mui/material";
import dayjs from "dayjs";
import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";

/* ================= DUMMY DATA GENERATOR ================= */
interface ReportData {
    purchase: number;
    acGodownPurchase: number;
    otherGodownPurchase: number;
    wtChkPending: number;
    cleaningPending: number;
    wtChkToday: number;
    cleaningToday: number;
    attyToday: number;
    sales: number;
    acGodownSales: number;
    otherGodownSales: number;
}

const generateReportData = (dateStr: string): ReportData => {
    // Exact values from the shared image for 19-08-2026
    if (dateStr === "2026-08-19") {
        return {
            purchase: 56.34,
            acGodownPurchase: 16.26,
            otherGodownPurchase: 0.27,
            wtChkPending: 97.73,
            cleaningPending: 44.55,
            wtChkToday: 16.72,
            cleaningToday: 7.81,
            attyToday: 48.66,
            sales: 34.1,
            acGodownSales: 0.0,
            otherGodownSales: 20.51,
        };
    }

    // Seeded pseudo-random generation to remain consistent per selected date
    let seed = 0;
    for (let i = 0; i < dateStr.length; i++) {
        seed += dateStr.charCodeAt(i);
    }

    const random = (min: number, max: number, offset: number) => {
        const r = Math.sin(seed + offset) * 10000;
        const val = (r - Math.floor(r)) * (max - min) + min;
        return Math.round(val * 100) / 100;
    };

    const purchase = random(30, 80, 1);
    const acGodownPurchase = random(10, 30, 2);
    const otherGodownPurchase = random(0.1, 5, 3);
    const wtChkPending = random(50, 120, 4);
    const cleaningPending = random(20, 60, 5);
    const wtChkToday = random(10, 30, 6);
    const cleaningToday = random(3, 15, 7);
    const attyToday = random(20, 60, 8);
    const sales = random(20, 50, 9);
    const acGodownSales = random(0, 10, 10) > 7 ? random(1, 10, 11) : 0.0;
    const otherGodownSales = random(10, 30, 12);

    return {
        purchase,
        acGodownPurchase,
        otherGodownPurchase,
        wtChkPending,
        cleaningPending,
        wtChkToday,
        cleaningToday,
        attyToday,
        sales,
        acGodownSales,
        otherGodownSales,
    };
};

/* ================= COMPONENT ================= */
const SalesStockGodown: React.FC = () => {
    const parentReportName = "Sales Stock Godown Tonnage Report";
    const [selectedDate, setSelectedDate] = useState("2026-08-19");
    const [drawerOpen, setDrawerOpen] = useState(false);

    const data = useMemo(() => generateReportData(selectedDate), [selectedDate]);

    // Subsection totals
    const inwardsTotal = useMemo(() => {
        return Math.round((data.purchase + data.acGodownPurchase + data.otherGodownPurchase) * 100) / 100;
    }, [data]);

    const todayInternalTotal = useMemo(() => {
        return Math.round((data.wtChkToday + data.cleaningToday + data.attyToday) * 100) / 100;
    }, [data]);

    const outwardsTotal = useMemo(() => {
        return Math.round((data.sales + data.acGodownSales + data.otherGodownSales) * 100) / 100;
    }, [data]);

    const overallTotal = useMemo(() => {
        return Math.round((inwardsTotal + todayInternalTotal + outwardsTotal) * 100) / 100;
    }, [inwardsTotal, todayInternalTotal, outwardsTotal]);


    /* ================= EXPORT LOGIC ================= */
    const exportHeaders = ["Section", "Particulars", "Tonnage"];

    const getExportRows = () => [
        ["TOTAL TONNAGE", `All Sections (as of ${dayjs(selectedDate).format("DD-MM-YYYY")})`, overallTotal],
        ["IN WARDS", "PURCHASE", data.purchase],
        ["IN WARDS", "AC GODOWN'S", data.acGodownPurchase],
        ["IN WARDS TOTAL", "OTHER GODOWN'S (Total: " + inwardsTotal + ")", data.otherGodownPurchase],
        ["INTERNAL PENDING", "WT CHK PENDING", data.wtChkPending],
        ["INTERNAL PENDING", "CLEANING PENDING", data.cleaningPending],
        ["INTERNAL TODAY", "WT CHK TODAY", data.wtChkToday],
        ["INTERNAL TODAY", "CLEANING TODAY", data.cleaningToday],
        ["INTERNAL TODAY TOTAL", "ATTY TODAY (Total: " + todayInternalTotal + ")", data.attyToday],
        ["OUT WARDS", "SALES", data.sales],
        ["OUT WARDS", "AC GODOWN'S", data.acGodownSales || "-"],
        ["OUT WARDS TOTAL", "OTHER GODOWN'S (Total: " + outwardsTotal + ")", data.otherGodownSales]
    ];

    const handleExportPDF = () => {
        const rows = getExportRows().map(row => [row[0], row[1], String(row[2])]);
        exportToPDF(`Sales_Stock_Godown_${selectedDate}`, exportHeaders, rows);
    };

    const handleExportExcel = () => {
        const rows = getExportRows().map(row => [row[0], row[1], row[2]]);
        exportToExcel(`Sales_Stock_Godown_${selectedDate}`, exportHeaders, rows);
    };

    /* ================= STYLING TOKENS (BLUISH THEME) ================= */
    const colors = {
        primaryAccent: "#1E3A8A",
        lightAccent: "#3B82F6",
        bgBlueLight: "#eff6ff",
        bgBlueMedium: "#dbeafe",
        dangerText: "#b91c1c",
        dangerBg: "#fef2f2",
        tableBorder: "#cbd5e1",
    };

    return (
        <>
            <PageHeader
                parentReportName={parentReportName}
                onExportPDF={handleExportPDF}
                onExportExcel={handleExportExcel}
            />

            {/* FLOATING FILTER DRAWER */}
            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((p) => !p)}
                onClose={() => setDrawerOpen(false)}
                fromDate={selectedDate}
                onFromDateChange={setSelectedDate}
                onApply={() => { }}
            />

            <AppLayout fullWidth>
                <Box
                    p={3}
                    sx={{
                        display: "flex",
                        flexDirection: "column",
                        gap: 2.5,
                        boxSizing: "border-box"
                    }}
                >
                    {/* QUICK METRICS CARDS */}
                    <Grid container spacing={2}>
                        <Grid item xs={12} sm={8} md={4}>
                            <Card sx={{ borderRadius: 2, boxShadow: "0 2px 8px rgba(0,0,0,0.06)", borderLeft: `5px solid ${colors.primaryAccent}` }}>
                                <CardContent sx={{ py: 1.5, px: 2, "&:last-child": { pb: 1.5 } }}>
                                    <Typography variant="caption" color="text.secondary" fontWeight={700}>
                                        INWARDS TONNAGE
                                    </Typography>
                                    <Typography variant="h5" fontWeight={800} sx={{ mt: 0.5, color: colors.primaryAccent }}>
                                        {inwardsTotal}
                                    </Typography>
                                </CardContent>
                            </Card>
                        </Grid>
                        <Grid item xs={12} sm={8} md={4}>
                            <Card sx={{ borderRadius: 2, boxShadow: "0 2px 8px rgba(0,0,0,0.06)", borderLeft: `5px solid ${colors.lightAccent}` }}>
                                <CardContent sx={{ py: 1.5, px: 2, "&:last-child": { pb: 1.5 } }}>
                                    <Typography variant="caption" color="text.secondary" fontWeight={700}>
                                        TODAY'S INTERNAL
                                    </Typography>
                                    <Typography variant="h5" fontWeight={800} sx={{ mt: 0.5, color: colors.lightAccent }}>
                                        {todayInternalTotal}
                                    </Typography>
                                </CardContent>
                            </Card>
                        </Grid>
                        <Grid item xs={12} sm={8} md={4}>
                            <Card sx={{ borderRadius: 2, boxShadow: "0 2px 8px rgba(0,0,0,0.06)", borderLeft: "5px solid #16a34a" }}>
                                <CardContent sx={{ py: 1.5, px: 2, "&:last-child": { pb: 1.5 } }}>
                                    <Typography variant="caption" color="text.secondary" fontWeight={700}>
                                        OUTWARDS TONNAGE
                                    </Typography>
                                    <Typography variant="h5" fontWeight={800} sx={{ mt: 0.5, color: "#16a34a" }}>
                                        {outwardsTotal}
                                    </Typography>
                                </CardContent>
                            </Card>
                        </Grid>
                    </Grid>

                    {/* TABLE LAYOUT */}
                    <TableContainer
                        component={Paper}
                        elevation={2}
                        sx={{
                            borderRadius: 3,
                            border: `1px solid ${colors.tableBorder}`,
                            maxHeight: "calc(100vh - 190px)",
                            overflow: "auto",
                            boxShadow: "0 4px 18px rgba(0, 0, 0, 0.08)"
                        }}
                    >
                        <Table
                            stickyHeader
                            sx={{
                                tableLayout: "fixed",
                                "& td, & th": {
                                    border: `1px solid ${colors.tableBorder}`,
                                    fontWeight: 600,
                                    fontSize: "0.9rem",
                                    py: 1.2,
                                    px: 2,
                                    textAlign: "center",
                                    overflow: "hidden",
                                    textOverflow: "ellipsis",
                                    whiteSpace: "nowrap"
                                }
                            }}
                        >
                            <TableHead>
                                <TableRow>
                                    <TableCell
                                        sx={{
                                            width: "30%",
                                            bgcolor: colors.primaryAccent,
                                            color: "#fff",
                                            fontWeight: 800,
                                            fontSize: "1.05rem",
                                            letterSpacing: 0.5,
                                            borderBottom: `2px solid ${colors.primaryAccent}`,
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 5
                                        }}
                                    >
                                        {overallTotal}
                                    </TableCell>
                                    <TableCell
                                        sx={{
                                            width: "40%",
                                            bgcolor: colors.primaryAccent,
                                            color: "#fff",
                                            fontWeight: 800,
                                            letterSpacing: 0.5,
                                            borderBottom: `2px solid ${colors.primaryAccent}`,
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 5
                                        }}
                                    >
                                        {dayjs(selectedDate).format("DD-MM-YYYY")}
                                    </TableCell>
                                    <TableCell
                                        sx={{
                                            width: "30%",
                                            bgcolor: colors.primaryAccent,
                                            color: "#fff",
                                            fontWeight: 800,
                                            letterSpacing: 1,
                                            borderBottom: `2px solid ${colors.primaryAccent}`,
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 5
                                        }}
                                    >
                                        TONNAGE
                                    </TableCell>
                                </TableRow>
                            </TableHead>
                            <TableBody>
                                {/* Row 2-4: Inwards Section */}
                                <TableRow>
                                    <TableCell rowSpan={3} sx={{ bgcolor: "#f8fafc", verticalAlign: "middle", borderBottom: `2px solid ${colors.primaryAccent}` }}>
                                        <Box display="flex" flexDirection="column" gap={0.5} alignItems="center">
                                            <Typography variant="body2" fontWeight={700} color="text.secondary">
                                                IN WARDS
                                            </Typography>
                                            <Typography variant="body1" fontWeight={800} sx={{ color: colors.primaryAccent }}>
                                                {inwardsTotal}
                                            </Typography>
                                        </Box>
                                    </TableCell>
                                    <TableCell>
                                        PURCHASE
                                    </TableCell>
                                    <TableCell>
                                        {data.purchase}
                                    </TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell>AC GODOWN'S</TableCell>
                                    <TableCell>{data.acGodownPurchase}</TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>OTHER GODOWN'S</TableCell>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>{data.otherGodownPurchase}</TableCell>
                                </TableRow>

                                {/* Row 5-9: Internal Section */}
                                <TableRow>
                                    <TableCell rowSpan={5} sx={{ bgcolor: "#f8fafc", verticalAlign: "middle", borderBottom: `2px solid ${colors.primaryAccent}` }}>
                                        <Box display="flex" flexDirection="column" gap={0.5} alignItems="center">
                                            <Typography variant="body2" fontWeight={700} color="text.secondary">
                                                INTERNAL
                                            </Typography>
                                            <Typography variant="body1" fontWeight={800} sx={{ color: colors.primaryAccent }}>
                                                {todayInternalTotal}
                                            </Typography>
                                        </Box>
                                    </TableCell>
                                    <TableCell>
                                        WT CHK PENDING
                                    </TableCell>
                                    <TableCell>
                                        {data.wtChkPending}
                                    </TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell>
                                        CLEANING PENDING
                                    </TableCell>
                                    <TableCell>
                                        {data.cleaningPending}
                                    </TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell>WT CHK TODAY</TableCell>
                                    <TableCell>{data.wtChkToday}</TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell>CLEANING TODAY</TableCell>
                                    <TableCell>{data.cleaningToday}</TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>ATTY TODAY</TableCell>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>{data.attyToday}</TableCell>
                                </TableRow>

                                {/* Row 10-12: Outwards Section */}
                                <TableRow>
                                    <TableCell rowSpan={3} sx={{ bgcolor: "#f8fafc", verticalAlign: "middle", borderBottom: `2px solid ${colors.primaryAccent}` }}>
                                        <Box display="flex" flexDirection="column" gap={0.5} alignItems="center">
                                            <Typography variant="body2" fontWeight={700} color="text.secondary">
                                                OUT WARDS
                                            </Typography>
                                            <Typography variant="body1" fontWeight={800} sx={{ color: colors.primaryAccent }}>
                                                {outwardsTotal}
                                            </Typography>
                                        </Box>
                                    </TableCell>
                                    <TableCell>
                                        SALES
                                    </TableCell>
                                    <TableCell>
                                        {data.sales}
                                    </TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell>AC GODOWN'S</TableCell>
                                    <TableCell>{data.acGodownSales || "-"}</TableCell>
                                </TableRow>

                                <TableRow>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>OTHER GODOWN'S</TableCell>
                                    <TableCell sx={{ borderBottom: `2px solid ${colors.primaryAccent}` }}>{data.otherGodownSales}</TableCell>
                                </TableRow>
                            </TableBody>
                        </Table>
                    </TableContainer>
                </Box>
            </AppLayout>
        </>
    );
};

export default SalesStockGodown;
