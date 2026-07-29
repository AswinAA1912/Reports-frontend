import React, { useState, useEffect } from "react";
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
    CircularProgress,
} from "@mui/material";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import AppLayout from "../../Layout/appLayout";
import { SalesDeliveryReportService, SalesDeliveryItem } from "../../services/salesDeliveryReport.service";

const getFunnelValue = (
    data: SalesDeliveryItem[],
    metricType: "Count" | "Tonnage",
    field: keyof Omit<SalesDeliveryItem, "Metric">
): string => {
    if (!data.length) return "-";
    
    const current = data.find((r) => r.Metric === metricType);
    const overall = data.find((r) => r.Metric === `Overall - ${metricType}`);
    
    if (!current || !overall) return "-";
    
    const curVal = current[field] as number;
    const overVal = overall[field] as number;
    
    if (metricType === "Tonnage") {
        return `${(curVal || 0).toFixed(2)}/${(overVal || 0).toFixed(2)} Ton`;
    }
    return `${Math.round(curVal || 0)}/${Math.round(overVal || 0)}`;
};

const SalesDeliveryReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const [fromDate, setFromDate] = useState(today);
    const [toDate, setToDate] = useState(today);
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [data, setData] = useState<SalesDeliveryItem[]>([]);
    const [loading, setLoading] = useState(false);
    const [filters, setFilters] = useState({
        Date: {
            from: today,
            to: today,
        },
    });

    useEffect(() => {
        loadData();
    }, [filters]);

    const loadData = async () => {
        try {
            setLoading(true);
            const res = await SalesDeliveryReportService.getSalesDeliveryCumulative({
                Fromdate: filters.Date.from,
                Todate: filters.Date.to,
            });
            setData(res.data.data || []);
        } catch (err: any) {
            console.error("Error loading sales delivery data:", err);
            toast.error(err?.response?.data?.message || "Failed to load report data ❌");
            setData([]);
        } finally {
            setLoading(false);
        }
    };

    // Excel Export
    const handleExportExcel = () => {
        try {
            if (!data.length) {
                toast.warning("No data to export");
                return;
            }

            const excelData: any[][] = [];
            excelData.push(["SALES DELIVERY FUNNEL TRACKING"]);
            excelData.push([]);
            excelData.push([
                "Sales Order",
                "Sales Invoice",
                "Printed",
                "Taken",
                "Check",
                "Dispatch",
                "Delivery",
                "Shed Sheet",
                "Metric"
            ]);

            excelData.push([
                getFunnelValue(data, "Count", "SalesOrder"),
                getFunnelValue(data, "Count", "SalesInvoice"),
                getFunnelValue(data, "Count", "Printed"),
                getFunnelValue(data, "Count", "Others1"),
                getFunnelValue(data, "Count", "Others2"),
                getFunnelValue(data, "Count", "Dispatch"),
                getFunnelValue(data, "Count", "Delivery"),
                getFunnelValue(data, "Count", "ShedSheet"),
                "Count"
            ]);

            excelData.push([
                getFunnelValue(data, "Tonnage", "SalesOrder"),
                getFunnelValue(data, "Tonnage", "SalesInvoice"),
                getFunnelValue(data, "Tonnage", "Printed"),
                getFunnelValue(data, "Tonnage", "Others1"),
                getFunnelValue(data, "Tonnage", "Others2"),
                getFunnelValue(data, "Tonnage", "Dispatch"),
                getFunnelValue(data, "Tonnage", "Delivery"),
                getFunnelValue(data, "Tonnage", "ShedSheet"),
                "Tonnage"
            ]);

            const ws = XLSX.utils.aoa_to_sheet(excelData);

            // Styling excel cells
            if (ws && ws["!ref"]) {
                const range = XLSX.utils.decode_range(ws["!ref"]);
                const borderStyle = {
                    top: { style: "thin", color: { rgb: "CFCFCF" } },
                    bottom: { style: "thin", color: { rgb: "CFCFCF" } },
                    left: { style: "thin", color: { rgb: "CFCFCF" } },
                    right: { style: "thin", color: { rgb: "CFCFCF" } }
                };

                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
                        if (!cell) continue;

                        cell.s = {
                            font: { name: "Arial", sz: 10 },
                            border: borderStyle
                        };

                        // Main Title Row
                        if (R === 0) {
                            cell.s.font = { name: "Arial", sz: 12, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                            cell.s.alignment = { horizontal: "center" };
                        }
                        // Header Row
                        else if (R === 2) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                            cell.s.fill = { fgColor: { rgb: "1E3A8A" } };
                        }
                        // Counts and Tonnage Row
                        else if (R === 3 || R === 4) {
                            cell.s.font = { name: "Arial", sz: 10, bold: true };
                            cell.s.fill = { fgColor: { rgb: "FFE5D9" } }; // Soft peach background
                            if (C === 8) {
                                cell.s.font = { name: "Arial", sz: 10, bold: true, color: { rgb: "1E3A8A" } };
                                cell.s.fill = { fgColor: { rgb: "F1F5F9" } };
                            }
                        }
                    }
                }
            }

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "SalesDelivery");
            XLSX.writeFile(wb, `SalesDelivery_Report_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
            toast.success("Excel Exported ✅");
        } catch (err) {
            console.error(err);
            toast.error("Excel Export Failed ❌");
        }
    };

    // PDF Export
    const handleExportPDF = () => {
        try {
            if (!data.length) {
                toast.warning("No data to export");
                return;
            }

            const doc = new jsPDF("landscape", "mm", "a4");
            doc.setFontSize(14);
            doc.setFont("helvetica", "bold");
            doc.text("SALES DELIVERY FUNNEL TRACKING REPORT", 148, 15, { align: "center" });

            const funnelHead = [["Sales Order", "Sales Invoice", "Printed", "Taken", "Check", "Dispatch", "Delivery", "Shed Sheet", "Metric"]];
            const funnelBody = [
                [
                    getFunnelValue(data, "Count", "SalesOrder"),
                    getFunnelValue(data, "Count", "SalesInvoice"),
                    getFunnelValue(data, "Count", "Printed"),
                    getFunnelValue(data, "Count", "Others1"),
                    getFunnelValue(data, "Count", "Others2"),
                    getFunnelValue(data, "Count", "Dispatch"),
                    getFunnelValue(data, "Count", "Delivery"),
                    getFunnelValue(data, "Count", "ShedSheet"),
                    "Count"
                ],
                [
                    getFunnelValue(data, "Tonnage", "SalesOrder"),
                    getFunnelValue(data, "Tonnage", "SalesInvoice"),
                    getFunnelValue(data, "Tonnage", "Printed"),
                    getFunnelValue(data, "Tonnage", "Others1"),
                    getFunnelValue(data, "Tonnage", "Others2"),
                    getFunnelValue(data, "Tonnage", "Dispatch"),
                    getFunnelValue(data, "Tonnage", "Delivery"),
                    getFunnelValue(data, "Tonnage", "ShedSheet"),
                    "Tonnage"
                ]
            ];

            autoTable(doc, {
                startY: 25,
                head: funnelHead,
                body: funnelBody,
                theme: "grid",
                headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255], fontStyle: "bold" },
                bodyStyles: { fontStyle: "bold" },
                styles: { fontSize: 9, cellPadding: 2.5 },
                didParseCell: (cellData) => {
                    if (cellData.column.index === 8) {
                        cellData.cell.styles.fillColor = [241, 245, 249];
                        cellData.cell.styles.textColor = [30, 58, 138];
                    } else {
                        cellData.cell.styles.fillColor = [255, 229, 204]; // soft peach background
                    }
                }
            });

            doc.save(`SalesDelivery_Report_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`);
            toast.success("PDF Exported ✅");
        } catch (err) {
            console.error(err);
            toast.error("PDF Export Failed ❌");
        }
    };

    return (
        <>
            <PageHeader
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                showPages={true}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((prev) => !prev)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                onFromDateChange={setFromDate}
                toDate={toDate}
                onToDateChange={setToDate}
                onApply={() => {
                    setFilters({
                        Date: { from: fromDate, to: toDate },
                    });
                    setDrawerOpen(false);
                }}
            />

            <AppLayout fullWidth>
                <Box px={3} pb={4} pt={4}>
                    <Typography variant="subtitle1" fontWeight="bold" color="#1e3a8a" mb={2} sx={{ letterSpacing: 0.5 }}>
                        SALES DELIVERY FUNNEL TRACKING
                    </Typography>

                    {loading ? (
                        <Box sx={{ display: "flex", justifyContent: "center", py: 8 }}>
                            <CircularProgress size={40} />
                        </Box>
                    ) : (
                        <TableContainer component={Paper} elevation={1} sx={{ borderRadius: 2, border: "1px solid #e2e8f0" }}>
                            <Table size="medium">
                                <TableHead>
                                    <TableRow sx={{ bgcolor: "#1E3A8A" }}>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Sales Order</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Sales Invoice</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Printed</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Taken</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Check</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Dispatch</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Delivery</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, borderRight: "1px solid #2448b2", fontSize: "0.9rem" }}>Shed Sheet</TableCell>
                                        <TableCell sx={{ color: "#ffffff", fontWeight: 700, py: 1.5, bgcolor: "#172d6c", width: "120px", fontSize: "0.9rem" }} align="center">Metric</TableCell>
                                    </TableRow>
                                </TableHead>
                                <TableBody>
                                    {data.length > 0 ? (
                                        <>
                                            {/* COUNT ROW */}
                                            <TableRow sx={{ "&:hover": { bgcolor: "#ecedeeff" } }}>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "SalesOrder")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "SalesInvoice")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "Printed")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "Others1")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "Others2")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "Dispatch")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "Delivery")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Count", "ShedSheet")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 800, py: 2, bgcolor: "#f1f5f9", color: "#1e3a8a", fontSize: "0.85rem" }} align="center">
                                                    Count
                                                </TableCell>
                                            </TableRow>

                                            {/* TONNAGE ROW */}
                                            <TableRow sx={{ "&:hover": { bgcolor: "#ecedeeff" } }}>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "SalesOrder")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "SalesInvoice")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "Printed")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "Others1")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "Others2")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "Dispatch")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "Delivery")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>
                                                    {getFunnelValue(data, "Tonnage", "ShedSheet")}
                                                </TableCell>
                                                <TableCell sx={{ fontWeight: 800, py: 2, bgcolor: "#f1f5f9", color: "#1e3a8a", fontSize: "0.85rem" }} align="center">
                                                    Tonnage
                                                </TableCell>
                                            </TableRow>
                                        </>
                                    ) : (
                                        <TableRow>
                                            <TableCell colSpan={9} align="center" sx={{ py: 6, color: "text.secondary" }}>
                                                No records found
                                            </TableCell>
                                        </TableRow>
                                    )}
                                </TableBody>
                            </Table>
                        </TableContainer>
                    )}
                </Box>
            </AppLayout>
        </>
    );
};

export default SalesDeliveryReport;
