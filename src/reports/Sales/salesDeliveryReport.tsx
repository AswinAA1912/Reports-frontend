import React from "react";
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
} from "@mui/material";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";
import PageHeader from "../../Layout/PageHeader";

// Static metrics matching the funnel tracking screenshot precisely
const METRICS = {
    count: {
        order: "1000/1000",
        invoice: "890/1000",
        printed: "780/890",
        taken: "750/780",
        check: "700/750",
        dispatch: "680/700",
        delivery: "650/680",
        shedSheet: ""
    },
    tonnage: {
        order: "45/45 Ton",
        invoice: "38/45 Ton",
        printed: "35/38 Ton",
        taken: "33/35 Ton",
        check: "31/33 Ton",
        dispatch: "29/31 Ton",
        delivery: "28/29 Ton",
        shedSheet: ""
    }
};

const SalesDeliveryReport: React.FC = () => {
    // Excel Export
    const handleExportExcel = () => {
        try {
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
                METRICS.count.order,
                METRICS.count.invoice,
                METRICS.count.printed,
                METRICS.count.taken,
                METRICS.count.check,
                METRICS.count.dispatch,
                METRICS.count.delivery,
                METRICS.count.shedSheet,
                "Count"
            ]);
            excelData.push([
                METRICS.tonnage.order,
                METRICS.tonnage.invoice,
                METRICS.tonnage.printed,
                METRICS.tonnage.taken,
                METRICS.tonnage.check,
                METRICS.tonnage.dispatch,
                METRICS.tonnage.delivery,
                METRICS.tonnage.shedSheet,
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
            const doc = new jsPDF("landscape", "mm", "a4");
            doc.setFontSize(14);
            doc.setFont("helvetica", "bold");
            doc.text("SALES DELIVERY FUNNEL TRACKING REPORT", 148, 15, { align: "center" });

            const funnelHead = [["Sales Order", "Sales Invoice", "Printed", "Taken", "Check", "Dispatch", "Delivery", "Shed Sheet", "Metric"]];
            const funnelBody = [
                [
                    METRICS.count.order,
                    METRICS.count.invoice,
                    METRICS.count.printed,
                    METRICS.count.taken,
                    METRICS.count.check,
                    METRICS.count.dispatch,
                    METRICS.count.delivery,
                    METRICS.count.shedSheet,
                    "Count"
                ],
                [
                    METRICS.tonnage.order,
                    METRICS.tonnage.invoice,
                    METRICS.tonnage.printed,
                    METRICS.tonnage.taken,
                    METRICS.tonnage.check,
                    METRICS.tonnage.dispatch,
                    METRICS.tonnage.delivery,
                    METRICS.tonnage.shedSheet,
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
                didParseCell: (data) => {
                    if (data.column.index === 8) {
                        data.cell.styles.fillColor = [241, 245, 249];
                        data.cell.styles.textColor = [30, 58, 138];
                    } else {
                        data.cell.styles.fillColor = [255, 229, 204]; // soft peach background
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
        <Box sx={{ width: "100%", overflowX: "hidden", minHeight: "100vh", bgcolor: "#f8fafc" }}>
            <PageHeader
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
                showPages={true}
            />

            <Box px={3} pb={4} pt={4}>
                {/* DYNAMIC FUNNEL SUMMARY GRID */}
                <Typography variant="subtitle1" fontWeight="bold" color="#1e3a8a" mb={2} sx={{ letterSpacing: 0.5 }}>
                    SALES DELIVERY FUNNEL TRACKING
                </Typography>
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
                            {/* COUNT ROW */}
                            <TableRow sx={{ "&:hover": { bgcolor: "#ecedeeff" } }}>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.order}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.invoice}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.printed}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.taken}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.check}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.dispatch}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.delivery}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.count.shedSheet}</TableCell>
                                <TableCell sx={{ fontWeight: 800, py: 2, bgcolor: "#f1f5f9", color: "#1e3a8a", fontSize: "0.85rem" }} align="center">Count</TableCell>
                            </TableRow>
                            {/* TONNAGE ROW */}
                            <TableRow sx={{ "&:hover": { bgcolor: "#ecedeeff" } }}>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.order}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.invoice}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.printed}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.taken}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.check}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.dispatch}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.delivery}</TableCell>
                                <TableCell sx={{ fontWeight: 700, borderRight: "1px solid #e2e8f0", py: 2, color: "#1e293b", fontSize: "0.85rem" }}>{METRICS.tonnage.shedSheet}</TableCell>
                                <TableCell sx={{ fontWeight: 800, py: 2, bgcolor: "#f1f5f9", color: "#1e3a8a", fontSize: "0.85rem" }} align="center">Tonnage</TableCell>
                            </TableRow>
                        </TableBody>
                    </Table>
                </TableContainer>
            </Box>
        </Box>
    );
};

export default SalesDeliveryReport;
