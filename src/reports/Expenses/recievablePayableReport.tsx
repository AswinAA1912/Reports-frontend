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
    TextField,
    CircularProgress,
    Menu,
    MenuItem,
    Checkbox,
} from "@mui/material";

import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import { toast } from "react-toastify";


import AppLayout from "../../Layout/appLayout";
import PageHeader, { ToggleMode } from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import CommonPagination from "../../Components/CommonPagination";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import { RecievablePayableReportService, RecievablePayableItem } from "../../services/recievablePayableReport.service";

/* ================= HELPERS ================= */

const formatINR = (value: number) =>
    new Intl.NumberFormat("en-IN", {
        style: "currency",
        currency: "INR",
    }).format(value || 0);

/* ================= COMPONENT ================= */

const RecievablePayableReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    const [rows, setRows] = useState<RecievablePayableItem[]>([]);
    const [loading, setLoading] = useState(false);
    const [toggleMode, setToggleMode] = useState<ToggleMode>("Abstract"); // Abstract = Receivable, Expanded = Payable

    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);
    const [drawerOpen, setDrawerOpen] = useState(false);

    const [fromDate, setFromDate] = useState(today);
    const [accountSearchText, setAccountSearchText] = useState("");
    const [selectedAccountNames, setSelectedAccountNames] = useState<string[]>([]);
    const [accountFilterAnchor, setAccountFilterAnchor] = useState<null | HTMLElement>(null);

    const [filters, setFilters] = useState({
        Date: {
            from: today,
        },
    });

    /* ================= DATA LOADING ================= */

    useEffect(() => {
        loadReport();
    }, [filters.Date.from, toggleMode]);

    const loadReport = async () => {
        try {
            setLoading(true);
            const apiCall =
                toggleMode === "Abstract"
                    ? RecievablePayableReportService.getReceivables
                    : RecievablePayableReportService.getPayables;

            const res = await apiCall({ Fromdate: filters.Date.from });
            setRows(res.data.data || []);
            setSelectedAccountNames([]);
            setAccountSearchText("");
            setPage(1);
        } catch (err: any) {
            console.error("Error loading report data:", err);
            toast.error(err?.response?.data?.message || "Failed to load report data ❌");
            setRows([]);
        } finally {
            setLoading(false);
        }
    };

    /* ================= FILTERS & SEARCH ================= */

    // Unique account names list
    const uniqueAccountNames = useMemo(() => {
        const names = new Set(rows.map((r) => r.Account_name).filter(Boolean));
        return Array.from(names).sort((a, b) => a.localeCompare(b));
    }, [rows]);

    // Account list filtered by search input inside dropdown menu
    const filteredAccountOptions = useMemo(() => {
        return uniqueAccountNames.filter((name) =>
            name.toLowerCase().includes(accountSearchText.toLowerCase())
        );
    }, [uniqueAccountNames, accountSearchText]);

    // Apply header filter selection on account names
    const filteredBySearch = useMemo(() => {
        if (selectedAccountNames.length === 0) return rows;
        return rows.filter(
            (row) =>
                row.Account_name && selectedAccountNames.includes(row.Account_name)
        );
    }, [rows, selectedAccountNames]);

    const {
        sortConfig,
        rangeFilter,
        setRangeFilter,
        filterAnchor,
        setFilterAnchor,
        activeHeader,
        clearRangeFilter,
        getMinMax,
        handleSort,
        openFilter,
        filteredAndSortedData,
    } = useNumericalFilter(filteredBySearch, ["Bal_Amount"]);

    const finalData = useMemo(() => {
        let result = [...filteredAndSortedData];

        if (!sortConfig) {
            // Default order: invoice_date descending (from today's date to lowest)
            result.sort((a, b) => {
                const dateA = a.invoice_date ? dayjs(a.invoice_date).valueOf() : 0;
                const dateB = b.invoice_date ? dayjs(b.invoice_date).valueOf() : 0;
                return dateB - dateA;
            });
        }

        return result;
    }, [filteredAndSortedData, sortConfig]);

    /* ================= PAGINATION ================= */

    const paginatedRows = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        const end = start + rowsPerPage;
        return finalData.slice(start, end);
    }, [finalData, page, rowsPerPage]);

    useEffect(() => {
        setPage(1);
    }, [finalData.length]);

    /* ================= CALCULATIONS ================= */

    const totalAmount = useMemo(() => {
        return finalData.reduce((sum, r) => sum + (Number(r.Bal_Amount) || 0), 0);
    }, [finalData]);

    /* ================= EXPORT ================= */

    const reportTitle = toggleMode === "Abstract" ? "Receivable Report" : "Payable Report";

    const handleExportExcel = () => {
        if (!finalData.length) {
            toast.warning("No data to export");
            return;
        }
        const exportRows = finalData.map((row, index) => ({
            "S.No": index + 1,
            "Account Name": row.Account_name || "",
            "Invoice Date": row.invoice_date ? dayjs(row.invoice_date).format("DD/MM/YYYY") : "",
            "Invoice No": row.invoice_no || "",
            "Balance Amount": row.Bal_Amount || 0,
        }));

        const worksheet = XLSX.utils.json_to_sheet(exportRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, reportTitle);
        XLSX.writeFile(
            workbook,
            `${reportTitle.replace(" ", "_")}_${dayjs().format("DDMMYYYY")}.xlsx`
        );
        toast.success("Excel exported successfully 📊");
    };

    const handleExportPDF = () => {
        if (!finalData.length) {
            toast.warning("No data to export");
            return;
        }
        const doc = new jsPDF("p", "mm", "a4");
        doc.setFontSize(14);
        doc.text(`${reportTitle} - As of ${dayjs(filters.Date.from).format("DD/MM/YYYY")}`, 14, 12);

        const tableHeaders = [["S.No", "Account Name", "Invoice Date", "Invoice No", "Balance Amount"]];
        const tableBody = finalData.map((row, index) => [
            index + 1,
            row.Account_name || "",
            row.invoice_date ? dayjs(row.invoice_date).format("DD/MM/YYYY") : "",
            row.invoice_no || "",
            row.Bal_Amount ? row.Bal_Amount.toFixed(2) : "0.00",
        ]);

        // Add Total Row to PDF body
        tableBody.push([
            "",
            "Total",
            "",
            "",
            totalAmount.toFixed(2),
        ]);

        autoTable(doc, {
            startY: 18,
            head: tableHeaders,
            body: tableBody,
            styles: { fontSize: 8 },
            headStyles: { fillColor: [30, 58, 138] },
            didParseCell: (data) => {
                if (data.row.index === tableBody.length - 1) {
                    data.cell.styles.fontStyle = "bold";
                    if (data.column.index === 4) {
                        data.cell.styles.halign = "right";
                    }
                } else if (data.column.index === 4) {
                    data.cell.styles.halign = "right";
                }
            }
        });

        doc.save(`${reportTitle.replace(" ", "_")}_${dayjs().format("DDMMYYYY")}.pdf`);
        toast.success("PDF exported successfully 📄");
    };

    /* ================= INTERACTION HANDLERS ================= */

    const handleToggleAccount = (name: string) => {
        setSelectedAccountNames((prev) => {
            const next = prev.includes(name)
                ? prev.filter((x) => x !== name)
                : [...prev, name];
            return next;
        });
        setPage(1);
    };

    const handleSelectAllAccounts = () => {
        if (selectedAccountNames.length === uniqueAccountNames.length) {
            setSelectedAccountNames([]);
        } else {
            setSelectedAccountNames(uniqueAccountNames);
        }
        setPage(1);
    };

    /* ================= RENDER ================= */

    const headerStyle = {
        color: "#fff",
        fontWeight: 600,
        backgroundColor: "#1E3A8A",
    };

    return (
        <>
            <PageHeader
                toggleMode={toggleMode}
                onToggleChange={(mode) => setToggleMode(mode)}
                onExportExcel={handleExportExcel}
                onExportPDF={handleExportPDF}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((prev) => !prev)}
                onClose={() => setDrawerOpen(false)}
                fromDate={fromDate}
                onFromDateChange={setFromDate}
                onApply={() =>
                    setFilters({
                        Date: { from: fromDate },
                    })
                }
            />

            <AppLayout fullWidth>
                <Box sx={{ mt: 1, overflow: "hidden" }}>
                    {loading ? (
                        <Box sx={{ display: "flex", justifyContent: "center", py: 8 }}>
                            <CircularProgress size={40} />
                        </Box>
                    ) : (
                        <>
                            <TableContainer
                                component={Paper}
                                sx={{
                                    maxHeight: "calc(100vh - 115px)",
                                    borderRadius: 0,
                                    border: "1px solid #cbd5e1",
                                    overflow: "auto",
                                    "& th, & td": {
                                        fontSize: "0.75rem",
                                        py: 1,
                                    },
                                }}
                            >
                                <Table size="small" stickyHeader sx={{ tableLayout: "fixed", width: "100%" }}>
                                    <TableHead>
                                        <TableRow>
                                            <TableCell sx={{ ...headerStyle, width: "60px" }}>S.No</TableCell>
                                            <TableCell
                                                sx={{ ...headerStyle, cursor: "pointer" }}
                                                onClick={(e) => setAccountFilterAnchor(e.currentTarget)}
                                            >
                                                Account Name
                                            </TableCell>
                                            <TableCell sx={{ ...headerStyle, width: "110px" }}>Invoice Date</TableCell>
                                            <TableCell sx={{ ...headerStyle, width: "160px" }}>Invoice No</TableCell>
                                            <TableCell sx={{ ...headerStyle, width: "160px" }} align="right">
                                                <SortableHeaderLabel
                                                    label="Balance Amount"
                                                    columnKey="Bal_Amount"
                                                    sortConfig={sortConfig}
                                                    onSort={handleSort}
                                                    onOpenFilter={(e) => openFilter(e, "Bal_Amount")}
                                                />
                                            </TableCell>
                                        </TableRow>
                                    </TableHead>

                                    <TableBody>
                                        {/* Total Row (Top) */}
                                        {finalData.length > 0 && (
                                            <TableRow sx={{ backgroundColor: "#F8FAFC", borderBottom: "2px solid #cbd5e1" }}>
                                                <TableCell colSpan={4} sx={{ fontWeight: 700, fontSize: "0.8rem", color: "#1E3A8A" }}>
                                                    Total
                                                </TableCell>
                                                <TableCell align="right" sx={{ fontWeight: 700, fontSize: "0.8rem", color: "#1E3A8A" }}>
                                                    {formatINR(totalAmount)}
                                                </TableCell>
                                            </TableRow>
                                        )}

                                        {paginatedRows.length > 0 ? (
                                            paginatedRows.map((row, i) => (
                                                <TableRow key={row.Acc_Id + i} hover>
                                                    <TableCell>{(page - 1) * rowsPerPage + i + 1}</TableCell>
                                                    <TableCell sx={{ fontWeight: 500, color: "#1E293B", wordBreak: "break-word", whiteSpace: "normal" }}>
                                                        {row.Account_name}
                                                    </TableCell>
                                                    <TableCell>
                                                        {row.invoice_date
                                                            ? dayjs(row.invoice_date).format("DD/MM/YYYY")
                                                            : "-"}
                                                    </TableCell>
                                                    <TableCell sx={{ wordBreak: "break-all", whiteSpace: "normal" }}>{row.invoice_no || "-"}</TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 600 }}>
                                                        {formatINR(row.Bal_Amount)}
                                                    </TableCell>
                                                </TableRow>
                                            ))
                                        ) : (
                                            <TableRow>
                                                <TableCell colSpan={5} align="center" sx={{ py: 6, color: "text.secondary" }}>
                                                    No records found.
                                                </TableCell>
                                            </TableRow>
                                        )}
                                    </TableBody>
                                </Table>
                            </TableContainer>

                            {/* Pagination */}
                            {finalData.length > 0 && (
                                <CommonPagination
                                    totalRows={finalData.length}
                                    page={page}
                                    rowsPerPage={rowsPerPage}
                                    onPageChange={setPage}
                                    onRowsPerPageChange={setRowsPerPage}
                                />
                            )}
                        </>
                    )}
                </Box>
            </AppLayout>

            {/* Account Name Header Filter Menu */}
            <Menu
                anchorEl={accountFilterAnchor}
                open={Boolean(accountFilterAnchor)}
                onClose={() => setAccountFilterAnchor(null)}
            >
                <Box p={1.5} sx={{ minWidth: 260 }}>
                    {/* Search Field */}
                    <TextField
                        size="small"
                        fullWidth
                        placeholder="Search Account Name..."
                        value={accountSearchText}
                        onChange={(e) => setAccountSearchText(e.target.value)}
                        sx={{ mb: 1 }}
                        autoFocus
                    />

                    {/* All option */}
                    <MenuItem dense onClick={handleSelectAllAccounts}>
                        <Checkbox
                            size="small"
                            checked={
                                selectedAccountNames.length === 0 ||
                                selectedAccountNames.length === uniqueAccountNames.length
                            }
                        />
                        All
                    </MenuItem>

                    {/* Checkbox options list */}
                    <Box sx={{ maxHeight: 250, overflow: "auto" }}>
                        {filteredAccountOptions.map((name) => {
                            const isChecked = selectedAccountNames.includes(name);
                            return (
                                <MenuItem key={name} dense onClick={() => handleToggleAccount(name)}>
                                    <Checkbox size="small" checked={isChecked} />
                                    {name}
                                </MenuItem>
                            );
                        })}
                    </Box>
                </Box>
            </Menu>

            {/* Numerical Filter Menu */}
            <NumericalFilterMenu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor)}
                onClose={() => setFilterAnchor(null)}
                activeHeader={activeHeader}
                min={activeHeader ? getMinMax(activeHeader).min : 0}
                max={activeHeader ? getMinMax(activeHeader).max : 100}
                rangeFilter={rangeFilter}
                onRangeChange={(key, range) => setRangeFilter((p) => ({ ...p, [key]: range }))}
                onClear={clearRangeFilter}
            />
        </>
    );
};

export default RecievablePayableReport;
