import React, { useEffect, useState, useMemo } from "react";
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
    IconButton,
    Collapse,
    Autocomplete,
    TextField,
    MenuItem,
    Menu,
    Button,
} from "@mui/material";

import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import KeyboardArrowUpIcon from "@mui/icons-material/KeyboardArrowUp";
import FilterAltIcon from "@mui/icons-material/FilterAlt";
import dayjs from "dayjs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx-js-style";
import { toast } from "react-toastify";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { bankboxService } from "../../services/bankbox.service";
import CommonPagination from "../../Components/CommonPagination";

interface ChequeRowProps {
    row: any;
    index: number;
}

const ChequeRow: React.FC<ChequeRowProps> = ({ row, index }) => {
    const [open, setOpen] = useState(false);

    const formatChequeNum = (v: any) => {
        const num = Number(v);
        if (isNaN(num) || num === 0) return "-";
        return "₹" + new Intl.NumberFormat("en-IN", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
        }).format(num);
    };

    return (
        <>
            <TableRow hover onClick={() => setOpen(!open)} sx={{ cursor: "pointer", "& > *": { borderBottom: "unset" } }}>
                <TableCell align="center" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600 }}>
                    <IconButton size="small">
                        {open ? <KeyboardArrowUpIcon /> : <KeyboardArrowDownIcon />}
                    </IconButton>
                    {index}
                </TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>
                    {row.receipt_date ? dayjs(row.receipt_date).format("DD-MM-YYYY") : ""}
                </TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>{row.receipt_invoice_no}</TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>{row.voucherTypeGet}</TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 700 }}>{row.creditAccountGet}</TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>{row.check_no}</TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>
                    {row.check_date ? dayjs(row.check_date).format("DD-MM-YYYY") : ""}
                </TableCell>
                <TableCell sx={{ borderRight: "1px solid #e2e8f0" }}>
                    {row.bank_date ? dayjs(row.bank_date).format("DD-MM-YYYY") : ""}
                </TableCell>
                <TableCell align="right" sx={{ borderRight: "1px solid #e2e8f0", fontWeight: 600 }}>
                    {formatChequeNum(row.debit_amount)}
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 600 }}>
                    {formatChequeNum(row.credit_amount)}
                </TableCell>
            </TableRow>
            <TableRow>
                <TableCell style={{ paddingBottom: 0, paddingTop: 0 }} colSpan={10}>
                    <Collapse in={open} timeout="auto" unmountOnExit>
                        <Box sx={{ margin: 2 }}>
                            {/* Sales Against Reference */}
                            <Typography variant="subtitle2" fontWeight={700} gutterBottom component="div" color="#1E3A8A">
                                Sales Against Reference
                            </Typography>
                            <TableContainer component={Paper} variant="outlined" sx={{ mb: 3 }}>
                                <Table size="small">
                                    <TableHead>
                                        <TableRow sx={{ bgcolor: "#f8fafc" }}>
                                            <TableCell sx={{ fontWeight: 600 }}>S.No</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Date</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Sales Invoice No</TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 600 }}>Invoice Value</TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 600 }}>Payment Amount</TableCell>
                                        </TableRow>
                                    </TableHead>
                                    <TableBody>
                                        {!row.billRef || row.billRef.length === 0 ? (
                                            <TableRow>
                                                <TableCell colSpan={5} align="center" sx={{ color: "#94a3b8", py: 2 }}>
                                                    No sales reference found.
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            <>
                                                {row.billRef.map((bill: any, idx: number) => (
                                                    <TableRow key={idx}>
                                                        <TableCell>{idx + 1}</TableCell>
                                                        <TableCell>
                                                            {bill.billDate ? dayjs(bill.billDate).format("DD-MM-YYYY") : ""}
                                                        </TableCell>
                                                        <TableCell>{bill.invoiceVoucherNumber}</TableCell>
                                                        <TableCell align="right">{formatChequeNum(bill.invoiceValue)}</TableCell>
                                                        <TableCell align="right">{formatChequeNum(bill.paidAmount)}</TableCell>
                                                    </TableRow>
                                                ))}
                                                <TableRow sx={{ bgcolor: "#f1f5f9" }}>
                                                    <TableCell colSpan={4} align="right" sx={{ fontWeight: 700 }}>Total</TableCell>
                                                    <TableCell align="right" sx={{ fontWeight: 700 }}>
                                                        {formatChequeNum(row.billRef.reduce((sum: number, b: any) => sum + Number(b.paidAmount || 0), 0))}
                                                    </TableCell>
                                                </TableRow>
                                            </>
                                        )}
                                    </TableBody>
                                </Table>
                            </TableContainer>

                            {/* Contra Reference */}
                            <Typography variant="subtitle2" fontWeight={700} gutterBottom component="div" color="#1E3A8A">
                                Contra Reference
                            </Typography>
                            <TableContainer component={Paper} variant="outlined">
                                <Table size="small">
                                    <TableHead>
                                        <TableRow sx={{ bgcolor: "#f8fafc" }}>
                                            <TableCell sx={{ fontWeight: 600 }}>Voucher No</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Date</TableCell>
                                            <TableCell align="right" sx={{ fontWeight: 600 }}>Amount</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Cheque Date</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Bank Date</TableCell>
                                            <TableCell sx={{ fontWeight: 600 }}>Narration</TableCell>
                                        </TableRow>
                                    </TableHead>
                                    <TableBody>
                                        {!row.contraRef || row.contraRef.length === 0 ? (
                                            <TableRow>
                                                <TableCell colSpan={6} align="center" sx={{ color: "#94a3b8", py: 2 }}>
                                                    No contra reference found.
                                                </TableCell>
                                            </TableRow>
                                        ) : (
                                            row.contraRef.map((contra: any, idx: number) => (
                                                <TableRow key={idx}>
                                                    <TableCell>{contra.contraVoucherNumber}</TableCell>
                                                    <TableCell>
                                                        {contra.contraDate ? dayjs(contra.contraDate).format("DD-MM-YYYY") : ""}
                                                    </TableCell>
                                                    <TableCell align="right">{formatChequeNum(contra.contraAmount)}</TableCell>
                                                    <TableCell>
                                                        {contra.chequeDate ? dayjs(contra.chequeDate).format("DD-MM-YYYY") : ""}
                                                    </TableCell>
                                                    <TableCell>
                                                        {contra.bankDate ? dayjs(contra.bankDate).format("DD-MM-YYYY") : ""}
                                                    </TableCell>
                                                    <TableCell>{contra.narration}</TableCell>
                                                </TableRow>
                                            ))
                                        )}
                                    </TableBody>
                                </Table>
                            </TableContainer>
                        </Box>
                    </Collapse>
                </TableCell>
            </TableRow>
        </>
    );
};

const ChequeTransactionReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");
    const initialFromDate = dayjs().subtract(365, "day").format("YYYY-MM-DD");

    const [loading, setLoading] = useState(false);
    const [drawerOpen, setDrawerOpen] = useState(false);
    const [chequeData, setChequeData] = useState<any[]>([]);

    // Header Filters (Party Name, Voucher Type, Cheque No)
    const [activeHeader, setActiveHeader] = useState<string | null>(null);
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [searchText, setSearchText] = useState("");
    const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({
        creditAccountGet: [],
        voucherTypeGet: [],
        check_no: []
    });

    const [creditAccounts, setCreditAccounts] = useState<{ value: any; label: string }[]>([]);
    const [voucherTypes, setVoucherTypes] = useState<{ value: any; label: string }[]>([]);

    const handleHeaderClick = (e: React.MouseEvent<HTMLElement>, columnKey: string) => {
        setActiveHeader(columnKey);
        setSearchText("");
        setFilterAnchor(e.currentTarget);
    };

    const [chequeFilters, setChequeFilters] = useState({
        Fromdate: initialFromDate,
        Todate: today,
        debitAccount: { value: "", label: "ALL" } as { value: any; label: string },
        creditAccount: { value: "", label: "ALL" } as { value: any; label: string },
        voucherType: { value: "", label: "ALL" } as { value: any; label: string },
        partyType: "Pending Party",
        chequeAccounts: [] as { value: any; label: string }[]
    });

    useEffect(() => {
        bankboxService.getChequeAccounts()
            .then(data => {
                if (data && Array.isArray(data)) {
                    setChequeFilters(pre => ({ ...pre, chequeAccounts: data }));
                }
            })
            .catch(e => console.error("Failed to load cheque accounts", e));

        bankboxService.getChequeCreditAccounts()
            .then(data => {
                if (data && Array.isArray(data)) {
                    setCreditAccounts(data);
                }
            })
            .catch(e => console.error("Failed to load credit accounts", e));

        bankboxService.getChequeVoucherTypes()
            .then(data => {
                if (data && Array.isArray(data)) {
                    setVoucherTypes(data);
                }
            })
            .catch(e => console.error("Failed to load voucher types", e));
    }, []);

    const fetchChequeData = async () => {
        const { Fromdate, Todate, debitAccount } = chequeFilters;
        if (!debitAccount.value) {
            toast.error("Please Select Debit Account");
            return;
        }

        try {
            setLoading(true);
            const data = await bankboxService.getChequeTransactions({
                Fromdate,
                Todate,
                debitAccount: debitAccount.value
            });
            if (data && Array.isArray(data)) {
                setChequeData(data);
            } else {
                setChequeData([]);
            }
        } catch (e) {
            console.error(e);
            toast.error("Failed to load cheque transaction data ❌");
        } finally {
            setLoading(false);
        }
    };

    const filteredChequeData = useMemo(() => {
        return chequeData.filter((item: any) => {
            const matchesCredit = chequeFilters.creditAccount.value === "" ||
                Number(item.credit_ledger) === Number(chequeFilters.creditAccount.value);
            const matchesVoucher = chequeFilters.voucherType.value === "" ||
                Number(item.receipt_voucher_type_id) === Number(chequeFilters.voucherType.value);
            const matchesPartyType =
                chequeFilters.partyType === "ALL" ||
                (chequeFilters.partyType === "Pending Party" && (!item.contraRef || item.contraRef.length === 0)) ||
                (chequeFilters.partyType === "Payed Party" && item.contraRef && item.contraRef.length > 0);

            // Filter for Pending Cheques: Bank Date is null or upcoming, or Cheque Date is upcoming
            const bankDate = item.bank_date ? dayjs(item.bank_date) : null;
            const chequeDate = item.check_date ? dayjs(item.check_date) : null;
            const today = dayjs().startOf("day");

            const isBankDateUpcoming = bankDate && bankDate.isAfter(today);
            const isChequeDateUpcoming = chequeDate && chequeDate.isAfter(today);
            const isNoBankDate = !bankDate || !item.bank_date;

            const isPending = isNoBankDate || isBankDateUpcoming || isChequeDateUpcoming;

            const matchesPending = chequeFilters.partyType !== "Pending Party" || isPending;

            if (!(matchesCredit && matchesVoucher && matchesPartyType && matchesPending)) return false;

            // Header Column Filters
            const creditFilter = columnFilters.creditAccountGet || [];
            if (creditFilter.length > 0 && !creditFilter.includes(item.creditAccountGet)) return false;

            const voucherFilter = columnFilters.voucherTypeGet || [];
            if (voucherFilter.length > 0 && !voucherFilter.includes(item.voucherTypeGet)) return false;

            const checkNoFilter = columnFilters.check_no || [];
            if (checkNoFilter.length > 0 && !checkNoFilter.includes(item.check_no)) return false;

            return true;
        });
    }, [chequeData, chequeFilters, columnFilters]);

    const chequeTotals = useMemo(() => {
        let debitSum = 0;
        let creditSum = 0;
        filteredChequeData.forEach((item: any) => {
            debitSum += Number(item.debit_amount || 0);
            creditSum += Number(item.credit_amount || 0);
        });
        return {
            debit: debitSum,
            credit: creditSum
        };
    }, [filteredChequeData]);

    const [chequePage, setChequePage] = useState(1);
    const [chequeRowsPerPage, setChequeRowsPerPage] = useState(100);

    const paginatedChequeData = useMemo(() => {
        return filteredChequeData.slice((chequePage - 1) * chequeRowsPerPage, chequePage * chequeRowsPerPage);
    }, [filteredChequeData, chequePage, chequeRowsPerPage]);

    useEffect(() => {
        setChequePage(1);
    }, [filteredChequeData]);

    const formatNum = (v: any) => {
        const num = Number(v);
        if (isNaN(num) || num === 0) return "-";
        return "₹" + new Intl.NumberFormat("en-IN", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
        }).format(num);
    };

    const handleChequeExportExcel = () => {
        try {
            const excelData: any[][] = [];
            excelData.push(["CHEQUE TRANSACTION DETAILS"]);
            excelData.push([]);
            excelData.push([
                "S.No",
                "Rec.Date",
                "Rec.No",
                "VchType",
                "Party Name",
                "Chq.No",
                "Chq.Date",
                "Bank.Date",
                "Debit",
                "Credit"
            ]);

            filteredChequeData.forEach((row: any, idx: number) => {
                excelData.push([
                    idx + 1,
                    row.receipt_date ? dayjs(row.receipt_date).format("DD-MM-YYYY") : "",
                    row.receipt_invoice_no || "",
                    row.voucherTypeGet || "",
                    row.creditAccountGet || "",
                    row.check_no || "",
                    row.check_date ? dayjs(row.check_date).format("DD-MM-YYYY") : "",
                    row.bank_date ? dayjs(row.bank_date).format("DD-MM-YYYY") : "",
                    row.debit_amount || 0,
                    row.credit_amount || 0
                ]);
            });

            const ws = XLSX.utils.aoa_to_sheet(excelData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Cheque Transactions");
            XLSX.writeFile(wb, `Cheque_Transactions_${dayjs().format("YYYYMMDD_HHmmss")}.xlsx`);
            toast.success("Excel Exported ✅");
        } catch (err) {
            console.error(err);
            toast.error("Excel Export Failed ❌");
        }
    };

    const handleChequeExportPDF = () => {
        try {
            const doc = new jsPDF("landscape", "mm", "a4");
            doc.setFontSize(14);
            doc.setFont("helvetica", "bold");
            doc.text("CHEQUE TRANSACTION DETAILS", 148, 12, { align: "center" });

            const headers = [["S.No", "Rec.Date", "Rec.No", "VchType", "Party Name", "Chq.No", "Chq.Date", "Bank.Date", "Debit", "Credit"]];
            const body = filteredChequeData.map((row: any, idx: number) => [
                idx + 1,
                row.receipt_date ? dayjs(row.receipt_date).format("DD-MM-YYYY") : "",
                row.receipt_invoice_no || "",
                row.voucherTypeGet || "",
                row.creditAccountGet || "",
                row.check_no || "",
                row.check_date ? dayjs(row.check_date).format("DD-MM-YYYY") : "",
                row.bank_date ? dayjs(row.bank_date).format("DD-MM-YYYY") : "",
                row.debit_amount ? Number(row.debit_amount).toLocaleString() : "-",
                row.credit_amount ? Number(row.credit_amount).toLocaleString() : "-"
            ]);

            autoTable(doc, {
                startY: 20,
                head: headers,
                body: body,
                styles: { fontSize: 8, cellPadding: 1.5 },
                headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255] },
                theme: "grid"
            });

            doc.save(`Cheque_Transactions_${dayjs().format("YYYYMMDD_HHmmss")}.pdf`);
            toast.success("PDF Exported ✅");
        } catch (err) {
            console.error(err);
            toast.error("PDF Export Failed ❌");
        }
    };

    return (
        <Box sx={{ width: "100%", overflowX: "hidden", minHeight: "100vh", bgcolor: "#f1f5f9" }}>
            <PageHeader
                onExportExcel={handleChequeExportExcel}
                onExportPDF={handleChequeExportPDF}
                showPages={true}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onClose={() => setDrawerOpen(false)}
                onToggle={() => setDrawerOpen((p) => !p)}
                fromDate={chequeFilters.Fromdate}
                toDate={chequeFilters.Todate}
                onFromDateChange={(val) => setChequeFilters(p => ({ ...p, Fromdate: val }))}
                onToDateChange={(val) => setChequeFilters(p => ({ ...p, Todate: val }))}
                onApply={fetchChequeData}
            >
                <Box mt={2}>
                    <Autocomplete
                        options={chequeFilters.chequeAccounts}
                        getOptionLabel={(option) => option.label || ""}
                        value={chequeFilters.debitAccount.value ? chequeFilters.debitAccount : null}
                        onChange={(_, newValue) => {
                            setChequeFilters(p => ({
                                ...p,
                                debitAccount: newValue || { value: "", label: "ALL" }
                            }));
                        }}
                        renderInput={(params) => <TextField {...params} label="Debit Account" sx={{ mb: 2 }} />}
                    />
                    <Autocomplete
                        options={creditAccounts}
                        getOptionLabel={(option) => option.label || ""}
                        value={chequeFilters.creditAccount.value ? chequeFilters.creditAccount : null}
                        onChange={(_, newValue) => {
                            setChequeFilters(p => ({
                                ...p,
                                creditAccount: newValue || { value: "", label: "ALL" }
                            }));
                        }}
                        renderInput={(params) => <TextField {...params} label="Credit Account" sx={{ mb: 2 }} />}
                    />
                    <Autocomplete
                        options={voucherTypes}
                        getOptionLabel={(option) => option.label || ""}
                        value={chequeFilters.voucherType.value ? chequeFilters.voucherType : null}
                        onChange={(_, newValue) => {
                            setChequeFilters(p => ({
                                ...p,
                                voucherType: newValue || { value: "", label: "ALL" }
                            }));
                        }}
                        renderInput={(params) => <TextField {...params} label="Voucher Type" sx={{ mb: 2 }} />}
                    />
                    <TextField
                        select
                        label="Party Type"
                        fullWidth
                        value={chequeFilters.partyType}
                        onChange={(e) => setChequeFilters(p => ({ ...p, partyType: e.target.value }))}
                        sx={{ mb: 2 }}
                    >
                        <MenuItem value="ALL">ALL</MenuItem>
                        <MenuItem value="Pending Party">Pending Party</MenuItem>
                        <MenuItem value="Payed Party">Payed Party</MenuItem>
                    </TextField>
                    <Box mt={2}>
                        <Button variant="contained" fullWidth onClick={fetchChequeData}>
                            Apply
                        </Button>
                    </Box>
                </Box>
            </ReportFilterDrawer>

            <Box px={2} pb={4} pt={2}>
                <Box
                    sx={{
                        border: "1px solid #cbd5e1",
                        borderRadius: 1.5,
                        py: 1.2,
                        textAlign: "center",
                        mb: 3,
                        background: "#fff",
                        boxShadow: 1,
                    }}
                >
                    <Typography variant="body1" fontWeight={700} sx={{ letterSpacing: 0.5, color: "#1e293b" }}>
                        CHEQUE TRANSACTION DETAILS  {" "}
                        {chequeFilters.Fromdate === chequeFilters.Todate
                            ? dayjs(chequeFilters.Fromdate).format("DD-MM-YYYY")
                            : `${dayjs(chequeFilters.Fromdate).format("DD-MM-YYYY")} - ${dayjs(chequeFilters.Todate).format("DD-MM-YYYY")}`}
                    </Typography>
                </Box>

                {loading ? (
                    <Box display="flex" justifyContent="center" py={10}>
                        <CircularProgress size={40} sx={{ color: "#1E3A8A" }} />
                    </Box>
                ) : (
                    <>
                        <TableContainer component={Paper} elevation={2} sx={{ borderRadius: 2, border: "1px solid #cbd5e1", overflow: "auto", maxHeight: "calc(100vh - 190px)" }}>
                            <Table size="small" sx={{ minWidth: 1000 }} stickyHeader>
                        <TableHead>
                            <TableRow>
                                <TableCell align="center" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, width: 80, border: "1px solid #cbd5e1" }}></TableCell>
                                <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Rec.Date</TableCell>
                                <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Rec.No</TableCell>
                                <TableCell
                                    onClick={(e) => handleHeaderClick(e, "voucherTypeGet")}
                                    sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1", cursor: "pointer" }}
                                >
                                    <Box display="flex" alignItems="center" gap={0.5}>
                                        VchType
                                        {columnFilters.voucherTypeGet.length > 0 && <FilterAltIcon fontSize="small" sx={{ color: "#ffffffff" }} />}
                                    </Box>
                                </TableCell>
                                <TableCell
                                    onClick={(e) => handleHeaderClick(e, "creditAccountGet")}
                                    sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1", cursor: "pointer" }}
                                >
                                    <Box display="flex" alignItems="center" gap={0.5}>
                                        Party Name
                                        {columnFilters.creditAccountGet.length > 0 && <FilterAltIcon fontSize="small" sx={{ color: "#ffffffff" }} />}
                                    </Box>
                                </TableCell>
                                <TableCell
                                    onClick={(e) => handleHeaderClick(e, "check_no")}
                                    sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1", cursor: "pointer" }}
                                >
                                    <Box display="flex" alignItems="center" gap={0.5}>
                                        Chq.No
                                        {columnFilters.check_no.length > 0 && <FilterAltIcon fontSize="small" sx={{ color: "#ffffffff" }} />}
                                    </Box>
                                </TableCell>
                                <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Chq.Date</TableCell>
                                <TableCell sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Bank.Date</TableCell>
                                <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Debit</TableCell>
                                <TableCell align="right" sx={{ backgroundColor: "#1E3A8A", color: "#fff", fontWeight: 700, py: 1.5, border: "1px solid #cbd5e1" }}>Credit</TableCell>
                            </TableRow>
                        </TableHead>
                        <TableBody>
                            {paginatedChequeData.length > 0 && (
                                <TableRow>
                                    <TableCell align="center" sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800 }}>
                                        Total
                                    </TableCell>
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1" }} />
                                    <TableCell align="right" sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", borderRight: "1px solid #cbd5e1", fontWeight: 800, color: "#1e3a8a" }}>
                                        {formatNum(chequeTotals.debit)}
                                    </TableCell>
                                    <TableCell align="right" sx={{ position: "sticky", top: "45px", zIndex: 10, backgroundColor: "#f1f5f9", fontWeight: 800, color: "#1e3a8a" }}>
                                        {formatNum(chequeTotals.credit)}
                                    </TableCell>
                                </TableRow>
                            )}
                            {paginatedChequeData.length === 0 ? (
                                <TableRow>
                                    <TableCell colSpan={10} align="center" sx={{ py: 6, color: "#94a3b8" }}>
                                        {!chequeFilters.debitAccount.value
                                            ? "Please select a Debit Account in filters to load data."
                                            : "No cheque transactions match your filter criteria."
                                        }
                                    </TableCell>
                                </TableRow>
                            ) : (
                                paginatedChequeData.map((row: any, idx: number) => (
                                    <ChequeRow key={row.receipt_id || idx} row={row} index={(chequePage - 1) * chequeRowsPerPage + idx + 1} />
                                ))
                            )}
                        </TableBody>
                    </Table>
                </TableContainer>

                <CommonPagination
                    totalRows={filteredChequeData.length}
                    page={chequePage}
                    rowsPerPage={chequeRowsPerPage}
                    onPageChange={setChequePage}
                    onRowsPerPageChange={setChequeRowsPerPage}
                />
            </>
        )}
    </Box>

            {/* Header Column Filters Menu Popup */}
            <Menu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor)}
                onClose={() => setFilterAnchor(null)}
                PaperProps={{
                    sx: {
                        maxHeight: 300,
                        width: 250,
                        boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                        borderRadius: 1.5,
                        border: "1px solid #e2e8f0"
                    }
                }}
            >
                {activeHeader && (
                    <Box p={1.5}>
                        <TextField
                            size="small"
                            fullWidth
                            placeholder="Search..."
                            value={searchText}
                            onChange={(e) => setSearchText(e.target.value)}
                            sx={{ mb: 1.5 }}
                        />
                        <Box sx={{ maxHeight: 180, overflowY: "auto" }}>
                            {(() => {
                                const allValues = Array.from(
                                    new Set(chequeData.map((x) => x[activeHeader]))
                                ).filter((v) => v && String(v).toLowerCase().includes(searchText.toLowerCase()));

                                const selectedValues = columnFilters[activeHeader] || [];

                                // Sort selected values first
                                const sortedValues = [
                                    ...allValues.filter((v) => selectedValues.includes(v)),
                                    ...allValues.filter((v) => !selectedValues.includes(v)),
                                ];

                                return sortedValues.map((v) => {
                                    const isSelected = selectedValues.includes(v);

                                    return (
                                        <MenuItem
                                            key={String(v)}
                                            onClick={() => {
                                                setColumnFilters((prev) => {
                                                    const prevValues = prev[activeHeader] || [];
                                                    const newValues = prevValues.includes(v)
                                                        ? prevValues.filter((x: any) => x !== v)
                                                        : [...prevValues, v];

                                                    return {
                                                        ...prev,
                                                        [activeHeader]: newValues,
                                                    };
                                                });
                                            }}
                                            sx={{
                                                backgroundColor: isSelected ? "#e0e7ff" : "transparent",
                                                fontWeight: isSelected ? 600 : 400,
                                                "&:hover": {
                                                    backgroundColor: isSelected ? "#c7d2fe" : "#f1f5f9",
                                                },
                                            }}
                                        >
                                            {String(v)}
                                        </MenuItem>
                                    );
                                });
                            })()}
                        </Box>
                    </Box>
                )}
            </Menu>
        </Box>
    );
};

export default ChequeTransactionReport;
