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
    CircularProgress,
    Menu,
    MenuItem,
    TextField,
    Chip,
    Stack
} from "@mui/material";
import dayjs from "dayjs";
import PageHeader from "../../Layout/PageHeader";
import AppLayout from "../../Layout/appLayout";
import CommonPagination from "../../Components/CommonPagination";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import {
    stockValueReportService,
    PosRate,
    StockValue,
} from "../../services/stockValueRateMaster.service";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";

/* ================= COMPONENT ================= */

const StockValueRateMasterReport: React.FC = () => {
    const today = dayjs().format("YYYY-MM-DD");

    const [loading, setLoading] = useState(false);

    const [drawerOpen, setDrawerOpen] = useState(false);

    const [columnFilters, setColumnFilters] = useState<Record<string, any>>({
        Ledger_Date: {
            from: today,
            to: today,
        },
    });

    const [posRates, setPosRates] = useState<PosRate[]>([]);
    const [stockValues, setStockValues] = useState<StockValue[]>([]);
    const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
    const [activeHeader, setActiveHeader] = useState<string | null>(null);
    const [searchText, setSearchText] = useState("");
    const [selectedBrand, setSelectedBrand] = useState<string>("All");

    const [page, setPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(100);

    /* ================= LOAD DATA ================= */
    useEffect(() => {
        const loadData = async () => {
            try {
                setLoading(true);

                const fromDate =
                    columnFilters?.Ledger_Date?.from || dayjs().format("YYYY-MM-DD");

                const [posRes, stockRes] = await Promise.all([
                    stockValueReportService.getPosRates({ FromDate: fromDate }),
                    stockValueReportService.getStockValues({ FromDate: fromDate }),
                ]);

                setPosRates(posRes.data?.data?.posRateMaster || []);
                setStockValues(stockRes.data?.data || []);
            } catch (err) {
                console.error("Error loading stock value report:", err);
                setPosRates([]);
                setStockValues([]);
            } finally {
                setLoading(false);
            }
        };

        loadData();
    }, [columnFilters["Ledger_Date"]]);

    /* ================= MERGE LOGIC ================= */
    const mergedData = useMemo(() => {
        const stockMap = new Map<string, number>();

        stockValues.forEach((s) => {
            stockMap.set(String(s.Product_Id), Number(s.CL_Rate));
        });

        return posRates.map((p) => {
            const clRate = stockMap.get(String(p.Item_Id)) || 0;

            return {
                POS_Brand_Name: p.POS_Brand_Name || "Others",
                Product_Name: p.Product_Name,
                Min_Rate: p.Min_Rate,
                Rate: p.Rate,
                CL_Rate: clRate,
            };
        });
    }, [posRates, stockValues]);

    // ****FILTER***** //

    const handleHeaderClick = (
        e: React.MouseEvent<HTMLElement>,
        columnKey: string
    ) => {
        if (columnKey !== "Product_Name") return;

        setActiveHeader(columnKey);
        setSearchText("");
        setFilterAnchor(e.currentTarget);
    };

    const applyFilters = (rows: any[]) => {
        return rows.filter((row) => {
            // ✅ PRODUCT NAME FILTER
            const selectedProducts = columnFilters.Product_Name || [];

            if (selectedProducts.length > 0) {
                if (!selectedProducts.includes(row.Product_Name)) {
                    return false;
                }
            }

            return true;
        });
    };

    /* ================= FILTERED DATA ================= */
    const filteredData = useMemo(() => {
        let rows = applyFilters(mergedData);

        if (selectedBrand !== "All") {
            rows = rows.filter(
                (row) => row.POS_Brand_Name === selectedBrand
            );
        }

        return rows.sort((a, b) => {
            const brandCompare = a.POS_Brand_Name.localeCompare(b.POS_Brand_Name);
            if (brandCompare !== 0) return brandCompare;

            return a.Product_Name.localeCompare(b.Product_Name);
        });
    }, [mergedData, columnFilters, selectedBrand]);

    const formatINR = (value: number) => {
        return new Intl.NumberFormat("en-IN", {
            style: "currency",
            currency: "INR",
        }).format(value || 0);
    };

    /* ================= PAGINATION ================= */
    const totalRows = filteredData.length;

    const paginatedData = useMemo(() => {
        const start = (page - 1) * rowsPerPage;
        const end = start + rowsPerPage;

        return filteredData.slice(start, end);
    }, [filteredData, page, rowsPerPage]);

    useEffect(() => {
        setPage(1);
    }, [selectedBrand, columnFilters.Product_Name, rowsPerPage]);

    /* ================= BRAND LIST ================= */
    const brandOptions = useMemo(() => {
        const brands = [...new Set(mergedData.map((x) => x.POS_Brand_Name))]
            .filter(Boolean)
            .sort();

        return ["All", ...brands];
    }, [mergedData]);


    /* ================= EXPORT ================= */
    const handleExportPDF = () => {
        const headers = ["S.No", "Product Name", "Min Rate", "List Rate", "CL Rate"];

        const data: any[] = [];
        let sno = 1;

        filteredData.forEach((row, index) => {
            const prev = filteredData[index - 1];

            const showGroup =
                index === 0 ||
                prev?.POS_Brand_Name !== row.POS_Brand_Name;

            if (showGroup) {
                data.push([row.POS_Brand_Name, "", "", "", ""]);
            }

            data.push([
                sno++,
                row.Product_Name,
                formatINR(row.Min_Rate),
                formatINR(row.Rate),
                formatINR(row.CL_Rate),
            ]);
        });

        exportToPDF("Rate Master Report", headers, data);
    };

    const handleExportExcel = () => {
        const headers = ["S.No", "Product Name", "Min Rate", "List Rate", "CL Rate"];

        const data: any[] = [];
        let sno = 1;

        filteredData.forEach((row, index) => {
            const prev = filteredData[index - 1];

            const showGroup =
                index === 0 ||
                prev?.POS_Brand_Name !== row.POS_Brand_Name;

            if (showGroup) {
                data.push([row.POS_Brand_Name, "", "", "", ""]);
            }

            data.push([
                sno++,
                row.Product_Name,
                formatINR(row.Min_Rate),
                formatINR(row.Rate),
                formatINR(row.CL_Rate),
            ]);
        });

        exportToExcel("Rate Master Report", headers, data);
    };

    /* ================= UI ================= */
    return (
        <>
            {/* 🔥 CLEAN HEADER (EXPORT ONLY) */}
            <PageHeader
                onExportPDF={handleExportPDF}
                onExportExcel={handleExportExcel}
            />

            <ReportFilterDrawer
                open={drawerOpen}
                onToggle={() => setDrawerOpen((p) => !p)}
                onClose={() => setDrawerOpen(false)}
                fromDate={columnFilters["Ledger_Date"]?.from || ""}
                onFromDateChange={(v) =>
                    setColumnFilters((prev) => ({
                        ...prev,
                        Ledger_Date: {
                            from: v,
                        },
                    }))
                }
                onApply={() => setDrawerOpen(false)}
            />

            <AppLayout fullWidth>
                <Box sx={{ mt: 1 }}>

                    {loading ? (
                        <CircularProgress />
                    ) : (
                        <>
                            <Box sx={{ mb: 1 }}>
                                <Stack
                                    direction="row"
                                    spacing={1}
                                    useFlexGap
                                    flexWrap="wrap"
                                >
                                    {brandOptions.map((brand) => (
                                        <Chip
                                            key={brand}
                                            label={brand}
                                            clickable
                                            onClick={() => {
                                                setSelectedBrand(brand);
                                                setPage(1);
                                            }}
                                            color={
                                                selectedBrand === brand
                                                    ? "primary"
                                                    : "default"
                                            }
                                            variant={
                                                selectedBrand === brand
                                                    ? "filled"
                                                    : "outlined"
                                            }
                                            sx={{
                                                fontWeight: 600,
                                                height: 30,
                                            }}
                                        />
                                    ))}
                                </Stack>
                            </Box>

                            <TableContainer
                                component={Paper}
                                sx={{
                                    maxHeight: "calc(100vh - 210px)",
                                }}
                            >
                                <Table size="small">
                                    <TableHead
                                        sx={{
                                            background: "#1E3A8A",
                                            position: "sticky",
                                            top: 0,
                                            zIndex: 2,
                                        }}
                                    >
                                        <TableRow>
                                            <TableCell sx={{ color: "#fff" }}>S.No</TableCell>
                                            <TableCell
                                                sx={{ color: "#fff", cursor: "pointer" }}
                                                onClick={(e) => handleHeaderClick(e, "Product_Name")}
                                            >
                                                Product Name
                                            </TableCell>
                                            <TableCell sx={{ color: "#fff" }}>Min Rate</TableCell>
                                            <TableCell sx={{ color: "#fff" }}>List Rate</TableCell>
                                            <TableCell sx={{ color: "#fff" }}>CL Rate</TableCell>
                                        </TableRow>
                                    </TableHead>

                                    <TableBody>

                                        {paginatedData.map((row, i) => {
                                            const previousRow = paginatedData[i - 1];
                                            const showGroupHeader =
                                                i === 0 ||
                                                previousRow?.POS_Brand_Name !== row.POS_Brand_Name;

                                            return (
                                                <React.Fragment key={i}>
                                                    {/* GROUP HEADER */}
                                                    {showGroupHeader && (
                                                        <TableRow
                                                            sx={{
                                                                background: "#eff2f7",
                                                            }}
                                                        >
                                                            <TableCell
                                                                colSpan={5}
                                                                sx={{
                                                                    fontWeight: 800,
                                                                    color: "#000000",
                                                                    fontSize: "13px",
                                                                }}
                                                            >
                                                                {row.POS_Brand_Name}
                                                            </TableCell>
                                                        </TableRow>
                                                    )}

                                                    {/* DATA ROW */}
                                                    <TableRow>
                                                        <TableCell>
                                                            {(page - 1) * rowsPerPage + i + 1}
                                                        </TableCell>

                                                        <TableCell>{row.Product_Name}</TableCell>

                                                        <TableCell>{formatINR(row.Min_Rate)}</TableCell>

                                                        <TableCell>{formatINR(row.Rate)}</TableCell>

                                                        <TableCell>{formatINR(row.CL_Rate)}</TableCell>
                                                    </TableRow>
                                                </React.Fragment>
                                            );
                                        })}

                                        {paginatedData.length === 0 && (
                                            <TableRow>
                                                <TableCell colSpan={4} align="center">
                                                    No Data Found
                                                </TableCell>
                                            </TableRow>
                                        )}
                                    </TableBody>
                                </Table>
                            </TableContainer>

                            <CommonPagination
                                totalRows={totalRows}
                                page={page}
                                rowsPerPage={rowsPerPage}
                                onPageChange={setPage}
                                onRowsPerPageChange={setRowsPerPage}
                            />
                        </>
                    )}
                </Box>
            </AppLayout>

            <Menu
                anchorEl={filterAnchor}
                open={Boolean(filterAnchor) && activeHeader === "Product_Name"}
                onClose={() => setFilterAnchor(null)}
            >
                <Box p={2} sx={{ minWidth: 220 }}>

                    {/* SEARCH */}
                    <TextField
                        size="small"
                        fullWidth
                        placeholder="Search"
                        value={searchText}
                        onChange={(e) => setSearchText(e.target.value)}
                        sx={{ mb: 1 }}
                    />

                    {/* CLEAR */}
                    <MenuItem
                        sx={{ fontWeight: 600 }}
                        onClick={() => {
                            setColumnFilters((prev) => ({
                                ...prev,
                                Product_Name: [],
                            }));
                            setFilterAnchor(null);
                        }}
                    >
                        All
                    </MenuItem>

                    {/* OPTIONS */}
                    {[...new Set(mergedData.map((r) => r.Product_Name))]
                        .filter(Boolean)
                        .filter((v) =>
                            String(v).toLowerCase().includes(searchText.toLowerCase())
                        )
                        .map((v) => {
                            const selected = columnFilters.Product_Name || [];
                            const isSelected = selected.includes(v);

                            return (
                                <MenuItem
                                    key={v}
                                    onClick={() => {
                                        setColumnFilters((prev) => {
                                            const prevValues = prev.Product_Name || [];

                                            const newValues = prevValues.includes(v)
                                                ? prevValues.filter((x: any) => x !== v)
                                                : [...prevValues, v];

                                            return {
                                                ...prev,
                                                Product_Name: newValues,
                                            };
                                        });
                                    }}
                                    sx={{
                                        backgroundColor: isSelected ? "#e0e7ff" : "transparent",
                                        fontWeight: isSelected ? 600 : 400,
                                    }}
                                >
                                    {v}
                                </MenuItem>
                            );
                        })}
                </Box>
            </Menu>
        </>
    );
};

export default StockValueRateMasterReport;