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
} from "@mui/material";
import dayjs from "dayjs";
import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import { mapForExport } from "../../utils/exportMapper";
import {
  UnitEconomicsReport,
  UnitEconomicsReportService,
} from "../../services/unitEconomicsReport.service";


/* ================= STYLES ================= */
const headStyle = {
  color: "#fff",
  fontWeight: 600,
  fontSize: "0.75rem",
  cursor: "pointer",
};

const UnitEconomicsReportPage: React.FC = () => {
  const today = dayjs().format("YYYY-MM-DD");

  const [data, setData] = useState<UnitEconomicsReport[]>([]);
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(100);

  /* -------- FILTERS -------- */
  const [filters, setFilters] = useState({
    Date: { from: today, to: today },
    Product: "",
  });

  const [tempDate, setTempDate] = useState(filters.Date);

  /* -------- HEADER FILTER -------- */
  const [filterAnchor, setFilterAnchor] = useState<null | HTMLElement>(null);
  const [activeHeader, setActiveHeader] = useState<string | null>(null);
  const [searchText, setSearchText] = useState("");
  const [lastSyncDate, setLastSyncDate] = useState<string | null>(null);

  /* -------- SUMMARY -------- */
  const [summaryColumn, setSummaryColumn] = useState<keyof UnitEconomicsReport | null>(null);

  const EXPORT_COLUMNS = [
    { label: "S.No", key: "sno" },
    { label: "Date", key: "Trans_Date", type: "date" },
    { label: "Product", key: "Product_Name" },
    { label: "Quantity", key: "Bill_Qty", type: "number" },
    { label: "Rate", key: "Rate", type: "number" },
    { label: "Amount", key: "Amount", type: "number" },
    { label: "Min Rate", key: "Min_Rate", type: "number" },
    { label: "List Rate", key: "List_Rate", type: "number" },
    { label: "COGS", key: "COGS", type: "number" },
  ];

  /* ================= LOAD DATA ================= */
  useEffect(() => {
    const loadData = async () => {
      const res = await UnitEconomicsReportService.getReports({
        Fromdate: filters.Date.from,
        Todate: filters.Date.to,
      });

      const responseData = res.data.data;

      // ✅ correct access
      let rows: UnitEconomicsReport[] = responseData.rows || [];

      if (filters.Product) {
        rows = rows.filter(
          (r) => r.Product_Name === filters.Product
        );
      }

      setData(rows);

      // ✅ second dataset
      setLastSyncDate(
        responseData.lastStockValueDate?.Last_Stock_Value_Date ?? null
      );

      setPage(1);
      setSummaryColumn(null);
    };

    loadData();
  }, [filters.Date, filters.Product]);

  /* ================= PAGINATION ================= */
  const paginatedData = data.slice(
    (page - 1) * rowsPerPage,
    page * rowsPerPage
  );

  /* ================= DROPDOWNS ================= */
  const products = useMemo(
    () =>
      [...new Set(data.map((d) => d.Product_Name))]
        .filter(Boolean)
        .filter((p) =>
          p.toLowerCase().includes(searchText.toLowerCase())
        ),
    [data, searchText]
  );

  /* ================= SUMMARY ================= */
  const getTotal = (key: keyof UnitEconomicsReport) => {
    return data.reduce((sum, r) => sum + (Number(r[key]) || 0), 0);
  };

  /* ================= HEADER CLICK ================= */
  const openFilter = (e: React.MouseEvent<HTMLElement>, column: string) => {
    setActiveHeader(column);
    setFilterAnchor(e.currentTarget);
    setSearchText("");

    if (["Bill_Qty", "Rate", "Amount", "COGS"].includes(column)) {
      if (summaryColumn === column) {
        setSummaryColumn(null);
      } else {
        setSummaryColumn(column as keyof UnitEconomicsReport);
      }
    }
  };

  useEffect(() => {
    setSummaryColumn(null);
  }, [filters.Date, filters.Product]);



  /* ================= EXPORT ================= */
  const handleExportPDF = () => {
    const { headers, data: exportData } = mapForExport(EXPORT_COLUMNS, data);
    exportToPDF("Unit Economics Report", headers, exportData);
  };

  const handleExportExcel = () => {
    const { headers, data: exportData } = mapForExport(EXPORT_COLUMNS, data);
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
          lastSyncDate && (
            <Box
              sx={{
                display: "inline-flex",
                alignItems: "center",
                gap: 0.25,
                px: 0.5,
                py: 0.35,
                fontSize: "0.75rem",
                backgroundColor: "#e2eef0",
                color: "#000000",
                border: "1px solid #0b78f5",
                borderRadius: 1,
                whiteSpace: "nowrap",
              }}
            >
              {/* Icon */}
              <Box
                sx={{
                  width: 14,
                  height: 14,
                  borderRadius: "3px",
                  backgroundColor: "#1d0a72",
                  color: "#f1f5f5",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: "0.6rem",
                  fontWeight: 600,
                  flexShrink: 0,
                }}
              >
                ⟳
              </Box>

              {/* Text */}
              <Box sx={{ fontSize: "0.65rem" }}>
                Last Sync:
                <Box component="span" sx={{ fontWeight: 600, ml: 0.25 }}>
                  {dayjs(lastSyncDate).format("DD/MM/YYYY")}
                </Box>
              </Box>
            </Box>
          )
        }
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
                  <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Date")}>Date</TableCell>
                  <TableCell sx={headStyle} onClick={(e) => openFilter(e, "Product")}>Product</TableCell>
                  <TableCell align="right" sx={headStyle} > Quantity </TableCell>
                  <TableCell align="right" sx={headStyle} >Rate</TableCell>
                  <TableCell align="right" sx={headStyle} >Amount</TableCell>
                  <TableCell align="right" sx={headStyle} >Min Rate</TableCell>
                  <TableCell align="right" sx={headStyle} >List Rate</TableCell>
                  <TableCell align="right" sx={headStyle} >COGS</TableCell>
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

                  {/* Date */}
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

                  {/* Min Rate (empty) */}
                  <TableCell />

                  {/* List Rate (empty) */}
                  <TableCell />

                  {/* COGS */}
                  <TableCell align="right" sx={{ fontWeight: 700 }}>
                    {getTotal("COGS").toFixed(2)}
                  </TableCell>
                </TableRow>
              </TableBody>

              {/* ===== BODY ===== */}
              <TableBody>
                {paginatedData.map((row, i) => (
                  <TableRow key={`${row.Product_Id}-${i}`}>
                    <TableCell sx={{ fontSize: "0.75rem" }}>
                      {(page - 1) * rowsPerPage + i + 1}
                    </TableCell>
                    <TableCell sx={{ fontSize: "0.75rem" }}>
                      {dayjs(row.Trans_Date).format("DD/MM/YYYY")}
                    </TableCell>
                    <TableCell sx={{ fontSize: "0.75rem" }}>{row.Product_Name}</TableCell>
                    <TableCell align="right">{Number(row.Bill_Qty).toFixed(2)}</TableCell>
                    <TableCell align="right">{formatINR(Number(row.Rate))}</TableCell>
                    <TableCell align="right">{formatINR(Number(row.Amount))}</TableCell>
                    <TableCell align="right">{formatINR(Number(row.Min_Rate))}</TableCell>
                    <TableCell align="right">{formatINR(Number(row.List_Rate))}</TableCell>
                    <TableCell align="right">{Number(row.COGS).toFixed(2)}</TableCell>
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
              <Box p={2} sx={{ minWidth: 220 }}>
                <TextField
                  size="small"
                  fullWidth
                  placeholder="Search Product"
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                  sx={{ mb: 1 }}
                />

                {/* ✅ ALL – ALWAYS FIRST */}
                <MenuItem
                  sx={{ fontWeight: 600 }}
                  onClick={() => {
                    setFilters((f) => ({ ...f, Product: "" }));
                    setFilterAnchor(null);
                  }}
                >
                  All
                </MenuItem>

                {/* ✅ PRODUCT LIST BELOW ALL */}
                {products.map((p) => (
                  <MenuItem
                    key={p}
                    onClick={() => {
                      setFilters((f) => ({ ...f, Product: p }));
                      setFilterAnchor(null);
                    }}
                  >
                    {p}
                  </MenuItem>
                ))}
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
    </>
  );
};

export default UnitEconomicsReportPage;
