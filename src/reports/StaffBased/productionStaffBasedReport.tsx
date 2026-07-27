import React, { useState, useMemo, useEffect } from "react";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
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
} from "@mui/material";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import dayjs from "dayjs";
import { useNavigate } from "react-router-dom";
import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import ReportFilterDrawer from "../../Components/ReportFilterDrawer";
import { mockProductionDetails } from "./mockProductionData";

/* ================= STYLES ================= */
const tableHeaderStyle = {
  backgroundColor: "#1E3A8A", // Premium slate/dark blue
  color: "#fff",
  fontWeight: 700,
  fontSize: "0.78rem",
  textTransform: "uppercase",
  letterSpacing: "0.5px",
  borderBottom: "2px solid #0D1B3E",
  padding: "10px 12px",
};

const categoryHeaderStyle = {
  backgroundColor: "#FFE0B2", // Warm light peach
  color: "#E65100", // Deep orange/amber text
  fontWeight: 800,
  fontSize: "0.85rem",
  textAlign: "center",
  borderRight: "2px solid #FFCC80",
  borderBottom: "1px solid #FFE0B2",
  textTransform: "uppercase",
  letterSpacing: "0.5px",
};

const rowHeaderStyle = {
  backgroundColor: "#FFF3E0", // Even lighter peach
  color: "#F57C00",
  fontWeight: 700,
  fontSize: "0.75rem",
  borderRight: "1px solid #FFE0B2",
};

const bodyStyle = {
  fontSize: "0.75rem",
  fontWeight: 500,
  color: "#374151",
};

const valueCellStyle = {
  fontSize: "0.75rem",
  cursor: "pointer",
  fontWeight: 600,
  transition: "all 0.2s",
  "&:hover": {
    backgroundColor: "#FEF3C7", // Faint gold on hover
    color: "#D97706",
  },
};

const CONFIG = [
  { category: "Inwards", vouType: "Pur / Ret", godown: "Main Godown", categorySpan: 3, vouTypeSpan: 2, isFirstInCategory: true, isFirstInVouType: true },
  { category: "Inwards", vouType: "Pur / Ret", godown: "Cold Storage", categorySpan: 0, vouTypeSpan: 0, isFirstInCategory: false, isFirstInVouType: false },
  { category: "Inwards", vouType: "Int Trf", godown: "Cold Storage", categorySpan: 0, vouTypeSpan: 1, isFirstInCategory: false, isFirstInVouType: true },
  { category: "Process", vouType: "Adj", godown: "Main Godown", categorySpan: 4, vouTypeSpan: 1, isFirstInCategory: true, isFirstInVouType: true },
  { category: "Process", vouType: "Atty", godown: "Cold Storage", categorySpan: 0, vouTypeSpan: 1, isFirstInCategory: false, isFirstInVouType: true },
  { category: "Process", vouType: "Wt.Check", godown: "Main Godown", categorySpan: 0, vouTypeSpan: 1, isFirstInCategory: false, isFirstInVouType: true },
  { category: "Process", vouType: "Cleaning", godown: "Cold Storage", categorySpan: 0, vouTypeSpan: 1, isFirstInCategory: false, isFirstInVouType: true },
  { category: "Outwards", vouType: "Sales", godown: "Main Godown", categorySpan: 2, vouTypeSpan: 1, isFirstInCategory: true, isFirstInVouType: true },
  { category: "Outwards", vouType: "Int Trf", godown: "Secondary Warehouse", categorySpan: 0, vouTypeSpan: 1, isFirstInCategory: false, isFirstInVouType: true },
] as const;

const ROLES = ["Atten by", "Created", "Printed By", "Taken", "Check", "Check 1", "Delivery"] as const;

const ProductionStaffBasedReport: React.FC = () => {
  const today = dayjs().format("YYYY-MM-DD");
  const navigate = useNavigate();

  const [fromDate, setFromDate] = useState(today);
  const [toDate, setToDate] = useState(today);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [displayMode, setDisplayMode] = useState<"tonnage" | "count">("tonnage");

  const [filters, setFilters] = useState({
    Date: { from: today, to: today },
  });

  useEffect(() => {
    const saved = sessionStorage.getItem("productionStaffBasedReportState");
    if (saved) {
      try {
        const state = JSON.parse(saved);
        if (state.fromDate) setFromDate(state.fromDate);
        if (state.toDate) setToDate(state.toDate);
        if (state.filters) setFilters(state.filters);
        if (state.displayMode) setDisplayMode(state.displayMode);
      } catch (e) {
        console.error("Error restoring productionStaffBasedReportState", e);
      }
      sessionStorage.removeItem("productionStaffBasedReportState");
    }
  }, []);

  // Load and filter raw data based on Date range
  const filteredData = useMemo(() => {
    return mockProductionDetails.filter(
      (d) => d.Date >= filters.Date.from && d.Date <= filters.Date.to
    );
  }, [filters.Date]);

  // Compute row values dynamically from mock database
  const gridData = useMemo(() => {
    return CONFIG.map((row) => {
      const rowRecords = filteredData.filter(
        (d) => d.Category === row.category && d.VouType === row.vouType && d.Godown === row.godown
      );

      const rolesValues = {} as Record<typeof ROLES[number], number>;
      ROLES.forEach((role) => {
        const roleRecords = rowRecords.filter((d) => d.Role === role);
        rolesValues[role] =
          displayMode === "tonnage"
            ? roleRecords.reduce((sum, r) => sum + r.Qty, 0)
            : new Set(roleRecords.map((r) => r.JournalNo)).size;
      });

      const total = Object.values(rolesValues).reduce((sum, v) => sum + v, 0);

      // overallCount: unique JournalNo count of all records in this row (regardless of role)
      const overallCount = new Set(rowRecords.map((r) => r.JournalNo)).size;

      // Tonnage: sum of Qty for all records matching this Category/VouType/Godown
      const tonnage = rowRecords.reduce((sum, r) => sum + r.Qty, 0);

      return {
        ...row,
        ...rolesValues,
        total,
        overallCount,
        tonnage,
      };
    });
  }, [filteredData, displayMode]);

  const numericKeys = useMemo(() => ["total", "overallCount", "tonnage", ...ROLES], []);

  const {
    sortConfig,
    rangeFilter,
    filterAnchor,
    setFilterAnchor,
    activeHeader,
    handleSort,
    openFilter,
    filteredAndSortedData: sortedGridData,
    getMinMax,
    setRangeFilter,
    clearRangeFilter,
  } = useNumericalFilter(gridData, numericKeys);

  // Compute column totals
  const colTotals = useMemo(() => {
    const totals = {} as Record<typeof ROLES[number] | "total" | "overallCount" | "tonnage", number>;

    ROLES.forEach((role) => {
      totals[role] = gridData.reduce((sum, row) => sum + row[role], 0);
    });

    totals.total = gridData.reduce((sum, row) => sum + row.total, 0);
    totals.overallCount = gridData.reduce((sum, row) => sum + row.overallCount, 0);
    totals.tonnage = gridData.reduce((sum, row) => sum + row.tonnage, 0);

    return totals;
  }, [gridData]);

  const formatValue = (val: number) => {
    if (val === 0) return "";
    return displayMode === "tonnage" ? val.toFixed(2) : val.toFixed(0);
  };

  const handleCellClick = (
    row: typeof CONFIG[number],
    role?: typeof ROLES[number] | "total" | "overall_count" | "overall_tonnage"
  ) => {
    sessionStorage.setItem(
      "productionStaffBasedReportState",
      JSON.stringify({
        fromDate,
        toDate,
        filters,
        displayMode,
      })
    );
    navigate("/productionStaffBasedDetails", {
      state: {
        filters: {
          DateRange: filters.Date,
          Category: row.category,
          VouType: row.vouType,
          Godown: row.godown,
          Role: role === "total" || role === "overall_count" || role === "overall_tonnage" ? undefined : role,
        },
        displayMode,
      },
    });
  };

  const handleHeaderTotalClick = (role: typeof ROLES[number]) => {
    sessionStorage.setItem(
      "productionStaffBasedReportState",
      JSON.stringify({
        fromDate,
        toDate,
        filters,
        displayMode,
      })
    );
    navigate("/productionStaffBasedDetails", {
      state: {
        filters: {
          DateRange: filters.Date,
          Role: role,
        },
        displayMode,
      },
    });
  };

  return (
    <>
      <PageHeader />

      <ReportFilterDrawer
        open={drawerOpen}
        onToggle={() => setDrawerOpen((p) => !p)}
        onClose={() => setDrawerOpen(false)}
        fromDate={fromDate}
        toDate={toDate}
        onFromDateChange={setFromDate}
        onToDateChange={setToDate}
        showProductionDisplayMode={true}
        productionDisplayModeValue={displayMode}
        onProductionDisplayModeChange={setDisplayMode}
        onApply={() => {
          setFilters({
            Date: { from: fromDate, to: toDate },
          });
          setDrawerOpen(false);
        }}
      />

      <AppLayout fullWidth>
        <Box sx={{ p: 2, display: "flex", flexDirection: "column", gap: 1.5 }}>
          {/* Dashboard Title & Meta */}
          <Box display="flex" justifyContent="space-between" alignItems="center">
            <Typography variant="h6" sx={{ fontWeight: 700, color: "#1E3A8A", display: "flex", alignItems: "center", gap: 1 }}>
              Production Staff Based Report
              <ChevronRightIcon sx={{ color: "#9CA3AF" }} />
              <span style={{ fontSize: "0.85rem", color: "#6B7280", fontWeight: 500 }}>
                {dayjs(filters.Date.from).format("DD/MM/YYYY")} — {dayjs(filters.Date.to).format("DD/MM/YYYY")}
              </span>
            </Typography>
          </Box>

          <TableContainer
            component={Paper}
            sx={{
              maxHeight: "calc(100vh - 160px)",
              boxShadow: "0 10px 25px rgba(0,0,0,0.08)",
              borderRadius: "12px",
              overflow: "auto",
            }}
          >
            <Table size="small" stickyHeader>
              <TableHead>
                {/* Row 1 Headers */}
                <TableRow>
                   <TableCell sx={tableHeaderStyle} rowSpan={2} align="center">Category</TableCell>
                  <TableCell sx={tableHeaderStyle} rowSpan={2}>Vou Type / Group</TableCell>
                  <TableCell sx={tableHeaderStyle} rowSpan={2}>Godown</TableCell>
                  {ROLES.map((r) => (
                    <TableCell key={r} sx={tableHeaderStyle} align="right">
                      <SortableHeaderLabel
                        label={r}
                        columnKey={r}
                        sortConfig={sortConfig}
                        onSort={handleSort}
                        onOpenFilter={(e) => openFilter(e, r)}
                      />
                    </TableCell>
                  ))}
                  <TableCell
                    sx={tableHeaderStyle}
                    align="right"
                    style={{ borderRight: "1px solid #E5E7EB" }}
                  >
                    <SortableHeaderLabel
                      label="Total"
                      columnKey="total"
                      sortConfig={sortConfig}
                      onSort={handleSort}
                      onOpenFilter={(e) => openFilter(e, "total")}
                    />
                  </TableCell>

                  {/* Spacer Cell */}
                  <TableCell
                    rowSpan={2}
                    style={{
                      width: 16,
                      border: "none",
                      backgroundColor: "#fff",
                      zIndex: 3,
                    }}
                  />

                  <TableCell
                    colSpan={2}
                    align="center"
                    sx={{
                      ...tableHeaderStyle,
                      backgroundColor: "#D97706", // Amber contrast header
                      borderBottom: "1px solid #B45309",
                    }}
                  >
                    Over All - Staff involved
                  </TableCell>
                </TableRow>

                {/* Row 2 Headers */}
                <TableRow>
                  <TableCell
                    align="right"
                    sx={{
                      ...tableHeaderStyle,
                      backgroundColor: "#F59E0B",
                      borderBottom: "2px solid #B45309",
                    }}
                  >
                    <SortableHeaderLabel
                      label="Count"
                      columnKey="overallCount"
                      sortConfig={sortConfig}
                      onSort={handleSort}
                      onOpenFilter={(e) => openFilter(e, "overallCount")}
                    />
                  </TableCell>
                  <TableCell
                    align="right"
                    sx={{
                      ...tableHeaderStyle,
                      backgroundColor: "#F59E0B",
                      borderBottom: "2px solid #B45309",
                    }}
                  >
                    <SortableHeaderLabel
                      label="Tonnage"
                      columnKey="tonnage"
                      sortConfig={sortConfig}
                      onSort={handleSort}
                      onOpenFilter={(e) => openFilter(e, "tonnage")}
                    />
                  </TableCell>
                </TableRow>

                {/* SUMMARY TOTAL ROW */}
                <TableRow
                  sx={{
                    background: "#FEF3C7", // Light yellow/gold summary row
                    borderBottom: "2px solid #F59E0B",
                  }}
                >
                  <TableCell
                    colSpan={3}
                    sx={{
                      fontWeight: 800,
                      color: "#92400E",
                      fontSize: "0.78rem",
                      textTransform: "uppercase",
                    }}
                  >
                    Total
                  </TableCell>

                  {ROLES.map((role) => (
                    <TableCell
                      key={role}
                      align="right"
                      onClick={() => handleHeaderTotalClick(role)}
                      sx={{
                        fontWeight: 800,
                        color: "#92400E",
                        fontSize: "0.75rem",
                        cursor: "pointer",
                        "&:hover": {
                          backgroundColor: "#FDE68A",
                        },
                      }}
                    >
                      {formatValue(colTotals[role])}
                    </TableCell>
                  ))}

                  <TableCell
                    align="right"
                    sx={{
                      fontWeight: 800,
                      color: "#92400E",
                      fontSize: "0.75rem",
                      borderRight: "1px solid #E5E7EB",
                      backgroundColor: "#FDE68A",
                    }}
                  >
                    {formatValue(colTotals.total)}
                  </TableCell>

                  {/* Spacer */}
                  <TableCell style={{ border: "none", backgroundColor: "transparent" }} />

                  <TableCell
                    align="right"
                    sx={{
                      fontWeight: 800,
                      color: "#92400E",
                      fontSize: "0.75rem",
                      backgroundColor: "#FDE68A",
                    }}
                  >
                    {colTotals.overallCount > 0 ? colTotals.overallCount.toFixed(0) : ""}
                  </TableCell>

                  <TableCell
                    align="right"
                    sx={{
                      fontWeight: 800,
                      color: "#92400E",
                      fontSize: "0.75rem",
                      backgroundColor: "#FDE68A",
                    }}
                  >
                    {colTotals.tonnage > 0 ? colTotals.tonnage.toFixed(2) : ""}
                  </TableCell>
                </TableRow>
              </TableHead>

              <TableBody>
                {sortedGridData.map((row, i) => (
                  <TableRow
                    key={i}
                    sx={{
                      transition: "background-color 0.2s",
                      "&:hover": { backgroundColor: "#F8FAFC" },
                    }}
                  >
                    {!sortConfig && row.isFirstInCategory ? (
                      <TableCell rowSpan={row.categorySpan} sx={categoryHeaderStyle}>
                        {row.category}
                      </TableCell>
                    ) : sortConfig ? (
                      <TableCell sx={categoryHeaderStyle}>
                        {row.category}
                      </TableCell>
                    ) : null}
                    {!sortConfig && row.isFirstInVouType ? (
                      <TableCell rowSpan={row.vouTypeSpan} sx={rowHeaderStyle}>
                        {row.vouType}
                      </TableCell>
                    ) : sortConfig ? (
                      <TableCell sx={rowHeaderStyle}>
                        {row.vouType}
                      </TableCell>
                    ) : null}
                    <TableCell sx={bodyStyle}>{row.godown}</TableCell>

                    {/* Role values cells */}
                    {ROLES.map((role) => (
                      <TableCell
                        key={role}
                        align="right"
                        onClick={() => handleCellClick(row, role)}
                        sx={valueCellStyle}
                      >
                        {formatValue(row[role])}
                      </TableCell>
                    ))}

                    {/* Row Total */}
                    <TableCell
                      align="right"
                      onClick={() => handleCellClick(row, "total")}
                      sx={{
                        ...valueCellStyle,
                        borderRight: "1px solid #E5E7EB",
                        fontWeight: 700,
                        backgroundColor: "#F1F5F9",
                      }}
                    >
                      {formatValue(row.total)}
                    </TableCell>

                    {/* Empty spacer column */}
                    <TableCell style={{ border: "none", backgroundColor: "transparent" }} />

                    {/* Overall Count */}
                    <TableCell
                      align="right"
                      onClick={() => handleCellClick(row, "overall_count")}
                      sx={{
                        ...valueCellStyle,
                        backgroundColor: "#FFFBEB",
                        fontWeight: 600,
                      }}
                    >
                      {row.overallCount > 0 ? row.overallCount.toFixed(0) : ""}
                    </TableCell>

                    {/* Overall Tonnage */}
                    <TableCell
                      align="right"
                      onClick={() => handleCellClick(row, "overall_tonnage")}
                      sx={{
                        ...valueCellStyle,
                        backgroundColor: "#FFFBEB",
                        fontWeight: 600,
                      }}
                    >
                      {row.tonnage > 0 ? row.tonnage.toFixed(2) : ""}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        </Box>
      </AppLayout>

      <NumericalFilterMenu
        anchorEl={filterAnchor}
        open={Boolean(filterAnchor)}
        onClose={() => setFilterAnchor(null)}
        activeHeader={activeHeader}
        min={activeHeader ? getMinMax(activeHeader).min : 0}
        max={activeHeader ? getMinMax(activeHeader).max : 100}
        rangeFilter={rangeFilter}
        onRangeChange={(key, range) => setRangeFilter(p => ({ ...p, [key]: range }))}
        onClear={clearRangeFilter}
      />
    </>
  );
};

export default ProductionStaffBasedReport;
