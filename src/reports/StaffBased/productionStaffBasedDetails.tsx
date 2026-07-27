import React, { useState, useMemo } from "react";
import {
  Box,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Button,
  Typography,
  IconButton,
} from "@mui/material";
import ArrowBackIcon from "@mui/icons-material/ArrowBack";
import { useLocation, useNavigate } from "react-router-dom";
import dayjs from "dayjs";
import { useNumericalFilter } from "../../hooks/useNumericalFilter";
import { NumericalFilterMenu } from "../../Components/NumericalFilterMenu";
import { SortableHeaderLabel } from "../../Components/SortableHeaderLabel";
import AppLayout from "../../Layout/appLayout";
import PageHeader from "../../Layout/PageHeader";
import CommonPagination from "../../Components/CommonPagination";
import { exportToPDF } from "../../utils/exportToPDF";
import { exportToExcel } from "../../utils/exportToExcel";
import { mapForExport } from "../../utils/exportMapper";
import { mockProductionDetails } from "./mockProductionData";

/* ================= STYLES ================= */
const headStyle = {
  backgroundColor: "#1E3A8A",
  color: "#fff",
  fontWeight: 600,
  fontSize: "0.75rem",
};

const bodyStyle = {
  fontSize: "0.72rem",
};

const EXPORT_COLUMNS = [
  { label: "S.No", key: "sno" },
  { label: "Date", key: "Date", type: "date" },
  { label: "Journal No", key: "JournalNo" },
  { label: "Voucher Type", key: "VouType" },
  { label: "Staff Name", key: "StaffName" },
  { label: "Item Name", key: "ItemName" },
  { label: "Total Qty", key: "Qty", type: "number" },
  { label: "Atten by", key: "role_Atten_by", type: "number" },
  { label: "Created", key: "role_Created", type: "number" },
  { label: "Printed By", key: "role_Printed_By", type: "number" },
  { label: "Taken", key: "role_Taken", type: "number" },
  { label: "Check", key: "role_Check", type: "number" },
  { label: "Check 1", key: "role_Check_1", type: "number" },
  { label: "Delivery", key: "role_Delivery", type: "number" },
];

const ROLES = ["Atten by", "Created", "Printed By", "Taken", "Check", "Check 1", "Delivery"] as const;

const ProductionStaffBasedDetails: React.FC = () => {
  const navigate = useNavigate();
  const location = useLocation();

  // Retrieve states passed from the previous screen
  const stateData = location.state || {};
  const queryFilters = stateData.filters || {};
  const displayMode = stateData.displayMode || "tonnage";

  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(25);
  // Filter records based on selected cell details
  const filteredRecords = useMemo(() => {
    return mockProductionDetails.filter((d) => {
      // 1. Date Range
      if (queryFilters.DateRange) {
        if (d.Date < queryFilters.DateRange.from || d.Date > queryFilters.DateRange.to) {
          return false;
        }
      }

      // 2. Category
      if (queryFilters.Category && d.Category !== queryFilters.Category) {
        return false;
      }

      // 3. Voucher Type
      if (queryFilters.VouType && d.VouType !== queryFilters.VouType) {
        return false;
      }

      // 4. Godown
      if (queryFilters.Godown && d.Godown !== queryFilters.Godown) {
        return false;
      }

      // 5. Role
      if (queryFilters.Role && d.Role !== queryFilters.Role) {
        return false;
      }

      // 6. Specific Role Type (Printed By)
      if (queryFilters.SpecificRoleType && d.Role !== queryFilters.SpecificRoleType) {
        return false;
      }

      return true;
    });
  }, [queryFilters]);

  // Group and pivot records by Date, JournalNo, StaffName, and ItemName
  const groupedRecords = useMemo(() => {
    const groups: Record<string, {
      Date: string;
      JournalNo: string;
      VouType: string;
      StaffName: string;
      ItemName: string;
      Qty: number;
      roles: Record<typeof ROLES[number], number>;
    }> = {};

    filteredRecords.forEach((rec) => {
      const key = `${rec.Date}_${rec.JournalNo}_${rec.StaffName}_${rec.ItemName}`;
      if (!groups[key]) {
        groups[key] = {
          Date: rec.Date,
          JournalNo: rec.JournalNo,
          VouType: rec.VouType,
          StaffName: rec.StaffName,
          ItemName: rec.ItemName,
          Qty: 0,
          roles: {
            "Atten by": 0,
            "Created": 0,
            "Printed By": 0,
            "Taken": 0,
            "Check": 0,
            "Check 1": 0,
            "Delivery": 0,
          },
        };
      }

      groups[key].Qty += rec.Qty;
      groups[key].roles[rec.Role] += displayMode === "tonnage" ? rec.Qty : 1;
    });

    return Object.values(groups).map((g) => ({
      Date: g.Date,
      JournalNo: g.JournalNo,
      VouType: g.VouType,
      StaffName: g.StaffName,
      ItemName: g.ItemName,
      Qty: g.Qty,
      "Atten by": g.roles["Atten by"] || 0,
      "Created": g.roles["Created"] || 0,
      "Printed By": g.roles["Printed By"] || 0,
      "Taken": g.roles["Taken"] || 0,
      "Check": g.roles["Check"] || 0,
      "Check 1": g.roles["Check 1"] || 0,
      "Delivery": g.roles["Delivery"] || 0,
    }));
  }, [filteredRecords, displayMode]);

  const numericKeys = useMemo(() => ["Qty", ...ROLES], []);

  const {
    sortConfig,
    rangeFilter,
    filterAnchor,
    setFilterAnchor,
    activeHeader,
    handleSort,
    openFilter,
    filteredAndSortedData,
    getMinMax,
    setRangeFilter,
    clearRangeFilter,
  } = useNumericalFilter(groupedRecords, numericKeys);

  // Paginated records
  const paginatedRecords = useMemo(() => {
    return filteredAndSortedData.slice((page - 1) * rowsPerPage, page * rowsPerPage);
  }, [filteredAndSortedData, page, rowsPerPage]);

  const totalQty = useMemo(() => {
    return filteredAndSortedData.reduce((sum, r) => sum + r.Qty, 0);
  }, [filteredAndSortedData]);

  // Format values inside the pivoted table cells
  const formatRoleValue = (val: number) => {
    if (val === 0 || val === undefined) return "";
    return displayMode === "tonnage" ? val.toFixed(2) : val.toFixed(0);
  };

  // Flattened data format for exports
  const exportData = useMemo(() => {
    return filteredAndSortedData.map((r) => ({
      Date: r.Date,
      JournalNo: r.JournalNo,
      VouType: r.VouType,
      StaffName: r.StaffName,
      ItemName: r.ItemName,
      Qty: r.Qty,
      role_Atten_by: r["Atten by"] || 0,
      role_Created: r["Created"] || 0,
      role_Printed_By: r["Printed By"] || 0,
      role_Taken: r["Taken"] || 0,
      role_Check: r["Check"] || 0,
      role_Check_1: r["Check 1"] || 0,
      role_Delivery: r["Delivery"] || 0,
    }));
  }, [filteredAndSortedData]);

  // Export handlers
  const handleExportPDF = () => {
    const { headers, data: mapped } = mapForExport(EXPORT_COLUMNS, exportData);
    exportToPDF("Production Staff Working Details", headers, mapped);
  };

  const handleExportExcel = () => {
    const { headers, data: mapped } = mapForExport(EXPORT_COLUMNS, exportData);
    exportToExcel("Production Staff Working Details", headers, mapped);
  };

  return (
    <>
      <PageHeader
        onExportPDF={handleExportPDF}
        onExportExcel={handleExportExcel}
        infoSlot={
          <Button
            variant="outlined"
            size="small"
            startIcon={<ArrowBackIcon />}
            onClick={() => navigate(-1)}
            sx={{
              height: 32,
              fontSize: "0.75rem",
              textTransform: "none",
              borderColor: "#1E3A8A",
              color: "#1E3A8A",
              fontWeight: 600,
              "&:hover": {
                borderColor: "#172E6D",
                backgroundColor: "#f1f5f9",
              },
            }}
          >
            Back to Summary
          </Button>
        }
      />

      <AppLayout fullWidth>
        <Box sx={{ p: 2.5, display: "flex", flexDirection: "column", gap: 2 }}>
          {/* Header Section */}
          <Box display="flex" justifyContent="space-between" alignItems="center">
            <Box display="flex" alignItems="center" gap={1}>
              <IconButton
                onClick={() => navigate(-1)}
                sx={{
                  color: "#1E3A8A",
                  p: 0.5,
                  "&:hover": {
                    backgroundColor: "rgba(30, 58, 138, 0.08)",
                  },
                }}
              >
                <ArrowBackIcon />
              </IconButton>
              <Typography variant="h6" sx={{ fontWeight: 700, color: "#1E3A8A" }}>
                Staff Working Details
              </Typography>
            </Box>

            <Box display="flex" gap={3}>
              <Box>
                <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
                  Total Tonnage Qty
                </Typography>
                <Typography variant="h6" sx={{ color: "#D97706", fontWeight: 700 }}>
                  {totalQty.toFixed(2)} Tons
                </Typography>
              </Box>
              <Box>
                <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
                  Record Count
                </Typography>
                <Typography variant="h6" sx={{ color: "#1E3A8A", fontWeight: 700 }}>
                  {filteredAndSortedData.length} Items
                </Typography>
              </Box>
            </Box>
          </Box>

          {/* Details Table */}
          <TableContainer
            component={Paper}
            sx={{
              borderRadius: "10px",
              boxShadow: "0 4px 12px rgba(0,0,0,0.05)",
              border: "1px solid #E2E8F0",
              overflow: "auto",
            }}
          >
            <Table size="small">
              <TableHead>
                <TableRow>
                  <TableCell sx={headStyle}>S.No</TableCell>
                  <TableCell sx={headStyle}>Date</TableCell>
                  <TableCell sx={headStyle}>Journal / Voucher No</TableCell>
                  <TableCell sx={headStyle}>Voucher Type</TableCell>
                  <TableCell sx={headStyle}>Staff Name</TableCell>
                  <TableCell sx={headStyle}>Item Name</TableCell>
                  <TableCell sx={headStyle} align="right">
                    <SortableHeaderLabel
                      label="Total Qty"
                      columnKey="Qty"
                      sortConfig={sortConfig}
                      onSort={handleSort}
                      onOpenFilter={(e) => openFilter(e, "Qty")}
                    />
                  </TableCell>
                  {ROLES.map((r) => (
                    <TableCell key={r} sx={headStyle} align="right">
                      <SortableHeaderLabel
                        label={r}
                        columnKey={r}
                        sortConfig={sortConfig}
                        onSort={handleSort}
                        onOpenFilter={(e) => openFilter(e, r)}
                      />
                    </TableCell>
                  ))}
                </TableRow>
              </TableHead>

              <TableBody>
                {paginatedRecords.length > 0 ? (
                  paginatedRecords.map((row, i) => (
                    <TableRow key={i}>
                      <TableCell sx={bodyStyle}>{(page - 1) * rowsPerPage + i + 1}</TableCell>
                      <TableCell sx={bodyStyle}>{dayjs(row.Date).format("DD/MM/YYYY")}</TableCell>
                      <TableCell sx={bodyStyle}>{row.JournalNo}</TableCell>
                      <TableCell sx={bodyStyle}>{row.VouType}</TableCell>
                      <TableCell sx={bodyStyle} style={{ fontWeight: 600, color: "#1E3A8A" }}>
                        {row.StaffName}
                      </TableCell>
                      <TableCell sx={bodyStyle}>{row.ItemName}</TableCell>
                      <TableCell sx={bodyStyle} align="right" style={{ fontWeight: 700, backgroundColor: "#F8FAFC" }}>
                        {row.Qty.toFixed(2)}
                      </TableCell>
                      {ROLES.map((r) => (
                        <TableCell key={r} sx={bodyStyle} align="right">
                          {formatRoleValue(row[r as keyof typeof row] as number)}
                        </TableCell>
                      ))}
                    </TableRow>
                  ))
                ) : (
                  <TableRow>
                    <TableCell colSpan={14} align="center" sx={{ py: 3, color: "text.secondary" }}>
                      No detailed records found matching active filters.
                    </TableCell>
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </TableContainer>

          {/* Pagination */}
          {filteredAndSortedData.length > 0 && (
            <CommonPagination
              totalRows={filteredAndSortedData.length}
              page={page}
              rowsPerPage={rowsPerPage}
              onPageChange={setPage}
              onRowsPerPageChange={setRowsPerPage}
            />
          )}
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

export default ProductionStaffBasedDetails;
