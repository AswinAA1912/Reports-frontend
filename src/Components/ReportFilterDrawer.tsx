import React from "react";
import {
    Box,
    Drawer,
    Typography,
    TextField,
    MenuItem,
    Button,
    IconButton,
    RadioGroup,
    FormControlLabel,
    Radio,
    FormControl,
    FormLabel,
    Autocomplete
} from "@mui/material";
import KeyboardArrowLeftIcon from "@mui/icons-material/KeyboardArrowLeft";

export interface DropdownOption {
    label: string;
    value: string | number;
}

interface ReportFilterDrawerProps {
    open: boolean;
    onToggle: () => void;
    onClose: () => void;

    fromDate?: string;
    onFromDateChange?: (value: string) => void;
    hideFromDate?: boolean;

    // OLD (keep for compatibility)
    dropdownLabel?: string;
    dropdownValue?: string | number;
    dropdownOptions?: DropdownOption[];
    onDropdownChange?: (value: string | number) => void;

    // NEW ✅
    filterLevels?: Record<number, any[]>;
    selectedFilters?: Record<string, any>;
    onFilterChange?: (column: string, value: any) => void;

    stockFilter?: "hasValues" | "zero" | "all";
    onStockFilterChange?: (val: "hasValues" | "zero" | "all") => void;

    onApply: () => void;

    toDate?: string;
    onToDateChange?: (value: string) => void;

    // SALE ORDER STATUS FILTER PROPS SPECIFIC TO PENDING SALE ORDER SCREEN
    showSaleOrderStatusFilter?: boolean;
    saleOrderStatusValue?: "pending" | "cancelled";
    onSaleOrderStatusChange?: (value: "pending" | "cancelled") => void;

    // STAFF BASED REPORT VALUE DISPLAY MODE
    showStaffBasedDisplayMode?: boolean;
    staffBasedDisplayMode?: "qty" | "count";
    onStaffBasedDisplayModeChange?: (value: "qty" | "count") => void;

    // RATE TYPE FILTER
    showRateTypeFilter?: boolean;
    rateTypeValue?: "cogs" | "min";
    onRateTypeChange?: (value: "cogs" | "min") => void;

    // UNIT FILTER
    showUnitFilter?: boolean;
    unitValue?: "kgs_ton" | "bags";
    onUnitChange?: (value: "kgs_ton" | "bags") => void;

    // QTY MODE FILTER (Qty vs Act Qty)
    showQtyModeFilter?: boolean;
    qtyModeValue?: "qty" | "actQty";
    onQtyModeChange?: (value: "qty" | "actQty") => void;
    children?: React.ReactNode;
}

const ReportFilterDrawer: React.FC<ReportFilterDrawerProps> = ({
    open,
    onToggle,
    onClose,
    fromDate,
    toDate,
    onFromDateChange,
    onToDateChange,
    hideFromDate = false,

    // OLD
    dropdownLabel,
    dropdownValue,
    dropdownOptions,
    onDropdownChange,

    // NEW ✅
    filterLevels,
    selectedFilters,
    onFilterChange,

    stockFilter,
    onStockFilterChange,
    onApply,

    // SALE ORDER STATUS FILTER PROPS SPECIFIC TO PENDING SALE ORDER SCREEN
    showSaleOrderStatusFilter,
    saleOrderStatusValue,
    onSaleOrderStatusChange,

    // STAFF BASED REPORT VALUE DISPLAY MODE
    showStaffBasedDisplayMode,
    staffBasedDisplayMode,
    onStaffBasedDisplayModeChange,

        showRateTypeFilter,
    rateTypeValue,
    onRateTypeChange,

    // UNIT FILTER
    showUnitFilter,
    unitValue,
    onUnitChange,

    // QTY MODE FILTER
    showQtyModeFilter,
    qtyModeValue,
    onQtyModeChange,
    children,
}) => {
    return (
        <>
            {/* 🔥 FIXED TOGGLE ARROW */}
            <IconButton
                onClick={onToggle}
                sx={{
                    position: "fixed",
                    right: 1,
                    top: "50%",
                    transform: "translateY(-50%)",
                    width: 25,
                    height: 30,
                    bgcolor: "#1E3A8A",
                    color: "#fff",
                    zIndex: 1300,
                    borderRadius: "6px",
                    boxShadow: 2,
                    "&:hover": {
                        bgcolor: "#162E6E",
                    },
                }}
            >
                <KeyboardArrowLeftIcon
                    sx={{
                        fontSize: 22,
                        transition: "0.25s",
                        transform: open ? "rotate(180deg)" : "rotate(0deg)",
                    }}
                />
            </IconButton>

            {/* DRAWER */}
            <Drawer anchor="right" open={open} onClose={onClose}>
                <Box
                    width={320}
                    p={2}
                    sx={{ backgroundColor: "#F1F5F9", height: "100%" }}
                >
                    <Typography variant="h6" mb={2} fontWeight={700}>
                        Filters
                    </Typography>

                    {!hideFromDate && fromDate !== undefined && onFromDateChange && (
                        <TextField
                            type="date"
                            label="From Date"
                            fullWidth
                            InputLabelProps={{ shrink: true }}
                            value={fromDate}
                            onChange={(e) => onFromDateChange(e.target.value)}
                            sx={{ mb: 2 }}
                        />
                    )}

                    {toDate !== undefined && onToDateChange && (
                        <TextField
                            type="date"
                            label="To Date"
                            fullWidth
                            InputLabelProps={{ shrink: true }}
                            value={toDate}
                            onChange={(e) => onToDateChange(e.target.value)}
                            sx={{ mb: 2 }}
                        />
                    )}

                    {/* SALE ORDER STATUS FILTER FOR PENDING SALE ORDER REPORT */}
                    {showSaleOrderStatusFilter && onSaleOrderStatusChange && (
                        <FormControl sx={{ mb: 2, display: "block" }}>
                            <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                                Sale Order Status
                            </FormLabel>
                            <RadioGroup
                                value={saleOrderStatusValue || "pending"}
                                onChange={(e) =>
                                    onSaleOrderStatusChange(e.target.value as "pending" | "cancelled")
                                }
                            >
                                <FormControlLabel
                                    value="pending"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Pending Sale Order</Typography>}
                                />
                                <FormControlLabel
                                    value="cancelled"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Cancelled</Typography>}
                                />
                            </RadioGroup>
                        </FormControl>
                    )}

                    {/* ✅ DYNAMIC FILTERS */}

                    {filterLevels && filterLevels[1] && (
                        <Box mb={2}>
                            <Typography fontWeight={600} mb={1}>
                                Filter Level-1
                            </Typography>

                            {filterLevels[1].map((filter: any) => (
                                <Autocomplete
                                    multiple
                                    options={filter.options || []}
                                    getOptionLabel={(option: any) => option.label}

                                    filterOptions={(options, { inputValue }) => {
                                        const search = inputValue
                                            .toLowerCase()
                                            .replace(/\s+/g, "")
                                            .replace(/[^a-z0-9]/gi, "");

                                        return options.filter((option: any) => {
                                            const label = option.label
                                                .toLowerCase()
                                                .replace(/\s+/g, "")
                                                .replace(/[^a-z0-9]/gi, "");

                                            return label.includes(search);
                                        });
                                    }}

                                    value={
                                        filter.options?.filter((opt: any) =>
                                            (selectedFilters?.[filter.columnName] || []).includes(opt.value)
                                        ) || []
                                    }

                                    onChange={(_, newValue) => {
                                        const values = newValue.map((opt: any) => opt.value);
                                        onFilterChange?.(filter.columnName, values);
                                    }}

                                    disableCloseOnSelect

                                    renderInput={(params) => (
                                        <TextField
                                            {...params}
                                            label={filter.columnName}
                                            placeholder="Search..."
                                        />
                                    )}

                                    sx={{ mb: 2 }}
                                />
                            ))}
                        </Box>
                    )}

                    {/* ✅ FALLBACK (OLD SUPPORT) */}
                    {!filterLevels &&
                        dropdownOptions &&
                        dropdownOptions.length > 0 && (
                            <TextField
                                select
                                label={dropdownLabel}
                                fullWidth
                                value={dropdownValue ?? ""}
                                onChange={(e) => {
                                    const value = e.target.value;
                                    const option = dropdownOptions.find(
                                        (opt) => String(opt.value) === value
                                    );
                                    onDropdownChange?.(option ? option.value : "");
                                }}
                                sx={{ mb: 2 }}
                            >
                                <MenuItem value="">All</MenuItem>
                                {dropdownOptions.map((opt) => (
                                    <MenuItem key={opt.value} value={String(opt.value)}>
                                        {opt.label}
                                    </MenuItem>
                                ))}
                            </TextField>
                        )}

                    {stockFilter && onStockFilterChange && (
                        <FormControl sx={{ mb: 2 }}>
                            <FormLabel sx={{ fontWeight: 600 }}>
                                Stock Filter
                            </FormLabel>

                            <RadioGroup
                                value={stockFilter}
                                onChange={(e) =>
                                    onStockFilterChange(e.target.value as any)
                                }
                            >
                                <FormControlLabel
                                    value="hasValues"
                                    control={<Radio size="small" />}
                                    label="Data only has values"
                                />
                                <FormControlLabel
                                    value="zero"
                                    control={<Radio size="small" />}
                                    label="Data with 0"
                                />
                                <FormControlLabel
                                    value="all"
                                    control={<Radio size="small" />}
                                    label="All"
                                />
                            </RadioGroup>
                        </FormControl>
                    )}
                    {showStaffBasedDisplayMode && onStaffBasedDisplayModeChange && (
                        <FormControl sx={{ mb: 2, display: "block" }}>
                            <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                                Value Display Mode
                            </FormLabel>
                            <RadioGroup
                                value={staffBasedDisplayMode || "qty"}
                                onChange={(e) =>
                                    onStaffBasedDisplayModeChange(e.target.value as "qty" | "count")
                                }
                            >
                                <FormControlLabel
                                    value="qty"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Quantity</Typography>}
                                />
                                <FormControlLabel
                                    value="count"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Invoice Count</Typography>}
                                />
                            </RadioGroup>
                        </FormControl>
                    )}

                                        {showRateTypeFilter && onRateTypeChange && (
                        <FormControl sx={{ mb: 2, display: "block" }}>
                            <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                                Rate Type
                            </FormLabel>
                            <RadioGroup
                                value={rateTypeValue || "cogs"}
                                onChange={(e) =>
                                    onRateTypeChange(e.target.value as "cogs" | "min")
                                }
                            >
                                <FormControlLabel
                                    value="cogs"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>COGS Rate</Typography>}
                                />
                                <FormControlLabel
                                    value="min"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Min Rate</Typography>}
                                />
                            </RadioGroup>
                        </FormControl>
                    )}

                    {showUnitFilter && onUnitChange && (
                        <FormControl sx={{ mb: 2, display: "block" }}>
                            <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                                Quantity Unit
                            </FormLabel>
                            <RadioGroup
                                value={unitValue || "kgs_ton"}
                                onChange={(e) =>
                                    onUnitChange(e.target.value as "kgs_ton" | "bags")
                                }
                            >
                                <FormControlLabel
                                    value="kgs_ton"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Qty in Kgs/Ton</Typography>}
                                />
                                <FormControlLabel
                                    value="bags"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Qty in Bags</Typography>}
                                />
                            </RadioGroup>
                        </FormControl>
                    )}

                    {showQtyModeFilter && onQtyModeChange && (
                        <FormControl sx={{ mb: 2, display: "block" }}>
                            <FormLabel sx={{ fontWeight: 600, color: "#1E3A8A", fontSize: "0.875rem", display: "block", mb: 0.5 }}>
                                Quantity Mode
                            </FormLabel>
                            <RadioGroup
                                value={qtyModeValue || "qty"}
                                onChange={(e) =>
                                    onQtyModeChange(e.target.value as "qty" | "actQty")
                                }
                            >
                                <FormControlLabel
                                    value="qty"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Qty</Typography>}
                                />
                                <FormControlLabel
                                    value="actQty"
                                    control={<Radio size="small" sx={{ color: "#1E3A8A", "&.Mui-checked": { color: "#1E3A8A" } }} />}
                                    label={<Typography sx={{ fontSize: "0.825rem" }}>Act Qty</Typography>}
                                />
                            </RadioGroup>
                        </FormControl>
                    )}

                    {children}

                    <Button
                        fullWidth
                        sx={{
                            backgroundColor: "#1E3A8A",
                            color: "#fff",
                            fontWeight: 600,
                            "&:hover": { backgroundColor: "#162E6E" },
                        }}
                        onClick={() => {
                            onApply();
                            onClose();
                        }}
                    >
                        Apply Filter
                    </Button>

                </Box>
            </Drawer>
        </>
    );
};

export default ReportFilterDrawer;
